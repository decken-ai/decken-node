import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import { DeckenClient, DeckenError } from '../src/index.mjs';

const execFileAsync = promisify(execFile);
const MOCK_PPTX = Buffer.from('PK\x03\x04mock-openxml-pptx-payload');
const MOCK_PNG = Buffer.from('\x89PNG\r\n\x1a\nmock-preview');

test('DeckenClient enforces HTTPS origin safety, sync/async deck polling, model comparison, and CLI output', async () => {
  // 1. Origin-pinning invariant: refuse sending API key over plaintext HTTP to remote host
  assert.throws(
    () => new DeckenClient({ apiKey: 'dk_secret_123', baseUrl: 'http://evil.example.com' }),
    (err) => err instanceof DeckenError && /insecure non-HTTPS/.test(err.message),
  );

  let pollCalls = 0;
  const jobs = new Map();

  const server = createServer(async (req, res) => {
    if (req.headers.authorization !== 'Bearer dk_test_token') {
      res.writeHead(401, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Unauthorized' }));
      return;
    }

    const url = new URL(req.url, 'http://127.0.0.1');
    if (url.pathname === '/v1/capabilities') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(
        JSON.stringify({
          service: 'decken-reconstruct-api',
          version: '1.0.0',
          defaultModelId: 'qwen/qwen3.8-flash',
          limits: { maxSlidesPerDeck: 24 },
          nativePrimitives: { inpainting: 'stroke-selective-harmonic-laplace', shapes: ['roundRect', 'trapezoid'] },
        }),
      );
      return;
    }

    if (url.pathname === '/v1/openapi.json') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ openapi: '3.1.0', paths: { '/v1/reconstruct': {} } }));
      return;
    }

    if (url.pathname === '/v1/decks' && req.method === 'POST') {
      const chunks = [];
      for await (const c of req) chunks.push(c);
      const body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      const scenario = body.source
        ? 'improve_material'
        : body.reference
          ? 'from_outline_with_reference'
          : 'from_outline';
      const jobId = `deck_${jobs.size + 1}`;
      const record = {
        jobId,
        scenario,
        status: 'completed',
        mode: body.mode || 'sync',
        engineMode: body.engineMode || 'vlm',
        modelId: body.modelId || 'qwen/qwen3.8-flash',
        createdAt: new Date().toISOString(),
        pptxUrl: `/v1/decks/${jobId}/pptx`,
        pollUrl: `/v1/decks/${jobId}`,
        pptxBase64: MOCK_PPTX.toString('base64'),
        metrics: { slideCount: 2, totalObjectCount: 20, totalTableObjectCount: 1 },
        styleConsistency: { referenceSourceKind: body.reference ? 'pptx_template' : 'default_editorial', score: 0.98 },
        agentDiagnostics: { editabilityVerdict: 'native_editable', nativeEditableRatio: 1 },
        usage: { physicalRequests: 1, totalTokens: 2800, costUsd: 0.0011 },
        slides: [{ slideId: 'slide-01', objectCount: 10, previewUrl: `/v1/decks/${jobId}/slides/slide-01/preview.png` }],
      };
      jobs.set(jobId, record);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(record));
      return;
    }

    if (url.pathname === '/v1/reconstruct' && req.method === 'POST') {
      const chunks = [];
      for await (const c of req) chunks.push(c);
      const body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      const slides = body.slides || [{ slideId: body.slideId || 'slide-01', imageDataUrl: body.imageDataUrl }];
      const jobId = `job_${jobs.size + 1}`;
      const record = {
        jobId,
        status: body.mode === 'async' ? 'running' : 'completed',
        mode: body.mode || 'sync',
        engineMode: body.engineMode || 'vlm',
        modelId: body.modelId || 'qwen/qwen3.8-flash',
        createdAt: new Date().toISOString(),
        durationMs: 1420,
        pptxUrl: `/v1/jobs/${jobId}/pptx`,
        pollUrl: `/v1/jobs/${jobId}`,
        ...(body.includePptxBase64 ? { pptxBase64: MOCK_PPTX.toString('base64') } : {}),
        metrics: {
          slideCount: slides.length,
          meanWholePageMae: body.modelId === 'openai/gpt-5.6-luna' ? 0.017 : 0.019,
          totalObjectCount: slides.length * 14,
          totalTextObjectCount: slides.length * 8,
          totalShapeObjectCount: slides.length * 5,
          totalPictureObjectCount: 0,
          totalConnectorObjectCount: 0,
          totalTableObjectCount: slides.length,
          totalGroupCount: slides.length * 2,
        },
        usage: {
          physicalRequests: slides.length,
          inputTokens: 1200 * slides.length,
          outputTokens: 1800 * slides.length,
          totalTokens: 3000 * slides.length,
          costUsd: 0.0012 * slides.length,
        },
        slides: slides.map((s) => ({
          slideId: s.slideId,
          objectCount: 14,
          textObjectCount: 8,
          shapeObjectCount: 5,
          pictureObjectCount: 0,
          connectorObjectCount: 0,
          tableObjectCount: 1,
          groupCount: 2,
          wholePageMae: 0.019,
          previewUrl: `/v1/jobs/${jobId}/slides/${s.slideId}/preview.png`,
        })),
      };
      jobs.set(jobId, record);
      res.writeHead(body.mode === 'async' ? 202 : 200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(record));
      return;
    }

    if (url.pathname === '/v1/jobs') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jobs: [...jobs.values()] }));
      return;
    }

    const match = /^\/v1\/jobs\/([^/]+)(?:\/(pptx|slides\/([^/]+)\/preview\.png))?$/.exec(url.pathname);
    if (match) {
      const job = jobs.get(match[1]);
      if (!job) {
        res.writeHead(404);
        res.end(JSON.stringify({ error: 'Not found' }));
        return;
      }
      if (!match[2]) {
        pollCalls++;
        if (pollCalls >= 2) job.status = 'completed';
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(
          JSON.stringify({
            ...job,
            ...(url.searchParams.get('includeBase64') === '1' ? { pptxBase64: MOCK_PPTX.toString('base64') } : {}),
          }),
        );
        return;
      }
      if (match[2] === 'pptx') {
        res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
        res.end(MOCK_PPTX);
        return;
      }
      res.writeHead(200, { 'Content-Type': 'image/png' });
      res.end(MOCK_PNG);
    }
  });

  const scratch = await mkdtemp(join(tmpdir(), 'decken-sdk-test-'));
  try {
    await new Promise((r) => server.listen(0, '127.0.0.1', r));
    const port = server.address().port;
    const baseUrl = `http://127.0.0.1:${port}`;
    const imgPath = join(scratch, 'slide-a.png');
    await writeFile(imgPath, MOCK_PNG);

    const client = new DeckenClient({ apiKey: 'dk_test_token', baseUrl });
    const caps = await client.getCapabilities();
    assert.equal(caps.defaultModelId, 'qwen/qwen3.8-flash');

    // Unified 3-scenario client.generate(...) test
    const tplPath = join(scratch, 'brand.pptx');
    await writeFile(tplPath, MOCK_PPTX);
    const genOut1 = join(scratch, 'gen-outline.pptx');
    const gen1 = await client.generate({ outline: '## Slide 1\n- Point', outputPath: genOut1 });
    assert.equal(gen1.scenario, 'from_outline');
    assert.deepEqual(await readFile(genOut1), MOCK_PPTX);

    const genOut2 = join(scratch, 'gen-ref.pptx');
    const gen2 = await client.generate({ outline: '## Slide 1\n- Point', reference: tplPath, outputPath: genOut2 });
    assert.equal(gen2.scenario, 'from_outline_with_reference');
    assert.deepEqual(await readFile(genOut2), MOCK_PPTX);

    const genOut3 = join(scratch, 'gen-improve.pptx');
    const gen3 = await client.generate({ source: tplPath, prompt: 'Upgrade layout', outputPath: genOut3 });
    assert.equal(gen3.scenario, 'improve_material');
    assert.deepEqual(await readFile(genOut3), MOCK_PPTX);

    // Sync single slide write to disk + HTTPS imageUrl support & plaintext http:// rejection
    const outSingle = join(scratch, 'single.pptx');
    const singleJob = await client.reconstructSlide({ imagePath: imgPath, outputPath: outSingle });
    assert.equal(singleJob.status, 'completed');
    assert.deepEqual(await readFile(outSingle), MOCK_PPTX);

    const outUrlSlide = join(scratch, 'url-slide.pptx');
    const urlJob = await client.reconstructSlide({
      imageUrl: 'https://cdn.example.com/slides/01-kpi.png',
      outputPath: outUrlSlide,
    });
    assert.equal(urlJob.status, 'completed');
    assert.deepEqual(await readFile(outUrlSlide), MOCK_PPTX);

    await assert.rejects(
      () => client.reconstructSlide({ imageUrl: 'http://cdn.example.com/insecure.png' }),
      /HTTPS/,
    );

    // Async multi-slide deck with automatic polling & binary download
    const outDeck = join(scratch, 'deck.pptx');
    const deckJob = await client.reconstructDeck({
      slides: [
        { slideId: 's1', imagePath: imgPath },
        { slideId: 's2', imagePath: imgPath },
      ],
      mode: 'async',
      pollIntervalMs: 25,
      outputPath: outDeck,
    });
    assert.equal(deckJob.status, 'completed');
    assert.equal(deckJob.metrics.slideCount, 2);
    assert.deepEqual(await readFile(outDeck), MOCK_PPTX);

    // Multi-model comparison
    const comparison = await client.compareModels({
      imagePath: imgPath,
      modelIds: ['qwen/qwen3.8-flash', 'openai/gpt-5.6-luna'],
    });
    assert.equal(comparison.runs.length, 2);
    assert.equal(comparison.runs[1].wholePageMae, 0.017);

    // CLI end-to-end test
    const cliOut = join(scratch, 'cli.pptx');
    const previewDir = join(scratch, 'previews');
    const cliBin = new URL('../bin/decken.mjs', import.meta.url).pathname;
    const { stdout } = await execFileAsync(process.execPath, [
      cliBin,
      'reconstruct',
      imgPath,
      '-o',
      cliOut,
      '--preview-dir',
      previewDir,
    ], {
      env: { ...process.env, DECKEN_API_KEY: 'dk_test_token', DECKEN_BASE_URL: baseUrl },
    });
    assert.match(stdout, /Saved/);
    assert.deepEqual(await readFile(cliOut), MOCK_PPTX);
    assert.deepEqual(await readFile(join(previewDir, 'slide-a.png')), MOCK_PNG);
  } finally {
    server.close();
    await rm(scratch, { recursive: true, force: true });
  }
});
