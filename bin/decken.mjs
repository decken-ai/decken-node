#!/usr/bin/env node
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { DeckenClient } from '../src/index.mjs';

const HELP = `Decken CLI — The Native OpenXML (.pptx) Slide Generation Kernel

Usage:
  decken generate [--outline <file.md|text>] [--reference <template.pptx|ref.png>] [--source <rough.pptx|slide.png>] [--prompt <text>] -o <output.pptx>
  decken reconstruct <image...> -o <output.pptx> [--engine vlm|deterministic] [--model <id>] [--async] [--preview-dir <dir>] [--json]
  decken capabilities [--json]
  decken openapi [-o openapi.json]
  decken compare <image> --models <m1,m2> [--json]
  decken jobs [jobId] [--download <output.pptx>] [--json]

Environment:
  DECKEN_API_KEY     Bearer token for https://api.decken.ai
  DECKEN_BASE_URL    Override gateway origin (default: https://api.decken.ai)
`;

async function main() {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      output: { type: 'string', short: 'o' },
      outline: { type: 'string' },
      reference: { type: 'string' },
      source: { type: 'string' },
      prompt: { type: 'string' },
      engine: { type: 'string', default: 'vlm' },
      model: { type: 'string' },
      models: { type: 'string' },
      title: { type: 'string' },
      'max-tokens': { type: 'string' },
      'preview-dir': { type: 'string' },
      download: { type: 'string' },
      async: { type: 'boolean', default: false },
      json: { type: 'boolean', default: false },
      help: { type: 'boolean', short: 'h', default: false },
    },
  });

  const [cmd, ...rest] = positionals;
  if (values.help || !cmd) {
    process.stdout.write(HELP);
    return;
  }

  const client = new DeckenClient();

  if (cmd === 'generate') {
    const outputPath = values.output || 'deck.pptx';
    let outlineVal = values.outline;
    if (outlineVal && !outlineVal.includes('\n') && (outlineVal.endsWith('.md') || outlineVal.endsWith('.json'))) {
      const content = await readFile(outlineVal, 'utf8');
      outlineVal = outlineVal.endsWith('.json') ? JSON.parse(content) : content;
    }
    const job = await client.generate({
      deckTitle: values.title,
      outline: outlineVal,
      reference: values.reference,
      source: values.source,
      prompt: values.prompt,
      outputPath,
      mode: values.async ? 'async' : 'sync',
      engineMode: values.engine,
      modelId: values.model,
      maxOutputTokens: values['max-tokens'] ? Number(values['max-tokens']) : undefined,
    });

    if (values.json) {
      const { pptxBase64: _omit, ...summary } = job;
      console.log(JSON.stringify({ ...summary, outputPath }, null, 2));
    } else {
      console.log(`✓ Saved ${outputPath} (${job.scenario || 'deck'}, job: ${job.jobId})`);
      console.log(`  Slides: ${job.metrics?.slideCount ?? 1} | Native Objects: ${job.metrics?.totalObjectCount ?? 0}`);
    }
    return;
  }

  if (cmd === 'capabilities') {
    const caps = await client.getCapabilities();
    if (values.json) {
      console.log(JSON.stringify(caps, null, 2));
    } else {
      console.log(`Service: ${caps.service} v${caps.version}`);
      console.log(`Default Model: ${caps.defaultModelId}`);
      console.log(`Max Slides/Deck: ${caps.limits?.maxSlidesPerDeck}`);
      console.log(`Inpainting: ${caps.nativePrimitives?.inpainting}`);
      console.log(`Shapes (${caps.nativePrimitives?.shapes?.length ?? 0}): ${(caps.nativePrimitives?.shapes ?? []).join(', ')}`);
    }
    return;
  }

  if (cmd === 'openapi') {
    const spec = await client.getOpenApiSpec();
    const serialized = JSON.stringify(spec, null, 2);
    if (values.output) {
      await writeFile(values.output, serialized);
      console.log(`Wrote OpenAPI ${spec.openapi} spec to ${values.output}`);
    } else {
      console.log(serialized);
    }
    return;
  }

  if (cmd === 'reconstruct') {
    if (rest.length === 0) {
      throw new Error('Provide at least one slide image path: decken reconstruct <image...> -o out.pptx');
    }
    const outputPath = values.output || 'reconstructed.pptx';
    const maxOutputTokens = values['max-tokens'] ? Number(values['max-tokens']) : undefined;
    const mode = values.async || rest.length > 1 ? 'async' : 'sync';

    const job = rest.length === 1
      ? await client.reconstructSlide({
          imagePath: rest[0],
          outputPath,
          mode,
          engineMode: values.engine,
          modelId: values.model,
          deckTitle: values.title,
          maxOutputTokens,
        })
      : await client.reconstructDeck({
          slides: rest.map((imagePath) => ({ imagePath })),
          outputPath,
          mode,
          engineMode: values.engine,
          modelId: values.model,
          deckTitle: values.title,
          maxOutputTokens,
        });

    if (values['preview-dir'] && Array.isArray(job.slides)) {
      await mkdir(values['preview-dir'], { recursive: true });
      for (const sl of job.slides) {
        const png = await client.downloadSlidePreview(job.jobId, sl.slideId);
        await writeFile(join(values['preview-dir'], `${sl.slideId}.png`), png);
      }
    }

    if (values.json) {
      const { pptxBase64: _omit, ...summary } = job;
      console.log(JSON.stringify({ ...summary, outputPath }, null, 2));
    } else {
      console.log(`✓ Saved ${outputPath} (job: ${job.jobId})`);
      console.log(`  Slides: ${job.metrics?.slideCount ?? job.slides?.length ?? 1} | Native Objects: ${job.metrics?.totalObjectCount ?? 0} | Tables: ${job.metrics?.totalTableObjectCount ?? 0}`);
      if (job.agentDiagnostics?.editabilityVerdict) {
        console.log(`  Verdict: ${job.agentDiagnostics.editabilityVerdict} (Native Ratio: ${((job.agentDiagnostics.nativeEditableRatio ?? 0) * 100).toFixed(1)}%)`);
      }
      console.log(`  Tokens: ${job.usage?.totalTokens ?? 0} | Cost: $${job.usage?.costUsd ?? 0}`);
    }
    return;
  }

  if (cmd === 'compare') {
    const imagePath = rest[0];
    const modelIds = (values.models || '')
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    if (!imagePath || modelIds.length === 0) {
      throw new Error('Usage: decken compare <image> --models qwen/qwen3.8-flash,openai/gpt-5.6-luna');
    }
    const report = await client.compareModels({
      imagePath,
      modelIds,
      maxOutputTokens: values['max-tokens'] ? Number(values['max-tokens']) : undefined,
    });
    console.log(JSON.stringify(report, null, 2));
    return;
  }

  if (cmd === 'jobs') {
    const jobId = rest[0];
    if (!jobId) {
      const { jobs } = await client.listJobs();
      if (values.json) {
        console.log(JSON.stringify(jobs, null, 2));
      } else {
        for (const j of jobs) {
          console.log(`${j.jobId}  ${j.status.padEnd(10)}  slides=${j.metrics?.slideCount ?? j.slides?.length ?? 0}  cost=$${j.usage?.costUsd ?? 0}  ${j.createdAt}`);
        }
      }
      return;
    }
    const job = await client.getJob(jobId);
    if (values.download) {
      const bytes = await client.downloadPptx(job.pptxUrl || jobId);
      await writeFile(values.download, bytes);
      console.log(`Downloaded ${values.download} (${bytes.byteLength} bytes)`);
      return;
    }
    console.log(JSON.stringify(job, null, 2));
    return;
  }

  throw new Error(`Unknown command: ${cmd}\n\n${HELP}`);
}

main().catch((err) => {
  console.error(`Error: ${err.message || err}`);
  process.exitCode = 1;
});
