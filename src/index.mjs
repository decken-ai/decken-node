import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, extname, basename } from 'node:path';

export const DEFAULT_BASE_URL = 'https://api.decken.ai';

const MIME_BY_EXT = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.potx': 'application/vnd.openxmlformats-officedocument.presentationml.template',
};

export class DeckenError extends Error {
  constructor(message, statusCode, body) {
    super(message);
    this.name = 'DeckenError';
    this.statusCode = statusCode;
    this.body = body;
  }
}

function isLoopbackHost(hostname) {
  const h = hostname.toLowerCase();
  return h === 'localhost' || h === '127.0.0.1' || h === '[::1]' || h.endsWith('.localhost');
}

export function validateBaseUrl(baseUrl, hasApiKey, allowInsecureLocalhost = true) {
  let parsed;
  try {
    parsed = new URL(baseUrl);
  } catch {
    throw new DeckenError(`Invalid Decken baseUrl: ${baseUrl}`, 0, null);
  }
  if (parsed.protocol !== 'https:') {
    if (parsed.protocol === 'http:' && allowInsecureLocalhost && isLoopbackHost(parsed.hostname)) {
      return parsed.origin + parsed.pathname.replace(/\/+$/, '');
    }
    if (hasApiKey || parsed.protocol !== 'http:') {
      throw new DeckenError(
        `Refusing insecure non-HTTPS baseUrl (${baseUrl}) when transmitting credentials. Use https://api.decken.ai or localhost.`,
        0,
        null,
      );
    }
  }
  return parsed.origin + parsed.pathname.replace(/\/+$/, '');
}

export async function normalizeArtifactPayload(input, fieldName = 'artifact') {
  if (!input) return undefined;
  if (typeof input === 'string') {
    const trimmed = input.trim();
    if (!trimmed) return undefined;
    if (trimmed.startsWith('https://')) {
      return trimmed;
    }
    if (trimmed.startsWith('http://')) {
      throw new DeckenError(`\`${fieldName}\` URL must use HTTPS (\`https://...\`).`, 0, null);
    }
    if (trimmed.startsWith('data:')) {
      return trimmed;
    }
    const bytes = await readFile(trimmed);
    const ext = extname(trimmed).toLowerCase();
    const mime = MIME_BY_EXT[ext] || 'application/octet-stream';
    return `data:${mime};base64,${bytes.toString('base64')}`;
  }
  if (typeof input === 'object') {
    if (typeof input.url === 'string' && input.url.trim()) {
      const u = input.url.trim();
      if (!u.startsWith('https://')) {
        throw new DeckenError(`\`${fieldName}.url\` must use HTTPS (\`https://...\`).`, 0, null);
      }
      return { ...(input.kind ? { kind: input.kind } : {}), url: u };
    }
    if (typeof input.dataUrl === 'string' && input.dataUrl.trim()) {
      return { ...(input.kind ? { kind: input.kind } : {}), dataUrl: input.dataUrl.trim() };
    }
    if (typeof input.path === 'string' && input.path.trim()) {
      const dataUrl = await normalizeArtifactPayload(input.path, fieldName);
      return { ...(input.kind ? { kind: input.kind } : {}), dataUrl };
    }
  }
  throw new DeckenError(`Invalid \`${fieldName}\` input.`, 0, null);
}

export async function normalizeSlidePayload(input, fallbackSlideId = 'slide-01') {
  if (typeof input.imageUrl === 'string' && input.imageUrl.trim()) {
    const rawUrl = input.imageUrl.trim();
    let parsed;
    try {
      parsed = new URL(rawUrl);
    } catch {
      throw new DeckenError(`Invalid imageUrl: ${rawUrl}`, 0, null);
    }
    if (parsed.protocol !== 'https:') {
      throw new DeckenError('`imageUrl` must use HTTPS (`https://...`).', 0, null);
    }
    const stem = basename(parsed.pathname, extname(parsed.pathname)) || fallbackSlideId;
    return { slideId: input.slideId || stem, imageUrl: rawUrl };
  }
  if (typeof input.imagePath === 'string' && input.imagePath.trim().startsWith('https://')) {
    return normalizeSlidePayload({ ...input, imageUrl: input.imagePath, imagePath: undefined }, fallbackSlideId);
  }
  const imageDataUrl = await normalizeImageDataUrl(input);
  const slideId =
    input.slideId
    || (input.imagePath ? basename(input.imagePath, extname(input.imagePath)) : fallbackSlideId);
  return { slideId, imageDataUrl };
}

export async function normalizeImageDataUrl(input) {
  if (typeof input.imageDataUrl === 'string' && input.imageDataUrl.trim()) {
    const raw = input.imageDataUrl.trim();
    return raw.startsWith('data:image/') ? raw : `data:image/png;base64,${raw}`;
  }
  if (input.imageBytes instanceof Uint8Array || Buffer.isBuffer(input.imageBytes)) {
    const mime = input.mimeType || 'image/png';
    return `data:${mime};base64,${Buffer.from(input.imageBytes).toString('base64')}`;
  }
  if (typeof input.imagePath === 'string' && input.imagePath.trim()) {
    const bytes = await readFile(input.imagePath);
    const ext = extname(input.imagePath).toLowerCase();
    const mime = input.mimeType || MIME_BY_EXT[ext] || 'image/png';
    return `data:${mime};base64,${bytes.toString('base64')}`;
  }
  throw new DeckenError(
    'Each slide input must provide `imagePath`, `imageBytes`, or `imageDataUrl`.',
    0,
    null,
  );
}

export class DeckenClient {
  constructor(options = {}) {
    this.apiKey = (options.apiKey ?? process.env.DECKEN_API_KEY ?? process.env.DECKEN_API_TOKEN ?? '').trim();
    const rawBase = (options.baseUrl ?? process.env.DECKEN_BASE_URL ?? DEFAULT_BASE_URL).trim();
    this.baseUrl = validateBaseUrl(rawBase, Boolean(this.apiKey), options.allowInsecureLocalhost ?? true);
    this.timeoutMs = options.timeoutMs ?? 180_000;
    this.fetchImpl = options.fetch ?? globalThis.fetch;
  }

  _headers(extra = {}) {
    const headers = { ...extra };
    if (this.apiKey) {
      headers.Authorization = `Bearer ${this.apiKey}`;
    }
    return headers;
  }

  async _requestJson(method, path, body, signal) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await this.fetchImpl(`${this.baseUrl}${path}`, {
        method,
        headers: this._headers(body !== undefined ? { 'Content-Type': 'application/json' } : {}),
        body: body !== undefined ? JSON.stringify(body) : undefined,
        signal: signal ?? controller.signal,
      });
      const text = await response.text();
      let parsed = null;
      try {
        parsed = text ? JSON.parse(text) : null;
      } catch {
        parsed = { raw: text };
      }
      if (!response.ok) {
        const msg = parsed?.error || `Decken API HTTP ${response.status} on ${method} ${path}`;
        throw new DeckenError(msg, response.status, parsed);
      }
      return parsed;
    } finally {
      clearTimeout(timer);
    }
  }

  async _requestBytes(path, signal) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const url = path.startsWith('http') ? path : `${this.baseUrl}${path}`;
      const response = await this.fetchImpl(url, {
        method,
        headers: this._headers(),
        signal: signal ?? controller.signal,
      });
      if (!response.ok) {
        const text = await response.text().catch(() => '');
        throw new DeckenError(`Failed binary download (${response.status}): ${text}`, response.status, text);
      }
      return new Uint8Array(await response.arrayBuffer());
    } finally {
      clearTimeout(timer);
    }
  }

  async getCapabilities() {
    return this._requestJson('GET', '/v1/capabilities');
  }

  async getOpenApiSpec() {
    return this._requestJson('GET', '/v1/openapi.json');
  }

  async getAgentManifest() {
    return this._requestJson('GET', '/.well-known/agent.json');
  }

  async listJobs() {
    return this._requestJson('GET', '/v1/jobs');
  }

  async getJob(jobId, options = {}) {
    const query = options.includeBase64 ? '?includeBase64=1' : '';
    return this._requestJson('GET', `/v1/jobs/${encodeURIComponent(jobId)}${query}`);
  }

  async downloadPptx(jobIdOrUrl) {
    const path = jobIdOrUrl.startsWith('/') || jobIdOrUrl.startsWith('http')
      ? jobIdOrUrl
      : `/v1/jobs/${encodeURIComponent(jobIdOrUrl)}/pptx`;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const url = path.startsWith('http') ? path : `${this.baseUrl}${path}`;
      const response = await this.fetchImpl(url, {
        method: 'GET',
        headers: this._headers(),
        signal: controller.signal,
      });
      if (!response.ok) {
        const text = await response.text().catch(() => '');
        throw new DeckenError(`Failed binary download (${response.status}): ${text}`, response.status, text);
      }
      return new Uint8Array(await response.arrayBuffer());
    } finally {
      clearTimeout(timer);
    }
  }

  async downloadSlidePreview(jobId, slideId) {
    return this.downloadPptx(
      `/v1/jobs/${encodeURIComponent(jobId)}/slides/${encodeURIComponent(slideId)}/preview.png`,
    );
  }

  async waitForJob(jobId, options = {}) {
    const pollIntervalMs = options.pollIntervalMs ?? 2000;
    const maxWaitMs = options.maxWaitMs ?? 600_000;
    const includeBase64 = options.includeBase64 ?? false;
    const deadline = Date.now() + maxWaitMs;

    while (Date.now() < deadline) {
      const job = await this.getJob(jobId, { includeBase64 });
      options.onProgress?.(job);
      if (job.status === 'completed') return job;
      if (job.status === 'failed') {
        throw new DeckenError(job.error || `Job ${jobId} failed`, 500, job);
      }
      await new Promise((r) => setTimeout(r, pollIntervalMs));
    }
    throw new DeckenError(`Timed out waiting for job ${jobId} after ${maxWaitMs}ms`, 408, null);
  }

  /**
   * Unified Scenario-Driven Slide Kernel (`POST /v1/decks`):
   * - Scenario 1 (`from_outline`): pass `outline` or `prompt`
   * - Scenario 2 (`from_outline_with_reference`): pass `outline`/`prompt` + `reference` (.pptx or image)
   * - Scenario 3 (`improve_material`): pass `source` (rough .pptx or slide image(s)) + optional `prompt`/`reference`
   */
  async generate(options = {}) {
    const mode = options.mode ?? 'sync';
    const wantFile = Boolean(options.outputPath);
    const reference = await normalizeArtifactPayload(options.reference, 'reference');
    const source = Array.isArray(options.source)
      ? await Promise.all(options.source.map((item, idx) => normalizeArtifactPayload(item, `source[${idx}]`)))
      : await normalizeArtifactPayload(options.source, 'source');

    if (!source && !options.outline && !options.prompt) {
      throw new DeckenError(
        '`generate` requires `outline` / `prompt` (Scenarios 1 & 2) or `source` (Scenario 3).',
        0,
        null,
      );
    }

    const payload = {
      mode,
      engineMode: options.engineMode ?? 'vlm',
      ...(options.deckTitle ? { deckTitle: options.deckTitle } : {}),
      ...(options.prompt ? { prompt: options.prompt } : {}),
      ...(options.outline !== undefined ? { outline: options.outline } : {}),
      ...(reference !== undefined ? { reference } : {}),
      ...(source !== undefined ? { source } : {}),
      ...(options.modelId ? { modelId: options.modelId } : {}),
      ...(options.maxOutputTokens ? { maxOutputTokens: options.maxOutputTokens } : {}),
      ...(options.maxPhysicalRequests ? { maxPhysicalRequests: options.maxPhysicalRequests } : {}),
      includePptxBase64: options.includePptxBase64 ?? (mode === 'sync'),
    };

    let job = await this._requestJson('POST', '/v1/decks', payload, options.signal);
    if (mode === 'async' && (options.wait ?? wantFile)) {
      job = await this.waitForJob(job.jobId, {
        pollIntervalMs: options.pollIntervalMs,
        maxWaitMs: options.maxWaitMs,
        includeBase64: options.includePptxBase64 ?? false,
        onProgress: options.onProgress,
      });
    }

    if (options.outputPath && job.status === 'completed') {
      const pptxBytes = job.pptxBase64
        ? Buffer.from(job.pptxBase64, 'base64')
        : Buffer.from(await this.downloadPptx(job.pptxUrl || job.jobId));
      await mkdir(dirname(options.outputPath), { recursive: true });
      await writeFile(options.outputPath, pptxBytes);
    }

    return job;
  }

  async reconstructSlide(options) {
    const normalized = await normalizeSlidePayload(options, 'slide-01');
    const mode = options.mode ?? 'sync';
    const wantFile = Boolean(options.outputPath);

    const payload = {
      mode,
      engineMode: options.engineMode ?? 'vlm',
      slideId: normalized.slideId,
      ...(normalized.imageUrl
        ? { imageUrl: normalized.imageUrl }
        : { imageDataUrl: normalized.imageDataUrl }),
      ...(options.deckTitle ? { deckTitle: options.deckTitle } : {}),
      ...(options.modelId ? { modelId: options.modelId } : {}),
      ...(options.maxOutputTokens ? { maxOutputTokens: options.maxOutputTokens } : {}),
      ...(options.maxPhysicalRequests ? { maxPhysicalRequests: options.maxPhysicalRequests } : {}),
      includePptxBase64: options.includePptxBase64 ?? (mode === 'sync'),
    };

    let job = await this._requestJson('POST', '/v1/reconstruct', payload, options.signal);
    if (mode === 'async' && (options.wait ?? wantFile)) {
      job = await this.waitForJob(job.jobId, {
        pollIntervalMs: options.pollIntervalMs,
        maxWaitMs: options.maxWaitMs,
        includeBase64: options.includePptxBase64 ?? false,
        onProgress: options.onProgress,
      });
    }

    if (options.outputPath && job.status === 'completed') {
      const pptxBytes = job.pptxBase64
        ? Buffer.from(job.pptxBase64, 'base64')
        : Buffer.from(await this.downloadPptx(job.pptxUrl || job.jobId));
      await mkdir(dirname(options.outputPath), { recursive: true });
      await writeFile(options.outputPath, pptxBytes);
    }

    return job;
  }

  async reconstructDeck(options) {
    if (!Array.isArray(options.slides) || options.slides.length === 0) {
      throw new DeckenError('`reconstructDeck` requires a non-empty `slides` array.', 0, null);
    }
    const normalizedSlides = [];
    for (let i = 0; i < options.slides.length; i++) {
      const s = options.slides[i];
      normalizedSlides.push(
        await normalizeSlidePayload(s, `slide-${String(i + 1).padStart(2, '0')}`),
      );
    }

    const mode = options.mode ?? (normalizedSlides.length > 1 ? 'async' : 'sync');
    const wait = options.wait ?? true;
    const payload = {
      mode,
      engineMode: options.engineMode ?? 'vlm',
      slides: normalizedSlides,
      ...(options.deckTitle ? { deckTitle: options.deckTitle } : {}),
      ...(options.modelId ? { modelId: options.modelId } : {}),
      ...(options.maxOutputTokens ? { maxOutputTokens: options.maxOutputTokens } : {}),
      ...(options.maxPhysicalRequests ? { maxPhysicalRequests: options.maxPhysicalRequests } : {}),
      includePptxBase64: options.includePptxBase64 ?? (mode === 'sync'),
    };

    let job = await this._requestJson('POST', '/v1/reconstruct', payload, options.signal);
    if (mode === 'async' && wait) {
      job = await this.waitForJob(job.jobId, {
        pollIntervalMs: options.pollIntervalMs,
        maxWaitMs: options.maxWaitMs,
        includeBase64: options.includePptxBase64 ?? false,
        onProgress: options.onProgress,
      });
    }

    if (options.outputPath && job.status === 'completed') {
      const pptxBytes = job.pptxBase64
        ? Buffer.from(job.pptxBase64, 'base64')
        : Buffer.from(await this.downloadPptx(job.pptxUrl || job.jobId));
      await mkdir(dirname(options.outputPath), { recursive: true });
      await writeFile(options.outputPath, pptxBytes);
    }

    return job;
  }

  async compareModels(options) {
    if (!Array.isArray(options.modelIds) || options.modelIds.length === 0) {
      throw new DeckenError('`compareModels` requires explicit `modelIds` (never guesses models).', 0, null);
    }
    const imageDataUrl = await normalizeImageDataUrl(options);
    const runs = [];
    for (const modelId of options.modelIds) {
      const job = await this.reconstructSlide({
        slideId: options.slideId || 'comparison-slide',
        imageDataUrl,
        mode: 'sync',
        engineMode: 'vlm',
        modelId,
        maxOutputTokens: options.maxOutputTokens,
        includePptxBase64: options.includePptxBase64 ?? false,
      });
      runs.push({
        modelId,
        jobId: job.jobId,
        durationMs: job.durationMs ?? null,
        wholePageMae: job.metrics?.meanWholePageMae ?? null,
        totalObjectCount: job.metrics?.totalObjectCount ?? 0,
        totalTableObjectCount: job.metrics?.totalTableObjectCount ?? 0,
        usage: job.usage,
        pptxUrl: job.pptxUrl,
      });
    }
    return {
      slideId: options.slideId || 'comparison-slide',
      comparedAt: new Date().toISOString(),
      runs,
    };
  }
}
