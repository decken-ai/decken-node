export const DEFAULT_BASE_URL = 'https://api.decken.ai';

export class DeckenError extends Error {
  statusCode: number;
  body: unknown;
  constructor(message: string, statusCode: number, body: unknown);
}

export interface DeckenClientOptions {
  apiKey?: string;
  baseUrl?: string;
  timeoutMs?: number;
  allowInsecureLocalhost?: boolean;
  fetch?: typeof globalThis.fetch;
}

export type DeckScenario = 'from_outline' | 'from_outline_with_reference' | 'improve_material';

export interface ArtifactInput {
  kind?: 'pptx' | 'image';
  url?: string;
  dataUrl?: string;
  path?: string;
}

export interface DeckOutlineSlide {
  slideId?: string;
  title: string;
  subtitle?: string;
  points?: string[];
  table?: Array<Array<string | { text: string; bold?: boolean; colSpan?: number }>>;
  diagram?: {
    kind?: 'flow' | 'architecture' | 'comparison' | 'metrics' | 'matrix';
    nodes?: Array<{ id?: string; label: string; detail?: string }>;
  };
}

export interface GenerateDeckOptions {
  deckTitle?: string;
  prompt?: string;
  outline?: string | DeckOutlineSlide[];
  reference?: string | ArtifactInput;
  source?: string | ArtifactInput | Array<string | ArtifactInput>;
  mode?: 'sync' | 'async';
  engineMode?: 'vlm' | 'deterministic';
  modelId?: string;
  maxOutputTokens?: number;
  maxPhysicalRequests?: number;
  includePptxBase64?: boolean;
  outputPath?: string;
  wait?: boolean;
  pollIntervalMs?: number;
  maxWaitMs?: number;
  onProgress?: (job: ReconstructJob) => void;
  signal?: AbortSignal;
}

export interface SlideImageInput {
  slideId?: string;
  imageUrl?: string;
  imagePath?: string;
  imageBytes?: Uint8Array;
  imageDataUrl?: string;
  mimeType?: string;
}

export interface ReconstructAgentDiagnostics {
  editabilityVerdict: 'native_editable' | 'hybrid_editable' | 'needs_refinement';
  nativeEditableObjectCount: number;
  rasterFallbackCount: number;
  nativeEditableRatio: number;
  wholePageMaePass: boolean | null;
  editablePrimitives: {
    texts: number;
    shapes: number;
    tables: number;
    connectors: number;
    groups: number;
    rasterAssets: number;
  };
  agentHints: string[];
}

export interface ReconstructMetrics {
  slideCount: number;
  meanWholePageMae: number | null;
  totalObjectCount: number;
  totalTextObjectCount: number;
  totalShapeObjectCount: number;
  totalPictureObjectCount: number;
  totalConnectorObjectCount: number;
  totalTableObjectCount: number;
  totalGroupCount: number;
}

export interface ReconstructUsage {
  physicalRequests: number;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  costUsd: number;
  records?: unknown[];
}

export interface ReconstructSlideSummary {
  slideId: string;
  objectCount: number;
  textObjectCount: number;
  shapeObjectCount: number;
  pictureObjectCount: number;
  connectorObjectCount: number;
  tableObjectCount: number;
  groupCount: number;
  wholePageMae: number | null;
  previewUrl: string;
}

export interface ReconstructJob {
  jobId: string;
  scenario?: DeckScenario;
  sourceMaterialKind?: 'pptx' | 'image';
  status: 'queued' | 'running' | 'completed' | 'failed';
  mode: 'sync' | 'async';
  engineMode: 'vlm' | 'deterministic';
  modelId: string;
  createdAt: string;
  startedAt?: string;
  completedAt?: string;
  durationMs?: number;
  error?: string;
  pptxUrl: string;
  pollUrl: string;
  pptxBase64?: string;
  metrics?: ReconstructMetrics;
  styleConsistency?: {
    referenceSourceKind: string;
    backgroundHex: string;
    accentHex: string;
    headingFont: string;
    bodyFont: string;
    palette: string[];
    score?: number;
  };
  agentDiagnostics?: ReconstructAgentDiagnostics;
  usage: ReconstructUsage;
  slides: ReconstructSlideSummary[];
}

export interface ReconstructSlideOptions extends SlideImageInput {
  mode?: 'sync' | 'async';
  engineMode?: 'vlm' | 'deterministic';
  deckTitle?: string;
  modelId?: string;
  maxOutputTokens?: number;
  maxPhysicalRequests?: number;
  includePptxBase64?: boolean;
  outputPath?: string;
  wait?: boolean;
  pollIntervalMs?: number;
  maxWaitMs?: number;
  onProgress?: (job: ReconstructJob) => void;
  signal?: AbortSignal;
}

export interface ReconstructDeckOptions {
  slides: SlideImageInput[];
  deckTitle?: string;
  mode?: 'sync' | 'async';
  engineMode?: 'vlm' | 'deterministic';
  modelId?: string;
  maxOutputTokens?: number;
  maxPhysicalRequests?: number;
  includePptxBase64?: boolean;
  outputPath?: string;
  wait?: boolean;
  pollIntervalMs?: number;
  maxWaitMs?: number;
  onProgress?: (job: ReconstructJob) => void;
  signal?: AbortSignal;
}

export interface CompareModelsOptions extends SlideImageInput {
  modelIds: string[];
  maxOutputTokens?: number;
  includePptxBase64?: boolean;
}

export declare class DeckenClient {
  readonly baseUrl: string;
  constructor(options?: DeckenClientOptions);
  getCapabilities(): Promise<Record<string, unknown>>;
  getOpenApiSpec(): Promise<Record<string, unknown>>;
  getAgentManifest(): Promise<Record<string, unknown>>;
  listJobs(): Promise<{ jobs: ReconstructJob[] }>;
  getJob(jobId: string, options?: { includeBase64?: boolean }): Promise<ReconstructJob>;
  downloadPptx(jobIdOrUrl: string): Promise<Uint8Array>;
  downloadSlidePreview(jobId: string, slideId: string): Promise<Uint8Array>;
  waitForJob(
    jobId: string,
    options?: {
      pollIntervalMs?: number;
      maxWaitMs?: number;
      includeBase64?: boolean;
      onProgress?: (job: ReconstructJob) => void;
    },
  ): Promise<ReconstructJob>;
  generate(options: GenerateDeckOptions): Promise<ReconstructJob>;
  reconstructSlide(options: ReconstructSlideOptions): Promise<ReconstructJob>;
  reconstructDeck(options: ReconstructDeckOptions): Promise<ReconstructJob>;
  compareModels(options: CompareModelsOptions): Promise<{
    slideId: string;
    comparedAt: string;
    runs: Array<{
      modelId: string;
      jobId: string;
      durationMs: number | null;
      wholePageMae: number | null;
      totalObjectCount: number;
      totalTableObjectCount: number;
      usage: ReconstructUsage;
      pptxUrl: string;
    }>;
  }>;
}
