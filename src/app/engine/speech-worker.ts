/**
 * Speech recognition worker: runs the models (transformers.js / ONNX Runtime, WebGPU when available) on
 * 16 kHz mono audio sent by the page. Models are fetched once from Hugging Face and kept in the browser's
 * cache; nothing leaves the device.
 *
 * Roles (the page runs one or two of these workers):
 * - `run`: everything, one stage after the other (no GPU: the stages would only compete for the CPU);
 * - `diarize`: segmentation (CPU) → speech regions, then voice embeddings (GPU if any) → speaker turns;
 * - `recognize`: Whisper on the GPU over the regions, in batches — side by side with `diarize`.
 */
import { ASR_MODELS, createDiarizer, createEngine, createRecognizer, type Diarizer, type Recognizer } from '../../core/speech/engine';
import { recognizeStage, segmentStage, speakerStage, transcribeAudio, type PipelineResult, type Stage } from '../../core/speech/pipeline';
import type { SegWindow, SpeakerTurn } from '../../core/speech/diarize';
import type { Word } from '../../core/speech/transcript';

export type Preset = keyof typeof ASR_MODELS;
type Region = { start: number; end: number };

export type SpeechRequest =
  | { type: 'probe' }
  | { type: 'run'; id: number; audio: Float32Array; preset: Preset; language: string; speakers?: number; terms?: string }
  | { type: 'diarize'; id: number; audio: Float32Array; speakers?: number }
  | { type: 'recognize'; id: number; audio: Float32Array; preset: Preset; language: string; terms?: string }
  | { type: 'regions'; id: number; regions: Region[] }
  | { type: 'cancel'; id: number };

export type SpeechResponse =
  | { type: 'probe'; webgpu: boolean; fp16: boolean; adapter?: string }
  | { type: 'download'; id: number; file: string; loaded: number; total: number }
  | { type: 'progress'; id: number; stage: Stage | 'load'; done: number; total: number }
  | { type: 'result'; id: number; transcript: PipelineResult; device: string; seconds: number }
  | { type: 'regions'; id: number; regions: Region[] }
  | { type: 'turns'; id: number; turns: SpeakerTurn[] }
  | { type: 'words'; id: number; words: Word[]; device: string }
  | { type: 'note'; id: number; message: string }
  | { type: 'error'; id: number; message: string; cancelled?: boolean };

const ORT_VERSION = __ORT_VERSION__;
/** Pieces decoded together on a GPU: the decoder's cost per step hardly grows with the batch */
const GPU_BATCH = 4;

let tfPromise: Promise<typeof import('@huggingface/transformers')> | null = null;
/** transformers.js with ONNX Runtime from this site (the runtime is stored gzipped: it is over 25 MB). */
function loadTf(): Promise<typeof import('@huggingface/transformers')> {
  tfPromise ??= (async () => {
    const tf = await import('@huggingface/transformers');
    const base = new URL(`/ort/${ORT_VERSION}/`, self.location.origin).href;
    const res = await fetch(`${base}ort-wasm-simd-threaded.asyncify.wasm.gz`);
    if (!res.ok || !res.body) throw new Error('Не удалось загрузить среду выполнения моделей.');
    const wasm = await new Response(res.body.pipeThrough(new DecompressionStream('gzip'))).arrayBuffer();
    const onnx = tf.env.backends.onnx as { wasm?: { wasmPaths?: unknown; wasmBinary?: ArrayBuffer; numThreads?: number } };
    if (onnx.wasm) {
      onnx.wasm.wasmPaths = { mjs: `${base}ort-wasm-simd-threaded.asyncify.mjs` };
      onnx.wasm.wasmBinary = wasm;
      // threads need cross-origin isolation (COOP + COEP, see public/_headers; the model hosts send CORS headers)
      onnx.wasm.numThreads = self.crossOriginIsolated ? Math.max(1, Math.min(8, (navigator.hardwareConcurrency || 2) - 1)) : 1;
    }
    tf.env.allowLocalModels = false;
    tf.env.useBrowserCache = true;
    return tf;
  })();
  return tfPromise;
}

interface GpuInfo {
  webgpu: boolean;
  fp16: boolean;
  adapter?: string;
}
let gpuInfo: Promise<GpuInfo> | null = null;
function probe(): Promise<GpuInfo> {
  gpuInfo ??= (async () => {
    const gpu = (navigator as Navigator & { gpu?: { requestAdapter(o?: unknown): Promise<{ features: Set<string>; info?: { vendor?: string; architecture?: string } } | null> } }).gpu;
    if (!gpu) return { webgpu: false, fp16: false };
    try {
      const a = await gpu.requestAdapter({ powerPreference: 'high-performance' });
      if (!a) return { webgpu: false, fp16: false };
      return { webgpu: true, fp16: a.features.has('shader-f16'), adapter: [a.info?.vendor, a.info?.architecture].filter(Boolean).join(' ') || undefined };
    } catch {
      return { webgpu: false, fp16: false };
    }
  })();
  return gpuInfo;
}

/** Models kept between runs (keyed by what they were made for). */
let full: { key: string; value: Awaited<ReturnType<typeof createEngine>> } | null = null;
let recognizer: { key: string; value: Recognizer } | null = null;
let diarizer: { key: string; value: Diarizer } | null = null;
const running = new Map<number, AbortController>();
/** Speech regions for a `recognize` job, sent by the page once `diarize` has found them. */
const regionWaiters = new Map<number, (r: Region[]) => void>();

const post = (m: SpeechResponse): void => (self as unknown as Worker).postMessage(m);

async function usePreset(preset: Preset): Promise<{ preset: Preset; gpu: GpuInfo }> {
  const gpu = await probe();
  // half precision needs WebGPU with fp16 shaders; otherwise the 4-bit build of the same model
  return { preset: preset === 'accurate' && !(gpu.webgpu && gpu.fp16) ? 'balanced' : preset, gpu };
}

self.onmessage = async (e: MessageEvent<SpeechRequest>) => {
  const req = e.data;
  if (req.type === 'probe') {
    post({ type: 'probe', ...(await probe()) });
    return;
  }
  if (req.type === 'cancel') {
    running.get(req.id)?.abort();
    return;
  }
  if (req.type === 'regions') {
    regionWaiters.get(req.id)?.(req.regions);
    regionWaiters.delete(req.id);
    return;
  }
  const ac = new AbortController();
  running.set(req.id, ac);
  const t0 = performance.now();
  const onDownload = (p: { file: string; loaded: number; total: number }): void => post({ type: 'download', id: req.id, file: p.file, loaded: p.loaded, total: p.total });
  const onProgress = (stage: Stage, done: number, total: number): void => post({ type: 'progress', id: req.id, stage, done, total });
  const onNote = (message: string): void => post({ type: 'note', id: req.id, message });
  try {
    if (req.type === 'run') {
      const { preset, gpu } = await usePreset(req.preset);
      const device = gpu.webgpu ? 'webgpu' : 'wasm';
      const key = `${preset}|${device}|${req.language}`;
      if (full?.key !== key) {
        await full?.value.dispose();
        full = null;
        post({ type: 'progress', id: req.id, stage: 'load', done: 0, total: 1 });
        const tf = await loadTf();
        full = { key, value: await createEngine(tf, { asr: ASR_MODELS[preset], device, language: req.language, batch: gpu.webgpu ? GPU_BATCH : 1, progress: onDownload }) };
      }
      full.value.setTerms(req.terms ?? '');
      const transcript = await transcribeAudio(req.audio, full.value, { language: req.language, speakers: req.speakers, signal: ac.signal, onNote }, onProgress);
      post({ type: 'result', id: req.id, transcript, device: `${device}${preset !== req.preset ? ` (${preset})` : ''}`, seconds: (performance.now() - t0) / 1000 });
    } else if (req.type === 'diarize') {
      const gpu = await probe();
      const key = gpu.webgpu ? 'gpu' : 'cpu';
      if (diarizer?.key !== key) {
        await diarizer?.value.dispose();
        const tf = await loadTf();
        diarizer = { key, value: await createDiarizer(tf, { device: 'wasm', embedDevice: gpu.webgpu ? 'webgpu' : 'wasm', progress: onDownload }) };
      }
      const { windows, regions } = await segmentStage(req.audio, diarizer.value, { signal: ac.signal }, onProgress);
      post({ type: 'regions', id: req.id, regions });
      const turns = await speakerStage(req.audio, windows as SegWindow[], regions, diarizer.value, { speakers: req.speakers, signal: ac.signal }, onProgress);
      post({ type: 'turns', id: req.id, turns });
    } else if (req.type === 'recognize') {
      // registered before the first await: the page may send the regions as soon as the other worker has them
      const regionsP = new Promise<Region[]>((resolve) => regionWaiters.set(req.id, resolve));
      const { preset, gpu } = await usePreset(req.preset);
      const device = gpu.webgpu ? 'webgpu' : 'wasm';
      const key = `${preset}|${device}|${req.language}`;
      // the models load while the other worker looks for speech
      if (recognizer?.key !== key) {
        await recognizer?.value.dispose();
        recognizer = null;
        post({ type: 'progress', id: req.id, stage: 'load', done: 0, total: 1 });
        const tf = await loadTf();
        recognizer = { key, value: await createRecognizer(tf, { asr: ASR_MODELS[preset], device, language: req.language, batch: gpu.webgpu ? GPU_BATCH : 1, progress: onDownload }) };
      }
      recognizer.value.setTerms(req.terms ?? '');
      // the page leaves time spent waiting for a first download out of the measured speed
      post({ type: 'progress', id: req.id, stage: 'load', done: 1, total: 1 });
      const regions = await Promise.race([
        regionsP,
        new Promise<never>((_, reject) => ac.signal.addEventListener('abort', () => reject(new DOMException('Отменено', 'AbortError')), { once: true })),
      ]);
      const words = await recognizeStage(req.audio, regions, recognizer.value, { language: req.language, signal: ac.signal, onNote }, onProgress);
      post({ type: 'words', id: req.id, words, device: `${device}${preset !== req.preset ? ` (${preset})` : ''}` });
    }
  } catch (err) {
    const cancelled = err instanceof DOMException && err.name === 'AbortError';
    post({ type: 'error', id: req.id, message: cancelled ? 'Отменено.' : err instanceof Error ? err.message : String(err), cancelled });
  } finally {
    running.delete(req.id);
    regionWaiters.delete(req.id);
  }
};
