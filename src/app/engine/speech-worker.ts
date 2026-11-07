/**
 * Speech recognition worker: runs the models (transformers.js / ONNX Runtime, WebGPU when available)
 * on 16 kHz mono audio sent by the page, and answers with the transcript. Models are fetched once from
 * Hugging Face and kept in the browser's cache; nothing leaves the device.
 */
import { createEngine, ASR_MODELS } from '../../core/speech/engine';
import { transcribeAudio, type PipelineResult, type Stage } from '../../core/speech/pipeline';

export type Preset = keyof typeof ASR_MODELS;

export type SpeechRequest =
  | { type: 'probe' }
  | { type: 'run'; id: number; audio: Float32Array; preset: Preset; language: string; speakers?: number; beams?: number; terms?: string }
  | { type: 'cancel'; id: number };

export type SpeechResponse =
  | { type: 'probe'; webgpu: boolean; fp16: boolean; adapter?: string }
  | { type: 'download'; id: number; file: string; loaded: number; total: number }
  | { type: 'progress'; id: number; stage: Stage | 'load'; done: number; total: number }
  | { type: 'result'; id: number; transcript: PipelineResult; device: string; seconds: number }
  | { type: 'error'; id: number; message: string; cancelled?: boolean };

const ORT_VERSION = __ORT_VERSION__;

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

let engine: { key: string; value: Awaited<ReturnType<typeof createEngine>> } | null = null;
const running = new Map<number, AbortController>();

const post = (m: SpeechResponse): void => (self as unknown as Worker).postMessage(m);

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
  const ac = new AbortController();
  running.set(req.id, ac);
  const t0 = performance.now();
  try {
    const gpu = await probe();
    let preset = req.preset;
    // half precision needs WebGPU with fp16 shaders; otherwise the 4-bit build of the same model
    if (preset === 'accurate' && !(gpu.webgpu && gpu.fp16)) preset = 'balanced';
    const device = gpu.webgpu ? 'webgpu' : 'wasm';
    const key = `${preset}|${device}|${req.language}|${req.beams ?? 1}`;
    if (engine?.key !== key) {
      await engine?.value.dispose();
      engine = null;
      post({ type: 'progress', id: req.id, stage: 'load', done: 0, total: 1 });
      const tf = await loadTf();
      const value = await createEngine(tf, {
        asr: ASR_MODELS[preset],
        device,
        language: req.language,
        beams: req.beams,
        progress: (p) => post({ type: 'download', id: req.id, file: p.file, loaded: p.loaded, total: p.total }),
      });
      engine = { key, value };
    }
    engine.value.setTerms(req.terms ?? '');
    const transcript = await transcribeAudio(req.audio, engine.value, { language: req.language, speakers: req.speakers, signal: ac.signal }, (stage, done, total) => post({ type: 'progress', id: req.id, stage, done, total }));
    post({ type: 'result', id: req.id, transcript, device: `${device}${preset !== req.preset ? ` (${preset})` : ''}`, seconds: (performance.now() - t0) / 1000 });
  } catch (err) {
    const cancelled = err instanceof DOMException && err.name === 'AbortError';
    post({ type: 'error', id: req.id, message: cancelled ? 'Отменено.' : err instanceof Error ? err.message : String(err), cancelled });
  } finally {
    running.delete(req.id);
  }
};
