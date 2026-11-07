/**
 * The models of the speech pipeline on transformers.js (ONNX Runtime): Whisper large-v3-turbo with word
 * timestamps, pyannote segmentation-3.0 and WeSpeaker ResNet34 embeddings. The library module is passed
 * in (web build in the browser worker, Node build in tests).
 */
import { WHISPER_CODES } from './languages';
import type { SpeechEngine } from './pipeline';
import type { Word } from './transcript';

type TF = typeof import('@huggingface/transformers');

export type Dtype = 'fp32' | 'fp16' | 'q8' | 'int8' | 'uint8' | 'q4' | 'q4f16' | 'bnb4';

export interface ModelChoice {
  id: string;
  encoder: Dtype;
  decoder: Dtype;
}

/** Presets: the recognition model and its precision (download size in MB, approximate). */
export const ASR_MODELS = {
  /** best accuracy: half-precision encoder (needs WebGPU with fp16) */
  accurate: { id: 'onnx-community/whisper-large-v3-turbo_timestamped', encoder: 'fp16', decoder: 'q4', mb: 1610 },
  /** the same model with 4-bit weights: about half the download, nearly the same accuracy */
  balanced: { id: 'onnx-community/whisper-large-v3-turbo_timestamped', encoder: 'q4', decoder: 'q4', mb: 760 },
  /** small model for weak devices */
  fast: { id: 'onnx-community/whisper-small_timestamped', encoder: 'q4', decoder: 'q4', mb: 300 },
} as const;

/**
 * A short punctuated sample in the language, given to Whisper as "previous text": without it Whisper
 * now and then writes whole pieces in lower case without punctuation (and misses words more often).
 */
export const STYLE_PROMPTS: Record<string, string> = {
  russian: 'Здравствуйте! Да, конечно. Хорошо, давайте начнём: вот что я хотел сказать.',
  ukrainian: 'Добрий день! Так, звісно. Добре, почнімо: ось що я хотів сказати.',
  belarusian: 'Добры дзень! Так, вядома. Добра, пачнём: вось што я хацеў сказаць.',
  english: 'Hello! Yes, of course. Well, let us begin: here is what I wanted to say.',
  german: 'Guten Tag! Ja, natürlich. Gut, fangen wir an: Das wollte ich sagen.',
  french: 'Bonjour ! Oui, bien sûr. Bon, commençons : voici ce que je voulais dire.',
  spanish: '¡Hola! Sí, claro. Bueno, empecemos: esto es lo que quería decir.',
  italian: 'Buongiorno! Sì, certo. Bene, cominciamo: ecco cosa volevo dire.',
};

export const SEGMENTATION_MODEL = 'onnx-community/pyannote-segmentation-3.0';
export const EMBEDDING_MODEL = 'onnx-community/wespeaker-voxceleb-resnet34-LM';

export interface EngineOptions {
  asr: ModelChoice;
  /** where Whisper runs; the small models run on the CPU (WASM) */
  device: 'webgpu' | 'wasm' | 'cpu';
  smallDevice?: 'wasm' | 'cpu';
  language: string;
  /** beam search width (1: greedy) */
  beams?: number;
  /** names and terms to spell as given (added to Whisper's prompt) */
  terms?: string;
  /** no style prompt (measurements) */
  plain?: boolean;
  progress?: (info: { file: string; loaded: number; total: number; status: string }) => void;
}

interface Disposable {
  dispose(): Promise<unknown>;
}

export async function createEngine(tf: TF, o: EngineOptions): Promise<SpeechEngine & { dispose(): Promise<void>; setTerms(terms: string): void }> {
  const progress_callback = (p: { status: string; file?: string; loaded?: number; total?: number }): void => {
    if (o.progress && p.file) o.progress({ file: p.file, loaded: p.loaded ?? 0, total: p.total ?? 0, status: p.status });
  };
  const small = o.smallDevice ?? (o.device === 'cpu' ? 'cpu' : 'wasm');
  const [asr, segModel, segProc, embModel, embFe] = await Promise.all([
    tf.pipeline('automatic-speech-recognition', o.asr.id, { device: o.device, dtype: { encoder_model: o.asr.encoder, decoder_model_merged: o.asr.decoder }, progress_callback } as never),
    tf.AutoModelForAudioFrameClassification.from_pretrained(SEGMENTATION_MODEL, { device: small, dtype: 'fp32', progress_callback } as never),
    tf.AutoProcessor.from_pretrained(SEGMENTATION_MODEL, { progress_callback } as never),
    tf.AutoModel.from_pretrained(EMBEDDING_MODEL, { device: small, dtype: 'fp32', progress_callback } as never),
    tf.AutoFeatureExtractor.from_pretrained(EMBEDDING_MODEL, { progress_callback } as never),
  ]);
  const run = asr as unknown as (audio: Float32Array, opts: Record<string, unknown>) => Promise<{ text: string; chunks?: { text: string; timestamp: [number, number | null] }[] }>;
  // the decoder prompt: <|startofprev|> text <|startoftranscript|> <|lang|> <|transcribe|>
  const tok = (asr as unknown as { tokenizer: { convert_tokens_to_ids(t: string): number; encode(t: string, o: { add_special_tokens: boolean }): number[] } }).tokenizer;
  const code = WHISPER_CODES[o.language] ?? 'ru';
  const ids = { prev: tok.convert_tokens_to_ids('<|startofprev|>'), sot: tok.convert_tokens_to_ids('<|startoftranscript|>'), lang: tok.convert_tokens_to_ids(`<|${code}|>`), task: tok.convert_tokens_to_ids('<|transcribe|>') };
  const promptFor = (context: string): number[] | null => {
    if (o.plain) return null;
    const terms = o.terms?.trim() ? `${o.terms.trim().replace(/\s+/g, ' ')}. ` : '';
    const style = STYLE_PROMPTS[o.language] ?? '';
    const text = `${terms}${style} ${context}`.trim();
    if (!text || [ids.prev, ids.sot, ids.lang, ids.task].some((v) => typeof v !== 'number')) return null;
    // at most ~150 prompt tokens (Whisper's limit is half its 448-token window); the latest words kept
    let t = tok.encode(' ' + text, { add_special_tokens: false });
    if (t.length > 150) {
      const head = tok.encode(' ' + `${terms}${style}`.trim(), { add_special_tokens: false }).slice(0, 80);
      t = [...head, ...tok.encode(' ' + context, { add_special_tokens: false }).slice(-(150 - head.length))];
    }
    return [ids.prev, ...t, ids.sot, ids.lang, ids.task];
  };
  return {
    async segment(window: Float32Array): Promise<Float32Array> {
      const inputs = await (segProc as unknown as (a: Float32Array) => Promise<Record<string, unknown>>)(window);
      const out = (await (segModel as unknown as (i: Record<string, unknown>) => Promise<{ logits: { data: Float32Array } }>)(inputs)).logits;
      return out.data;
    },
    async embed(audio: Float32Array): Promise<Float32Array> {
      const feats = await (embFe as unknown as (a: Float32Array) => Promise<Record<string, unknown>>)(audio);
      const out = (await (embModel as unknown as (i: Record<string, unknown>) => Promise<Record<string, { data: Float32Array }>>)(feats)) as Record<string, { data: Float32Array }>;
      const t = out.embeddings ?? out.last_hidden_state ?? Object.values(out)[0];
      return Float32Array.from(t.data);
    },
    async transcribe(piece: Float32Array, context = ''): Promise<Word[]> {
      const prompt = promptFor(context);
      const r = await run(piece, { language: o.language, task: 'transcribe', return_timestamps: 'word', ...(prompt ? { decoder_input_ids: prompt } : {}), ...(o.beams && o.beams > 1 ? { num_beams: o.beams } : {}) });
      const words: Word[] = [];
      for (const c of r.chunks ?? []) {
        const [s, e] = c.timestamp;
        const start = Number.isFinite(s) ? s : (words[words.length - 1]?.end ?? 0);
        words.push({ text: c.text, start, end: e !== null && Number.isFinite(e) ? e : start + 0.3 });
      }
      return words;
    },
    setTerms(terms: string): void {
      o.terms = terms;
    },
    async dispose(): Promise<void> {
      await Promise.all([asr as unknown as Disposable, segModel as unknown as Disposable, embModel as unknown as Disposable].map((m) => m.dispose?.().catch(() => undefined)));
    },
  };
}
