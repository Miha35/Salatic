/**
 * Speech → speaker-labelled transcript: speech regions and local speakers (segmentation), Whisper on the
 * speech only (pieces cut in the pauses), voice embeddings, clustering, words given to speakers.
 * The models are behind SpeechEngine (transformers.js in the browser worker and in tests).
 */
import { asrChunks, cluster, decodePowerset, reconstruct, speakerAudio, speechRegions, WINDOW, WINDOW_FRAMES, windowStarts, type SegWindow, type SpeakerTurn } from './diarize';
import { assignSpeakers, cleanText, dropWordsOutside, finalize, smoothBySentence, toUtterances, type Transcript, type Word } from './transcript';

export interface SpeechEngine {
  /** Segmentation logits of a 10 s window (frames × 7, powerset classes). */
  segment(window: Float32Array): Promise<Float32Array>;
  /** A voice embedding of some speech of one speaker. */
  embed(audio: Float32Array): Promise<Float32Array>;
  /**
   * Words with times (seconds from the piece's start) of at most 30 s of speech; `context`: the text just
   * before it (keeps style and spellings consistent from piece to piece).
   */
  transcribe(piece: Float32Array, context?: string): Promise<Word[]>;
  /** Several pieces at once (a GPU decodes a batch about as fast as one piece), each with its context; absent: one by one. */
  transcribeBatch?(pieces: Float32Array[], contexts?: string[]): Promise<Word[][]>;
  /** Pieces per batch for transcribeBatch (1 or absent: one by one, with the previous text as context) */
  batchSize?: number;
}

export interface PipelineOptions {
  /** Exact number of speakers, when known */
  speakers?: number;
  /** Clustering threshold (Euclidean distance of unit embeddings); pyannote 3.1's for WeSpeaker ResNet34 */
  threshold?: number;
  language?: string;
  signal?: AbortSignal;
  /** How recognition ran (batches, or why it fell back to one piece at a time) */
  onNote?: (message: string) => void;
}

export type Stage = 'segment' | 'asr' | 'speakers' | 'done';
export type Progress = (stage: Stage, done: number, total: number) => void;

export interface PipelineResult extends Transcript {
  turns: SpeakerTurn[];
  regions: { start: number; end: number }[];
}

export const PYANNOTE_THRESHOLD = 0.7045654963945799;

const check = (signal?: AbortSignal): void => {
  if (signal?.aborted) throw new DOMException('Отменено', 'AbortError');
};

type Region = { start: number; end: number };

/** 1. Segmentation: 10 s windows (the step grows on long recordings to bound the embeddings) → windows, speech regions. */
export async function segmentStage(audio: Float32Array, engine: Pick<SpeechEngine, 'segment'>, opts: PipelineOptions = {}, progress: Progress = () => {}): Promise<{ windows: SegWindow[]; regions: Region[] }> {
  const n = audio.length;
  const stepSeconds = Math.max(2.5, n / 16000 / 1200);
  const stepFrames = Math.max(1, Math.round((stepSeconds * 16000) / 270));
  const starts = windowStarts(n, stepFrames);
  const windows: SegWindow[] = [];
  for (let i = 0; i < starts.length; i++) {
    check(opts.signal);
    const s = starts[i];
    const piece = new Float32Array(WINDOW);
    piece.set(audio.subarray(s, Math.min(n, s + WINDOW)));
    const logits = await engine.segment(piece);
    const frames = Math.min(WINDOW_FRAMES, Math.floor(logits.length / 7));
    windows.push({ start: s, active: decodePowerset(logits, frames) });
    progress('segment', i + 1, starts.length);
  }
  return { windows, regions: speechRegions(windows, n) };
}

/**
 * 2. Recognition of the speech regions, in pieces cut in the pauses. Each piece is given the text just
 * before it as context. A batching engine decodes the recording in `batchSize` lanes side by side (lane k
 * is the k-th stretch of consecutive pieces), so that every piece but the first of each lane still has
 * the previous piece's text.
 */
export async function recognizeStage(audio: Float32Array, regions: Region[], engine: Pick<SpeechEngine, 'transcribe' | 'transcribeBatch' | 'batchSize'>, opts: PipelineOptions = {}, progress: Progress = () => {}): Promise<Word[]> {
  const chunks = asrChunks(audio, regions);
  const n = chunks.length;
  const total = chunks.reduce((sum, c) => sum + (c.end - c.start), 0);
  let done = 0;
  /** words of each piece, times from the piece's start */
  const results: (Word[] | undefined)[] = new Array(n);
  const pieceOf = (c: Region): Float32Array => audio.subarray(Math.floor(c.start * 16000), Math.ceil(c.end * 16000));
  const store = (i: number, ws: Word[]): void => {
    results[i] = ws;
    done += chunks[i].end - chunks[i].start;
    progress('asr', done, total);
  };
  /** the last 40 words decoded before piece i (after a long pause it starts afresh) */
  const contextFor = (i: number): string => {
    if (i === 0 || chunks[i].start - chunks[i - 1].end >= 10) return '';
    const prev: string[] = [];
    for (let k = i - 1; k >= 0 && prev.length < 40 && results[k]; k--) {
      const r = results[k]!;
      for (let w = r.length - 1; w >= 0 && prev.length < 40; w--) prev.unshift(r[w].text);
    }
    return prev.length ? cleanText(prev.join('')) : '';
  };
  const batch = engine.transcribeBatch ? Math.max(1, engine.batchSize ?? 1) : 1;
  if (batch > 1 && n) {
    const lanes = Math.min(batch, n);
    const per = Math.ceil(n / lanes);
    for (let step = 0; step < per; step++) {
      check(opts.signal);
      const idx: number[] = [];
      for (let lane = 0; lane < lanes; lane++) if (lane * per + step < Math.min(n, (lane + 1) * per)) idx.push(lane * per + step);
      try {
        const res = await engine.transcribeBatch!(idx.map((i) => pieceOf(chunks[i])), idx.map(contextFor));
        idx.forEach((i, k) => store(i, res[k]));
        if (step === 0) opts.onNote?.(`batched decoding: ${idx.length} lanes`);
      } catch (e) {
        // out of GPU memory or an unsupported operation: the rest one piece at a time
        if (opts.signal?.aborted) throw e;
        opts.onNote?.(`batched decoding failed, one piece at a time: ${e instanceof Error ? e.message : e}`);
        break;
      }
    }
  }
  for (let i = 0; i < n; i++) {
    if (results[i]) continue;
    check(opts.signal);
    store(i, await engine.transcribe(pieceOf(chunks[i]), contextFor(i)));
  }
  const words: Word[] = [];
  chunks.forEach((c, i) => {
    for (const w of results[i]!) words.push({ text: w.text, start: c.start + w.start, end: c.start + Math.max(w.start, w.end) });
  });
  return dropWordsOutside(words, regions);
}

/** 3. Speakers: an embedding per local speaker of each window, clustered, turns reconstructed. */
export async function speakerStage(audio: Float32Array, windows: SegWindow[], regions: Region[], engine: Pick<SpeechEngine, 'embed'>, opts: PipelineOptions = {}, progress: Progress = () => {}): Promise<SpeakerTurn[]> {
  if (opts.speakers === 1) return regions.map((r) => ({ speaker: 0, start: r.start, end: r.end }));
  const embs: Float32Array[] = [];
  const owner: [number, number][] = [];
  const jobs: [number, number, Float32Array][] = [];
  windows.forEach((w, wi) => {
    for (let s = 0; s < 3; s++) {
      const a = speakerAudio(audio, w, s);
      if (a) jobs.push([wi, s, a]);
    }
  });
  for (let j = 0; j < jobs.length; j++) {
    check(opts.signal);
    const [wi, s, a] = jobs[j];
    const e = await engine.embed(a);
    if (e.every((v) => Number.isFinite(v))) {
      embs.push(e);
      owner.push([wi, s]);
    }
    progress('speakers', j + 1, jobs.length);
  }
  const labels = cluster(embs, opts.speakers ? { count: opts.speakers } : { threshold: opts.threshold ?? PYANNOTE_THRESHOLD });
  const byWindow = windows.map(() => [-1, -1, -1]);
  owner.forEach(([wi, s], i) => (byWindow[wi][s] = labels[i]));
  return reconstruct(windows, byWindow, audio.length);
}

/** 4. Words and turns → utterances of numbered speakers. */
export function combine(words: Word[], turns: SpeakerTurn[], regions: Region[], duration: number, language?: string): PipelineResult {
  const spk = smoothBySentence(words, assignSpeakers(words, turns));
  return { ...finalize(toUtterances(words, spk), duration, language), turns, regions };
}

/** The whole pipeline, one stage after the other (CPU; the browser runs stages 2 and 3 side by side on a GPU). */
export async function transcribeAudio(audio: Float32Array, engine: SpeechEngine, opts: PipelineOptions = {}, progress: Progress = () => {}): Promise<PipelineResult> {
  const { windows, regions } = await segmentStage(audio, engine, opts, progress);
  const words = await recognizeStage(audio, regions, engine, opts, progress);
  const turns = await speakerStage(audio, windows, regions, engine, opts, progress);
  const res = combine(words, turns, regions, audio.length / 16000, opts.language);
  progress('done', 1, 1);
  return res;
}
