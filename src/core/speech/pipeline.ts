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
}

export interface PipelineOptions {
  /** Exact number of speakers, when known */
  speakers?: number;
  /** Clustering threshold (Euclidean distance of unit embeddings); pyannote 3.1's for WeSpeaker ResNet34 */
  threshold?: number;
  language?: string;
  signal?: AbortSignal;
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

export async function transcribeAudio(audio: Float32Array, engine: SpeechEngine, opts: PipelineOptions = {}, progress: Progress = () => {}): Promise<PipelineResult> {
  const n = audio.length;
  const duration = n / 16000;
  // 1. segmentation: 10 s windows; the step grows on long recordings to bound the embeddings
  const stepSeconds = Math.max(2.5, duration / 1200);
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
  const regions = speechRegions(windows, n);
  // 2. recognition of the speech, in pieces cut in the pauses
  const chunks = asrChunks(audio, regions);
  const total = chunks.reduce((sum, c) => sum + (c.end - c.start), 0);
  let done = 0;
  let words: Word[] = [];
  let prevEnd = -Infinity;
  for (const c of chunks) {
    check(opts.signal);
    const piece = audio.subarray(Math.floor(c.start * 16000), Math.ceil(c.end * 16000));
    // the previous piece's last words, when it ended shortly before (a long pause starts afresh)
    const context = c.start - prevEnd < 10 ? cleanText(words.slice(-40).map((w) => w.text).join('')) : '';
    const ws = await engine.transcribe(piece, context);
    prevEnd = c.end;
    for (const w of ws) words.push({ text: w.text, start: c.start + w.start, end: c.start + Math.max(w.start, w.end) });
    done += c.end - c.start;
    progress('asr', done, total);
  }
  words = dropWordsOutside(words, regions);
  // 3. speakers
  let turns: SpeakerTurn[] = [];
  if (opts.speakers === 1) turns = regions.map((r) => ({ speaker: 0, start: r.start, end: r.end }));
  else {
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
    turns = reconstruct(windows, byWindow, n);
  }
  const spk = smoothBySentence(words, assignSpeakers(words, turns));
  const tr = finalize(toUtterances(words, spk), duration, opts.language);
  progress('done', 1, 1);
  return { ...tr, turns, regions };
}
