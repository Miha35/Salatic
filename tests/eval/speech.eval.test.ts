/**
 * Accuracy of the speech pipeline on real Russian speech (not part of `npm test`: it downloads ~1 GB of
 * models and runs for minutes). Data: FLEURS ru_ru dev and Russian LibriSpeech test, prepared in
 * $SPEECH_DATA (see README). Run: npm run eval:speech
 *
 * - WER/CER of the recognition (normalised: case, punctuation, ё);
 * - conversations assembled from utterances of different real readers: DER and the share of words given
 *   to the right speaker.
 */
import { afterAll, describe, expect, it } from 'vitest';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { decodeWav, toSpeechRate } from '../../src/core/speech/audio';
import { ASR_MODELS, createEngine } from '../../src/core/speech/engine';
import { transcribeAudio, type PipelineResult } from '../../src/core/speech/pipeline';

const DATA = process.env.SPEECH_DATA ?? '';
const PRESET = (process.env.SPEECH_PRESET ?? 'balanced') as keyof typeof ASR_MODELS;
const N = Number(process.env.SPEECH_N ?? 25);
const BEAMS = Number(process.env.SPEECH_BEAMS ?? 1);
const OUT = process.env.SPEECH_OUT ?? '';
const PLAIN = process.env.SPEECH_PLAIN === '1';
/** Pieces per batch in the conversations (4 = the browser's GPU path, which drops the previous piece's text as context) */
const BATCH = Number(process.env.SPEECH_BATCH ?? 1);

/** Lower case, ё → е, letters and digits only. */
export const normalize = (s: string): string[] =>
  s
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^\p{L}\p{N}\s-]/gu, ' ')
    .replace(/(^|\s)-+|-+(\s|$)/g, ' ')
    .split(/\s+/)
    .filter(Boolean);

/** Levenshtein distance between sequences. */
export function edits<T>(a: T[], b: T[]): number {
  const d = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    let prev = d[0];
    d[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const t = d[j];
      d[j] = Math.min(d[j] + 1, d[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = t;
    }
  }
  return d[b.length];
}

/** DER over 10 ms frames, best one-to-one mapping of speakers (greedy on overlap), no collar. */
export function der(ref: { speaker: number; start: number; end: number }[], hyp: { speaker: number; start: number; end: number }[], duration: number): number {
  const F = Math.ceil(duration * 100);
  const r = new Int16Array(F).fill(-1);
  const h = new Int16Array(F).fill(-1);
  for (const t of ref) for (let f = Math.floor(t.start * 100); f < Math.min(F, Math.ceil(t.end * 100)); f++) r[f] = t.speaker;
  for (const t of hyp) for (let f = Math.floor(t.start * 100); f < Math.min(F, Math.ceil(t.end * 100)); f++) if (h[f] < 0) h[f] = t.speaker;
  const co = new Map<string, number>();
  for (let f = 0; f < F; f++) if (r[f] >= 0 && h[f] >= 0) co.set(`${r[f]}|${h[f]}`, (co.get(`${r[f]}|${h[f]}`) ?? 0) + 1);
  const map = new Map<number, number>();
  const used = new Set<number>();
  for (const [k] of [...co.entries()].sort((a, b) => b[1] - a[1])) {
    const [rs, hs] = k.split('|').map(Number);
    if (map.has(hs) || used.has(rs)) continue;
    map.set(hs, rs);
    used.add(rs);
  }
  let errors = 0;
  let speech = 0;
  for (let f = 0; f < F; f++) {
    if (r[f] >= 0) speech++;
    if (r[f] < 0 && h[f] < 0) continue;
    if (r[f] < 0 || h[f] < 0 || map.get(h[f]) !== r[f]) errors++;
  }
  return errors / Math.max(1, speech);
}

interface Item {
  file: string;
  text: string;
  reader?: number;
}

function fleurs(): Item[] {
  const rows = readFileSync(join(DATA, 'fleurs/dev.tsv'), 'utf8').trim().split('\n').map((l) => l.split('\t'));
  // one recording per sentence, spread over the set
  const seen = new Set<string>();
  const out: Item[] = [];
  for (const r of rows) {
    if (seen.has(r[0])) continue;
    seen.add(r[0]);
    out.push({ file: join(DATA, 'fleurs/x/dev', r[1]), text: r[2] });
  }
  return out.filter((_, i) => i % Math.max(1, Math.floor(out.length / N)) === 0).slice(0, N);
}

/** RuLS chapters → readers (CHAPTERS_READERS.TXT, book 2145). */
const READER: Record<string, number> = { '01': 2671, '02': 2671, '03': 2671, '04': 3056, '05': 3056, '06': 4471, '07': 4471, '10': 4471, '11': 4471, '08': 2826, '09': 2826, '14': 5548, '36': 5548, '15': 4372, '25': 4091, '26': 4091, '27': 4091 };

function ruls(): Item[] {
  const idx = JSON.parse(readFileSync(join(DATA, 'ruls/x/index.json'), 'utf8')) as { id: string; file: string; raw: string }[];
  return idx.map((x) => ({ file: join(DATA, 'ruls/x', x.file), text: x.raw, reader: READER[/poemi_(\d\d)_/.exec(x.id)?.[1] ?? ''] }));
}

const load = (f: string): Float32Array => toSpeechRate(decodeWav(new Uint8Array(readFileSync(f))));

const has = DATA && existsSync(join(DATA, 'fleurs/dev.tsv')) && existsSync(join(DATA, 'ruls/x/index.json'));

describe.skipIf(!has)(`speech accuracy (${PRESET}, beams ${BEAMS}, batch ${BATCH}${PLAIN ? ', no prompt' : ''})`, () => {
  let engine: Awaited<ReturnType<typeof createEngine>>;
  const report: Record<string, unknown> = { preset: PRESET, beams: BEAMS, batch: BATCH, plain: PLAIN };

  it('loads the models', async () => {
    const tf = await import('@huggingface/transformers');
    tf.env.cacheDir = join(DATA, '../models');
    engine = await createEngine(tf, { asr: ASR_MODELS[PRESET], device: 'cpu', language: 'russian', beams: BEAMS, plain: PLAIN, batch: BATCH });
  }, 1_800_000);

  it('recognition: WER and CER on real speech', async () => {
    const items = [...fleurs(), ...ruls().filter((_, i) => i % 50 === 7).slice(0, N)];
    let wErr = 0;
    let wTot = 0;
    let cErr = 0;
    let cTot = 0;
    const samples: string[] = [];
    const t0 = Date.now();
    let audioS = 0;
    for (const it of items) {
      const audio = load(it.file);
      audioS += audio.length / 16000;
      const words = await engine.transcribe(audio);
      const hyp = words.map((w) => w.text).join('');
      const r = normalize(it.text);
      const h = normalize(hyp);
      wErr += edits(r, h);
      wTot += r.length;
      cErr += edits([...r.join(' ')], [...h.join(' ')]);
      cTot += r.join(' ').length;
      if (edits(r, h)) samples.push(`REF: ${it.text}\nHYP: ${hyp.trim()}`);
    }
    report.asr = { utterances: items.length, wer: wErr / wTot, cer: cErr / cTot, audio_s: audioS, compute_s: (Date.now() - t0) / 1000, errors: samples };
    console.log(`WER ${((100 * wErr) / wTot).toFixed(2)}%  CER ${((100 * cErr) / cTot).toFixed(2)}%  on ${items.length} utterances, ${audioS.toFixed(0)} s audio in ${((Date.now() - t0) / 1000).toFixed(0)} s`);
    expect(wErr / wTot).toBeLessThan(0.15);
  }, 7_200_000);

  /** A conversation: turns of 1–2 utterances from the given readers, with pauses; reference turns and text. */
  function conversation(readers: number[], turns: number, seed: number): { audio: Float32Array; ref: { speaker: number; start: number; end: number }[]; text: string } {
    const pool = ruls();
    const byReader = readers.map((r) => pool.filter((x) => x.reader === r));
    let rnd = seed;
    const next = (): number => ((rnd = (rnd * 1103515245 + 12345) % 2147483648) / 2147483648);
    const parts: Float32Array[] = [];
    const ref: { speaker: number; start: number; end: number }[] = [];
    let t = 0.5;
    parts.push(new Float32Array(8000));
    const texts: string[] = [];
    let last = -1;
    for (let k = 0; k < turns; k++) {
      let s = Math.floor(next() * readers.length);
      if (s === last) s = (s + 1) % readers.length;
      last = s;
      const list = byReader[s];
      const u = list[Math.floor(next() * list.length)];
      const a = load(u.file);
      ref.push({ speaker: s, start: t, end: t + a.length / 16000 });
      parts.push(a);
      texts.push(u.text);
      t += a.length / 16000;
      const pause = new Float32Array(Math.round((0.3 + next() * 0.9) * 16000));
      parts.push(pause);
      t += pause.length / 16000;
    }
    const audio = new Float32Array(parts.reduce((s, p) => s + p.length, 0));
    let o = 0;
    for (const p of parts) {
      audio.set(p, o);
      o += p.length;
    }
    return { audio, ref, text: texts.join(' ') };
  }

  const scoreWords = (res: PipelineResult, ref: { speaker: number; start: number; end: number }[]): number => {
    // each utterance's words, timed proportionally, against the reference speaker at that time
    const pairs = new Map<string, number>();
    let total = 0;
    const votes: [number, number, number][] = [];
    for (const u of res.utterances) {
      const n = u.text.split(/\s+/).length;
      for (let i = 0; i < n; i++) {
        const t = u.start + ((i + 0.5) / n) * (u.end - u.start);
        const r = ref.find((x) => t >= x.start - 0.3 && t <= x.end + 0.3);
        if (!r) continue;
        votes.push([r.speaker, u.speaker, 1]);
        pairs.set(`${r.speaker}|${u.speaker}`, (pairs.get(`${r.speaker}|${u.speaker}`) ?? 0) + 1);
        total++;
      }
    }
    const map = new Map<number, number>();
    const used = new Set<number>();
    for (const [k] of [...pairs.entries()].sort((a, b) => b[1] - a[1])) {
      const [rs, hs] = k.split('|').map(Number);
      if (map.has(hs) || used.has(rs)) continue;
      map.set(hs, rs);
      used.add(rs);
    }
    return votes.filter(([r, h]) => map.get(h) === r).length / Math.max(1, total);
  };

  for (const [name, readers, turns] of [
    ['two speakers (woman, man)', [2671, 4471], 12],
    ['three speakers (woman, two men)', [3056, 4471, 2826], 15],
    ['four speakers', [2671, 4372, 4091, 2826], 16],
  ] as const) {
    it(`diarization: ${name}`, async () => {
      const conv = conversation([...readers], turns, readers[0] + turns);
      const t0 = Date.now();
      const res = await transcribeAudio(conv.audio, engine, { language: 'russian' });
      const d = der(conv.ref, res.turns, conv.audio.length / 16000);
      const words = scoreWords(res, conv.ref);
      const wer = edits(normalize(conv.text), normalize(res.utterances.map((u) => u.text).join(' '))) / normalize(conv.text).length;
      const entry = { speakers_found: res.speakers.length, speakers_true: readers.length, der: d, word_speaker_accuracy: words, wer, audio_s: conv.audio.length / 16000, compute_s: (Date.now() - t0) / 1000, sample: res.utterances.slice(0, 4).map((u) => `${res.speakers[u.speaker]}: ${u.text}`) };
      (report.diarization ??= {} as Record<string, unknown>) && ((report.diarization as Record<string, unknown>)[name] = entry);
      console.log(name, JSON.stringify({ ...entry, sample: undefined }));
      expect(res.speakers.length).toBe(readers.length);
      expect(words).toBeGreaterThan(0.9);
    }, 3_600_000);
  }

  afterAll(() => {
    if (OUT) writeFileSync(OUT, JSON.stringify(report, null, 1));
  });
});
