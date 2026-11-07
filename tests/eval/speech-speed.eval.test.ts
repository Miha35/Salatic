/**
 * Speed of the speech pipeline (Node, CPU): stage timings on a long conversation assembled from real
 * readers of Russian LibriSpeech. Run: SPEECH_DATA=… SPEECH_MINUTES=10 npm run eval:speech -- speech-speed
 */
import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { decodeWav, toSpeechRate } from '../../src/core/speech/audio';
import { ASR_MODELS, createEngine } from '../../src/core/speech/engine';
import { transcribeAudio, type SpeechEngine } from '../../src/core/speech/pipeline';

const DATA = process.env.SPEECH_DATA ?? '';
const MINUTES = Number(process.env.SPEECH_MINUTES ?? 10);
const PRESET = (process.env.SPEECH_PRESET ?? 'balanced') as keyof typeof ASR_MODELS;
const OUT = process.env.SPEECH_OUT ?? '';
const has = DATA && existsSync(join(DATA, 'ruls/x/index.json'));

/** A conversation of three readers (turns of 1–3 utterances, 0.5 s pauses) of about `minutes`. */
function conversation(minutes: number): Float32Array {
  const idx = JSON.parse(readFileSync(join(DATA, 'ruls/x/index.json'), 'utf8')) as { id: string; file: string }[];
  const readers = [['01', '02', '03'], ['06', '07', '10', '11'], ['15']];
  const pools = readers.map((chs) => idx.filter((x) => chs.some((c) => x.id.startsWith(`poemi_${c}_`))));
  const parts: Float32Array[] = [];
  let total = 0;
  let turn = 0;
  const used = [0, 0, 0];
  let seed = 11;
  const rnd = (): number => ((seed = (seed * 16807) % 2147483647) / 2147483647);
  while (total < minutes * 60 * 16000) {
    const s = turn % 3;
    const n = 1 + Math.floor(rnd() * 3);
    for (let i = 0; i < n; i++) {
      const u = pools[s][used[s]++ % pools[s].length];
      const a = toSpeechRate(decodeWav(new Uint8Array(readFileSync(join(DATA, 'ruls/x', u.file)))));
      parts.push(a, new Float32Array(8000));
      total += a.length + 8000;
    }
    turn++;
  }
  const audio = new Float32Array(total);
  let o = 0;
  for (const p of parts) {
    audio.set(p, o);
    o += p.length;
  }
  return audio;
}

describe.skipIf(!has)('batched Whisper decoding', () => {
  it('gives the words and times of one-by-one decoding', async () => {
    const tf = await import('@huggingface/transformers');
    tf.env.cacheDir = join(DATA, '../models');
    const engine = await createEngine(tf, { asr: ASR_MODELS[PRESET], device: 'cpu', language: 'russian', batch: 3 });
    // three pieces of continuous reading (different lengths), as the pipeline cuts them
    const idx = JSON.parse(readFileSync(join(DATA, 'ruls/x/index.json'), 'utf8')) as { id: string; file: string }[];
    const load = (f: string): Float32Array => toSpeechRate(decodeWav(new Uint8Array(readFileSync(join(DATA, 'ruls/x', f)))));
    const pieces: Float32Array[] = [];
    for (const [chapter, seconds] of [['08', 26], ['14', 18], ['26', 9]] as const) {
      const list = idx.filter((x) => x.id.startsWith(`poemi_${chapter}_`));
      const parts: Float32Array[] = [];
      let n = 0;
      for (const u of list) {
        const a = load(u.file);
        if (n + a.length > seconds * 16000) break;
        parts.push(a);
        n += a.length;
      }
      const p = new Float32Array(n);
      let o = 0;
      for (const a of parts) {
        p.set(a, o);
        o += a.length;
      }
      pieces.push(p);
    }
    const one = [];
    for (const p of pieces) one.push(await engine.transcribe(p));
    // transformers.js aligns a single piece over min(mel frames, 1500) encoder frames, i.e. mostly over the
    // whole 30 s window including the padding; with the same choice the batch must give the same times
    const { transcribeBatchWith } = await import('../../src/core/speech/whisper-batch');
    const asr = (engine as unknown as { _asr?: unknown })._asr;
    const t0 = performance.now();
    const many = await engine.transcribeBatch!(pieces);
    console.log(`batch of ${pieces.length}: ${((performance.now() - t0) / 1000).toFixed(1)} s`);
    const same = asr ? await transcribeBatchWith(tf, asr, pieces, (engine as unknown as { _init: number[] })._init, (n) => Math.floor(n / 160)) : null;
    const diffs: number[] = [];
    many.forEach((ws, k) => {
      const a = one[k].map((w) => w.text).join('');
      expect(ws.map((w) => w.text).join(''), `piece ${k}`).toBe(a);
      if (same) {
        expect(same[k].length).toBe(one[k].length);
        same[k].forEach((w, i) => {
          expect(Math.abs(w.start - one[k][i].start), `same window: start of «${w.text}»`).toBeLessThan(0.021);
          expect(Math.abs(w.end - one[k][i].end), `same window: end of «${w.text}»`).toBeLessThan(0.021);
        });
      }
      ws.forEach((w, i) => diffs.push(Math.abs(w.start - one[k][i].start), Math.abs(w.end - one[k][i].end)));
    });
    diffs.sort((x, y) => x - y);
    console.log(`times vs one by one (window = audio only): median ${diffs[diffs.length >> 1].toFixed(2)} s, max ${diffs[diffs.length - 1].toFixed(2)} s`);
    expect(diffs[Math.floor(diffs.length * 0.9)]).toBeLessThan(0.1);
    await engine.dispose();
  }, 1_800_000);
});

describe.skipIf(!has)(`speech speed (${PRESET}, ${MINUTES} min)`, () => {
  it('stage timings', async () => {
    const tf = await import('@huggingface/transformers');
    tf.env.cacheDir = join(DATA, '../models');
    const engine = await createEngine(tf, { asr: ASR_MODELS[PRESET], device: 'cpu', language: 'russian' });
    const audio = conversation(MINUTES);
    const t = { segment: 0, embed: 0, transcribe: 0, nSeg: 0, nEmb: 0, nAsr: 0 };
    const timed: SpeechEngine = {
      async segment(w) {
        const s = performance.now();
        const r = await engine.segment(w);
        t.segment += performance.now() - s;
        t.nSeg++;
        return r;
      },
      async embed(a) {
        const s = performance.now();
        const r = await engine.embed(a);
        t.embed += performance.now() - s;
        t.nEmb++;
        return r;
      },
      async transcribe(p, c) {
        const s = performance.now();
        const r = await engine.transcribe(p, c);
        t.transcribe += performance.now() - s;
        t.nAsr++;
        return r;
      },
      ...(engine.transcribeBatch
        ? {
            async transcribeBatch(ps: Float32Array[], c?: string[]) {
              const s = performance.now();
              const r = await engine.transcribeBatch!(ps, c);
              t.transcribe += performance.now() - s;
              t.nAsr += ps.length;
              return r;
            },
          }
        : {}),
    };
    const t0 = performance.now();
    const res = await transcribeAudio(audio, timed, { language: 'russian' });
    const wall = (performance.now() - t0) / 1000;
    const seconds = audio.length / 16000;
    const report = {
      preset: PRESET,
      audio_min: +(seconds / 60).toFixed(1),
      wall_s: +wall.toFixed(1),
      realtime_factor: +(wall / seconds).toFixed(3),
      segment_s: +(t.segment / 1000).toFixed(1),
      windows: t.nSeg,
      embed_s: +(t.embed / 1000).toFixed(1),
      embeddings: t.nEmb,
      asr_s: +(t.transcribe / 1000).toFixed(1),
      pieces: t.nAsr,
      speakers: res.speakers.length,
    };
    console.log(JSON.stringify(report));
    if (OUT) writeFileSync(OUT, JSON.stringify(report, null, 1));
    expect(res.speakers.length).toBe(3);
    await engine.dispose();
  }, 7_200_000);
});
