/** Speech pipeline: audio handling, diarization maths, word/speaker alignment, cleaning, exports, orchestration. */
import { describe, expect, it } from 'vitest';
import { decodeWav, downmix, encodeWav, resample, toSpeechRate } from '../../src/core/speech/audio';
import { asrChunks, cluster, decodePowerset, FRAME_STEP, reconstruct, speakerAudio, speechRegions, WINDOW, WINDOW_FRAMES, windowStarts, type SegWindow } from '../../src/core/speech/diarize';
import { assignSpeakers, cleanText, dropWordsOutside, finalize, smoothBySentence, toDoc, toJson, toMarkdown, toSrt, toText, toUtterances, toVtt, type Word } from '../../src/core/speech/transcript';
import { transcribeAudio, type SpeechEngine } from '../../src/core/speech/pipeline';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';

/** A WAV of the given layout (little-endian), from per-channel samples in -1..1. */
function wav(channels: Float32Array[], rate: number, bits: 8 | 16 | 24 | 32, float = false, extensible = false): Uint8Array {
  const n = channels[0].length;
  const bps = bits / 8;
  const fmtLen = extensible ? 40 : 16;
  const out = new Uint8Array(12 + 8 + fmtLen + 8 + n * bps * channels.length);
  const dv = new DataView(out.buffer);
  const s = (o: number, t: string): void => [...t].forEach((c, i) => (out[o + i] = c.charCodeAt(0)));
  s(0, 'RIFF');
  dv.setUint32(4, out.length - 8, true);
  s(8, 'WAVE');
  s(12, 'fmt ');
  dv.setUint32(16, fmtLen, true);
  dv.setUint16(20, extensible ? 0xfffe : float ? 3 : 1, true);
  dv.setUint16(22, channels.length, true);
  dv.setUint32(24, rate, true);
  dv.setUint32(28, rate * bps * channels.length, true);
  dv.setUint16(32, bps * channels.length, true);
  dv.setUint16(34, bits, true);
  if (extensible) {
    dv.setUint16(36, 22, true);
    dv.setUint16(38, bits, true);
    dv.setUint16(44, float ? 3 : 1, true); // sub-format GUID starts with the format tag
  }
  const d = 20 + fmtLen;
  s(d, 'data');
  dv.setUint32(d + 4, n * bps * channels.length, true);
  let o = d + 8;
  for (let i = 0; i < n; i++)
    for (const c of channels) {
      const v = c[i];
      if (float) dv.setFloat32(o, v, true);
      else if (bits === 8) out[o] = Math.round(v * 127 + 128);
      else if (bits === 16) dv.setInt16(o, Math.round(v * 32767), true);
      else if (bits === 24) {
        const x = Math.round(v * 8388607);
        out[o] = x & 255;
        out[o + 1] = (x >> 8) & 255;
        out[o + 2] = (x >> 16) & 255;
      } else dv.setInt32(o, Math.round(v * 2147483647), true);
      o += bps;
    }
  return out;
}
const sine = (f: number, rate: number, seconds: number, amp = 0.5): Float32Array => Float32Array.from({ length: Math.round(rate * seconds) }, (_, i) => amp * Math.sin((2 * Math.PI * f * i) / rate));
const rms = (x: Float32Array, from = 0, to = x.length): number => {
  let s = 0;
  for (let i = from; i < to; i++) s += x[i] * x[i];
  return Math.sqrt(s / Math.max(1, to - from));
};

describe('audio', () => {
  it('reads WAV in every common layout', () => {
    const x = sine(440, 16000, 0.1);
    for (const [bits, float, ext] of [
      [8, false, false],
      [16, false, false],
      [24, false, false],
      [32, false, false],
      [32, true, false],
      [16, false, true],
      [32, true, true],
    ] as const) {
      const a = decodeWav(wav([x], 16000, bits, float, ext));
      expect(a.rate).toBe(16000);
      expect(a.channels[0].length).toBe(x.length);
      const tol = bits === 8 ? 0.02 : 1e-3;
      expect(Math.max(...a.channels[0].map((v, i) => Math.abs(v - x[i]))), `${bits}${float ? 'f' : ''}${ext ? 'x' : ''}`).toBeLessThan(tol);
    }
  });

  it('a float WAV is not misread as integers (the bug that makes Whisper hallucinate)', () => {
    const a = decodeWav(wav([sine(300, 16000, 0.5, 0.3)], 16000, 32, true));
    expect(rms(a.channels[0])).toBeCloseTo(0.3 / Math.SQRT2, 2);
  });

  it('downmixes stereo and resamples 48 kHz to 16 kHz without aliasing', () => {
    const l = sine(1000, 48000, 1);
    const r = sine(1000, 48000, 1);
    const mono = toSpeechRate({ rate: 48000, channels: [l, r] });
    expect(mono.length).toBe(16000);
    expect(rms(mono, 1000, 15000)).toBeCloseTo(0.5 / Math.SQRT2, 2); // passband kept
    // a 10 kHz tone cannot exist at 16 kHz: it must be filtered out, not folded down to 6 kHz
    const high = resample(sine(10000, 48000, 1), 48000, 16000);
    expect(rms(high, 1000, 15000)).toBeLessThan(0.01);
    // 44.1 kHz (non-integer ratio) keeps a 3 kHz tone at its level
    expect(rms(resample(sine(3000, 44100, 1), 44100, 16000), 1000, 15000)).toBeCloseTo(0.5 / Math.SQRT2, 2);
    expect(downmix([Float32Array.of(1, 0), Float32Array.of(0, 1)])).toEqual(Float32Array.of(0.5, 0.5));
  });

  it('writes WAV that reads back', () => {
    const x = sine(500, 16000, 0.2);
    const back = decodeWav(encodeWav(x)).channels[0];
    expect(Math.max(...back.map((v, i) => Math.abs(v - x[i])))).toBeLessThan(1e-4);
  });

  it('rejects what is not WAV or not PCM', () => {
    expect(() => decodeWav(new Uint8Array(40))).toThrow();
    const adpcm = wav([sine(100, 8000, 0.01)], 8000, 16);
    new DataView(adpcm.buffer).setUint16(20, 2, true);
    expect(() => decodeWav(adpcm)).toThrow(/encoding/);
  });
});

/** A window with given activity runs per local speaker: [speaker, fromFrame, toFrame]. */
function win(start: number, runs: [number, number, number][]): SegWindow {
  const active = [new Uint8Array(WINDOW_FRAMES), new Uint8Array(WINDOW_FRAMES), new Uint8Array(WINDOW_FRAMES)];
  for (const [s, a, b] of runs) active[s].fill(1, a, b);
  return { start, active };
}

describe('diarization maths', () => {
  it('powerset classes decode to local speakers (pyannote segmentation-3.0)', () => {
    const logits = new Float32Array(7 * 7);
    for (let f = 0; f < 7; f++) logits[f * 7 + f] = 5; // frame f → class f
    const a = decodePowerset(logits, 7);
    expect([...a[0]]).toEqual([0, 1, 0, 0, 1, 1, 0]);
    expect([...a[1]]).toEqual([0, 0, 1, 0, 1, 0, 1]);
    expect([...a[2]]).toEqual([0, 0, 0, 1, 0, 1, 1]);
  });

  it('windows cover the whole recording on the frame grid', () => {
    expect(windowStarts(100000, 148)).toEqual([0]);
    const n = 16000 * 61;
    const s = windowStarts(n, 148);
    expect(s[0]).toBe(0);
    expect(s.every((v) => v % FRAME_STEP === 0)).toBe(true);
    expect(s[s.length - 1] + WINDOW).toBeGreaterThanOrEqual(n);
    for (let i = 1; i < s.length; i++) expect(s[i] - s[i - 1]).toBeLessThanOrEqual(148 * FRAME_STEP);
  });

  it('clusters embeddings by threshold, by count, and attaches small clusters', () => {
    // three voices: directions in 8-D with noise
    let seed = 7;
    const noise = (): number => ((seed = (seed * 16807) % 2147483647) / 2147483647 - 0.5) * 0.2;
    const centre = [0, 1, 2].map((k) => Float32Array.from({ length: 8 }, (_, i) => (i === k ? 1 : 0) + (i === 7 ? 0.3 : 0)));
    const truth: number[] = [];
    const embs: Float32Array[] = [];
    for (let i = 0; i < 60; i++) {
      const k = i % 3;
      truth.push(k);
      embs.push(Float32Array.from(centre[k], (v) => v + noise()));
    }
    const same = (labels: number[]): boolean => truth.every((t, i) => truth.every((u, j) => (t === u) === (labels[i] === labels[j])));
    expect(same(cluster(embs, { threshold: 0.7 }))).toBe(true);
    expect(new Set(cluster(embs, { count: 2 })).size).toBe(2);
    expect(new Set(cluster(embs, { count: 3 })).size).toBe(3);
    // an outlier forms no speaker of its own
    const labels = cluster([...embs, Float32Array.from({ length: 8 }, (_, i) => (i === 5 ? 1 : 0))], { threshold: 0.7 });
    expect(new Set(labels).size).toBe(3);
    expect(cluster([], {})).toEqual([]);
  });

  it('reconstructs global turns from overlapping windows', () => {
    // 20 s: speaker A 0–8 s, B 9–18 s; two windows overlapping 5–15 s, local labels swapped in the second
    const f = (t: number, start: number): number => Math.round((t * 16000 - start) / FRAME_STEP);
    const w1 = win(0, [
      [0, f(0, 0), f(8, 0)],
      [1, f(9, 0), WINDOW_FRAMES],
    ]);
    const s2 = 5 * 16000 - ((5 * 16000) % FRAME_STEP);
    const w2 = win(s2, [
      [2, 0, f(8, s2)],
      [0, f(9, s2), f(15, s2)],
    ]);
    const w3start = 10 * 16000 - ((10 * 16000) % FRAME_STEP);
    const w3 = win(w3start, [[1, 0, f(18, w3start)]]);
    const turns = reconstruct(
      [w1, w2, w3],
      [
        [0, 1, -1],
        [1, -1, 0],
        [-1, 1, -1],
      ],
      20 * 16000,
    );
    const a = turns.filter((t) => t.speaker === 0);
    const b = turns.filter((t) => t.speaker === 1);
    expect(a.length).toBe(1);
    expect(b.length).toBe(1);
    expect(a[0].start).toBeLessThan(0.1);
    expect(a[0].end).toBeCloseTo(8, 0);
    expect(b[0].start).toBeCloseTo(9, 0);
    expect(b[0].end).toBeCloseTo(18, 0);
  });

  it('speaker audio: the frames a speaker talks alone, null when too short', () => {
    const audio = Float32Array.from({ length: WINDOW }, (_, i) => (i < 80000 ? 0.5 : -0.5));
    const w = win(0, [
      [0, 0, 200],
      [1, 150, 400],
      [2, 0, 10],
    ]);
    const a = speakerAudio(audio, w, 0)!;
    // frames 0–149 are its alone, minus 0–9 where speaker 3 talks too
    expect(a.length).toBeCloseTo(140 * FRAME_STEP, -2);
    expect(speakerAudio(audio, w, 2)).toBeNull();
  });

  it('speech regions and Whisper pieces: at most 28 s, cut in pauses or at the quietest point', () => {
    const w = win(0, [[0, 30, 300]]);
    const regions = speechRegions([w], WINDOW);
    expect(regions.length).toBe(1);
    expect(regions[0].start).toBeCloseTo((30 * FRAME_STEP) / 16000 + 0.03 - 0.2, 1);
    // 70 s of speech with a quiet moment at 20 s and 45 s
    const audio = Float32Array.from({ length: 70 * 16000 }, (_, i) => (Math.abs(i / 16000 - 20) < 0.1 || Math.abs(i / 16000 - 45) < 0.1 ? 0 : Math.sin(i)));
    const chunks = asrChunks(audio, [{ start: 0, end: 70 }]);
    expect(chunks.every((c) => c.end - c.start <= 28.001)).toBe(true);
    expect(chunks[0].end).toBeCloseTo(20, 0);
    expect(chunks[1].end).toBeCloseTo(45, 0);
    expect(chunks[chunks.length - 1].end).toBe(70);
    // separate regions are packed into pieces up to 28 s
    expect(asrChunks(audio, [{ start: 0, end: 5 }, { start: 6, end: 12 }, { start: 30, end: 40 }]).map((c) => [c.start, c.end])).toEqual([
      [0, 12],
      [30, 40],
    ]);
  });
});

const W = (text: string, start: number, end: number): Word => ({ text, start, end });

describe('transcript', () => {
  it('removes subtitle credits and loops Whisper invents, keeps real speech', () => {
    expect(cleanText('Добрый день. Субтитры сделал DimaTorzok')).toBe('Добрый день.');
    expect(cleanText('Редактор субтитров А.Синецкая Корректор А.Егорова')).toBe('');
    expect(cleanText('Подписывайтесь на канал! Итак, начнём.')).toBe('Итак, начнём.');
    expect(cleanText('и вот так и вот так и вот так и вот так всё')).toBe('и вот так всё');
    expect(cleanText('да да да да да да да да да да')).toBe('да, да, да');
    // ordinary repetitions and words are left alone
    expect(cleanText('Нет, нет, нет! Это не так.')).toBe('Нет, нет, нет! Это не так.');
    expect(cleanText('Субтитры к фильму включены.')).toBe('Субтитры к фильму включены.');
  });

  it('drops words outside the speech regions', () => {
    const words = [W(' Привет', 1, 1.4), W(' Спасибо', 10, 10.5), W(' мир', 1.5, 1.9)];
    expect(dropWordsOutside(words, [{ start: 0.8, end: 2.2 }]).map((w) => w.text)).toEqual([' Привет', ' мир']);
  });

  it('assigns words to speakers and keeps sentences whole across a boundary a little off', () => {
    const words = [W(' Добрый', 0, 0.4), W(' день.', 0.4, 0.8), W(' Как', 1.0, 1.2), W(' дела?', 1.2, 1.6), W(' Хорошо,', 2.0, 2.4), W(' спасибо.', 2.4, 2.9)];
    // the diarization boundary between A and B falls 0.3 s too late
    const turns = [
      { speaker: 0, start: 0, end: 2.3 },
      { speaker: 1, start: 2.3, end: 3 },
    ];
    const raw = assignSpeakers(words, turns);
    expect(raw).toEqual([0, 0, 0, 0, 0, 1]);
    const smooth = smoothBySentence(words, raw);
    expect(smooth).toEqual([0, 0, 0, 0, 1, 1]);
    const tr = finalize(toUtterances(words, smooth), 3);
    expect(tr.speakers).toEqual(['Спикер 1', 'Спикер 2']);
    expect(tr.utterances.map((u) => `${u.speaker}:${u.text}`)).toEqual(['0:Добрый день. Как дела?', '1:Хорошо, спасибо.']);
  });

  it('exports: text, Markdown, SRT, WebVTT, JSON and a Word document', () => {
    const tr = finalize(
      toUtterances([W(' Здравствуйте.', 0.5, 1.2), W(' Добрый', 61.2, 61.5), W(' вечер!', 61.5, 62)], [0, 1, 1]),
      62,
      'ru',
    );
    tr.speakers = ['Анна', 'Борис <ведущий>'];
    expect(toText(tr)).toBe('Анна [00:00]: Здравствуйте.\n\nБорис <ведущий> [01:01]: Добрый вечер!\n');
    expect(toMarkdown(tr)).toContain('**Анна** _00:00_');
    expect(toSrt(tr)).toBe('1\n00:00:00,500 --> 00:00:01,200\nАнна: Здравствуйте.\n\n2\n00:01:01,200 --> 00:01:02,000\nБорис <ведущий>: Добрый вечер!\n');
    expect(toVtt(tr)).toContain('00:01:01.200 --> 00:01:02.000\n<v Борис &lt;ведущий&gt;>Добрый вечер!');
    expect(JSON.parse(toJson(tr)).utterances[1]).toEqual({ speaker: 'Борис <ведущий>', start: 61.2, end: 62, text: 'Добрый вечер!' });
    const back = docxToDoc(docToDocx(toDoc(tr)).data).doc;
    const texts = back.blocks.map((b) => ('runs' in b ? b.runs.map((r) => r.text).join('') : ''));
    expect(texts.some((t) => t.includes('Анна') && t.includes('Здравствуйте.'))).toBe(true);
  });

  it('long utterances become subtitle cues of two lines at most', () => {
    const text = Array.from({ length: 60 }, (_, i) => `слово${i}`).join(' ');
    const tr = finalize([{ speaker: 0, start: 0, end: 30, text }], 30);
    const cues = toSrt(tr).trim().split('\n\n');
    expect(cues.length).toBeGreaterThan(5);
    for (const c of cues) expect(c.split('\n')[2].length).toBeLessThanOrEqual(84 + 'Спикер 1: '.length);
  });
});

describe('pipeline (models stubbed)', () => {
  /**
   * 24 s: speaker A (a 200 Hz tone) 1–9 s, B (a 700 Hz tone) 11–19 s, A again 20–23 s. The stub models
   * "hear" the tones: segmentation marks frames with energy, the embedding is the tone's frequency, the
   * recogniser returns one word per second of signal.
   */
  const rate = 16000;
  const audio = new Float32Array(24 * rate);
  const put = (f: number, a: number, b: number): void => {
    for (let i = a * rate; i < b * rate; i++) audio[i] = 0.5 * Math.sin((2 * Math.PI * f * i) / rate);
  };
  put(200, 1, 9);
  put(700, 11, 19);
  put(200, 20, 23);
  /** The tone's frequency over the samples that carry it (zero crossings). */
  const freqOf = (x: Float32Array): number => {
    let z = 0;
    let voiced = 0;
    for (let i = 1; i < x.length; i++) {
      if (x[i] === 0 && x[i - 1] === 0) continue;
      voiced++;
      if ((x[i - 1] < 0) !== (x[i] < 0)) z++;
    }
    return (z / 2) * (rate / Math.max(1, voiced));
  };
  // which absolute window is being segmented: identified by matching the slice
  const engine: SpeechEngine = {
    async segment(w) {
      const frames = WINDOW_FRAMES;
      const out = new Float32Array(frames * 7);
      for (let f = 0; f < frames; f++) {
        const seg = w.subarray(f * FRAME_STEP, f * FRAME_STEP + 990);
        const e = rms(seg);
        const fr = freqOf(seg);
        // low tone → local speaker 1, high tone → local speaker 2 (as a real model might, per window)
        out[f * 7 + (e < 0.05 ? 0 : fr < 450 ? 1 : 2)] = 10;
      }
      return out;
    },
    async embed(a) {
      const fr = freqOf(a);
      return Float32Array.of(Math.cos(fr / 300), Math.sin(fr / 300), 0.1);
    },
    async transcribe(piece) {
      const words: Word[] = [];
      for (let s = 0; s + 1 <= piece.length / rate; s++) {
        const seg = piece.subarray(s * rate, (s + 1) * rate);
        if (rms(seg) > 0.1) words.push({ text: freqOf(seg) < 450 ? ' низкий' : ' высокий', start: s, end: s + 0.9 });
      }
      if (words.length) words[words.length - 1].text += '.';
      return words;
    },
  };

  it('finds both voices and gives every word to its speaker', async () => {
    const stages = new Set<string>();
    const res = await transcribeAudio(audio, engine, {}, (st) => stages.add(st));
    expect([...stages]).toEqual(['segment', 'asr', 'speakers', 'done']);
    expect(res.speakers).toEqual(['Спикер 1', 'Спикер 2']);
    expect(res.utterances.map((u) => u.speaker)).toEqual([0, 1, 0]);
    expect(res.utterances[0].text.startsWith('Низкий')).toBe(true);
    expect(res.utterances[1].text.split(' ').every((w) => w.startsWith('высокий') || w.startsWith('Высокий'))).toBe(true);
    expect(res.utterances[1].start).toBeGreaterThan(10.5);
    expect(res.utterances[1].end).toBeLessThan(19.8); // the stub's words last 0.9 s
  });

  it('one speaker when asked; stops when cancelled', async () => {
    const res = await transcribeAudio(audio, engine, { speakers: 1 });
    expect(res.speakers).toEqual(['Спикер 1']);
    const ac = new AbortController();
    ac.abort();
    await expect(transcribeAudio(audio, engine, { signal: ac.signal })).rejects.toThrow(/Отменено/);
  });

  it('decodes pieces in batches when the engine can, and falls back to one by one when a batch fails', async () => {
    // 12 stretches of speech 8 s apart → 6 pieces → 3 lanes of 2 pieces
    const long = new Float32Array(200 * rate);
    for (let k = 0; k < 12; k++) long.set(audio.subarray(rate, 9 * rate), (1 + k * 16) * rate);
    const batches: number[] = [];
    const contexts: string[][] = [];
    const seen: string[] = [];
    const single = await transcribeAudio(long, { ...engine, transcribe: (p, c) => (seen.push(c ?? ''), engine.transcribe(p)) }, { speakers: 1 });
    const batched = await transcribeAudio(long, { ...engine, batchSize: 4, transcribeBatch: async (ps, cs = []) => (batches.push(ps.length), contexts.push(cs), Promise.all(ps.map((p) => engine.transcribe(p)))) }, { speakers: 1 });
    expect(seen.length).toBe(6);
    expect(batches).toEqual([3, 3]);
    // lanes start afresh; then each piece has its lane's previous piece as context (the pieces are alike,
    // so that is the context piece 1 gets when decoding one by one)
    expect(contexts[0]).toEqual(['', '', '']);
    expect(seen[1].length).toBeGreaterThan(0);
    expect(contexts[1]).toEqual([seen[1], seen[1], seen[1]]);
    expect(batched.utterances.map((u) => u.text)).toEqual(single.utterances.map((u) => u.text));
    let calls = 0;
    const failing = await transcribeAudio(long, { ...engine, batchSize: 4, transcribeBatch: async () => (calls++, Promise.reject(new Error('out of memory'))) }, { speakers: 1 });
    expect(calls).toBe(1);
    expect(failing.utterances.map((u) => u.text)).toEqual(single.utterances.map((u) => u.text));
  });

  it('silence gives an empty transcript, not invented text', async () => {
    const res = await transcribeAudio(new Float32Array(5 * rate), engine);
    expect(res.utterances).toEqual([]);
  });
});
