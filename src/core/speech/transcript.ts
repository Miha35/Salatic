/**
 * Transcript: words (with times) + speaker turns → speaker-labelled utterances; cleaning of Whisper's
 * known failure modes; exports (TXT, Markdown, SRT, WebVTT, JSON, document model for DOCX/PDF/ODT).
 */
import type { Block, Doc } from '../model';
import type { SpeakerTurn } from './diarize';

export interface Word {
  text: string; // as Whisper gives it (leading space between words)
  start: number;
  end: number;
}

export interface Utterance {
  speaker: number;
  start: number;
  end: number;
  text: string;
}

export interface Transcript {
  /** Display names, by speaker index ("Спикер 1"…) */
  speakers: string[];
  utterances: Utterance[];
  duration: number;
  language?: string;
}

// ------------------------------------------------------------------------------------- cleaning

/** Credits and calls to action Whisper invents on silence, music and noise (it learnt them from subtitles). */
const HALLUCINATIONS: RegExp[] = [
  /субтитр\S*\s+(сделал|создавал|создал|подготовил|подогнал|делал|выполнил|редактировал)\S*[^.!?]*/giu,
  /редактор\s+субтитров(?:\s+(?!корректор)[\p{L}.]+){0,2}(?:\s+корректор(?:\s+[\p{L}.]+){0,2})?/giu,
  /корректор\s+[\p{L}]\.\s*[\p{L}]+/giu,
  /dimatorzok/giu,
  /продолжение\s+следует\.*/giu,
  /(подписывайтесь|подпишитесь)\s+на\s+(наш\s+)?канал[^.!?]*[.!?]?/giu,
  /ставьте\s+лайк\S*[^.!?]*[.!?]?/giu,
  /спасибо\s+за\s+просмотр[.!]?/giu,
  /amara\.org[^.!?]*/giu,
  /субтитры\s+(от|by)\s+\S+/giu,
];

/** Remove invented credits; collapse loops ("the same phrase over and over") Whisper falls into. */
export function cleanText(text: string): string {
  let t = text;
  for (const re of HALLUCINATIONS) t = t.replace(re, ' ');
  // a single word repeated 8+ times → 3 times
  t = t.replace(/(?<![\p{L}-])([\p{L}-]+)(?:,?\s+\1(?![\p{L}-])){7,}/giu, '$1, $1, $1');
  // loops: a phrase of 3+ words repeated 3+ times in a row → once
  t = t.replace(/((?:\S+\s+){2,8}?\S+?)([,.!?…]?\s+\1){2,}/gu, '$1');
  return t.replace(/\s+([,.!?…:;])/g, '$1').replace(/\s{2,}/g, ' ').trim();
}

/** Words Whisper put where no one speaks (outside the speech regions) are inventions. */
export function dropWordsOutside(words: Word[], regions: { start: number; end: number }[], tolerance = 0.4): Word[] {
  return words.filter((w) => {
    const mid = (w.start + w.end) / 2;
    return regions.some((r) => mid >= r.start - tolerance && mid <= r.end + tolerance);
  });
}

// ------------------------------------------------------------------------------------- alignment

/** The speaker of each word: the turn it overlaps most, else the nearest one (within 1.5 s), else the previous word's. */
export function assignSpeakers(words: Word[], turns: SpeakerTurn[]): number[] {
  const out: number[] = [];
  let prev = turns[0]?.speaker ?? 0;
  for (const w of words) {
    const s = Math.min(w.start, w.end);
    const e = Math.max(w.start, w.end, s + 0.02);
    const overlap = new Map<number, number>();
    let near = -1;
    let nearD = Infinity;
    for (const t of turns) {
      const o = Math.min(e, t.end) - Math.max(s, t.start);
      if (o > 0) overlap.set(t.speaker, (overlap.get(t.speaker) ?? 0) + o);
      const d = o > 0 ? 0 : Math.min(Math.abs(t.start - e), Math.abs(s - t.end));
      if (d < nearD) {
        nearD = d;
        near = t.speaker;
      }
    }
    let spk = prev;
    if (overlap.size) spk = [...overlap.entries()].sort((a, b) => b[1] - a[1])[0][0];
    else if (near >= 0 && nearD <= 1.5) spk = near;
    out.push(spk);
    prev = spk;
  }
  return out;
}

const SENTENCE_END = /[.!?…]["»)]*$/u;

/**
 * Speaker changes rarely fall inside a sentence: words at a sentence's edges given to another speaker
 * (a diarization boundary a few hundred ms off) go to the sentence's main speaker.
 */
export function smoothBySentence(words: Word[], spk: number[], maxEdge = 0.8): number[] {
  const out = [...spk];
  let i = 0;
  while (i < words.length) {
    let j = i;
    while (j < words.length - 1 && !SENTENCE_END.test(words[j].text.trim())) j++;
    // the sentence words[i..j]: its main speaker by spoken time
    const time = new Map<number, number>();
    for (let k = i; k <= j; k++) time.set(out[k], (time.get(out[k]) ?? 0) + Math.max(0.05, words[k].end - words[k].start));
    const main = [...time.entries()].sort((a, b) => b[1] - a[1])[0][0];
    // leading and trailing runs of other speakers, if short
    const fixEdge = (from: number, step: number): void => {
      let k = from;
      let dur = 0;
      const run: number[] = [];
      while (k >= i && k <= j && out[k] !== main) {
        dur += Math.max(0.05, words[k].end - words[k].start);
        run.push(k);
        k += step;
      }
      if (run.length && run.length < j - i + 1 && dur <= maxEdge) for (const r of run) out[r] = main;
    };
    fixEdge(i, 1);
    fixEdge(j, -1);
    i = j + 1;
  }
  return out;
}

/** Words + speakers → utterances (a new one at a speaker change or a long pause). */
export function toUtterances(words: Word[], spk: number[], pause = 4): Utterance[] {
  const out: Utterance[] = [];
  words.forEach((w, i) => {
    const last = out[out.length - 1];
    if (last && last.speaker === spk[i] && w.start - last.end < pause) {
      last.text += w.text;
      last.end = Math.max(last.end, w.end);
    } else out.push({ speaker: spk[i], start: w.start, end: w.end, text: w.text });
  });
  return out
    .map((u) => ({ ...u, text: tidy(u.text) }))
    .filter((u) => u.text);
}

/** Spacing and the first letter of an utterance. */
function tidy(t: string): string {
  const s = cleanText(t.replace(/\s+/g, ' ').trim());
  return s ? s[0].toLocaleUpperCase('ru') + s.slice(1) : s;
}

/** Speakers renumbered by first appearance; names "Спикер N". */
export function finalize(utts: Utterance[], duration: number, language?: string): Transcript {
  const order = new Map<number, number>();
  for (const u of utts) if (!order.has(u.speaker)) order.set(u.speaker, order.size);
  const utterances = utts.map((u) => ({ ...u, speaker: order.get(u.speaker)! }));
  // neighbouring utterances of one speaker (after cleaning) join
  const joined: Utterance[] = [];
  for (const u of utterances) {
    const last = joined[joined.length - 1];
    if (last && last.speaker === u.speaker && u.start - last.end < 4) {
      last.text = `${last.text} ${u.text}`;
      last.end = u.end;
    } else joined.push({ ...u });
  }
  return { speakers: Array.from({ length: order.size }, (_, i) => `Спикер ${i + 1}`), utterances: joined, duration, ...(language ? { language } : {}) };
}

// ------------------------------------------------------------------------------------- exports

const pad = (n: number, w = 2): string => String(Math.floor(n)).padStart(w, '0');
/** 1:02:03 or 02:03 */
export function clock(t: number): string {
  const s = Math.max(0, t);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  return h ? `${h}:${pad(m)}:${pad(s % 60)}` : `${pad(m)}:${pad(s % 60)}`;
}
const stamp = (t: number, sep: string): string => {
  const s = Math.max(0, t);
  return `${pad(s / 3600)}:${pad((s % 3600) / 60)}:${pad(s % 60)}${sep}${pad(Math.round((s % 1) * 1000) % 1000, 3)}`;
};

export function toText(tr: Transcript, times = true): string {
  return tr.utterances.map((u) => `${tr.speakers[u.speaker]}${times ? ` [${clock(u.start)}]` : ''}: ${u.text}`).join('\n\n') + '\n';
}

export function toMarkdown(tr: Transcript): string {
  return tr.utterances.map((u) => `**${tr.speakers[u.speaker]}** _${clock(u.start)}_  \n${u.text}`).join('\n\n') + '\n';
}

/** Subtitles: long utterances cut into cues of at most ~84 characters (two lines). */
function cues(tr: Transcript): { start: number; end: number; who: string; text: string }[] {
  const out: { start: number; end: number; who: string; text: string }[] = [];
  for (const u of tr.utterances) {
    const words = u.text.split(/\s+/);
    const parts: string[] = [];
    let cur = '';
    for (const w of words) {
      if (cur && (cur + ' ' + w).length > 84) {
        parts.push(cur);
        cur = w;
      } else cur = cur ? `${cur} ${w}` : w;
    }
    if (cur) parts.push(cur);
    const total = parts.reduce((s, p) => s + p.length, 0) || 1;
    let t = u.start;
    for (const p of parts) {
      const d = ((u.end - u.start) * p.length) / total;
      out.push({ start: t, end: t + d, who: tr.speakers[u.speaker], text: p });
      t += d;
    }
  }
  return out;
}

export function toSrt(tr: Transcript): string {
  return cues(tr)
    .map((c, i) => `${i + 1}\n${stamp(c.start, ',')} --> ${stamp(c.end, ',')}\n${c.who}: ${c.text}\n`)
    .join('\n');
}

export function toVtt(tr: Transcript): string {
  const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return 'WEBVTT\n\n' + cues(tr).map((c) => `${stamp(c.start, '.')} --> ${stamp(c.end, '.')}\n<v ${esc(c.who)}>${esc(c.text)}\n`).join('\n');
}

export function toJson(tr: Transcript): string {
  return JSON.stringify(
    {
      duration: tr.duration,
      language: tr.language,
      speakers: tr.speakers,
      utterances: tr.utterances.map((u) => ({ speaker: tr.speakers[u.speaker], start: Math.round(u.start * 100) / 100, end: Math.round(u.end * 100) / 100, text: u.text })),
    },
    null,
    2,
  );
}

/** The transcript as a document (for DOCX, ODT, PDF, RTF through the converters). */
export function toDoc(tr: Transcript, title = 'Расшифровка'): Doc {
  const blocks: Block[] = [{ type: 'heading', level: 1, runs: [{ text: title }] }];
  for (const u of tr.utterances)
    blocks.push({
      type: 'paragraph',
      runs: [
        { text: tr.speakers[u.speaker], bold: true },
        { text: `  ${clock(u.start)}`, color: '808080', size: 9 },
        { text: '\n' + u.text },
      ],
      spaceAfter: 8,
    });
  return { title, blocks };
}
