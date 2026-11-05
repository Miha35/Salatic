/** CommonMark/GFM-subset Markdown → Doc. */
import type { Align, Block, Doc, ListItem, TableCell, TextRun } from '../model';
import { normalizeRuns, trimRuns } from '../model';
import { imageSize, parseDataUri, sniffImageMime } from '../util';
import { htmlToDoc } from './html';

type Style = Omit<TextRun, 'text'>;

const RE_FENCE = /^ {0,3}(`{3,}|~{3,})\s*([^`]*)$/;
const RE_ATX = /^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$/;
const RE_HR = /^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$/;
const RE_LIST = /^( {0,3})([-*+]|\d{1,9}[.)])(?:([ \t]+)(.*)|$)/;
const RE_QUOTE = /^ {0,3}> ?(.*)$/;
const RE_TABLE_DELIM = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;
const RE_HTML_BLOCK = /^ {0,3}<(div|p|table|img|br|h[1-6]|ul|ol|pre|blockquote|hr|figure|section|center|details)\b/i;

const isBlank = (l: string): boolean => /^\s*$/.test(l);

function expandTabs(line: string): string {
  if (line.indexOf('\t') < 0) return line;
  let out = '';
  for (const ch of line) {
    if (ch === '\t') out += ' '.repeat(4 - (out.length % 4));
    else out += ch;
  }
  return out;
}

function startsBlock(line: string): boolean {
  return RE_FENCE.test(line) || RE_ATX.test(line) || RE_HR.test(line) || RE_QUOTE.test(line) || RE_LIST.test(line) || RE_HTML_BLOCK.test(line);
}

function splitRow(line: string): string[] {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1);
  const cells: string[] = [];
  let cur = '';
  let inCode = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '\\' && s[i + 1] === '|') {
      cur += '|';
      i++;
    } else if (ch === '`') {
      inCode = !inCode;
      cur += ch;
    } else if (ch === '|' && !inCode) {
      cells.push(cur.trim());
      cur = '';
    } else cur += ch;
  }
  cells.push(cur.trim());
  return cells;
}

function parseBlocks(lines: string[]): Block[] {
  const out: Block[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (isBlank(line)) {
      i++;
      continue;
    }
    // Fenced code
    const fence = RE_FENCE.exec(line);
    if (fence) {
      const marker = fence[1];
      const indent = line.indexOf(marker[0]);
      const body: string[] = [];
      i++;
      while (i < lines.length && !new RegExp(`^ {0,3}${marker[0] === '`' ? '`' : '~'}{${marker.length},}\\s*$`).test(lines[i])) {
        body.push(lines[i].replace(new RegExp(`^ {0,${indent}}`), ''));
        i++;
      }
      i++; // closing fence
      out.push({ type: 'code', text: body.join('\n') });
      continue;
    }
    const atx = RE_ATX.exec(line);
    if (atx) {
      out.push({ type: 'heading', level: atx[1].length as 1, runs: parseInline(atx[2] ?? '') });
      i++;
      continue;
    }
    if (RE_HR.test(line)) {
      out.push({ type: 'hr' });
      i++;
      continue;
    }
    if (RE_QUOTE.test(line)) {
      const body: string[] = [];
      while (i < lines.length && !isBlank(lines[i])) {
        const q = RE_QUOTE.exec(lines[i]);
        if (q) body.push(q[1]);
        else if (startsBlock(lines[i])) break;
        else body.push(lines[i]); // lazy continuation
        i++;
      }
      out.push({ type: 'quote', blocks: parseBlocks(body) });
      continue;
    }
    if (RE_HTML_BLOCK.test(line)) {
      const body: string[] = [];
      while (i < lines.length && !isBlank(lines[i])) body.push(lines[i++]);
      out.push(...htmlToDoc(body.join('\n')).blocks);
      continue;
    }
    const li = RE_LIST.exec(line);
    if (li) {
      const ordered = /\d/.test(li[2]);
      const bulletChar = ordered ? li[2].slice(-1) : li[2];
      const items: ListItem[] = [];
      while (i < lines.length) {
        const m = RE_LIST.exec(lines[i]);
        if (!m || /\d/.test(m[2]) !== ordered || (ordered ? m[2].slice(-1) : m[2]) !== bulletChar) break;
        const contentIndent = m[1].length + m[2].length + Math.min((m[3] ?? ' ').length, 4);
        const body: string[] = [m[4] ?? ''];
        i++;
        while (i < lines.length) {
          const l = lines[i];
          if (isBlank(l)) {
            // blank line: continue the item only if the next non-blank line is indented
            let k = i + 1;
            while (k < lines.length && isBlank(lines[k])) k++;
            if (k < lines.length && /^ */.exec(lines[k])![0].length >= contentIndent) {
              for (; i < k; i++) body.push('');
              continue;
            }
            break;
          }
          const ind = /^ */.exec(l)![0].length;
          if (ind >= contentIndent) body.push(l.slice(contentIndent));
          else if (RE_LIST.test(l) || startsBlock(l)) break;
          else body.push(l.trimStart()); // lazy continuation of a paragraph
          i++;
        }
        items.push({ blocks: parseBlocks(body) });
        // allow a single blank line between items of the same list
        if (i < lines.length && isBlank(lines[i])) {
          let k = i + 1;
          while (k < lines.length && isBlank(lines[k])) k++;
          const n = k < lines.length ? RE_LIST.exec(lines[k]) : null;
          if (n && /\d/.test(n[2]) === ordered && (ordered ? n[2].slice(-1) : n[2]) === bulletChar) i = k;
        }
      }
      for (const it of items) if (!it.blocks.length) it.blocks.push({ type: 'paragraph', runs: [] });
      out.push({ type: 'list', ordered, items });
      continue;
    }
    // GFM table
    if (line.includes('|') && i + 1 < lines.length && RE_TABLE_DELIM.test(lines[i + 1]) && lines[i + 1].includes('-')) {
      const header = splitRow(line);
      const aligns: (Align | undefined)[] = splitRow(lines[i + 1]).map((d) =>
        d.startsWith(':') && d.endsWith(':') ? 'center' : d.endsWith(':') ? 'right' : undefined,
      );
      const rows = [header];
      i += 2;
      while (i < lines.length && !isBlank(lines[i]) && lines[i].includes('|')) rows.push(splitRow(lines[i++]));
      const width = header.length;
      out.push({
        type: 'table',
        rows: rows.map((r, ri) => ({
          cells: Array.from({ length: width }, (_, ci): TableCell => {
            const p: Block = { type: 'paragraph', runs: parseInline(r[ci] ?? '', ri === 0 ? { bold: true } : {}) };
            if (aligns[ci]) (p as { align?: Align }).align = aligns[ci];
            return ri === 0 ? { blocks: [p], header: true } : { blocks: [p] };
          }),
        })),
      });
      continue;
    }
    // Indented code block
    if (/^( {4}|\t)/.test(line)) {
      const body: string[] = [];
      while (i < lines.length && (/^( {4}|\t)/.test(lines[i]) || isBlank(lines[i]))) {
        body.push(lines[i].replace(/^( {4}|\t)/, ''));
        i++;
      }
      while (body.length && isBlank(body[body.length - 1])) body.pop();
      out.push({ type: 'code', text: body.join('\n') });
      continue;
    }
    // Paragraph (with possible setext heading underline)
    const para: string[] = [line];
    i++;
    let setext = 0;
    while (i < lines.length && !isBlank(lines[i])) {
      const l = lines[i];
      if (/^ {0,3}=+\s*$/.test(l)) {
        setext = 1;
        i++;
        break;
      }
      if (/^ {0,3}-+\s*$/.test(l)) {
        setext = 2;
        i++;
        break;
      }
      if (startsBlock(l)) break;
      para.push(l);
      i++;
    }
    const text = para
      .map((l, k) => {
        const last = k === para.length - 1;
        const t = l.replace(/^\s+/, '');
        if (last) return t.replace(/\s+$/, '');
        if (/ {2,}$/.test(t) || /\\$/.test(t)) return t.replace(/\\$/, '').replace(/\s+$/, '') + '\u2028';
        return t.replace(/\s+$/, '') + ' ';
      })
      .join('');
    if (setext) {
      out.push({ type: 'heading', level: setext as 1, runs: parseInline(text.replace(/\u2028/g, ' ')) });
      continue;
    }
    out.push(...paragraphWithImages(text));
  }
  return out;
}

/** Split a paragraph around block-level images. */
function paragraphWithImages(text: string): Block[] {
  const images: Block[] = [];
  const runs = parseInline(text, {}, images);
  if (!images.length) return [{ type: 'paragraph', runs }];
  const blocks: Block[] = [];
  // Images are emitted after the paragraph text in their original order (good enough for documents).
  const r = trimRuns(runs);
  if (r.length) blocks.push({ type: 'paragraph', runs: r });
  blocks.push(...images);
  return blocks;
}

const PUNCT = /[!"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~]/;

/** Emphasis/link text rarely spans more than this; bounding keeps pathological input cheap. */
const EMPHASIS_WINDOW = 4000;

/**
 * Per-string index of valid closing delimiters and matching brackets, computed once so
 * that inline parsing stays O(n log n) even on hostile input like "****…" or "[[[[…".
 */
class InlineIndex {
  private closers = new Map<string, number[]>();
  private brackets: Int32Array | null = null;
  constructor(private s: string) {}

  private validCloser(idx: number, delim: string): boolean {
    const s = this.s;
    const before = s[idx - 1];
    const after = s[idx + delim.length];
    if (before === undefined || before === '\\' || /\s/.test(before)) return false;
    if (delim.length === 1 && (after === delim || before === delim)) return false;
    if (delim[0] === '_' && after && /[\p{L}\p{N}]/u.test(after)) return false;
    return true;
  }

  /** First valid closing delimiter strictly after `from` (within the window), or -1. */
  closing(from: number, delim: string): number {
    let list = this.closers.get(delim);
    if (!list) {
      list = [];
      for (let k = this.s.indexOf(delim); k >= 0; k = this.s.indexOf(delim, k + 1)) if (this.validCloser(k, delim)) list.push(k);
      this.closers.set(delim, list);
    }
    let lo = 0;
    let hi = list.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (list[mid] <= from) lo = mid + 1;
      else hi = mid;
    }
    const idx = lo < list.length ? list[lo] : -1;
    return idx >= 0 && idx - from <= EMPHASIS_WINDOW ? idx : -1;
  }

  /** Index of the "]" matching the "[" at `open`, or -1. */
  bracket(open: number): number {
    if (!this.brackets) {
      const s = this.s;
      const match = new Int32Array(s.length).fill(-1);
      const stack: number[] = [];
      for (let k = 0; k < s.length; k++) {
        const ch = s[k];
        if (ch === '\\') k++;
        else if (ch === '[') stack.push(k);
        else if (ch === ']' && stack.length) match[stack.pop()!] = k;
      }
      this.brackets = match;
    }
    const m = this.brackets[open];
    return m >= 0 && m - open <= EMPHASIS_WINDOW ? m : -1;
  }
}

function parseLinkTarget(s: string, start: number): { url: string; end: number } | null {
  if (s[start] !== '(') return null;
  let depth = 0;
  const limit = Math.min(s.length, start + EMPHASIS_WINDOW);
  for (let k = start; k < limit; k++) {
    if (s[k] === '(') depth++;
    else if (s[k] === ')') {
      depth--;
      if (depth === 0) {
        let inner = s.slice(start + 1, k).trim();
        if (inner.startsWith('<')) inner = inner.slice(1, inner.indexOf('>') < 0 ? undefined : inner.indexOf('>'));
        else inner = inner.replace(/\s+("[^"]*"|'[^']*'|\([^)]*\))\s*$/, '');
        return { url: inner.trim(), end: k + 1 };
      }
    }
  }
  return null;
}

/** Footnote definitions of the document being parsed ([^id]: text). */
let footnoteDefs: Map<string, string[]> | null = null;

/** Parse inline Markdown into runs. Block images are pushed into `images` when given. */
export function parseInline(src: string, style: Style = {}, images?: Block[]): TextRun[] {
  const index = new InlineIndex(src);
  const runs: TextRun[] = [];
  let buf = '';
  const emit = (): void => {
    if (buf) runs.push({ text: buf, ...style });
    buf = '';
  };
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '\u2028') {
      buf += '\n';
      i++;
      continue;
    }
    if (ch === '\\' && i + 1 < src.length && PUNCT.test(src[i + 1])) {
      buf += src[i + 1];
      i += 2;
      continue;
    }
    if (ch === '`') {
      let tl = 0;
      while (src[i + tl] === '`') tl++;
      const tick = '`'.repeat(tl);
      const end = src.indexOf(tick, i + tl);
      if (end > 0) {
        emit();
        let code = src.slice(i + tick.length, end).replace(/\u2028/g, ' ');
        if (/^ .* $/.test(code) && code.trim()) code = code.slice(1, -1);
        runs.push({ text: code, ...style, code: true });
        i = end + tick.length;
        continue;
      }
      buf += tick;
      i += tl;
      continue;
    }
    if (ch === '[' && src[i + 1] === '^' && footnoteDefs) {
      const m = /^\[\^([^\]\s]+)\]/.exec(src.slice(i, i + 200));
      const def = m ? footnoteDefs.get(m[1].toLowerCase()) : undefined;
      if (m && def) {
        emit();
        const defs = footnoteDefs;
        footnoteDefs = null; // no nested notes
        const note = parseBlocks(def);
        footnoteDefs = defs;
        const { sup: _s, ...rest } = style;
        runs.push({ text: '', ...rest, note });
        i += m[0].length;
        continue;
      }
    }
    if ((ch === '!' && src[i + 1] === '[') || ch === '[') {
      const open = ch === '!' ? i + 1 : i;
      const close = index.bracket(open);
      const target = close > 0 ? parseLinkTarget(src, close + 1) : null;
      if (target) {
        const label = src.slice(open + 1, close);
        emit();
        if (ch === '!') {
          const parsed = parseDataUri(target.url);
          const mime = parsed ? sniffImageMime(parsed.data) ?? parsed.mime : '';
          if (parsed && mime.startsWith('image/') && images) {
            const size = imageSize(parsed.data) ?? { width: 300, height: 200 };
            images.push({ type: 'image', data: parsed.data, mime, width: size.width, height: size.height, alt: label || undefined });
          } else if (/^https?:/i.test(target.url)) {
            runs.push({ text: label || target.url, ...style, link: target.url });
          } else if (label) runs.push({ text: label, ...style });
        } else {
          const link = /^(https?:|mailto:|tel:)/i.test(target.url) ? target.url : undefined;
          runs.push(...parseInline(label, link ? { ...style, link } : style, images));
        }
        i = target.end;
        continue;
      }
    }
    if (ch === '<') {
      const m = /^<((?:https?:\/\/|mailto:)[^>\s]+)>/i.exec(src.slice(i));
      if (m) {
        emit();
        runs.push({ text: m[1].replace(/^mailto:/i, ''), ...style, link: m[1] });
        i += m[0].length;
        continue;
      }
    }
    if (ch === 'h' && (src.startsWith('http://', i) || src.startsWith('https://', i)) && (i === 0 || /[\s(]/.test(src[i - 1]))) {
      const m = /^https?:\/\/[^\s<]*[^\s<.,:;"')\]!?]/.exec(src.slice(i));
      if (m) {
        emit();
        runs.push({ text: m[0], ...style, link: style.link ?? m[0] });
        i += m[0].length;
        continue;
      }
    }
    if (ch === '*' || ch === '_' || ch === '~') {
      const triple = src.startsWith(ch.repeat(3), i) && ch !== '~';
      const double = src.startsWith(ch + ch, i);
      const leftFlanking = i + 1 < src.length && !/\s/.test(src[i + (triple ? 3 : double ? 2 : 1)] ?? ' ');
      const intraword = ch === '_' && i > 0 && /[\p{L}\p{N}]/u.test(src[i - 1]);
      if (leftFlanking && !intraword) {
        if (triple) {
          const end = index.closing(i + 3, ch.repeat(3));
          if (end > 0) {
            emit();
            runs.push(...parseInline(src.slice(i + 3, end), { ...style, bold: true, italic: true }, images));
            i = end + 3;
            continue;
          }
        }
        if (double) {
          const end = index.closing(i + 2, ch + ch);
          if (end > 0) {
            emit();
            const add = ch === '~' ? { strike: true } : { bold: true };
            runs.push(...parseInline(src.slice(i + 2, end), { ...style, ...add }, images));
            i = end + 2;
            continue;
          }
        } else if (ch !== '~') {
          const end = index.closing(i + 1, ch);
          if (end > 0) {
            emit();
            runs.push(...parseInline(src.slice(i + 1, end), { ...style, italic: true }, images));
            i = end + 1;
            continue;
          }
        }
      }
    }
    buf += ch;
    i++;
  }
  emit();
  return normalizeRuns(runs);
}

export function markdownToDoc(md: string, title?: string): Doc {
  const lines = md
    .replace(/^﻿/, '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map(expandTabs);
  // Optional YAML front matter
  if (lines[0] === '---') {
    const end = lines.indexOf('---', 1);
    if (end > 0) {
      const fm = lines.slice(1, end).join('\n');
      const t = /^title:\s*["']?(.+?)["']?\s*$/m.exec(fm);
      if (t) title ??= t[1];
      lines.splice(0, end + 1);
    }
  }
  // Footnote definitions: "[^id]: text" plus indented continuation lines
  const defs = new Map<string, string[]>();
  for (let i = 0; i < lines.length; i++) {
    const m = /^ {0,3}\[\^([^\]\s]+)\]:\s?(.*)$/.exec(lines[i]);
    if (!m) continue;
    const body = [m[2]];
    let j = i + 1;
    while (j < lines.length && (/^( {4}|\t)/.test(lines[j]) || (lines[j].trim() === '' && j + 1 < lines.length && /^( {4}|\t)/.test(lines[j + 1])))) body.push(lines[j++].replace(/^( {4}|\t)/, ''));
    defs.set(m[1].toLowerCase(), body);
    lines.splice(i, j - i);
    i--;
  }
  const saved = footnoteDefs;
  footnoteDefs = defs.size ? defs : null;
  let blocks: Block[];
  try {
    blocks = parseBlocks(lines);
  } finally {
    footnoteDefs = saved;
  }
  const firstHeading = blocks.find((b) => b.type === 'heading');
  if (!title && firstHeading && firstHeading.type === 'heading') title = firstHeading.runs.map((r) => r.text).join('');
  return { title, blocks };
}
