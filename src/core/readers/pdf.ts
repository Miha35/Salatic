/**
 * PDF → Doc with layout reconstruction.
 *
 * pdf.js gives positioned text fragments; this module rebuilds lines, paragraphs,
 * headings (by font size), bullet/numbered lists, simple tables (aligned columns),
 * bold/italic (from font names) and embedded images, and drops repeating
 * headers/footers and page numbers.
 */
import type { Align, Block, Doc, TableRow, TextRun } from '../model';
import { normalizeRuns } from '../model';
import { encodePng } from '../png';
import { ConvertError, cleanText, tick } from '../util';

type PdfJs = typeof import('pdfjs-dist');
type Matrix = [number, number, number, number, number, number];

export interface PdfReadEnv {
  pdfjs: PdfJs;
  /** Extra getDocument() parameters (cMapUrl, wasmUrl, …). */
  docParams?: Record<string, unknown>;
  /** Encode raw pixels; defaults to the built-in PNG encoder. */
  encodeImage?: (pixels: Uint8Array | Uint8ClampedArray, width: number, height: number, channels: 1 | 3 | 4) => Promise<{ data: Uint8Array; mime: string }>;
  onProgress?: (fraction: number) => void;
  password?: string;
}

interface Frag {
  str: string;
  x: number;
  y: number; // baseline, device space (y grows downwards)
  w: number;
  fs: number;
  bold: boolean;
  italic: boolean;
  mono: boolean;
}

interface Seg {
  x0: number;
  x1: number;
  runs: TextRun[];
}

interface Line {
  y: number;
  x0: number;
  x1: number;
  fs: number;
  segs: Seg[];
}

interface Para {
  kind: 'para' | 'heading' | 'bullet' | 'number' | 'code';
  lines: Line[];
  runs: TextRun[];
  fs: number;
  top: number;
  x0: number;
  align?: Align;
  start?: number;
  level?: number;
}

interface PageImage {
  top: number;
  block: Block;
}

type PageItem = { type: 'para'; p: Para } | { type: 'table'; lines: Line[]; top: number; cols: number[] } | { type: 'image'; img: PageImage };

const BULLET_RE = /^\s*([•●○◦▪▫■□‣⁃∙·➢►▸▹✓✔❖]|[-–—*](?=\s))\s*/;
const NUMBER_RE = /^\s*(\d{1,3})[.)]\s+(?=\S)/;
const PAGE_NUMBER_RE = /^(?:(?:стр\.?|страница|page|p\.)\s*)?[-–—]?\s*\d{1,4}\s*[-–—]?(?:\s*(?:из|of|\/)\s*\d{1,4})?$/i;

const l0fs = (p: { lines: { fs: number }[] }): number => p.lines[0]?.fs ?? 10;

const median = (a: number[]): number => {
  if (!a.length) return 0;
  const s = [...a].sort((x, y) => x - y);
  return s[s.length >> 1];
};

function styleFromFontName(name: string | undefined, flags: { bold?: boolean; black?: boolean; italic?: boolean; isMonospace?: boolean }): { bold: boolean; italic: boolean; mono: boolean } {
  const n = (name ?? '').replace(/^[A-Z]{6}\+/, '');
  return {
    bold: !!flags.bold || !!flags.black || /bold|black|heavy|semibold|demibold|demi\b|extrabold|ultrabold|[-,]bd\b|w[6-9]\b/i.test(n),
    italic: !!flags.italic || /italic|oblique|[-,]it\b|kursiv/i.test(n),
    mono: !!flags.isMonospace || /courier|mono|consolas|menlo|code/i.test(n),
  };
}

function mul(m1: Matrix, m2: Matrix): Matrix {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

function bboxOf(m: Matrix): { x0: number; y0: number; x1: number; y1: number } {
  const pts = [
    [m[4], m[5]],
    [m[0] + m[4], m[1] + m[5]],
    [m[2] + m[4], m[3] + m[5]],
    [m[0] + m[2] + m[4], m[1] + m[3] + m[5]],
  ];
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
}

interface RawImage {
  width: number;
  height: number;
  kind?: number;
  data?: Uint8Array | Uint8ClampedArray;
  bitmap?: { width: number; height: number; close?: () => void };
}

function toPixels(img: RawImage): { pixels: Uint8Array | Uint8ClampedArray; channels: 1 | 3 | 4 } | null {
  const { width, height, kind, data } = img;
  if (!data || !width || !height) return null;
  if (kind === 2 && data.length >= width * height * 3) return { pixels: data, channels: 3 };
  if (kind === 3 && data.length >= width * height * 4) {
    // Drop alpha when fully opaque – smaller output and JPEG-friendly.
    let opaque = true;
    for (let i = 3; i < data.length; i += 4) {
      if (data[i] !== 255) {
        opaque = false;
        break;
      }
    }
    if (!opaque) return { pixels: data, channels: 4 };
    const rgb = new Uint8Array(width * height * 3);
    for (let i = 0, j = 0; j < rgb.length; i += 4, j += 3) {
      rgb[j] = data[i];
      rgb[j + 1] = data[i + 1];
      rgb[j + 2] = data[i + 2];
    }
    return { pixels: rgb, channels: 3 };
  }
  if (kind === 1) {
    const rowBytes = (width + 7) >> 3;
    const gray = new Uint8Array(width * height);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) gray[y * width + x] = data[y * rowBytes + (x >> 3)] & (128 >> (x & 7)) ? 255 : 0;
    }
    return { pixels: gray, channels: 1 };
  }
  return null;
}

class PdfLayout {
  bodyFs = 10;
  /** Code-block detection is disabled when the whole document is set in a monospace font. */
  detectCode = true;
  lineGap = new Map<number, number>();
  warnings: string[] = [];

  constructor(private env: PdfReadEnv) {}

  /**
   * Group text fragments into lines: fragments are banded by baseline in content-stream
   * order, then sorted by x inside a band (so markers drawn after their text still lead
   * the line), de-duplicated (fake-bold overprinting) and split into segments at wide gaps.
   */
  buildLines(frags: Frag[]): Line[] {
    const bands: { y: number; fs: number; frags: Frag[] }[] = [];
    let band: { y: number; fs: number; frags: Frag[] } | null = null;
    for (const f of frags) {
      if (band && Math.abs(f.y - band.y) <= Math.max(f.fs, band.fs) * 0.55) {
        band.frags.push(f);
        if (f.str.trim() && f.fs > band.fs * 1.05) {
          band.fs = f.fs;
          band.y = f.y;
        }
      } else {
        band = { y: f.y, fs: f.str.trim() ? f.fs : 0, frags: [f] };
        bands.push(band);
      }
    }
    const lines: Line[] = [];
    for (const b of bands) {
      const sorted = b.frags
        .map((f, i) => ({ f, i }))
        .sort((a, c) => a.f.x - c.f.x || a.i - c.i)
        .map((x) => x.f);
      const kept: Frag[] = [];
      for (const f of sorted) {
        if (f.str.trim() && kept.some((k) => k.str === f.str && Math.abs(k.x - f.x) < 1.5 && Math.abs(k.y - f.y) < 1)) continue;
        kept.push(f);
      }
      // Dominant size: the size carrying most characters
      const weight = new Map<number, number>();
      for (const f of kept) if (f.str.trim()) weight.set(f.fs, (weight.get(f.fs) ?? 0) + f.str.length);
      let fs = b.fs || 10;
      let best = -1;
      for (const [size, w] of weight) if (w > best || (w === best && size > fs)) [fs, best] = [size, w];
      const baseline = kept.find((f) => f.fs === fs)?.y ?? b.y;
      const line: Line = { y: baseline, x0: Infinity, x1: -Infinity, fs, segs: [] };
      let seg: Seg | null = null;
      let lastX1 = -Infinity;
      let forceSplit = false;
      for (const f of kept) {
        if (!f.str.trim()) {
          if (f.w > fs * 1.6) forceSplit = true;
          else if (seg && seg.runs.length && !/\s$/.test(seg.runs[seg.runs.length - 1].text)) seg.runs.push(this.run(' ', f));
          lastX1 = Math.max(lastX1, f.x + f.w);
          continue;
        }
        const gap = f.x - lastX1;
        if (!seg || forceSplit || gap > fs * 1.6) {
          seg = { x0: f.x, x1: f.x + f.w, runs: [] };
          line.segs.push(seg);
        } else if (gap > fs * 0.18 && seg.runs.length && !/\s$/.test(seg.runs[seg.runs.length - 1].text) && !/^\s/.test(f.str)) {
          seg.runs.push(this.run(' ', f));
        }
        forceSplit = false;
        seg.runs.push(this.run(f.str, f, f.fs < fs * 0.8 && f.y < baseline - fs * 0.15));
        seg.x1 = Math.max(seg.x1, f.x + f.w);
        line.x0 = Math.min(line.x0, f.x);
        line.x1 = Math.max(line.x1, f.x + f.w);
        lastX1 = Math.max(lastX1, f.x + f.w);
      }
      for (const sg of line.segs) sg.runs = normalizeRuns(sg.runs);
      line.segs = line.segs.filter((sg) => sg.runs.some((r) => r.text.trim()));
      if (line.segs.length) lines.push(line);
    }
    return lines;
  }

  private run(text: string, f: Frag, sup = false): TextRun {
    const r: TextRun = { text };
    if (f.bold) r.bold = true;
    if (f.italic) r.italic = true;
    if (f.mono) r.code = true;
    if (sup && /^\d{1,3}$|^[*†‡]$/.test(text.trim())) r.sup = true;
    return r;
  }

  static lineText(l: Line): string {
    return l.segs.map((s) => s.runs.map((r) => r.text).join('')).join('\t');
  }

  /** Detect runs of ≥2 consecutive lines whose segments align into ≥2 columns. */
  findTables(lines: Line[]): { start: number; end: number; cols: number[] }[] {
    const tables: { start: number; end: number; cols: number[] }[] = [];
    let i = 0;
    while (i < lines.length) {
      const l = lines[i];
      if (l.segs.length < 2) {
        i++;
        continue;
      }
      const cols = l.segs.map((s) => s.x0);
      const tol = l.fs * 1.2;
      let j = i + 1;
      let multi = 1;
      while (j < lines.length) {
        const m = lines[j];
        if (m.y - lines[j - 1].y > m.fs * 3.2 || m.y <= lines[j - 1].y) break;
        const aligned = m.segs.every((s) => cols.some((c) => Math.abs(c - s.x0) < tol) || (s.x0 > cols[0] && s.x1 < (cols[1] ?? Infinity) + tol));
        if (!aligned) break;
        if (m.segs.length >= 2) multi++;
        for (const s of m.segs) if (!cols.some((c) => Math.abs(c - s.x0) < tol)) cols.push(s.x0);
        j++;
      }
      if (multi >= 2 && j - i >= 2) {
        cols.sort((a, b) => a - b);
        tables.push({ start: i, end: j, cols });
        i = j;
      } else i++;
    }
    return tables;
  }

  tableRows(lines: Line[], cols: number[]): TableRow[] {
    const rows: TextRun[][][] = [];
    const colOf = (x: number): number => {
      let best = 0;
      for (let k = 0; k < cols.length; k++) if (x >= cols[k] - lines[0].fs * 1.2) best = k;
      return best;
    };
    for (const l of lines) {
      const cells: TextRun[][] = cols.map(() => []);
      for (const s of l.segs) {
        const c = colOf(s.x0);
        if (cells[c].length) cells[c].push({ text: ' ' });
        cells[c].push(...s.runs);
      }
      // Wrapped cell text: no content in the first column → continuation of the previous row
      const prev = rows[rows.length - 1];
      if (prev && !cells[0].length && l.segs.length < cols.length) {
        cells.forEach((runs, k) => {
          if (runs.length) prev[k].push(...(prev[k].length ? [{ text: ' ' }] : []), ...runs);
        });
      } else rows.push(cells);
    }
    return rows.map((r) => ({
      cells: r.map((runs) => ({ blocks: [{ type: 'paragraph', runs: normalizeRuns(runs) } as Block] })),
    }));
  }

  /** Split lines into paragraphs using spacing, indentation, font size and list markers. */
  buildParas(lines: Line[], contentLeft: number, contentRight: number): Para[] {
    const paras: Para[] = [];
    let cur: Para | null = null;
    const width = contentRight - contentLeft;
    for (let i = 0; i < lines.length; i++) {
      const l = lines[i];
      const text = PdfLayout.lineText(l);
      const isCode = this.detectCode && l.segs.every((sg) => sg.runs.every((r) => r.code || !r.text.trim()));
      const prev = cur?.lines[cur.lines.length - 1];
      if (isCode) {
        const normal = this.lineGap.get(Math.round(l.fs * 2) / 2) ?? l.fs * 1.25;
        if (cur && cur.kind === 'code' && prev && l.y - prev.y > 0 && l.y - prev.y < normal * 1.6) cur.lines.push(l);
        else {
          cur = { kind: 'code', lines: [l], runs: [], fs: l.fs, top: l.y - l.fs, x0: l.x0 };
          paras.push(cur);
        }
        continue;
      }
      const bullet = BULLET_RE.exec(text);
      const num = NUMBER_RE.exec(text);
      let newPara = !cur || !prev || cur.kind === 'code';
      if (cur && prev) {
        const gap = l.y - prev.y;
        const fsKey = Math.round(Math.max(l.fs, prev.fs) * 2) / 2;
        const normal = this.lineGap.get(fsKey) ?? Math.max(l.fs, prev.fs) * 1.2;
        const prevText = PdfLayout.lineText(prev);
        const prevShort = prev.x1 < contentRight - Math.max(width * 0.12, l.fs * 4);
        const indented = l.x0 > prev.x0 + l.fs * 0.9 && cur.lines.length > 0 && cur.kind === 'para';
        newPara =
          gap < 0 ||
          gap > normal * 1.35 ||
          Math.abs(l.fs - prev.fs) > Math.max(0.9, prev.fs * 0.12) ||
          !!bullet ||
          !!num ||
          (prevShort && /[.!?:;…»"”)]$/.test(prevText.trim()) && cur.kind !== 'heading') ||
          (indented && /[.!?:;…»"”]$/.test(prevText.trim())) ||
          (cur.kind !== 'para' && l.x0 < cur.x0 - l.fs * 0.5) ||
          // list item continuation lines are indented past the marker
          ((cur.kind === 'bullet' || cur.kind === 'number') && l.x0 < cur.lines[0].x0 + l.fs * 0.3 && prevShort);
      }
      if (newPara) {
        cur = { kind: bullet ? 'bullet' : num ? 'number' : 'para', lines: [l], runs: [], fs: l.fs, top: l.y - l.fs, x0: l.x0 };
        if (num) cur.start = +num[1];
        paras.push(cur);
      } else cur!.lines.push(l);
    }
    for (const p of paras) this.finishPara(p, contentLeft, contentRight);
    return paras;
  }

  private finishPara(p: Para, contentLeft: number, contentRight: number): void {
    if (p.kind === 'code') {
      const minX = Math.min(...p.lines.map((l) => l.x0));
      p.runs = [
        {
          text: p.lines
            .map((l) => ' '.repeat(Math.max(0, Math.round((l.x0 - minX) / (l.fs * 0.6)))) + l.segs.map((sg) => sg.runs.map((r) => r.text).join('')).join('    '))
            .join('\n'),
        },
      ];
      p.fs = l0fs(p);
      return;
    }
    const runs: TextRun[] = [];
    p.lines.forEach((l, idx) => {
      const lineRuns = l.segs.flatMap((s, k) => (k ? [{ text: '\t' } as TextRun, ...s.runs] : s.runs));
      if (idx > 0) {
        const last = runs[runs.length - 1];
        const first = lineRuns[0];
        if (last && first && /[\p{L}]-$/u.test(last.text) && /^\p{Ll}/u.test(first.text)) {
          last.text = last.text.slice(0, -1); // de-hyphenate
        } else if (last && !/\s$/.test(last.text)) runs.push({ text: ' ' });
      }
      runs.push(...lineRuns.map((r) => ({ ...r })));
    });
    let merged = normalizeRuns(runs);
    // Strip list markers
    if (p.kind === 'bullet' || p.kind === 'number') {
      const re = p.kind === 'bullet' ? BULLET_RE : NUMBER_RE;
      while (merged.length) {
        const m = re.exec(merged[0].text);
        if (m) {
          merged[0] = { ...merged[0], text: merged[0].text.slice(m[0].length) };
          if (!merged[0].text) merged.shift();
          break;
        }
        if (!merged[0].text.trim()) merged.shift();
        else break;
      }
      merged = normalizeRuns(merged.map((r, i) => (i === 0 ? { ...r, text: r.text.replace(/^[\s\t]+/, '') } : r)));
    }
    // Trim whitespace and convert tabs inside normal paragraphs to spaces when short
    if (merged.length) {
      merged[0].text = merged[0].text.replace(/^\s+/, '');
      merged[merged.length - 1].text = merged[merged.length - 1].text.replace(/\s+$/, '');
    }
    p.runs = normalizeRuns(merged);
    p.fs = median(p.lines.map((l) => l.fs));
    // Alignment
    const width = contentRight - contentLeft;
    if (width > 150 && p.kind === 'para') {
      const center = (contentLeft + contentRight) / 2;
      const centered = p.lines.every((l) => Math.abs((l.x0 + l.x1) / 2 - center) < Math.max(p.fs * 1.5, width * 0.03) && l.x0 - contentLeft > p.fs * 2);
      const right = p.lines.every((l) => Math.abs(l.x1 - contentRight) < p.fs * 1.2 && l.x0 - contentLeft > width * 0.3);
      if (centered) p.align = 'center';
      else if (right) p.align = 'right';
    }
  }

  /** Learn document-wide body font size and normal line spacing per size. */
  calibrate(pages: Line[][]): void {
    const sizeChars = new Map<number, number>();
    const gaps = new Map<number, number[]>();
    for (const lines of pages) {
      for (let i = 0; i < lines.length; i++) {
        const l = lines[i];
        const key = Math.round(l.fs * 2) / 2;
        sizeChars.set(key, (sizeChars.get(key) ?? 0) + PdfLayout.lineText(l).length);
        const next = lines[i + 1];
        if (next && Math.abs(next.fs - l.fs) < 0.5) {
          const g = next.y - l.y;
          if (g > l.fs * 0.8 && g < l.fs * 2.4) (gaps.get(key) ?? gaps.set(key, []).get(key)!).push(g);
        }
      }
    }
    let mono = 0;
    let total = 0;
    for (const lines of pages) {
      for (const l of lines) {
        for (const sg of l.segs) {
          for (const r of sg.runs) {
            total += r.text.length;
            if (r.code) mono += r.text.length;
          }
        }
      }
    }
    this.detectCode = mono < total * 0.5;
    let best = 0;
    for (const [size, chars] of sizeChars) {
      if (chars > best) {
        best = chars;
        this.bodyFs = size;
      }
    }
    for (const [k, g] of gaps) if (g.length >= 2) this.lineGap.set(k, median(g));
  }

  async pageImages(page: PdfPageLike, viewport: Matrix, pageArea: number, hasText: boolean, seen: Map<string, number>): Promise<PageImage[]> {
    const { OPS } = this.env.pdfjs;
    const ops = await page.getOperatorList();
    const out: PageImage[] = [];
    let ctm: Matrix = [1, 0, 0, 1, 0, 0];
    const stack: Matrix[] = [];
    for (let i = 0; i < ops.fnArray.length; i++) {
      const fn = ops.fnArray[i];
      const args = ops.argsArray[i] as unknown[];
      if (fn === OPS.save) stack.push(ctm);
      else if (fn === OPS.restore) ctm = stack.pop() ?? ctm;
      else if (fn === OPS.transform) ctm = mul(ctm, args as Matrix);
      else if (fn === OPS.paintFormXObjectBegin) {
        stack.push(ctm);
        const m = args?.[0] as Matrix | null;
        if (m && m.length === 6) ctm = mul(ctm, m);
      } else if (fn === OPS.paintFormXObjectEnd) ctm = stack.pop() ?? ctm;
      else if (fn === OPS.paintImageXObject || fn === OPS.paintInlineImageXObject) {
        const box = bboxOf(mul(viewport, ctm));
        const w = box.x1 - box.x0;
        const h = box.y1 - box.y0;
        if (w < 18 || h < 18) continue;
        if (hasText && w * h > pageArea * 0.8) continue; // page background or scanned layer behind OCR text
        let raw: RawImage | null = null;
        let key = '';
        if (fn === OPS.paintInlineImageXObject) raw = args[0] as RawImage;
        else {
          key = String(args[0]);
          if ((seen.get(key) ?? 0) > 0) continue; // repeated logo/watermark
          raw = (await resolveObj(key.startsWith('g_') ? page.commonObjs : page.objs, key)) as RawImage | null;
          seen.set(key, (seen.get(key) ?? 0) + 1);
        }
        const block = raw ? await this.imageBlock(raw, w, h) : null;
        if (block) out.push({ top: box.y0, block });
      }
    }
    return out;
  }

  private async imageBlock(raw: RawImage, w: number, h: number): Promise<Block | null> {
    const px = toPixels(raw);
    if (!px) return null;
    if (raw.width * raw.height > 40_000_000) return null;
    const enc = this.env.encodeImage
      ? await this.env.encodeImage(px.pixels, raw.width, raw.height, px.channels)
      : { data: encodePng(px.pixels, raw.width, raw.height, px.channels), mime: 'image/png' };
    // PDF points → CSS px
    return { type: 'image', data: enc.data, mime: enc.mime, width: Math.round((w * 96) / 72), height: Math.round((h * 96) / 72) };
  }
}

interface PdfObjs {
  get(id: string, callback?: (data: unknown) => void): unknown;
}

interface PdfPageLike {
  getOperatorList(): Promise<{ fnArray: number[]; argsArray: unknown[] }>;
  objs: PdfObjs;
  commonObjs: PdfObjs;
}

/** Image objects are resolved asynchronously by pdf.js – wait for them (with a timeout). */
function resolveObj(objs: PdfObjs, id: string, timeoutMs = 10_000): Promise<unknown> {
  return new Promise((resolve) => {
    const timer = setTimeout(() => resolve(null), timeoutMs);
    try {
      objs.get(id, (data) => {
        clearTimeout(timer);
        resolve(data ?? null);
      });
    } catch {
      clearTimeout(timer);
      resolve(null);
    }
  });
}

function headerFooterFilter(pages: { lines: Line[]; height: number }[]): (line: Line, pageIdx: number) => boolean {
  const zone = (l: Line, h: number): boolean => l.y < h * 0.08 || l.y > h * 0.92;
  const norm = (t: string): string => t.replace(/\d+/g, '#').replace(/\s+/g, ' ').trim().toLowerCase();
  const counts = new Map<string, number>();
  for (const p of pages) {
    const seenHere = new Set<string>();
    for (const l of p.lines) if (zone(l, p.height)) seenHere.add(norm(PdfLayout.lineText(l)));
    for (const k of seenHere) counts.set(k, (counts.get(k) ?? 0) + 1);
  }
  const repeated = new Set([...counts].filter(([, c]) => pages.length >= 3 && c >= pages.length * 0.5).map(([k]) => k));
  return (l, i) => {
    if (!zone(l, pages[i].height)) return true;
    const t = PdfLayout.lineText(l).trim();
    if (PAGE_NUMBER_RE.test(t)) return false;
    return !repeated.has(norm(t));
  };
}

export async function pdfToDoc(bytes: Uint8Array, env: PdfReadEnv): Promise<{ doc: Doc; warnings: string[] }> {
  const { pdfjs } = env;
  let pdf;
  let task;
  try {
    task = pdfjs.getDocument({
      data: bytes.slice(),
      isOffscreenCanvasSupported: false,
      isImageDecoderSupported: false,
      disableFontFace: true,
      fontExtraProperties: true,
      useSystemFonts: false,
      isEvalSupported: false,
      stopAtErrors: false,
      password: env.password,
      ...env.docParams,
    } as Parameters<PdfJs['getDocument']>[0]);
    pdf = await task.promise;
  } catch (e) {
    void task?.destroy();
    const name = (e as Error)?.name;
    if (name === 'PasswordException') throw new ConvertError('PDF защищён паролем. Снимите защиту и попробуйте снова.', 'password');
    if (name === 'InvalidPDFException') throw new ConvertError('Файл PDF повреждён или имеет неверный формат.');
    throw e;
  }
  const layout = new PdfLayout(env);
  const pages: { lines: Line[]; height: number; width: number; images: PageImage[] }[] = [];
  const seenImages = new Map<string, number>();
  let title: string | undefined;
  try {
    const meta = await pdf.getMetadata().catch(() => null);
    const t = (meta?.info as { Title?: string } | undefined)?.Title;
    if (t && t.trim() && !/^(untitled|microsoft word|документ)/i.test(t.trim())) title = t.trim();

    for (let n = 1; n <= pdf.numPages; n++) {
      const page = await pdf.getPage(n);
      const vp = page.getViewport({ scale: 1 });
      const vpm = vp.transform as Matrix;
      const tc = await page.getTextContent();
      // Operator list loads fonts (for bold/italic detection) and images.
      const images = await layout.pageImages(page as unknown as PdfPageLike, vpm, vp.width * vp.height, tc.items.some((it) => 'str' in it && it.str.trim()), seenImages);
      const fontCache = new Map<string, { bold: boolean; italic: boolean; mono: boolean }>();
      const frags: Frag[] = [];
      for (const it of tc.items) {
        if (!('str' in it) || !it.str) continue;
        const m = mul(vpm, it.transform as Matrix);
        if (Math.abs(m[1]) > Math.abs(m[0]) * 0.2 || Math.abs(m[2]) > Math.abs(m[3]) * 0.2) continue; // rotated text (watermarks, side notes)
        const fs = Math.hypot(m[2], m[3]);
        if (fs < 1) continue;
        let style = fontCache.get(it.fontName);
        if (!style) {
          let font: { name?: string; bold?: boolean; black?: boolean; italic?: boolean; isMonospace?: boolean } = {};
          try {
            font = page.commonObjs.get(it.fontName) as typeof font;
          } catch {
            /* font not loaded – fall back to plain style */
          }
          style = styleFromFontName(font?.name, font ?? {});
          fontCache.set(it.fontName, style);
        }
        const str = cleanText(it.str.normalize('NFC'));
        if (!str) continue;
        frags.push({ str, x: m[4], y: m[5], w: it.width * vp.scale, fs, ...style });
      }
      pages.push({ lines: layout.buildLines(frags), height: vp.height, width: vp.width, images });
      page.cleanup();
      env.onProgress?.(n / pdf.numPages);
      if (n % 4 === 0) await tick();
    }
  } finally {
    void task.destroy();
  }

  layout.calibrate(pages.map((p) => p.lines));
  const keep = headerFooterFilter(pages);

  // Build page items
  const pageItems: PageItem[][] = pages.map((p, pi) => {
    const lines = p.lines.filter((l) => keep(l, pi));
    if (!lines.length) return p.images.map((img) => ({ type: 'image' as const, img }));
    const left = Math.min(...lines.map((l) => l.x0));
    const right = Math.max(...lines.map((l) => l.x1));
    const items: PageItem[] = [];
    const tables = layout.findTables(lines);
    let cursor = 0;
    for (const t of [...tables, { start: lines.length, end: lines.length, cols: [] }]) {
      if (t.start > cursor) for (const para of layout.buildParas(lines.slice(cursor, t.start), left, right)) items.push({ type: 'para', p: para });
      if (t.end > t.start) {
        const tl = lines.slice(t.start, t.end);
        items.push({ type: 'table', lines: tl, top: tl[0].y - tl[0].fs, cols: t.cols });
      }
      cursor = t.end;
    }
    // Insert images by vertical position
    for (const img of p.images) {
      const idx = items.findIndex((it) => (it.type === 'para' ? it.p.top : it.type === 'table' ? it.top : img.top + 1) > img.top);
      if (idx < 0) items.push({ type: 'image', img });
      else items.splice(idx, 0, { type: 'image', img });
    }
    return items;
  });

  // Heading levels by distinct sizes larger than body text
  const headingSizes = new Set<number>();
  for (const items of pageItems) {
    for (const it of items) {
      if (it.type !== 'para') continue;
      const p = it.p;
      const len = p.runs.reduce((s, r) => s + r.text.length, 0);
      if (p.kind === 'para' && p.fs >= layout.bodyFs * 1.18 && len > 0 && len < 250 && p.lines.length <= 4) {
        p.kind = 'heading';
        headingSizes.add(Math.round(p.fs));
      }
    }
  }
  const sizes = [...headingSizes].sort((a, b) => b - a);
  const levelOf = (fs: number): 1 | 2 | 3 | 4 => (Math.min(sizes.indexOf(Math.round(fs)), 3) + 1) as 1 | 2 | 3 | 4;

  // Merge paragraphs split by page breaks
  const flat: PageItem[] = [];
  for (const items of pageItems) {
    const first = items[0];
    const last = flat[flat.length - 1];
    if (
      first?.type === 'para' &&
      last?.type === 'para' &&
      first.p.kind === 'para' &&
      last.p.kind === 'para' &&
      Math.abs(first.p.fs - last.p.fs) < 0.5 &&
      !/[.!?:…»"”]$/.test(last.p.runs.map((r) => r.text).join('').trim()) &&
      /^\p{Ll}/u.test(first.p.runs[0]?.text ?? '')
    ) {
      last.p.runs = normalizeRuns([...last.p.runs, { text: ' ' }, ...first.p.runs]);
      flat.push(...items.slice(1));
    } else flat.push(...items);
  }

  // Emit blocks
  const blocks: Block[] = [];
  type ListBlock = Extract<Block, { type: 'list' }>;
  let stack: { list: ListBlock; x0: number }[] = [];
  for (const it of flat) {
    if (it.type === 'image') {
      stack = [];
      blocks.push(it.img.block);
      continue;
    }
    if (it.type === 'table') {
      stack = [];
      const rows = layout.tableRows(it.lines, it.cols);
      if (rows.length) blocks.push({ type: 'table', rows });
      continue;
    }
    const p = it.p;
    if (!p.runs.length) continue;
    if (p.kind === 'bullet' || p.kind === 'number') {
      const ordered = p.kind === 'number';
      const item = { blocks: [{ type: 'paragraph', runs: p.runs } as Block] };
      while (stack.length && stack[stack.length - 1].x0 > p.x0 + p.fs * 0.6) stack.pop();
      const top = stack[stack.length - 1];
      const fresh = (): ListBlock => {
        const l: ListBlock = { type: 'list', ordered, items: [] };
        if (ordered && p.start && p.start !== 1) l.start = p.start;
        return l;
      };
      if (!top) {
        const l = fresh();
        blocks.push(l);
        stack.push({ list: l, x0: p.x0 });
      } else if (p.x0 > top.x0 + p.fs * 1.2) {
        const l = fresh();
        top.list.items[top.list.items.length - 1].blocks.push(l);
        stack.push({ list: l, x0: p.x0 });
      } else if (top.list.ordered !== ordered) {
        const l = fresh();
        stack.pop();
        const parent = stack[stack.length - 1];
        if (parent) parent.list.items[parent.list.items.length - 1].blocks.push(l);
        else blocks.push(l);
        stack.push({ list: l, x0: p.x0 });
      }
      stack[stack.length - 1].list.items.push(item);
      continue;
    }
    stack = [];
    if (p.kind === 'code') {
      blocks.push({ type: 'code', text: p.runs[0].text });
    } else if (p.kind === 'heading') {
      blocks.push({ type: 'heading', level: levelOf(p.fs), runs: p.runs.map(({ bold: _b, ...r }) => r), ...(p.align ? { align: p.align } : {}) });
    } else {
      blocks.push({ type: 'paragraph', runs: p.runs, ...(p.align ? { align: p.align } : {}) });
    }
  }

  const warnings = [...layout.warnings];
  const textPages = pages.filter((p) => p.lines.length).length;
  if (textPages < pages.length) {
    warnings.push(
      textPages === 0
        ? 'В PDF нет текстового слоя (скан). Страницы перенесены как изображения; распознавание текста (OCR) не выполняется.'
        : `На ${pages.length - textPages} стр. нет текстового слоя — они перенесены как изображения.`,
    );
  }
  if (!blocks.length) throw new ConvertError('В PDF не найдено ни текста, ни изображений для извлечения.');
  return { doc: { title, blocks }, warnings };
}
