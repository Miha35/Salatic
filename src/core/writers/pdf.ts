/**
 * Doc → PDF layout engine on top of pdf-lib.
 *
 * Blocks are laid out into "boxes" (lists of drawing ops with coordinates relative
 * to the box top). Boxes are then placed onto pages; a box that does not fit is
 * sliced at a position where no atomic op (text line, image) is cut, so paragraphs
 * and table rows flow naturally across pages.
 */
import { PDFDocument, PDFName, PDFString, rgb, type PDFFont, type PDFImage, type PDFPage } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import type { Align, Block, Doc, TextRun } from '../model';
import { cleanText } from '../util';

export interface FontSet {
  regular: Uint8Array;
  bold: Uint8Array;
  italic: Uint8Array;
  boldItalic: Uint8Array;
  mono: Uint8Array;
}

export interface PdfWriteEnv {
  fonts: () => Promise<FontSet>;
  /** Convert images pdf-lib cannot embed (GIF, BMP, WebP) to PNG/JPEG. */
  transcodeImage?: (data: Uint8Array, mime: string) => Promise<{ data: Uint8Array; mime: string } | null>;
  onProgress?: (fraction: number) => void;
}

export interface PdfOptions {
  pageSize?: 'a4' | 'letter';
}

type FontKey = 'regular' | 'bold' | 'italic' | 'boldItalic' | 'mono';
type Color = [number, number, number];

type Op =
  | { k: 'text'; x: number; y: number; top: number; bottom: number; text: string; font: FontKey; size: number; color: Color }
  | { k: 'line'; x1: number; y1: number; x2: number; y2: number; w: number; color: Color; top: number; bottom: number }
  | { k: 'rect'; x: number; y: number; w: number; h: number; fill?: Color; stroke?: Color; clip: true; top: number; bottom: number }
  | { k: 'image'; x: number; y: number; w: number; h: number; img: PDFImage; top: number; bottom: number }
  | { k: 'link'; x: number; y: number; w: number; h: number; url: string; top: number; bottom: number };

interface Box {
  height: number;
  ops: Op[];
  spaceBefore: number;
  spaceAfter: number;
  keepWithNext?: boolean;
  pageBreak?: boolean;
}

interface Token {
  text: string;
  font: FontKey;
  size: number;
  width: number;
  space: boolean;
  newline: boolean;
  run: TextRun;
}

const BLACK: Color = [0.1, 0.1, 0.12];
const GRAY: Color = [0.45, 0.45, 0.5];
const LINK: Color = [0.02, 0.39, 0.76];
const BODY = 11;
const LINE = 1.38;
const HEADING_SIZES = [0, 22, 18, 15, 13, 12, 11];
const FALLBACK_CHARS: Record<string, string> = {
  '→': '->', '←': '<-', '↔': '<->', '⇒': '=>', '✓': 'v', '✔': 'v', '✕': 'x', '✗': 'x',
  '☐': '[ ]', '☑': '[x]', '☒': '[x]', '◦': '-', '▪': '-', '●': '•', '○': 'o', '−': '-',
  ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', '‑': '-', '‐': '-', '\u2028': ' ',
};

const hexColor = (hex: string): Color => [parseInt(hex.slice(0, 2), 16) / 255, parseInt(hex.slice(2, 4), 16) / 255, parseInt(hex.slice(4, 6), 16) / 255];

class PdfWriter {
  private pdf!: PDFDocument;
  private fonts = new Map<FontKey, PDFFont>();
  private charsets = new Map<FontKey, Set<number>>();
  private widthCache = new Map<FontKey, Map<number, number>>();
  private images = new Map<Uint8Array, PDFImage | null>();
  warnings = new Set<string>();
  private pageW: number;
  private pageH: number;
  private margin = 56.7;
  private pages: PDFPage[] = [];
  private page!: PDFPage;
  private cursor = 0;

  constructor(
    private env: PdfWriteEnv,
    opts: PdfOptions,
  ) {
    const letter = opts.pageSize === 'letter';
    this.pageW = letter ? 612 : 595.28;
    this.pageH = letter ? 792 : 841.89;
  }

  private get contentW(): number {
    return this.pageW - 2 * this.margin;
  }

  private get bottom(): number {
    return this.pageH - this.margin;
  }

  async init(doc: Doc): Promise<void> {
    this.pdf = await PDFDocument.create();
    this.pdf.registerFontkit(fontkit);
    const needed = new Set<FontKey>(['regular', 'bold']);
    const scan = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.type === 'paragraph' || b.type === 'heading') for (const r of b.runs) needed.add(this.fontKey(r, b.type === 'heading'));
        else if (b.type === 'list') b.items.forEach((i) => scan(i.blocks));
        else if (b.type === 'table') b.rows.forEach((r) => r.cells.forEach((c) => scan(c.blocks)));
        else if (b.type === 'quote') {
          needed.add('italic');
          scan(b.blocks);
        } else if (b.type === 'code') needed.add('mono');
      }
    };
    scan(doc.blocks);
    const data = await this.env.fonts();
    for (const key of needed) {
      const font = await this.pdf.embedFont(data[key], { subset: true });
      this.fonts.set(key, font);
      this.charsets.set(key, new Set(font.getCharacterSet()));
      this.widthCache.set(key, new Map());
    }
    if (doc.title) this.pdf.setTitle(doc.title);
    this.pdf.setCreator('Salatic');
    this.pdf.setProducer('Salatic');
    this.pdf.setLanguage('ru-RU');
  }

  private fontKey(r: Omit<TextRun, 'text'>, forceBold = false): FontKey {
    if (r.code) return 'mono';
    const bold = forceBold || !!r.bold;
    if (bold && r.italic) return 'boldItalic';
    if (bold) return 'bold';
    if (r.italic) return 'italic';
    return 'regular';
  }

  /** Replace characters missing from the font. */
  private sanitize(text: string, font: FontKey): string {
    const cs = this.charsets.get(font)!;
    let out = '';
    for (const ch of text) {
      const cp = ch.codePointAt(0)!;
      if (cs.has(cp) || cp === 32 || ch === '\n') out += ch;
      else if (ch === '\t') out += '    ';
      else if (ch === ' ') out += cs.has(0xa0) ? ch : ' ';
      else if (FALLBACK_CHARS[ch] && [...FALLBACK_CHARS[ch]].every((c) => cs.has(c.codePointAt(0)!) || c === ' ')) out += FALLBACK_CHARS[ch];
      else if (cp < 32) continue;
      else {
        out += '?';
        this.warnings.add('Некоторые символы отсутствуют во встроенном шрифте и заменены на «?».');
      }
    }
    return out;
  }

  private measure(text: string, font: FontKey, size: number): number {
    const cache = this.widthCache.get(font)!;
    const f = this.fonts.get(font)!;
    let w = 0;
    for (const ch of text) {
      const cp = ch.codePointAt(0)!;
      let cw = cache.get(cp);
      if (cw === undefined) {
        cw = f.widthOfTextAtSize(ch, 1000);
        cache.set(cp, cw);
      }
      w += cw;
    }
    return (w * size) / 1000;
  }

  private tokens(runs: TextRun[], baseSize: number, forceBold: boolean): Token[] {
    const out: Token[] = [];
    for (const r of runs) {
      const font = this.fontKey(r, forceBold);
      const size = r.sup || r.sub ? baseSize * 0.68 : r.code ? baseSize * 0.92 : baseSize;
      const text = this.sanitize(cleanText(r.text).replace(/\r/g, ''), font);
      for (const part of text.split(/(\n| +)/)) {
        if (!part) continue;
        if (part === '\n') out.push({ text: '', font, size, width: 0, space: false, newline: true, run: r });
        else if (part[0] === ' ') out.push({ text: part, font, size, width: this.measure(part, font, size), space: true, newline: false, run: r });
        else out.push({ text: part, font, size, width: this.measure(part, font, size), space: false, newline: false, run: r });
      }
    }
    return out;
  }

  /** Break tokens into lines no wider than `width`. */
  private breakLines(tokens: Token[], width: number, keepIndent = false): Token[][] {
    const lines: Token[][] = [];
    let line: Token[] = [];
    let w = 0;
    let afterNewline = true;
    const push = (): void => {
      while (line.length && line[line.length - 1].space) line.pop();
      lines.push(line);
      line = [];
      w = 0;
    };
    for (const t of tokens) {
      if (t.newline) {
        push();
        afterNewline = true;
        continue;
      }
      if (t.space) {
        if (line.length || (keepIndent && afterNewline)) {
          line.push(t);
          w += t.width;
        }
        continue;
      }
      afterNewline = false;
      if (w + t.width <= width + 0.01 || !line.length || line.every((x) => x.space)) {
        if (t.width > width && !line.length) {
          // Hard-break a word that is wider than the line
          let chunk = '';
          for (const ch of t.text) {
            const cw = this.measure(chunk + ch, t.font, t.size);
            if (cw > width && chunk) {
              line.push({ ...t, text: chunk, width: this.measure(chunk, t.font, t.size) });
              push();
              chunk = ch;
            } else chunk += ch;
          }
          if (chunk) {
            const cw = this.measure(chunk, t.font, t.size);
            line.push({ ...t, text: chunk, width: cw });
            w = cw;
          }
          continue;
        }
        line.push(t);
        w += t.width;
      } else {
        push();
        line.push(t);
        w = t.width;
      }
    }
    if (line.length || !lines.length) push();
    return lines;
  }

  /** Lay out a run list as a paragraph box. */
  private textBox(runs: TextRun[], x: number, width: number, opts: { size?: number; bold?: boolean; align?: Align; color?: Color; spaceBefore?: number; spaceAfter?: number; italic?: boolean; keepIndent?: boolean }): Box {
    const size = opts.size ?? BODY;
    const effRuns = opts.italic ? runs.map((r) => ({ ...r, italic: true })) : runs;
    const lines = this.breakLines(this.tokens(effRuns, size, !!opts.bold), width, opts.keepIndent);
    const ops: Op[] = [];
    let y = 0;
    lines.forEach((line, li) => {
      const maxSize = line.reduce((m, t) => Math.max(m, t.run.sup || t.run.sub ? t.size / 0.68 : t.size), size);
      const lh = maxSize * LINE;
      const baseline = y + maxSize * 1.02;
      const lineW = line.reduce((s, t) => s + t.width, 0);
      let cx = x;
      let extraPerSpace = 0;
      if (opts.align === 'center') cx = x + (width - lineW) / 2;
      else if (opts.align === 'right') cx = x + width - lineW;
      else if (opts.align === 'justify' && li < lines.length - 1) {
        const spaces = line.filter((t) => t.space).length;
        if (spaces) extraPerSpace = (width - lineW) / spaces;
      }
      // Merge adjacent tokens with identical style into a single text op
      let pending: { x: number; text: string; t: Token; w: number } | null = null;
      const flush = (): void => {
        if (!pending) return;
        const t = pending.t;
        const r = t.run;
        let by = baseline;
        if (r.sup) by -= maxSize * 0.33;
        if (r.sub) by += maxSize * 0.14;
        const color = r.link ? LINK : r.color ? hexColor(r.color) : (opts.color ?? BLACK);
        if (pending.text.trim()) ops.push({ k: 'text', x: pending.x, y: by, top: y, bottom: y + lh, text: pending.text, font: t.font, size: t.size, color });
        if (r.underline || r.link) ops.push({ k: 'line', x1: pending.x, y1: by + t.size * 0.12, x2: pending.x + pending.w, y2: by + t.size * 0.12, w: Math.max(0.5, t.size * 0.05), color, top: y, bottom: y + lh });
        if (r.strike) ops.push({ k: 'line', x1: pending.x, y1: by - t.size * 0.3, x2: pending.x + pending.w, y2: by - t.size * 0.3, w: Math.max(0.5, t.size * 0.05), color, top: y, bottom: y + lh });
        if (r.link) ops.push({ k: 'link', x: pending.x, y: y, w: pending.w, h: lh, url: r.link, top: y, bottom: y + lh });
        pending = null;
      };
      for (const t of line) {
        const w = t.width + (t.space ? extraPerSpace : 0);
        if (pending && pending.t.run === t.run && pending.t.font === t.font && !extraPerSpace) {
          pending.text += t.text;
          pending.w += w;
        } else {
          flush();
          pending = { x: cx, text: t.text, t, w };
        }
        cx += w;
      }
      flush();
      y += lh;
    });
    return { height: y, ops, spaceBefore: opts.spaceBefore ?? 0, spaceAfter: opts.spaceAfter ?? size * 0.55 };
  }

  private async image(b: Extract<Block, { type: 'image' }>): Promise<PDFImage | null> {
    if (this.images.has(b.data)) return this.images.get(b.data)!;
    let data = b.data;
    let mime = b.mime;
    if (mime !== 'image/png' && mime !== 'image/jpeg') {
      const t = this.env.transcodeImage ? await this.env.transcodeImage(data, mime).catch(() => null) : null;
      if (!t) {
        this.warnings.add('Некоторые изображения пропущены: формат не поддерживается.');
        this.images.set(b.data, null);
        return null;
      }
      data = t.data;
      mime = t.mime;
    }
    let img: PDFImage | null = null;
    try {
      img = mime === 'image/png' ? await this.pdf.embedPng(data) : await this.pdf.embedJpg(data);
    } catch {
      this.warnings.add('Некоторые изображения повреждены и пропущены.');
    }
    this.images.set(b.data, img);
    return img;
  }

  private shift(box: Box, dx: number, dy: number): Op[] {
    return box.ops.map((o) => {
      switch (o.k) {
        case 'line':
          return { ...o, x1: o.x1 + dx, x2: o.x2 + dx, y1: o.y1 + dy, y2: o.y2 + dy, top: o.top + dy, bottom: o.bottom + dy };
        default:
          return { ...o, x: o.x + dx, y: o.y + dy, top: o.top + dy, bottom: o.bottom + dy } as Op;
      }
    });
  }

  private markerFor(ordered: boolean, index: number, depth: number): string {
    if (ordered) return `${index}.`;
    const cs = this.charsets.get('regular')!;
    const pick = ['•', '◦', '▪'][depth % 3];
    return cs.has(pick.codePointAt(0)!) ? pick : depth % 2 ? '–' : '•';
  }

  async layout(blocks: Block[], x: number, width: number, depth = 0): Promise<Box[]> {
    const boxes: Box[] = [];
    for (const b of blocks) {
      switch (b.type) {
        case 'paragraph':
          if (!b.runs.some((r) => r.text.trim())) boxes.push({ height: BODY * LINE * 0.6, ops: [], spaceBefore: 0, spaceAfter: 0 });
          else boxes.push(this.textBox(b.runs, x, width, { align: b.align }));
          break;
        case 'heading': {
          const size = HEADING_SIZES[b.level];
          const box = this.textBox(b.runs, x, width, { size, bold: true, align: b.align, spaceBefore: size * 0.85, spaceAfter: size * 0.45 });
          box.keepWithNext = true;
          boxes.push(box);
          break;
        }
        case 'list': {
          let n = b.start ?? 1;
          const markerW = b.ordered ? this.measure(`${n + b.items.length - 1}.`, 'regular', BODY) + 7 : 14;
          for (const it of b.items) {
            const inner = await this.layout(it.blocks.length ? it.blocks : [{ type: 'paragraph', runs: [{ text: ' ' }] }], x + markerW + 4, width - markerW - 4, depth + 1);
            inner.forEach((ib, k) => {
              if (k < inner.length - 1) ib.spaceAfter = Math.min(ib.spaceAfter, 3);
            });
            const first = inner[0];
            if (first) {
              const firstText = first.ops.find((o): o is Extract<Op, { k: 'text' }> => o.k === 'text');
              const baseline = firstText ? firstText.y : BODY * 1.02;
              const marker = this.markerFor(b.ordered, n, depth);
              const mw = this.measure(marker, 'regular', BODY);
              // The marker goes first in the content stream so text extractors read it before the item text
              first.ops.unshift({ k: 'text', x: b.ordered ? x + markerW - mw - 3 : x + 3, y: baseline, top: 0, bottom: firstText ? firstText.bottom : BODY * LINE, text: marker, font: 'regular', size: BODY, color: BLACK });
              const last = inner[inner.length - 1];
              last.spaceAfter = 3;
            }
            boxes.push(...inner);
            n++;
          }
          if (boxes.length) boxes[boxes.length - 1].spaceAfter = BODY * 0.55;
          break;
        }
        case 'table':
          boxes.push(...(await this.table(b, x, width, depth)));
          break;
        case 'image': {
          const img = await this.image(b);
          if (!img) break;
          let w = (b.width * 72) / 96;
          let h = (b.height * 72) / 96;
          const maxH = this.bottom - this.margin - 20;
          const k = Math.min(1, width / w, maxH / h);
          w *= k;
          h *= k;
          const ix = w < width ? x + (width - w) / 2 : x;
          boxes.push({ height: h, ops: [{ k: 'image', x: ix, y: 0, w, h, img, top: 0, bottom: h }], spaceBefore: 4, spaceAfter: 8 });
          break;
        }
        case 'code': {
          const size = 9.5;
          const pad = 6;
          const inner = this.textBox([{ text: b.text, code: true }], x + pad, width - 2 * pad, { size: size / 0.92, spaceAfter: 0, keepIndent: true });
          const h = inner.height + 2 * pad;
          boxes.push({
            height: h,
            ops: [{ k: 'rect', x, y: 0, w: width, h, fill: [0.955, 0.958, 0.965], clip: true, top: 0, bottom: h }, ...this.shift(inner, 0, pad)],
            spaceBefore: 2,
            spaceAfter: 8,
          });
          break;
        }
        case 'quote': {
          const inner = await this.layout(b.blocks, x + 14, width - 14, depth);
          for (const ib of inner) {
            ib.ops = ib.ops.map((o) => (o.k === 'text' && o.color === BLACK ? { ...o, color: [0.27, 0.27, 0.3] as Color } : o));
            ib.ops.unshift({ k: 'rect', x, y: -ib.spaceBefore, w: 3, h: ib.height + ib.spaceBefore + ib.spaceAfter, fill: [0.8, 0.8, 0.84], clip: true, top: 0, bottom: ib.height });
          }
          boxes.push(...inner);
          break;
        }
        case 'hr':
          boxes.push({ height: 1, ops: [{ k: 'line', x1: x, y1: 0.5, x2: x + width, y2: 0.5, w: 0.8, color: [0.75, 0.75, 0.78], top: 0, bottom: 1 }], spaceBefore: 8, spaceAfter: 10 });
          break;
        case 'pagebreak':
          boxes.push({ height: 0, ops: [], spaceBefore: 0, spaceAfter: 0, pageBreak: true });
          break;
      }
    }
    return boxes;
  }

  private async table(b: Extract<Block, { type: 'table' }>, x: number, width: number, depth: number): Promise<Box[]> {
    const cols = Math.max(1, ...b.rows.map((r) => r.cells.reduce((s, c) => s + (c.colspan ?? 1), 0)));
    const pad = 4;
    // Natural (max) and minimum widths per column from single-span cells
    const maxW = new Array(cols).fill(20);
    const minW = new Array(cols).fill(20);
    for (const r of b.rows) {
      let c = 0;
      for (const cell of r.cells) {
        const span = cell.colspan ?? 1;
        if (span === 1 && c < cols) {
          for (const blk of cell.blocks) {
            if (blk.type !== 'paragraph' && blk.type !== 'heading') continue;
            for (const run of blk.runs) {
              const font = this.fontKey(run, !!cell.header);
              const text = this.sanitize(run.text, font);
              for (const line of text.split('\n')) maxW[c] = Math.max(maxW[c], this.measure(line, font, BODY) + 2 * pad);
              for (const word of text.split(/\s+/)) minW[c] = Math.max(minW[c], Math.min(width / cols, this.measure(word, font, BODY) + 2 * pad));
            }
          }
        }
        c += span;
      }
    }
    const sumMax = maxW.reduce((a, v) => a + v, 0);
    let widths: number[];
    if (sumMax <= width) widths = maxW.map((w) => (w / sumMax) * width);
    else {
      const sumMin = minW.reduce((a, v) => a + v, 0);
      const extra = Math.max(0, width - sumMin);
      const flex = maxW.map((w, i) => Math.max(0, w - minW[i]));
      const sumFlex = flex.reduce((a, v) => a + v, 0) || 1;
      widths = minW.map((w, i) => w + (extra * flex[i]) / sumFlex);
      const total = widths.reduce((a, v) => a + v, 0);
      widths = widths.map((w) => (w / total) * width);
    }
    const boxes: Box[] = [];
    const border: Color = [0.62, 0.62, 0.66];
    for (const r of b.rows) {
      let c = 0;
      const cellBoxes: { x: number; w: number; boxes: Box[]; header: boolean }[] = [];
      for (const cell of r.cells) {
        if (c >= cols) break;
        const span = Math.min(cell.colspan ?? 1, cols - c);
        const cx = x + widths.slice(0, c).reduce((a, v) => a + v, 0);
        const cw = widths.slice(c, c + span).reduce((a, v) => a + v, 0);
        const blocks = cell.header
          ? cell.blocks.map((bl) => (bl.type === 'paragraph' ? { ...bl, runs: bl.runs.map((rr) => ({ ...rr, bold: true })) } : bl))
          : cell.blocks;
        cellBoxes.push({ x: cx, w: cw, boxes: await this.layout(blocks, cx + pad, cw - 2 * pad, depth + 1), header: !!cell.header });
        c += span;
      }
      if (c < cols) {
        const cx = x + widths.slice(0, c).reduce((a, v) => a + v, 0);
        cellBoxes.push({ x: cx, w: width - (cx - x), boxes: [], header: false });
      }
      // Stack each cell's boxes vertically
      let rowH = 0;
      const ops: Op[] = [];
      const cellOps: Op[][] = cellBoxes.map((cb) => {
        let y = pad;
        const out: Op[] = [];
        cb.boxes.forEach((bx, i) => {
          if (i > 0) y += bx.spaceBefore;
          out.push(...this.shift(bx, 0, y));
          y += bx.height + (i < cb.boxes.length - 1 ? Math.min(bx.spaceAfter, 4) : 0);
        });
        rowH = Math.max(rowH, y + pad);
        return out;
      });
      rowH = Math.max(rowH, BODY * LINE + 2 * pad);
      cellBoxes.forEach((cb, i) => {
        if (cb.header) ops.push({ k: 'rect', x: cb.x, y: 0, w: cb.w, h: rowH, fill: [0.95, 0.95, 0.96], clip: true, top: 0, bottom: rowH });
        ops.push({ k: 'rect', x: cb.x, y: 0, w: cb.w, h: rowH, stroke: border, clip: true, top: 0, bottom: rowH });
        ops.push(...cellOps[i]);
      });
      boxes.push({ height: rowH, ops, spaceBefore: 0, spaceAfter: 0 });
    }
    if (boxes.length) {
      boxes[0].spaceBefore = 4;
      boxes[boxes.length - 1].spaceAfter = 10;
    }
    return boxes;
  }

  private newPage(): void {
    this.page = this.pdf.addPage([this.pageW, this.pageH]);
    this.pages.push(this.page);
    this.cursor = this.margin;
  }

  private draw(ops: Op[], offset: number, from: number, to: number): void {
    const H = this.pageH;
    for (const o of ops) {
      if (o.k === 'rect') {
        const top = Math.max(o.y, from);
        const bot = Math.min(o.y + o.h, to);
        if (bot <= top) continue;
        const y = H - (offset + bot - from);
        const h = bot - top;
        this.page.drawRectangle({
          x: o.x,
          y,
          width: o.w,
          height: h,
          color: o.fill ? rgb(...o.fill) : undefined,
          borderColor: o.stroke ? rgb(...o.stroke) : undefined,
          borderWidth: o.stroke ? 0.6 : 0,
        });
        continue;
      }
      if (o.top < from - 0.01 || o.bottom > to + 0.01) continue;
      switch (o.k) {
        case 'text':
          this.page.drawText(o.text, { x: o.x, y: H - (offset + o.y - from), size: o.size, font: this.fonts.get(o.font)!, color: rgb(...o.color) });
          break;
        case 'line':
          this.page.drawLine({ start: { x: o.x1, y: H - (offset + o.y1 - from) }, end: { x: o.x2, y: H - (offset + o.y2 - from) }, thickness: o.w, color: rgb(...o.color) });
          break;
        case 'image':
          this.page.drawImage(o.img, { x: o.x, y: H - (offset + o.y + o.h - from), width: o.w, height: o.h });
          break;
        case 'link': {
          const y1 = H - (offset + o.y + o.h - from);
          const annot = this.pdf.context.obj({
            Type: 'Annot',
            Subtype: 'Link',
            Rect: [o.x, y1, o.x + o.w, y1 + o.h],
            Border: [0, 0, 0],
            A: { Type: 'Action', S: 'URI', URI: PDFString.of(o.url) },
          });
          this.page.node.addAnnot(this.pdf.context.register(annot));
          break;
        }
      }
    }
  }

  /** Largest y ≤ limit where no atomic op is cut. */
  private cutAt(box: Box, from: number, limit: number): number {
    const candidates = new Set<number>();
    for (const o of box.ops) if (o.k !== 'rect' && o.bottom > from && o.bottom <= limit) candidates.add(o.bottom);
    const sorted = [...candidates].sort((a, b) => b - a);
    for (const y of sorted) {
      if (!box.ops.some((o) => o.k !== 'rect' && o.top < y - 0.01 && o.bottom > y + 0.01)) return y;
    }
    return from;
  }

  place(boxes: Box[]): void {
    this.newPage();
    const contentH = this.bottom - this.margin;
    for (let i = 0; i < boxes.length; i++) {
      const box = boxes[i];
      if (box.pageBreak) {
        if (this.cursor > this.margin) this.newPage();
        continue;
      }
      const atTop = this.cursor <= this.margin + 0.01;
      if (!atTop) this.cursor += box.spaceBefore;
      let need = box.height;
      if (box.keepWithNext && boxes[i + 1] && !boxes[i + 1].pageBreak) {
        const next = boxes[i + 1];
        const firstLine = next.ops.filter((o) => o.k !== 'rect').reduce((m, o) => Math.min(m, o.bottom), next.height);
        need += next.spaceBefore + Math.min(firstLine, contentH / 3);
      }
      if (this.cursor + need > this.bottom && !atTop && need <= contentH) this.newPage();
      // Place, slicing across pages when needed
      let from = 0;
      while (from < box.height - 0.01) {
        const avail = this.bottom - this.cursor;
        if (box.height - from <= avail + 0.01) {
          this.draw(box.ops, this.cursor, from, box.height);
          this.cursor += box.height - from;
          break;
        }
        let cut = this.cutAt(box, from, from + avail);
        if (cut <= from) {
          if (this.cursor > this.margin + 0.01) {
            this.newPage();
            continue;
          }
          cut = from + avail; // a single atomic op taller than a page: clip it
        }
        this.draw(box.ops, this.cursor, from, cut);
        from = cut;
        this.newPage();
      }
      this.cursor += box.spaceAfter;
    }
  }

  async finish(): Promise<Uint8Array> {
    if (this.pages.length > 1) {
      const font = this.fonts.get('regular')!;
      this.pages.forEach((p, i) => {
        const label = String(i + 1);
        const w = font.widthOfTextAtSize(label, 9);
        p.drawText(label, { x: (this.pageW - w) / 2, y: this.margin / 2 - 3, size: 9, font, color: rgb(...GRAY) });
      });
    }
    // Hint viewers to open with page-width zoom
    this.pdf.catalog.set(PDFName.of('PageLayout'), PDFName.of('OneColumn'));
    return this.pdf.save({ useObjectStreams: true });
  }
}

export async function docToPdf(doc: Doc, env: PdfWriteEnv, opts: PdfOptions = {}): Promise<{ data: Uint8Array; warnings: string[] }> {
  const w = new PdfWriter(env, opts);
  await w.init(doc);
  const margin = 56.7;
  const pageW = opts.pageSize === 'letter' ? 612 : 595.28;
  const boxes = await w.layout(doc.blocks.length ? doc.blocks : [{ type: 'paragraph', runs: [] }], margin, pageW - 2 * margin);
  env.onProgress?.(0.6);
  w.place(boxes);
  const data = await w.finish();
  env.onProgress?.(1);
  return { data, warnings: [...w.warnings] };
}
