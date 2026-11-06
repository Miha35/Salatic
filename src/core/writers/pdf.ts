/**
 * Doc → PDF layout engine on top of pdf-lib.
 *
 * Blocks are laid out into "boxes" (drawing ops with coordinates relative to the box
 * top). Boxes are placed onto pages; a box that does not fit is sliced where no atomic
 * op (text line, image) is cut, so paragraphs and table rows flow across pages.
 * Footnotes are reserved at the bottom of the page that references them; headers and
 * footers (with page-number fields) are drawn on every page once the total is known.
 *
 * Typography: metric-compatible fonts per family (Liberation Sans/Serif/Mono, Carlito),
 * per-character fallback (DejaVu Sans → Noto Sans CJK → Noto Emoji), Word-style line
 * spacing from the fonts' Windows metrics, subsetting with HarfBuzz.
 */
import { clip, concatTransformationMatrix, degrees, drawSvgPath, endPath, PDFDict, PDFDocument, PDFName, PDFNumber, PDFOperator, PDFOperatorNames, PDFString, popGraphicsState, pushGraphicsState, rectangle, rgb, setGraphicsState, type PDFContext, type PDFImage, type PDFObject, type PDFPage, type PDFRef } from 'pdf-lib';
import { parseMetafile, type Metafile } from '../metafile';
import { parseSvg } from '../svg-render';
import type { Align, Block, Doc, FontFamily, Gradient, ParaFormat, TextRun } from '../model';
import { FontManager, styleOf, type FontKey, type FontLoader } from '../pdf-fonts';
import type { Subsetter } from '../hb-subset';
import { cleanText, maxOf } from '../util';

export type { FontKey, FontLoader };

export interface PdfWriteEnv {
  fonts: FontLoader;
  /** HarfBuzz subsetter; without it fonts are embedded whole (correct but larger). */
  subsetter?: Subsetter;
  /** Convert images pdf-lib cannot embed (GIF, BMP, WebP, TIFF…) to PNG/JPEG. */
  transcodeImage?: (data: Uint8Array, mime: string) => Promise<{ data: Uint8Array; mime: string } | null>;
  onProgress?: (fraction: number) => void;
}

export interface PdfOptions {
  /** 'auto' keeps the document's own page setup (DOCX), A4 otherwise. */
  pageSize?: 'auto' | 'a4' | 'letter';
  /** Subset embedded fonts (default). Disabled in tests to compare renderings. */
  subsetFonts?: boolean;
}

export type Color = [number, number, number];

export type Op =
  | { k: 'text'; x: number; y: number; top: number; bottom: number; text: string; font: FontKey; size: number; color: Color; field?: 'page' | 'pages'; fieldW?: number; anchor?: Align; noteRef?: number }
  | { k: 'line'; x1: number; y1: number; x2: number; y2: number; w: number; color: Color; top: number; bottom: number }
  | { k: 'rect'; x: number; y: number; w: number; h: number; fill?: Color; stroke?: Color; clip: true; top: number; bottom: number }
  | { k: 'image'; x: number; y: number; w: number; h: number; img: PDFImage; top: number; bottom: number }
  | { k: 'meta'; x: number; y: number; w: number; h: number; meta: Metafile; fonts: ({ key: FontKey; text: string }[] | null)[]; imgs: (PDFImage | null)[]; top: number; bottom: number }
  | { k: 'link'; x: number; y: number; w: number; h: number; url: string; top: number; bottom: number };

export interface Box {
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
  /** No line-break opportunity before this token (same word, different font). */
  join: boolean;
  run: TextRun;
  noteRef?: number;
}

/** Inherited layout context. */
export interface Ctx {
  family: FontFamily;
  size: number;
  bold?: boolean;
  italic?: boolean;
  color?: Color;
}

interface Note {
  num: number;
  blocks: Block[];
  boxes: Box[];
  height: number;
}

const BLACK: Color = [0.1, 0.1, 0.12];
const GRAY: Color = [0.45, 0.45, 0.5];
const LINK: Color = [0.02, 0.39, 0.76];
const HEADING_SCALE = [0, 2, 1.64, 1.36, 1.18, 1.09, 1];
const MM = 72 / 25.4;
const FALLBACK_CHARS: Record<string, string> = { '\u2028': ' ', '\u2029': ' ' };

export const hexColor = (hex: string): Color => [parseInt(hex.slice(0, 2), 16) / 255, parseInt(hex.slice(2, 4), 16) / 255, parseInt(hex.slice(4, 6), 16) / 255];

/** A PDF function mapping 0…1 to the gradient colours: one exponential segment per stop pair, stitched. */
function gradientFunction(ctx: PDFContext, g: Gradient, channel: 'rgb' | 'alpha' = 'rgb'): PDFObject {
  const rgb = (c: string): number[] => [0, 2, 4].map((o) => Math.round(((parseInt(c.slice(o, o + 2), 16) || 0) / 255) * 1000) / 1000);
  const value = (st: { color: string; alpha?: number }): number[] => (channel === 'rgb' ? rgb(st.color) : [Math.round((st.alpha ?? 1) * 1000) / 1000]);
  const stops = g.stops.map((s) => ({ pos: Math.min(1, Math.max(0, s.pos)), color: s.color, alpha: s.alpha })).sort((a, b) => a.pos - b.pos);
  if (!stops.length) stops.push({ pos: 0, color: '000000', alpha: 1 });
  if (stops[0].pos > 0) stops.unshift({ ...stops[0], pos: 0 });
  if (stops[stops.length - 1].pos < 1) stops.push({ ...stops[stops.length - 1], pos: 1 });
  // bounds must increase strictly: nudge hard stops apart
  for (let i = 1; i < stops.length; i++) if (stops[i].pos <= stops[i - 1].pos) stops[i].pos = Math.min(1, stops[i - 1].pos + 1e-4);
  const seg = (a: { color: string; alpha?: number }, b: { color: string; alpha?: number }): PDFObject => ctx.obj({ FunctionType: 2, Domain: [0, 1], C0: value(a), C1: value(b), N: 1 });
  if (stops.length === 2) return seg(stops[0], stops[1]);
  const fns = stops.slice(1).map((s, i) => seg(stops[i], s));
  return ctx.obj({ FunctionType: 3, Domain: [0, 1], Functions: fns, Bounds: stops.slice(1, -1).map((s) => Math.round(s.pos * 10000) / 10000), Encode: fns.flatMap(() => [0, 1]) });
}

export class PdfWriter {
  protected pdf!: PDFDocument;
  private shadings = new Map<Gradient, PDFRef>();
  private pageShadings = new WeakMap<PDFPage, number>();

  /**
   * Fill SVG path `d` (drawn at x, y like drawSvgPath, y down) with a gradient: the path becomes the
   * clip and an axial/radial shading paints it, so colours blend smoothly in any viewer and printer.
   */
  protected fillGradient(page: PDFPage, d: string, g: Gradient, x: number, y: number, opacity = 1): void {
    const ctx = this.pdf.context;
    let ref = this.shadings.get(g);
    if (!ref) {
      const c = g.coords.map((v) => Math.round(v * 1000) / 1000);
      ref = ctx.register(ctx.obj({ ShadingType: g.kind === 'linear' ? 2 : 3, ColorSpace: 'DeviceRGB', Coords: c, Function: gradientFunction(ctx, g), Extend: [true, true] }));
      this.shadings.set(g, ref);
    }
    const res = page.node.normalizedEntries().Resources;
    let dict = res.lookupMaybe(PDFName.of('Shading'), PDFDict);
    if (!dict) {
      dict = ctx.obj({});
      res.set(PDFName.of('Shading'), dict);
    }
    const n = (this.pageShadings.get(page) ?? 0) + 1;
    this.pageShadings.set(page, n);
    const name = PDFName.of(`Sh${n}`);
    dict.set(name, ref);
    const alphas = g.stops.map((s) => s.alpha ?? 1);
    const varying = Math.max(...alphas) - Math.min(...alphas) > 0.005;
    const alpha = opacity * (varying ? 1 : alphas[0]);
    const gsDict = ctx.obj({ Type: 'ExtGState' });
    let gsUsed = false;
    if (alpha < 0.999) {
      gsDict.set(PDFName.of('ca'), PDFNumber.of(Math.round(alpha * 1000) / 1000));
      gsUsed = true;
    }
    if (varying) {
      // Stop opacity varies: a luminosity soft mask painted with the same geometry, alpha as gray
      const c = g.coords.map((v) => Math.round(v * 1000) / 1000);
      const alphaSh = ctx.register(ctx.obj({ ShadingType: g.kind === 'linear' ? 2 : 3, ColorSpace: 'DeviceGray', Coords: c, Function: gradientFunction(ctx, g, 'alpha'), Extend: [true, true] }));
      const m = g.matrix ? `${g.matrix.map((v) => Math.round(v * 10000) / 10000).join(' ')} cm ` : '';
      const form = ctx.register(
        ctx.stream(`${m}/Sa sh`, {
          Type: 'XObject',
          Subtype: 'Form',
          BBox: [-100000, -100000, 100000, 100000],
          Group: { S: 'Transparency', CS: 'DeviceGray' },
          Resources: { Shading: { Sa: alphaSh } },
        }),
      );
      gsDict.set(PDFName.of('SMask'), ctx.obj({ Type: 'Mask', S: 'Luminosity', G: form }));
      gsUsed = true;
    }
    // The mask is defined in the user space current when the state is set: after the path transforms
    const gs = gsUsed ? page.node.newExtGState('GS', gsDict) : undefined;
    // drawSvgPath gives: q, transforms, path…, h, Q — keep the transforms and the path
    const body = drawSvgPath(d, { x, y, scale: undefined, color: undefined, borderColor: undefined, borderWidth: 0 }).slice(1, -2);
    page.pushOperators(
      pushGraphicsState(),
      ...body,
      clip(),
      endPath(),
      ...(gs ? [setGraphicsState(gs)] : []),
      ...(g.matrix ? [concatTransformationMatrix(...g.matrix)] : []),
      PDFOperator.of(PDFOperatorNames.ShadingFill, [name]),
      popGraphicsState(),
    );
  }
  protected fm: FontManager;
  protected images = new Map<Uint8Array, PDFImage | null>();
  warnings = new Set<string>();
  protected pageW = 595.28;
  protected pageH = 841.89;
  private margins = { top: 20 * MM, right: 20 * MM, bottom: 20 * MM, left: 20 * MM };
  private headerDist = 12.5 * MM;
  private footerDist = 12.5 * MM;
  private contentTop = 0;
  private contentBottom = 0;
  private pages: PDFPage[] = [];
  private page!: PDFPage;
  private cursor = 0;
  protected base: Ctx = { family: 'sans', size: 11 };
  private defaults: ParaFormat = {};
  private notes: Note[] = [];
  private placedNotes = new Set<number>();
  private pageNotes: Note[] = [];
  private headerBoxes: Box[] = [];
  private footerBoxes: Box[] = [];
  private headerBlocks: Block[] = [];
  private footerBlocks: Block[] = [];

  constructor(
    protected env: PdfWriteEnv,
    protected opts: PdfOptions,
  ) {
    this.fm = new FontManager(env.fonts, env.subsetter);
  }

  get contentW(): number {
    return this.pageW - this.margins.left - this.margins.right;
  }

  async init(doc: Doc): Promise<void> {
    this.pdf = await PDFDocument.create();
    this.pdf.registerFontkit((await import('@pdf-lib/fontkit')).default);
    const size = this.opts.pageSize ?? 'auto';
    if (size === 'auto' && doc.page) {
      this.pageW = doc.page.width;
      this.pageH = doc.page.height;
      this.margins = { ...doc.page.margins };
    } else if (size === 'letter') {
      this.pageW = 612;
      this.pageH = 792;
    }
    if (doc.page) this.margins = { ...doc.page.margins };
    this.contentTop = this.margins.top;
    this.contentBottom = this.pageH - this.margins.bottom;
    this.base = { family: doc.font ?? 'sans', size: doc.size ?? 11 };
    this.defaults = doc.defaults ?? {};
    if (doc.title) this.pdf.setTitle(doc.title);
    this.pdf.setCreator('Salatic');
    this.pdf.setProducer('Salatic');
    this.pdf.setLanguage('ru-RU');
  }

  // ------------------------------------------------------------------ text

  protected keyFor(r: TextRun, ctx: Ctx): FontKey {
    const family: FontFamily = r.code ? 'mono' : (r.font ?? ctx.family);
    return `${family}-${styleOf(!!(r.bold || ctx.bold), !!(r.italic || ctx.italic))}`;
  }

  protected sizeFor(r: TextRun, ctx: Ctx): number {
    const s = r.size ?? (r.code && !r.font ? ctx.size * 0.92 : ctx.size);
    return r.sup || r.sub || r.note ? s * 0.68 : s;
  }

  /** Split runs into tokens; each token is drawn with a single font. */
  protected async tokens(runs: TextRun[], ctx: Ctx): Promise<Token[]> {
    const out: Token[] = [];
    for (const r of runs) {
      const primary = this.keyFor(r, ctx);
      const size = this.sizeFor(r, ctx);
      await this.fm.load([primary]);
      if (!this.fm.loaded(primary)) throw new Error(`Шрифт ${primary} недоступен`);
      if (r.note) {
        const num = this.notes.length + 1;
        this.notes.push({ num, blocks: r.note, boxes: [], height: 0 });
        const text = String(num);
        this.fm.markUsed(primary, text);
        out.push({ text, font: primary, size, width: this.fm.width(primary, text, size), space: false, newline: false, join: true, run: r, noteRef: num });
        continue;
      }
      if (r.field) {
        const text = '00';
        this.fm.markUsed(primary, '0123456789');
        out.push({ text, font: primary, size, width: this.fm.width(primary, text, size), space: false, newline: false, join: false, run: r });
        continue;
      }
      let text = cleanText(r.text).replace(/\r/g, '').replace(/\t/g, '    ');
      text = text.replace(/[\u2028\u2029]/g, (c) => FALLBACK_CHARS[c]);
      // Load fallback fonts needed for this text before measuring
      const need = new Set<FontKey>();
      for (const ch of text) for (const k of this.fm.fallbacksFor(primary, ch.codePointAt(0)!)) need.add(k);
      if (need.size) await this.fm.load(need);
      for (const part of text.split(/(\n| +)/)) {
        if (!part) continue;
        if (part === '\n') {
          out.push({ text: '', font: primary, size, width: 0, space: false, newline: true, join: false, run: r });
          continue;
        }
        if (part[0] === ' ') {
          const font = this.fm.has(primary, 32) ? primary : (this.fm.resolve(primary, 32) ?? primary);
          this.fm.markUsed(font, part);
          out.push({ text: part, font, size, width: this.fm.width(font, part, size), space: true, newline: false, join: false, run: r });
          continue;
        }
        // Split the word where the font changes; CJK characters are break opportunities
        let seg = '';
        let segFont: FontKey | null = null;
        // A word glued to the previous run (e.g. "**bold**text") must not break between them
        const prev = out[out.length - 1];
        let join = !!prev && !prev.space && !prev.newline;
        const flush = (): void => {
          if (!seg || !segFont) return;
          this.fm.markUsed(segFont, seg);
          out.push({ text: seg, font: segFont, size, width: this.fm.width(segFont, seg, size), space: false, newline: false, join, run: r });
          seg = '';
          join = true;
        };
        for (const ch of part) {
          const cp = ch.codePointAt(0)!;
          let font = this.fm.resolve(primary, cp);
          let c = ch;
          if (!font) {
            this.fm.missing = true;
            font = primary;
            c = '?';
          }
          const cjk = font === 'cjk';
          if (font !== segFont || cjk) {
            flush();
            segFont = font;
            if (cjk) join = false;
          }
          seg += c;
          if (cjk) {
            flush();
            join = false; // a CJK character may be followed by a line break
          }
        }
        flush();
      }
    }
    return out;
  }

  /** Group tokens into unbreakable units (a word possibly made of several fonts). */
  protected breakLines(tokens: Token[], firstWidth: number, width: number, keepIndent = false): Token[][] {
    const lines: Token[][] = [];
    let line: Token[] = [];
    let w = 0;
    let afterNewline = true;
    const lineWidth = (): number => (lines.length ? width : firstWidth);
    const push = (): void => {
      while (line.length && line[line.length - 1].space) line.pop();
      lines.push(line);
      line = [];
      w = 0;
    };
    let i = 0;
    while (i < tokens.length) {
      const t = tokens[i];
      if (t.newline) {
        push();
        afterNewline = true;
        i++;
        continue;
      }
      if (t.space) {
        if (line.length || (keepIndent && afterNewline)) {
          line.push(t);
          w += t.width;
        }
        i++;
        continue;
      }
      afterNewline = false;
      // Collect the unit
      const unit: Token[] = [t];
      let j = i + 1;
      while (j < tokens.length && tokens[j].join && !tokens[j].space && !tokens[j].newline) unit.push(tokens[j++]);
      const uw = unit.reduce((s, x) => s + x.width, 0);
      const limit = lineWidth();
      if (w + uw <= limit + 0.01 || !line.length || line.every((x) => x.space)) {
        if (uw > limit && !line.some((x) => !x.space)) {
          // Hard-break a unit wider than the line, character by character
          line = [];
          w = 0;
          for (const u of unit) {
            let chunk = '';
            for (const ch of u.text) {
              const cw = this.fm.width(u.font, chunk + ch, u.size);
              if (w + cw > lineWidth() && (chunk || line.length)) {
                if (chunk) line.push({ ...u, text: chunk, width: this.fm.width(u.font, chunk, u.size) });
                push();
                chunk = ch;
              } else chunk += ch;
            }
            if (chunk) {
              const cw = this.fm.width(u.font, chunk, u.size);
              line.push({ ...u, text: chunk, width: cw });
              w += cw;
            }
          }
        } else {
          line.push(...unit);
          w += uw;
        }
      } else {
        push();
        line.push(...unit);
        w = uw;
      }
      i = j;
    }
    if (line.length || !lines.length) push();
    return lines;
  }

  private para(b: ParaFormat | undefined, fallback: ParaFormat): Required<Omit<ParaFormat, 'align'>> & { align?: Align } {
    const d = this.defaults;
    return {
      align: b?.align ?? fallback.align,
      indent: b?.indent ?? fallback.indent ?? 0,
      firstLine: b?.firstLine ?? fallback.firstLine ?? d.firstLine ?? 0,
      spaceBefore: b?.spaceBefore ?? fallback.spaceBefore ?? d.spaceBefore ?? 0,
      spaceAfter: b?.spaceAfter ?? fallback.spaceAfter ?? d.spaceAfter ?? this.base.size * 0.55,
      lineHeight: b?.lineHeight ?? fallback.lineHeight ?? d.lineHeight ?? 1.2,
    };
  }

  /** Lay out runs as a paragraph box. */
  protected async textBox(runs: TextRun[], x: number, width: number, ctx: Ctx, p: ReturnType<PdfWriter['para']>, keepIndent = false): Promise<Box> {
    const tokens = await this.tokens(runs, ctx);
    const left = x + p.indent;
    const firstX = left + p.firstLine;
    const lines = this.breakLines(tokens, width - p.indent - p.firstLine, width - p.indent, keepIndent);
    const ops: Op[] = [];
    let y = 0;
    const baseKey = this.keyFor({ text: '' }, ctx);
    await this.fm.load([baseKey]);
    lines.forEach((line, li) => {
      const lx = li === 0 ? firstX : left;
      const lw = li === 0 ? width - p.indent - p.firstLine : width - p.indent;
      // Line box from the tallest font on the line
      let asc = 0;
      let desc = 0;
      let natural = 0;
      const measureSet = line.length ? line : [{ font: baseKey, size: ctx.size, run: {} as TextRun } as Token];
      for (const t of measureSet) {
        const m = this.fm.metrics(t.font);
        const s = t.run.sup || t.run.sub || t.run.note ? t.size / 0.68 : t.size;
        asc = Math.max(asc, m.ascent * s);
        desc = Math.max(desc, m.descent * s);
        natural = Math.max(natural, m.line * s);
      }
      const lh = Math.max(natural * p.lineHeight, natural * 0.8);
      const baseline = y + asc;
      const lineW = line.reduce((s, t) => s + t.width, 0);
      let cx = lx;
      let extraPerSpace = 0;
      if (p.align === 'center') cx = lx + (lw - lineW) / 2;
      else if (p.align === 'right') cx = lx + lw - lineW;
      else if (p.align === 'justify' && li < lines.length - 1) {
        const spaces = line.filter((t) => t.space).length;
        if (spaces && lineW > lw * 0.6) extraPerSpace = Math.max(0, (lw - lineW) / spaces);
      }
      let pending: { x: number; text: string; t: Token; w: number } | null = null;
      const flush = (): void => {
        if (!pending) return;
        const t = pending.t;
        const r = t.run;
        let by = baseline;
        const full = t.size / (r.sup || r.sub || r.note ? 0.68 : 1);
        if (r.sup || r.note) by -= full * 0.33;
        if (r.sub) by += full * 0.14;
        const color = r.link ? LINK : r.color ? hexColor(r.color) : (ctx.color ?? BLACK);
        if (r.highlight) ops.push({ k: 'rect', x: pending.x, y, w: pending.w, h: lh, fill: hexColor(r.highlight), clip: true, top: y, bottom: y + lh });
        if (pending.text.trim()) {
          const op: Op = { k: 'text', x: pending.x, y: by, top: y, bottom: y + lh, text: pending.text, font: t.font, size: t.size, color };
          if (r.field) Object.assign(op, { field: r.field, fieldW: pending.w, anchor: p.align ?? 'left' });
          if (t.noteRef) op.noteRef = t.noteRef;
          ops.push(op);
        }
        const thick = Math.max(0.5, t.size * 0.05);
        if (r.underline || r.link) ops.push({ k: 'line', x1: pending.x, y1: by + t.size * 0.12, x2: pending.x + pending.w, y2: by + t.size * 0.12, w: thick, color, top: y, bottom: y + lh });
        if (r.strike) ops.push({ k: 'line', x1: pending.x, y1: by - t.size * 0.3, x2: pending.x + pending.w, y2: by - t.size * 0.3, w: thick, color, top: y, bottom: y + lh });
        if (r.link) ops.push({ k: 'link', x: pending.x, y, w: pending.w, h: lh, url: r.link, top: y, bottom: y + lh });
        pending = null;
      };
      for (const t of line) {
        const w = t.width + (t.space ? extraPerSpace : 0);
        if (pending && pending.t.run === t.run && pending.t.font === t.font && !extraPerSpace && !t.noteRef && !pending.t.noteRef) {
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
    return { height: y, ops, spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter };
  }

  // ------------------------------------------------------------------ blocks

  /** Vector picture (EMF/WMF): parse, prepare its fonts and bitmaps; drawn by draw(). */
  protected async metaOp(data: Uint8Array, x: number, y: number, w: number, h: number, mime = 'image/emf'): Promise<Extract<Op, { k: 'meta' }> | null> {
    const meta = mime === 'image/svg+xml' ? parseSvg(data) : parseMetafile(data);
    if (!meta || !meta.cmds.length) {
      this.warnings.add(mime === 'image/svg+xml' ? 'Некоторые рисунки SVG не удалось прочитать и они пропущены.' : 'Некоторые векторные рисунки (EMF/WMF) не удалось прочитать и они пропущены.');
      return null;
    }
    const fonts: ({ key: FontKey; text: string }[] | null)[] = [];
    const imgs: (PDFImage | null)[] = [];
    for (const c of meta.cmds) {
      if (c.k === 'text') {
        const primary: FontKey = `${c.family}-${styleOf(c.bold, c.italic)}`;
        await this.fm.load([primary]);
        const need = new Set<FontKey>();
        for (const ch of c.text) for (const k of this.fm.fallbacksFor(primary, ch.codePointAt(0)!)) need.add(k);
        if (need.size) await this.fm.load(need);
        const segs: { key: FontKey; text: string }[] = [];
        for (const ch of c.text) {
          const key = this.fm.resolve(primary, ch.codePointAt(0)!) ?? primary;
          const t = this.fm.resolve(primary, ch.codePointAt(0)!) ? ch : '?';
          this.fm.markUsed(key, t);
          if (segs.length && segs[segs.length - 1].key === key) segs[segs.length - 1].text += t;
          else segs.push({ key, text: t });
        }
        fonts.push(segs);
        imgs.push(null);
      } else if (c.k === 'image') {
        fonts.push(null);
        let img: PDFImage | null = null;
        try {
          img = c.png[0] === 0xff && c.png[1] === 0xd8 ? await this.pdf.embedJpg(c.png) : await this.pdf.embedPng(c.png);
        } catch {
          img = null;
        }
        imgs.push(img);
      } else {
        fonts.push(null);
        imgs.push(null);
      }
    }
    return { k: 'meta', x, y, w, h, meta, fonts, imgs, top: y, bottom: y + h };
  }

  /** Draw a prepared metafile with its top-left corner at (x, top) in PDF coordinates (y up). */
  protected drawMeta(page: PDFPage, o: Extract<Op, { k: 'meta' }>, x: number, top: number): void {
    const m = o.meta;
    const sx = o.w / m.width;
    const sy = o.h / m.height;
    page.pushOperators(pushGraphicsState(), rectangle(x, top - o.h, o.w, o.h), clip(), endPath(), concatTransformationMatrix(sx, 0, 0, sy, x, top));
    let flowX = 0;
    m.cmds.forEach((c, i) => {
      if (c.k === 'path') {
        if (c.grad && c.fill) {
          this.fillGradient(page, c.d, c.grad, 0, 0, c.opacity ?? 1);
          if (c.stroke) page.drawSvgPath(c.d, { x: 0, y: 0, borderColor: rgb(...hexColor(c.stroke)), borderWidth: c.width, borderDashArray: c.dash, borderOpacity: c.strokeOpacity !== undefined && c.strokeOpacity < 1 ? c.strokeOpacity : undefined });
          return;
        }
        page.drawSvgPath(c.d, {
          x: 0,
          y: 0,
          color: c.fill ? rgb(...hexColor(c.fill)) : undefined,
          borderColor: c.stroke ? rgb(...hexColor(c.stroke)) : undefined,
          borderWidth: c.stroke ? c.width : 0,
          borderDashArray: c.dash,
          opacity: c.opacity !== undefined && c.opacity < 1 ? c.opacity : undefined,
          borderOpacity: c.strokeOpacity !== undefined && c.strokeOpacity < 1 ? c.strokeOpacity : undefined,
        });
      } else if (c.k === 'image') {
        const img = o.imgs[i];
        if (img) page.drawImage(img, { x: c.x, y: -(c.y + c.h), width: c.w, height: c.h });
      } else {
        const segs = o.fonts[i] ?? [];
        const width = segs.reduce((s, sg) => s + this.fm.width(sg.key, sg.text, c.size), 0);
        let cx = c.flow ? flowX : c.anchor === 'middle' ? c.x - width / 2 : c.anchor === 'end' ? c.x - width : c.x;
        const startX = cx;
        const base = c.baseline === 'top' ? c.y + c.size * 0.88 : c.baseline === 'bottom' ? c.y - c.size * 0.21 : c.y;
        const rad = (c.angle * Math.PI) / 180;
        for (const sg of segs) {
          // rotation is around the reference point, counter-clockwise like GDI escapement
          const dx = cx - c.x;
          const px = c.x + dx * Math.cos(rad) + (base - c.y) * Math.sin(rad);
          const py = c.y - dx * Math.sin(rad) + (base - c.y) * Math.cos(rad);
          page.drawText(sg.text, { x: px, y: -py, size: c.size, font: this.fm.pdfFont(sg.key), color: rgb(...hexColor(c.color)), rotate: c.angle ? degrees(c.angle) : undefined });
          cx += this.fm.width(sg.key, sg.text, c.size);
        }
        flowX = startX + width;
      }
    });
    page.pushOperators(popGraphicsState());
  }

  protected async image(b: Extract<Block, { type: 'image' }>): Promise<PDFImage | null> {
    if (this.images.has(b.data)) return this.images.get(b.data)!;
    let data = b.data;
    let mime = b.mime;
    if (mime !== 'image/png' && mime !== 'image/jpeg') {
      const t = this.env.transcodeImage ? await this.env.transcodeImage(data, mime).catch(() => null) : null;
      if (!t) {
        this.warnings.add('Некоторые изображения пропущены: формат не поддерживается в этой среде.');
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
    return box.ops.map((o) =>
      o.k === 'line'
        ? { ...o, x1: o.x1 + dx, x2: o.x2 + dx, y1: o.y1 + dy, y2: o.y2 + dy, top: o.top + dy, bottom: o.bottom + dy }
        : ({ ...o, x: o.x + dx, y: o.y + dy, top: o.top + dy, bottom: o.bottom + dy } as Op),
    );
  }

  private async layout(blocks: Block[], x: number, width: number, ctx: Ctx, depth = 0): Promise<Box[]> {
    const boxes: Box[] = [];
    for (const b of blocks) {
      switch (b.type) {
        case 'paragraph': {
          const p = this.para(b, {});
          if (!b.runs.some((r) => r.text.trim() || r.note || r.field)) {
            const key = this.keyFor(b.runs[0] ?? { text: '' }, ctx);
            await this.fm.load([key]);
            const s = b.runs[0]?.size ?? ctx.size;
            boxes.push({ height: this.fm.metrics(key).line * s * p.lineHeight, ops: [], spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter });
          } else boxes.push(await this.textBox(b.runs, x, width, ctx, p));
          break;
        }
        case 'heading': {
          const size = b.runs.find((r) => r.size)?.size ?? Math.round(this.base.size * HEADING_SCALE[b.level] * 2) / 2;
          const p = this.para(b, { spaceBefore: size * 0.8, spaceAfter: size * 0.4, lineHeight: 1.1 });
          const box = await this.textBox(b.runs, x, width, { ...ctx, size, bold: true }, p);
          box.keepWithNext = true;
          boxes.push(box);
          break;
        }
        case 'list': {
          let n = b.start ?? 1;
          const key = this.keyFor({ text: '' }, ctx);
          await this.fm.load([key]);
          const markerW = b.ordered ? this.fm.width(key, `${n + b.items.length - 1}.`, ctx.size) + 7 : ctx.size * 1.3;
          for (const it of b.items) {
            const inner = await this.layout(it.blocks.length ? it.blocks : [{ type: 'paragraph', runs: [{ text: ' ' }] }], x + markerW + 4, width - markerW - 4, ctx, depth + 1);
            inner.forEach((ib, k) => {
              if (k < inner.length - 1) ib.spaceAfter = Math.min(ib.spaceAfter, 3);
            });
            const first = inner[0];
            if (first) {
              const firstText = first.ops.find((o): o is Extract<Op, { k: 'text' }> => o.k === 'text');
              const baseline = firstText ? firstText.y : ctx.size;
              const pick = b.ordered ? `${n}.` : ['•', '◦', '▪'][depth % 3];
              const marker = this.fm.has(key, pick.codePointAt(0)!) ? pick : '•';
              this.fm.markUsed(key, marker);
              const mw = this.fm.width(key, marker, ctx.size);
              // Marker first in the content stream so text extractors read it before the item text
              first.ops.unshift({ k: 'text', x: b.ordered ? x + markerW - mw - 3 : x + 3, y: baseline, top: 0, bottom: firstText ? firstText.bottom : ctx.size * 1.3, text: marker, font: key, size: ctx.size, color: ctx.color ?? BLACK });
              inner[inner.length - 1].spaceAfter = 3;
            }
            boxes.push(...inner);
            n++;
          }
          if (boxes.length) boxes[boxes.length - 1].spaceAfter = this.para(undefined, {}).spaceAfter;
          break;
        }
        case 'table':
          boxes.push(...(await this.table(b, x, width, ctx, depth)));
          break;
        case 'image': {
          if (b.mime === 'image/emf' || b.mime === 'image/wmf' || b.mime === 'image/svg+xml') {
            let w = (b.width * 72) / 96;
            let h = (b.height * 72) / 96;
            const maxH = this.contentBottom - this.contentTop - 20;
            const k = Math.min(1, width / w, maxH / h);
            w *= k;
            h *= k;
            const ix = w < width ? x + (width - w) / 2 : x;
            const op = await this.metaOp(b.data, ix, 0, w, h, b.mime);
            if (op) boxes.push({ height: h, ops: [op], spaceBefore: 4, spaceAfter: 8 });
            break;
          }
          const img = await this.image(b);
          if (!img) break;
          let w = (b.width * 72) / 96;
          let h = (b.height * 72) / 96;
          const maxH = this.contentBottom - this.contentTop - 20;
          const k = Math.min(1, width / w, maxH / h);
          w *= k;
          h *= k;
          const ix = w < width ? x + (width - w) / 2 : x;
          boxes.push({ height: h, ops: [{ k: 'image', x: ix, y: 0, w, h, img, top: 0, bottom: h }], spaceBefore: 4, spaceAfter: 8 });
          break;
        }
        case 'code': {
          const pad = 6;
          const inner = await this.textBox([{ text: b.text, font: 'mono', size: Math.max(7, ctx.size - 1.5) }], x + pad, width - 2 * pad, ctx, this.para(undefined, { spaceAfter: 0, spaceBefore: 0, lineHeight: 1.1, firstLine: 0 }), true);
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
          const inner = await this.layout(b.blocks, x + 14, width - 14, { ...ctx, color: [0.27, 0.27, 0.3] }, depth);
          for (const ib of inner) ib.ops.unshift({ k: 'rect', x, y: -ib.spaceBefore, w: 3, h: ib.height + ib.spaceBefore + ib.spaceAfter, fill: [0.8, 0.8, 0.84], clip: true, top: 0, bottom: ib.height });
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

  private async table(b: Extract<Block, { type: 'table' }>, x: number, width: number, ctx: Ctx, depth: number): Promise<Box[]> {
    const cols = maxOf(b.rows.map((r) => r.cells.reduce((s, c) => s + (c.colspan ?? 1), 0)), 1);
    const pad = 4;
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
              const cctx = { ...ctx, bold: ctx.bold || cell.header };
              const key = this.keyFor(run, cctx);
              await this.fm.load([key]);
              const s = this.sizeFor(run, cctx);
              const measure = (t: string): number => {
                let w = 0;
                for (const ch of t) {
                  const k = this.fm.resolve(key, ch.codePointAt(0)!) ?? key;
                  w += this.fm.advance(k, this.fm.has(k, ch.codePointAt(0)!) ? ch.codePointAt(0)! : 63);
                }
                return (w * s) / 1000;
              };
              for (const line of run.text.split('\n')) maxW[c] = Math.max(maxW[c], measure(line) + 2 * pad);
              for (const word of run.text.split(/\s+/)) minW[c] = Math.max(minW[c], Math.min(width / cols, measure(word) + 2 * pad));
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
    const cellPara = { spaceAfter: 2, spaceBefore: 0, firstLine: 0, indent: 0 };
    for (const r of b.rows) {
      let c = 0;
      const cellBoxes: { x: number; w: number; boxes: Box[]; header: boolean }[] = [];
      for (const cell of r.cells) {
        if (c >= cols) break;
        const span = Math.min(cell.colspan ?? 1, cols - c);
        const cx = x + widths.slice(0, c).reduce((a, v) => a + v, 0);
        const cw = widths.slice(c, c + span).reduce((a, v) => a + v, 0);
        const blocks = cell.blocks.map((bl) => (bl.type === 'paragraph' ? { ...cellPara, ...bl } : bl));
        cellBoxes.push({ x: cx, w: cw, boxes: await this.layout(blocks, cx + pad, cw - 2 * pad, { ...ctx, bold: ctx.bold || cell.header }, depth + 1), header: !!cell.header });
        c += span;
      }
      if (c < cols) {
        const cx = x + widths.slice(0, c).reduce((a, v) => a + v, 0);
        cellBoxes.push({ x: cx, w: width - (cx - x), boxes: [], header: false });
      }
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
      rowH = Math.max(rowH, ctx.size * 1.3 + 2 * pad);
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

  // ------------------------------------------------------------------ headers, footers, notes

  async prepare(doc: Doc): Promise<Box[]> {
    const ctx = { ...this.base };
    const w = this.contentW;
    const x = this.margins.left;
    const strip = (boxes: Box[]): Box[] => {
      if (boxes.length) {
        boxes[0].spaceBefore = 0;
        boxes[boxes.length - 1].spaceAfter = 0;
      }
      return boxes;
    };
    const small = { ...ctx, size: Math.max(7, ctx.size - 1) };
    this.headerBlocks = doc.header ?? [];
    this.footerBlocks = doc.footer ?? [];
    if (this.headerBlocks.length) this.headerBoxes = strip(await this.layout(this.headerBlocks, x, w, small));
    if (this.footerBlocks.length) this.footerBoxes = strip(await this.layout(this.footerBlocks, x, w, small));
    const stack = (boxes: Box[]): number => boxes.reduce((s, b, i) => s + b.height + (i ? b.spaceBefore : 0) + (i < boxes.length - 1 ? b.spaceAfter : 0), 0);
    this.contentTop = Math.max(this.margins.top, this.headerBoxes.length ? this.headerDist + stack(this.headerBoxes) + 6 : 0);
    this.contentBottom = Math.min(this.pageH - this.margins.bottom, this.footerBoxes.length ? this.pageH - this.footerDist - stack(this.footerBoxes) - 6 : Infinity);
    // Main flow (footnotes are numbered while laying it out)
    const main = await this.layout(doc.blocks.length ? doc.blocks : [{ type: 'paragraph', runs: [] }], x, w, ctx);
    // Footnote bodies
    const noteCtx = { ...ctx, size: Math.max(7, ctx.size * 0.8) };
    for (let i = 0; i < this.notes.length; i++) {
      const n = this.notes[i];
      const blocks = n.blocks.length ? [...n.blocks] : [{ type: 'paragraph', runs: [] } as Block];
      const first = blocks[0];
      const mark: TextRun = { text: `${n.num} `, sup: false };
      blocks[0] = first.type === 'paragraph' ? { ...first, runs: [mark, ...first.runs], firstLine: 0, indent: 0, spaceBefore: 0 } : first;
      if (first.type !== 'paragraph') blocks.unshift({ type: 'paragraph', runs: [mark] });
      n.boxes = strip(await this.layout(blocks, x, w, noteCtx));
      n.boxes.forEach((b) => (b.spaceAfter = Math.min(b.spaceAfter, 2)));
      n.height = stack(n.boxes) + 2;
    }
    if (this.fm.missing) this.warnings.add('Некоторые символы отсутствуют во всех встроенных шрифтах и заменены на «?».');
    // Page numbers when the document has no footer of its own
    this.fm.markUsed(this.keyFor({ text: '' }, ctx), '0123456789');
    await this.fm.embedAll(this.pdf, this.opts.subsetFonts !== false);
    return main;
  }

  // ------------------------------------------------------------------ placement

  private newPage(): void {
    if (this.page) this.flushNotes();
    this.page = this.pdf.addPage([this.pageW, this.pageH]);
    this.pages.push(this.page);
    this.cursor = this.contentTop;
    this.pageNotes = [];
  }

  private notesHeight(list: Note[]): number {
    if (!list.length) return 0;
    return list.reduce((s, n) => s + n.height, 0) + (this.pageNotes.length ? 0 : 10);
  }

  private get limit(): number {
    return this.contentBottom - this.notesHeight(this.pageNotes);
  }

  private notesIn(box: Box, from: number, to: number): Note[] {
    const out: Note[] = [];
    for (const o of box.ops) {
      if (o.k === 'text' && o.noteRef && o.top >= from - 0.01 && o.bottom <= to + 0.01 && !this.placedNotes.has(o.noteRef)) {
        const n = this.notes[o.noteRef - 1];
        if (n && !out.includes(n)) out.push(n);
      }
    }
    return out;
  }

  private flushNotes(): void {
    if (!this.pageNotes.length) return;
    const total = this.notesHeight(this.pageNotes);
    let y = this.contentBottom - total;
    this.page.drawLine({ start: { x: this.margins.left, y: this.pageH - (y + 4) }, end: { x: this.margins.left + this.contentW / 3, y: this.pageH - (y + 4) }, thickness: 0.5, color: rgb(...GRAY) });
    y += 10;
    for (const n of this.pageNotes) {
      let by = y;
      n.boxes.forEach((b, i) => {
        if (i) by += b.spaceBefore;
        this.draw(b.ops, by, 0, b.height);
        by += b.height + b.spaceAfter;
      });
      y += n.height;
    }
  }

  protected draw(ops: Op[], offset: number, from: number, to: number, page = this.page, pageNo = 0, total = 0): void {
    const H = this.pageH;
    for (const o of ops) {
      if (o.k === 'rect') {
        const top = Math.max(o.y, from);
        const bot = Math.min(o.y + o.h, to);
        if (bot <= top) continue;
        page.drawRectangle({
          x: o.x,
          y: H - (offset + bot - from),
          width: o.w,
          height: bot - top,
          color: o.fill ? rgb(...o.fill) : undefined,
          borderColor: o.stroke ? rgb(...o.stroke) : undefined,
          borderWidth: o.stroke ? 0.6 : 0,
        });
        continue;
      }
      if (o.top < from - 0.01 || o.bottom > to + 0.01) continue;
      switch (o.k) {
        case 'text': {
          let text = o.text;
          let x = o.x;
          if (o.field) {
            text = String(o.field === 'page' ? pageNo : total);
            const aw = this.fm.width(o.font, text, o.size);
            const pw = o.fieldW ?? aw;
            if (o.anchor === 'center') x += (pw - aw) / 2;
            else if (o.anchor === 'right') x += pw - aw;
          }
          page.drawText(text, { x, y: H - (offset + o.y - from), size: o.size, font: this.fm.pdfFont(o.font), color: rgb(...o.color) });
          break;
        }
        case 'line':
          page.drawLine({ start: { x: o.x1, y: H - (offset + o.y1 - from) }, end: { x: o.x2, y: H - (offset + o.y2 - from) }, thickness: o.w, color: rgb(...o.color) });
          break;
        case 'image':
          page.drawImage(o.img, { x: o.x, y: H - (offset + o.y + o.h - from), width: o.w, height: o.h });
          break;
        case 'meta':
          this.drawMeta(page, o, o.x, H - (offset + o.y - from));
          break;
        case 'link': {
          const y1 = H - (offset + o.y + o.h - from);
          const annot = this.pdf.context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [o.x, y1, o.x + o.w, y1 + o.h], Border: [0, 0, 0], A: { Type: 'Action', S: 'URI', URI: PDFString.of(o.url) } });
          page.node.addAnnot(this.pdf.context.register(annot));
          break;
        }
      }
    }
  }

  private candidates(box: Box, from: number, limit: number): number[] {
    const set = new Set<number>();
    for (const o of box.ops) if (o.k !== 'rect' && o.bottom > from && o.bottom <= limit) set.add(o.bottom);
    return [...set].sort((a, b) => b - a).filter((y) => !box.ops.some((o) => o.k !== 'rect' && o.top < y - 0.01 && o.bottom > y + 0.01));
  }

  place(boxes: Box[]): void {
    this.newPage();
    const contentH = this.contentBottom - this.contentTop;
    let pendingBreak = false;
    for (let i = 0; i < boxes.length; i++) {
      const box = boxes[i];
      if (box.pageBreak) {
        pendingBreak = true;
        continue;
      }
      if (pendingBreak) {
        pendingBreak = false;
        if (this.cursor > this.contentTop) this.newPage();
      }
      const atTop = this.cursor <= this.contentTop + 0.01;
      if (!atTop) this.cursor += box.spaceBefore;
      let need = box.height;
      if (box.keepWithNext && boxes[i + 1] && !boxes[i + 1].pageBreak) {
        const next = boxes[i + 1];
        const firstLine = next.ops.filter((o) => o.k !== 'rect').reduce((m, o) => Math.min(m, o.bottom), next.height);
        need += next.spaceBefore + Math.min(firstLine, contentH / 3);
      }
      if (this.cursor + need > this.limit && !atTop && need <= contentH) this.newPage();
      let from = 0;
      while (from < box.height - 0.01) {
        const rest = this.notesIn(box, from, box.height);
        if (this.cursor + (box.height - from) + this.notesHeight([...this.pageNotes, ...rest]) - this.notesHeight(this.pageNotes) <= this.limit + 0.01) {
          this.draw(box.ops, this.cursor, from, box.height);
          this.addNotes(rest);
          this.cursor += box.height - from;
          break;
        }
        let cut = from;
        for (const c of this.candidates(box, from, from + (this.limit - this.cursor))) {
          const notes = this.notesIn(box, from, c);
          if (this.cursor + (c - from) + this.notesHeight([...this.pageNotes, ...notes]) - this.notesHeight(this.pageNotes) <= this.limit + 0.01) {
            cut = c;
            break;
          }
        }
        if (cut <= from) {
          if (this.cursor > this.contentTop + 0.01 || this.pageNotes.length) {
            this.newPage();
            continue;
          }
          cut = Math.min(box.height, from + (this.limit - this.cursor)); // an atomic op taller than the page: clip
          if (cut <= from) cut = box.height;
        }
        const notes = this.notesIn(box, from, cut);
        this.draw(box.ops, this.cursor, from, cut);
        this.addNotes(notes);
        const drawn = cut - from;
        from = cut;
        if (from < box.height - 0.01) this.newPage();
        else this.cursor += drawn;
      }
      this.cursor += box.spaceAfter;
    }
  }

  private addNotes(list: Note[]): void {
    for (const n of list) {
      this.placedNotes.add(n.num);
      this.pageNotes.push(n);
    }
  }

  async finish(): Promise<Uint8Array> {
    this.flushNotes();
    this.pageNotes = [];
    const total = this.pages.length;
    const small = { ...this.base, size: Math.max(7, this.base.size - 1) };
    const strip = (boxes: Box[]): Box[] => {
      if (boxes.length) {
        boxes[0].spaceBefore = 0;
        boxes[boxes.length - 1].spaceAfter = 0;
      }
      return boxes;
    };
    // Headers/footers with page fields are laid out per page once the page count is known
    const perPage = async (blocks: Block[], boxes: Box[], no: number): Promise<Box[]> =>
      hasFields(blocks) ? strip(await this.layout(fillFields(blocks, no, total), this.margins.left, this.contentW, small)) : boxes;
    const headers: Box[][] = [];
    const footers: Box[][] = [];
    for (let i = 0; i < total; i++) {
      headers.push(await perPage(this.headerBlocks, this.headerBoxes, i + 1));
      footers.push(await perPage(this.footerBlocks, this.footerBoxes, i + 1));
    }
    const drawStack = (boxes: Box[], top: number, page: PDFPage, no: number): void => {
      let y = top;
      boxes.forEach((b, i) => {
        if (i) y += b.spaceBefore;
        this.draw(b.ops, y, 0, b.height, page, no, total);
        y += b.height + b.spaceAfter;
      });
    };
    const stackH = (boxes: Box[]): number => boxes.reduce((s, b, i) => s + b.height + (i ? b.spaceBefore : 0) + (i < boxes.length - 1 ? b.spaceAfter : 0), 0);
    this.pages.forEach((p, i) => {
      if (headers[i].length) drawStack(headers[i], this.headerDist, p, i + 1);
      if (footers[i].length) drawStack(footers[i], this.pageH - this.footerDist - stackH(footers[i]), p, i + 1);
      else if (total > 1) {
        const key = this.keyFor({ text: '' }, this.base);
        const label = String(i + 1);
        const w = this.fm.width(key, label, 9);
        p.drawText(label, { x: (this.pageW - w) / 2, y: Math.min(this.margins.bottom / 2, 28) - 3, size: 9, font: this.fm.pdfFont(key), color: rgb(...GRAY) });
      }
    });
    this.pdf.catalog.set(PDFName.of('PageLayout'), PDFName.of('OneColumn'));
    return this.pdf.save({ useObjectStreams: true });
  }
}

function mapRuns(blocks: Block[], fn: (r: TextRun) => TextRun): Block[] {
  return blocks.map((b): Block => {
    switch (b.type) {
      case 'paragraph':
      case 'heading':
        return { ...b, runs: b.runs.map(fn) };
      case 'list':
        return { ...b, items: b.items.map((it) => ({ blocks: mapRuns(it.blocks, fn) })) };
      case 'table':
        return { ...b, rows: b.rows.map((r) => ({ cells: r.cells.map((c) => ({ ...c, blocks: mapRuns(c.blocks, fn) })) })) };
      case 'quote':
        return { ...b, blocks: mapRuns(b.blocks, fn) };
      default:
        return b;
    }
  });
}

function hasFields(blocks: Block[]): boolean {
  let found = false;
  mapRuns(blocks, (r) => ((found ||= !!r.field), r));
  return found;
}

const fillFields = (blocks: Block[], page: number, total: number): Block[] =>
  mapRuns(blocks, (r) => (r.field ? { ...r, field: undefined, text: String(r.field === 'page' ? page : total) } : r));

export async function docToPdf(doc: Doc, env: PdfWriteEnv, opts: PdfOptions = {}): Promise<{ data: Uint8Array; warnings: string[] }> {
  const w = new PdfWriter(env, opts);
  await w.init(doc);
  const boxes = await w.prepare(doc);
  env.onProgress?.(0.6);
  w.place(boxes);
  const data = await w.finish();
  env.onProgress?.(1);
  return { data, warnings: [...w.warnings] };
}
