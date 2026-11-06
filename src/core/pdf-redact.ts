/**
 * True removal of page content inside rectangles (redaction), working on the content streams:
 *
 * - text: every glyph whose centre lies in a rectangle is removed from its string and replaced by a
 *   TJ adjustment of exactly its advance, so the remaining glyphs keep their positions. Invisible
 *   text (OCR layers) is removed too. Text inside form XObjects is handled on private copies.
 *   Vertical writing (Identity-V, WMode 1) follows the column with /W2 + /DW2 metrics.
 * - in `redact` mode also: vector paths lying entirely inside a rectangle, inline images centred in
 *   one, the pixels of image XObjects under a rectangle (samples zeroed, with their soft and stencil
 *   masks), and annotations overlapping. Image data is decoded through filter chains (pdf-lib's
 *   filters, PNG/TIFF predictors) and the codecs (DCT, JPX, JBIG2, CCITT via `decodeImage`), and
 *   rewritten losslessly; alternates, OPI and XMP of such images, and the page's thumbnail and
 *   PieceInfo are dropped.
 *
 * Glyph advances come from the font: /Widths (simple), /W + /DW (CID), /Widths × /FontMatrix (Type 3),
 * or the standard-14 metrics for base fonts without widths.
 */
import { Encodings, Font, FontNames } from '@pdf-lib/standard-fonts';
import { decodePDFRawStream, PDFArray, PDFBool, PDFDict, PDFName, PDFNumber, PDFRawStream, PDFRef, PDFStream, type PDFContext, type PDFDocument, type PDFObject, type PDFPage } from 'pdf-lib';
import { zlibSync } from 'fflate';

export interface Rect {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export interface RedactOptions {
  /** Also remove graphics, image pixels and annotations (a redaction), not only text. */
  redact?: boolean;
  /** Decodes JPEG image data to RGBA (browser hook); a fallback when `decodeImage` is not given. */
  decodeJpeg?: (data: Uint8Array) => Promise<{ width: number; height: number; rgba: Uint8Array } | null>;
  /** Decodes DCT, JPX, JBIG2 and CCITT fax image data (see pdf-image-filters.ts). */
  decodeImage?: ImageFilterDecoder;
}

/** What an image filter decoder is told about the image. */
export interface ImageFilterInfo {
  width: number;
  height: number;
  /** Colour components the image's colour space has (0: none given, JPX decides) */
  comps: number;
  indexed: boolean;
  colorTransform?: number;
  smaskInData?: boolean;
  /** JBIG2 global segments */
  globals?: Uint8Array;
  /** CCITTFaxDecode parameters */
  params?: Record<string, number | boolean>;
}

/**
 * Decodes the data of an image filter into samples: `comps` components of `bpc` bits per pixel, rows padded
 * to whole bytes: the filter's output, in the image's own colour space and before its Decode array.
 */
export type ImageFilterDecoder = (
  filter: 'DCTDecode' | 'JPXDecode' | 'JBIG2Decode' | 'CCITTFaxDecode',
  data: Uint8Array,
  info: ImageFilterInfo,
) => Promise<{ data: Uint8Array; comps: number; bpc: number } | null>;

export interface RedactResult {
  glyphs: number;
  paths: number;
  images: number;
  annotations: number;
  warnings: string[];
}

type M = [number, number, number, number, number, number];
const IDENT: M = [1, 0, 0, 1, 0, 0];
/** a then b (row-vector convention of PDF: p' = p × a × b) */
const mul = (a: M, b: M): M => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
const apply = (m: M, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const inside = (rects: Rect[], x: number, y: number): boolean => rects.some((r) => x >= r.x0 && x <= r.x1 && y >= r.y0 && y <= r.y1);

// ---------------------------------------------------------------------------
// Tokenizer

type Operand = { t: 'num'; v: number } | { t: 'name'; v: string } | { t: 'str'; v: Uint8Array } | { t: 'arr'; v: Operand[] } | { t: 'other' };

interface Op {
  op: string;
  args: Operand[];
  /** byte range of the operands and operator */
  start: number;
  end: number;
}

const WS = new Set([0, 9, 10, 12, 13, 32]);
const DELIM = new Set([40, 41, 60, 62, 91, 93, 123, 125, 47, 37]);

function tokenize(b: Uint8Array): Op[] {
  const ops: Op[] = [];
  let i = 0;
  let args: Operand[] = [];
  let argStart = -1;
  const stack: Operand[][] = [];
  const skipWs = (): void => {
    for (;;) {
      while (i < b.length && WS.has(b[i])) i++;
      if (b[i] === 37) while (i < b.length && b[i] !== 10 && b[i] !== 13) i++;
      else break;
    }
  };
  const push = (o: Operand, at: number): void => {
    if (stack.length) stack[stack.length - 1].push(o);
    else {
      if (argStart < 0) argStart = at;
      args.push(o);
    }
  };
  let guard = 0;
  while (guard++ < 50_000_000) {
    skipWs();
    if (i >= b.length) break;
    const at = i;
    const c = b[i];
    if (c === 40) {
      // literal string
      const out: number[] = [];
      let depth = 1;
      i++;
      while (i < b.length && depth) {
        const ch = b[i++];
        if (ch === 92) {
          const e = b[i++];
          if (e === 110) out.push(10);
          else if (e === 114) out.push(13);
          else if (e === 116) out.push(9);
          else if (e === 98) out.push(8);
          else if (e === 102) out.push(12);
          else if (e === 13) {
            if (b[i] === 10) i++;
          } else if (e === 10) {
            /* line continuation */
          } else if (e >= 48 && e <= 55) {
            let v = e - 48;
            for (let k = 0; k < 2 && b[i] >= 48 && b[i] <= 55; k++) v = v * 8 + (b[i++] - 48);
            out.push(v & 0xff);
          } else out.push(e);
        } else {
          if (ch === 40) depth++;
          else if (ch === 41 && --depth === 0) break;
          out.push(ch);
        }
      }
      push({ t: 'str', v: Uint8Array.from(out) }, at);
    } else if (c === 60 && b[i + 1] === 60) {
      // dictionary: skip to the matching >> (contents are not needed)
      let depth = 0;
      while (i < b.length) {
        if (b[i] === 60 && b[i + 1] === 60) {
          depth++;
          i += 2;
        } else if (b[i] === 62 && b[i + 1] === 62) {
          depth--;
          i += 2;
          if (!depth) break;
        } else if (b[i] === 40) {
          // a string inside the dictionary may contain >>
          let d = 1;
          i++;
          while (i < b.length && d) {
            if (b[i] === 92) i++;
            else if (b[i] === 40) d++;
            else if (b[i] === 41) d--;
            i++;
          }
        } else i++;
      }
      push({ t: 'other' }, at);
    } else if (c === 60) {
      i++;
      let hex = '';
      while (i < b.length && b[i] !== 62) {
        if (!WS.has(b[i])) hex += String.fromCharCode(b[i]);
        i++;
      }
      i++;
      if (hex.length % 2) hex += '0';
      const v = new Uint8Array(hex.length / 2);
      for (let k = 0; k < v.length; k++) v[k] = parseInt(hex.slice(k * 2, k * 2 + 2), 16) || 0;
      push({ t: 'str', v }, at);
    } else if (c === 91) {
      i++;
      if (!stack.length && argStart < 0) argStart = at;
      stack.push([]);
    } else if (c === 93) {
      i++;
      const arr = stack.pop() ?? [];
      push({ t: 'arr', v: arr }, at);
    } else if (c === 47) {
      i++;
      let n = '';
      while (i < b.length && !WS.has(b[i]) && !DELIM.has(b[i])) n += String.fromCharCode(b[i++]);
      push({ t: 'name', v: n.replace(/#([0-9a-f]{2})/gi, (_, h: string) => String.fromCharCode(parseInt(h, 16))) }, at);
    } else if ((c >= 48 && c <= 57) || c === 43 || c === 45 || c === 46) {
      let n = '';
      while (i < b.length && ((b[i] >= 48 && b[i] <= 57) || b[i] === 43 || b[i] === 45 || b[i] === 46)) n += String.fromCharCode(b[i++]);
      push({ t: 'num', v: parseFloat(n) || 0 }, at);
    } else if (c === 41 || c === 62 || c === 123 || c === 125) {
      i++; // stray delimiter
    } else {
      let word = '';
      while (i < b.length && !WS.has(b[i]) && !DELIM.has(b[i])) word += String.fromCharCode(b[i++]);
      if (word === 'true' || word === 'false' || word === 'null') {
        push({ t: 'other' }, at);
        continue;
      }
      if (stack.length) {
        // an operator inside an array is malformed: close the arrays
        while (stack.length) push({ t: 'arr', v: stack.pop()! }, at);
      }
      if (word === 'BI') {
        // inline image: dictionary up to ID, then binary data up to EI
        const id = indexOf(b, [73, 68], i);
        let p = id < 0 ? b.length : id + 2;
        if (WS.has(b[p])) p++;
        let end = p;
        for (;;) {
          end = indexOf(b, [69, 73], end);
          if (end < 0) {
            end = b.length;
            break;
          }
          if (WS.has(b[end - 1]) && (end + 2 >= b.length || WS.has(b[end + 2]) || DELIM.has(b[end + 2]))) break;
          end += 2;
        }
        i = Math.min(b.length, end + 2);
        ops.push({ op: 'BI', args: [], start: argStart >= 0 ? argStart : at, end: i });
      } else ops.push({ op: word, args, start: argStart >= 0 ? argStart : at, end: i });
      args = [];
      argStart = -1;
    }
  }
  return ops;
}

function indexOf(b: Uint8Array, pat: number[], from: number): number {
  outer: for (let i = from; i <= b.length - pat.length; i++) {
    for (let k = 0; k < pat.length; k++) if (b[i + k] !== pat[k]) continue outer;
    return i;
  }
  return -1;
}

// ---------------------------------------------------------------------------
// Fonts

interface FontInfo {
  /** bytes of the code starting at i */
  codeLen(s: Uint8Array, i: number): number;
  /** advance in text space units per unit font size */
  width(code: number, len: number): number;
  vertical: boolean;
  /** vertical writing: displacement w1 (negative: down) and position vector v of a glyph, per unit font size */
  vmetric(code: number, len: number): { w1: number; vx: number; vy: number };
}

const WIN_ANSI: string[] = (() => {
  const names: string[] = [];
  for (const cp of Encodings.WinAnsi.supportedCodePoints) {
    const { code, name } = Encodings.WinAnsi.encodeUnicodeCodePoint(cp);
    names[code] ??= name;
  }
  return names;
})();
/** StandardEncoding differs from WinAnsi mainly in the quotes of the ASCII range. */
const STANDARD = WIN_ANSI.map((n, c) => (c === 39 ? 'quoteright' : c === 96 ? 'quoteleft' : c < 128 ? n : undefined));

function standardFont(base: string): Font | null {
  const n = base.replace(/^[A-Z]{6}\+/, '').replace(/[,-]/g, '').toLowerCase();
  const bold = /bold|black|heavy|semibold/.test(n);
  const italic = /italic|oblique/.test(n);
  const pick = (r: FontNames, b: FontNames, i: FontNames, bi: FontNames): FontNames => (bold && italic ? bi : bold ? b : italic ? i : r);
  let name: FontNames | null = null;
  if (/^(helvetica|arial)/.test(n)) name = pick(FontNames.Helvetica, FontNames.HelveticaBold, FontNames.HelveticaOblique, FontNames.HelveticaBoldOblique);
  else if (/^(times|timesnewroman)/.test(n)) name = pick(FontNames.TimesRoman, FontNames.TimesRomanBold, FontNames.TimesRomanItalic, FontNames.TimesRomanBoldItalic);
  else if (/^courier/.test(n)) name = pick(FontNames.Courier, FontNames.CourierBold, FontNames.CourierOblique, FontNames.CourierBoldOblique);
  else if (/^symbol/.test(n)) name = FontNames.Symbol;
  else if (/^zapfdingbats/.test(n)) name = FontNames.ZapfDingbats;
  return name ? Font.load(name) : null;
}

const num = (o: PDFObject | undefined): number | undefined => (o instanceof PDFNumber ? o.asNumber() : undefined);

class Fonts {
  private cache = new Map<PDFDict, FontInfo>();
  constructor(private ctx: PDFContext) {}

  private look<T extends PDFObject>(o: PDFObject | undefined, type: { prototype: T }): T | undefined {
    const v = o instanceof PDFRef ? this.ctx.lookup(o) : o;
    return v instanceof (type as unknown as abstract new () => T) ? v : undefined;
  }

  get(dict: PDFDict): FontInfo {
    let f = this.cache.get(dict);
    if (!f) {
      f = this.build(dict);
      this.cache.set(dict, f);
    }
    return f;
  }

  private build(d: PDFDict): FontInfo {
    const sub = d.lookup(PDFName.of('Subtype'))?.toString();
    const L = (k: string): PDFObject | undefined => d.lookup(PDFName.of(k));
    if (sub === '/Type0') {
      const enc = L('Encoding');
      const desc = this.look(this.look(L('DescendantFonts'), PDFArray)?.get(0), PDFDict);
      const dw = num(desc?.lookup(PDFName.of('DW'))) ?? 1000;
      const widths = new Map<number, number>();
      const w = this.look(desc?.lookup(PDFName.of('W')), PDFArray);
      if (w) {
        for (let i = 0; i < w.size(); ) {
          const first = num(w.lookup(i));
          const next = w.lookup(i + 1);
          if (first === undefined) break;
          if (next instanceof PDFArray) {
            for (let k = 0; k < next.size(); k++) widths.set(first + k, num(next.lookup(k)) ?? dw);
            i += 2;
          } else {
            const last = num(next) ?? first;
            const val = num(w.lookup(i + 2)) ?? dw;
            for (let c = first; c <= last && c - first < 65536; c++) widths.set(c, val);
            i += 3;
          }
        }
      }
      // Vertical metrics: W2 entries (w1y vx vy per CID) and the DW2 default [vy w1y]
      const dw2 = this.look(desc?.lookup(PDFName.of('DW2')), PDFArray);
      const dvy = num(dw2?.lookup(0)) ?? 880;
      const dw1 = num(dw2?.lookup(1)) ?? -1000;
      const vm = new Map<number, [number, number, number]>();
      const w2 = this.look(desc?.lookup(PDFName.of('W2')), PDFArray);
      if (w2) {
        for (let i = 0; i < w2.size(); ) {
          const first = num(w2.lookup(i));
          const next = w2.lookup(i + 1);
          if (first === undefined) break;
          if (next instanceof PDFArray) {
            for (let k = 0; k + 2 < next.size(); k += 3) vm.set(first + k / 3, [num(next.lookup(k)) ?? dw1, num(next.lookup(k + 1)) ?? 0, num(next.lookup(k + 2)) ?? dvy]);
            i += 2;
          } else {
            const last = num(next) ?? first;
            const t: [number, number, number] = [num(w2.lookup(i + 2)) ?? dw1, num(w2.lookup(i + 3)) ?? 0, num(w2.lookup(i + 4)) ?? dvy];
            for (let c = first; c <= last && c - first < 65536; c++) vm.set(c, t);
            i += 5;
          }
        }
      }
      // Code → CID: Identity, or the cidrange/cidchar maps of an embedded CMap
      let ranges: { lo: number; hi: number; len: number }[] = [{ lo: 0, hi: 0xffff, len: 2 }];
      let cidOf: (code: number) => number = (c) => c;
      let vertical = false;
      if (enc instanceof PDFName) vertical = /-V$/.test(enc.asString());
      else {
        const cm = this.look(enc, PDFStream);
        if (cm) {
          const text = new TextDecoder('latin1').decode(streamBytes(cm));
          vertical = /\/WMode\s+1/.test(text);
          const cs: { lo: number; hi: number; len: number }[] = [];
          for (const block of text.matchAll(/begincodespacerange([\s\S]*?)endcodespacerange/g))
            for (const m of block[1].matchAll(/<([0-9a-f]+)>\s*<([0-9a-f]+)>/gi)) cs.push({ lo: parseInt(m[1], 16), hi: parseInt(m[2], 16), len: m[1].length / 2 });
          if (cs.length) ranges = cs;
          const map = new Map<number, number>();
          const rangesCid: [number, number, number][] = [];
          for (const block of text.matchAll(/begincidrange([\s\S]*?)endcidrange/g))
            for (const m of block[1].matchAll(/<([0-9a-f]+)>\s*<([0-9a-f]+)>\s*(\d+)/gi)) rangesCid.push([parseInt(m[1], 16), parseInt(m[2], 16), +m[3]]);
          for (const block of text.matchAll(/begincidchar([\s\S]*?)endcidchar/g)) for (const m of block[1].matchAll(/<([0-9a-f]+)>\s*(\d+)/gi)) map.set(parseInt(m[1], 16), +m[2]);
          if (map.size || rangesCid.length)
            cidOf = (c) => {
              const v = map.get(c);
              if (v !== undefined) return v;
              const r = rangesCid.find(([lo, hi]) => c >= lo && c <= hi);
              return r ? r[2] + c - r[0] : 0;
            };
        }
      }
      return {
        vertical,
        codeLen: (s, i) => {
          for (let len = 1; len <= 4; len++) {
            if (i + len > s.length) break;
            let v = 0;
            for (let k = 0; k < len; k++) v = (v << 8) | s[i + k];
            if (ranges.some((r) => r.len === len && v >= r.lo && v <= r.hi)) return len;
          }
          return Math.min(ranges[0]?.len ?? 2, s.length - i) || 1;
        },
        width: (code) => (widths.get(cidOf(code)) ?? dw) / 1000,
        vmetric: (code) => {
          const cid = cidOf(code);
          const m = vm.get(cid);
          // default position vector: half the horizontal width, DW2's vy
          return m ? { w1: m[0] / 1000, vx: m[1] / 1000, vy: m[2] / 1000 } : { w1: dw1 / 1000, vx: (widths.get(cid) ?? dw) / 2000, vy: dvy / 1000 };
        },
      };
    }
    // Simple fonts (Type1, TrueType, Type3, MMType1)
    const first = num(L('FirstChar')) ?? 0;
    const wArr = this.look(L('Widths'), PDFArray);
    const fd = this.look(L('FontDescriptor'), PDFDict);
    const missing = num(fd?.lookup(PDFName.of('MissingWidth'))) ?? 0;
    let scale = 1 / 1000;
    if (sub === '/Type3') {
      const fm = this.look(L('FontMatrix'), PDFArray);
      scale = num(fm?.lookup(0)) ?? 0.001;
    }
    const widths: number[] = [];
    if (wArr) for (let k = 0; k < wArr.size(); k++) widths[first + k] = num(wArr.lookup(k)) ?? missing;
    let std: Font | null = null;
    let names: (string | undefined)[] = [];
    if (!wArr) {
      std = standardFont(L('BaseFont')?.toString().slice(1) ?? '');
      const enc = L('Encoding');
      const encDict = this.look(enc, PDFDict);
      const base = (enc instanceof PDFName ? enc : encDict?.lookup(PDFName.of('BaseEncoding')))?.toString();
      names = base === '/WinAnsiEncoding' || base === '/MacRomanEncoding' ? [...WIN_ANSI] : [...STANDARD];
      const diff = this.look(encDict?.lookup(PDFName.of('Differences')), PDFArray);
      if (diff) {
        let code = 0;
        for (let k = 0; k < diff.size(); k++) {
          const v = diff.lookup(k);
          if (v instanceof PDFNumber) code = v.asNumber();
          else if (v instanceof PDFName) names[code++] = v.asString().slice(1);
        }
      }
    }
    return {
      vertical: false,
      vmetric: () => ({ w1: -1, vx: 0.5, vy: 0.88 }),
      codeLen: () => 1,
      width: (code) => {
        if (wArr) return (widths[code] ?? missing) * scale;
        const name = names[code];
        const w = std && name ? std.getWidthOfGlyph(name) : undefined;
        return (typeof w === 'number' ? w : missing || 500) / 1000;
      },
    };
  }
}

function streamBytes(s: PDFStream): Uint8Array {
  if (s instanceof PDFRawStream) {
    try {
      return decodePDFRawStream(s).decode();
    } catch {
      return new Uint8Array(0);
    }
  }
  const any = s as unknown as { getUnencodedContents?: () => Uint8Array };
  return any.getUnencodedContents ? any.getUnencodedContents() : s.getContents();
}

// ---------------------------------------------------------------------------
// Interpreter

interface GState {
  ctm: M;
  fontDict?: PDFDict;
  fs: number;
  tc: number;
  tw: number;
  th: number;
  tl: number;
  rise: number;
}

const ascii = (s: string): Uint8Array => Uint8Array.from(s, (c) => c.charCodeAt(0) & 0xff);
const fmt = (n: number): string => {
  const r = Math.round(n * 1000) / 1000;
  return Object.is(r, -0) ? '0' : String(r);
};
const hex = (b: Uint8Array): string => '<' + Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('') + '>';

class Redactor {
  result: RedactResult = { glyphs: 0, paths: 0, images: 0, annotations: 0, warnings: [] };
  private fonts: Fonts;
  private warned = new Set<string>();
  /** image streams already redacted (shared between uses) */
  private imageDone = new Map<PDFRawStream, PDFRef>();

  constructor(
    private ctx: PDFContext,
    private rects: Rect[],
    private opts: RedactOptions,
  ) {
    this.fonts = new Fonts(ctx);
  }

  private warn(s: string): void {
    if (!this.warned.has(s)) {
      this.warned.add(s);
      this.result.warnings.push(s);
    }
  }

  private dict(o: PDFObject | undefined): PDFDict | undefined {
    const v = o instanceof PDFRef ? this.ctx.lookup(o) : o;
    return v instanceof PDFDict ? v : undefined;
  }

  /**
   * Process one content stream. Returns the new bytes (or null if unchanged) and resource
   * replacements (XObject name → new object) made for forms and images.
   */
  async run(bytes: Uint8Array, res: PDFDict | undefined, ctm: M, depth = 0): Promise<{ bytes: Uint8Array | null; xobjects: Map<string, PDFRef> }> {
    const ops = tokenize(bytes);
    const edits: { start: number; end: number; text: string }[] = [];
    const xobjects = new Map<string, PDFRef>();
    const fontRes = this.dict(res?.lookup(PDFName.of('Font')));
    const xres = this.dict(res?.lookup(PDFName.of('XObject')));
    const stack: GState[] = [];
    let g: GState = { ctm, fs: 0, tc: 0, tw: 0, th: 1, tl: 0, rise: 0 };
    let tm: M = IDENT;
    let tlm: M = IDENT;
    // path tracking
    let pathStart = -1;
    let pathPts: [number, number][] = [];
    let clipped = false;
    let cur: [number, number] = [0, 0];
    // Marked content: a span that loses glyphs loses its properties too (ActualText/Alt repeat the text)
    const marked: { start: number; end: number; tag: string; props: boolean; removed: boolean }[] = [];

    const show = (op: Op, items: Operand[], prefix: string): void => {
      const font = g.fontDict ? this.fonts.get(g.fontDict) : undefined;
      if (!font) return;
      const fs = g.fs;
      const out: (number | Uint8Array)[] = [];
      let removed = 0;
      for (const it of items) {
        if (it.t === 'num') {
          // vertical writing moves down the column, without horizontal scaling
          tm = font.vertical ? mul([1, 0, 0, 1, 0, (-it.v / 1000) * fs], tm) : mul([1, 0, 0, 1, (-it.v / 1000) * fs * g.th, 0], tm);
          out.push(it.v);
          continue;
        }
        if (it.t !== 'str') continue;
        const s = it.v;
        let keep: number[] = [];
        for (let i = 0; i < s.length; ) {
          const len = font.codeLen(s, i);
          let code = 0;
          for (let k = 0; k < len; k++) code = (code << 8) | s[i + k];
          const w = font.width(code, len);
          const space = len === 1 && code === 32 ? g.tw : 0;
          const trm = mul([fs * g.th, 0, 0, fs, 0, g.rise], mul(tm, g.ctm));
          if (font.vertical) {
            // the glyph hangs from the current point: its horizontal origin is at -v; centre of the em box
            const v = font.vmetric(code, len);
            const [cx, cy] = apply(trm, w / 2 - v.vx, 0.38 - v.vy);
            const ty = v.w1 * fs + g.tc + space;
            if (fs !== 0 && inside(this.rects, cx, cy)) {
              if (keep.length) out.push(Uint8Array.from(keep));
              keep = [];
              out.push(-(ty * 1000) / fs);
              removed++;
            } else for (let k = 0; k < len; k++) keep.push(s[i + k]);
            tm = mul([1, 0, 0, 1, 0, ty], tm);
            i += len;
            continue;
          }
          // glyph centre in user space
          const [cx, cy] = apply(trm, w / 2, 0.3);
          const advance = (w * fs + g.tc + space) * g.th;
          if (fs !== 0 && inside(this.rects, cx, cy)) {
            if (keep.length) out.push(Uint8Array.from(keep));
            keep = [];
            out.push(-((w * fs + g.tc + space) * 1000) / fs);
            removed++;
          } else for (let k = 0; k < len; k++) keep.push(s[i + k]);
          tm = mul([1, 0, 0, 1, advance, 0], tm);
          i += len;
        }
        if (keep.length) out.push(Uint8Array.from(keep));
      }
      if (!removed) return;
      for (const m of marked) m.removed = true;
      this.result.glyphs += removed;
      // merge adjacent numbers
      const merged: (number | Uint8Array)[] = [];
      for (const x of out) {
        const last = merged[merged.length - 1];
        if (typeof x === 'number' && typeof last === 'number') merged[merged.length - 1] = last + x;
        else merged.push(x);
      }
      const arr = merged.map((x) => (typeof x === 'number' ? fmt(x) : hex(x))).join(' ');
      edits.push({ start: op.start, end: op.end, text: `${prefix}[${arr}] TJ` });
    };

    for (let k = 0; k < ops.length; k++) {
      const o = ops[k];
      const a = o.args;
      const n = (i: number): number => (a[i]?.t === 'num' ? (a[i] as { v: number }).v : 0);
      switch (o.op) {
        case 'q':
          stack.push({ ...g });
          break;
        case 'Q':
          g = stack.pop() ?? g;
          break;
        case 'cm':
          g.ctm = mul([n(0), n(1), n(2), n(3), n(4), n(5)], g.ctm);
          break;
        case 'BDC':
        case 'BMC':
          marked.push({ start: o.start, end: o.end, tag: a[0]?.t === 'name' ? a[0].v : 'Span', props: o.op === 'BDC', removed: false });
          break;
        case 'EMC': {
          const m = marked.pop();
          if (m?.removed && m.props) edits.push({ start: m.start, end: m.end, text: `/${m.tag.replace(/[^\w.-]/g, '')} BMC` });
          break;
        }
        case 'BT':
          tm = IDENT;
          tlm = IDENT;
          break;
        case 'Tf': {
          const name = a[0]?.t === 'name' ? a[0].v : '';
          g.fontDict = this.dict(fontRes?.lookup(PDFName.of(name)));
          g.fs = n(1);
          break;
        }
        case 'Tc':
          g.tc = n(0);
          break;
        case 'Tw':
          g.tw = n(0);
          break;
        case 'Tz':
          g.th = n(0) / 100;
          break;
        case 'TL':
          g.tl = n(0);
          break;
        case 'Ts':
          g.rise = n(0);
          break;
        case 'Td':
          tlm = mul([1, 0, 0, 1, n(0), n(1)], tlm);
          tm = tlm;
          break;
        case 'TD':
          g.tl = -n(1);
          tlm = mul([1, 0, 0, 1, n(0), n(1)], tlm);
          tm = tlm;
          break;
        case 'Tm':
          tlm = [n(0), n(1), n(2), n(3), n(4), n(5)];
          tm = tlm;
          break;
        case 'T*':
          tlm = mul([1, 0, 0, 1, 0, -g.tl], tlm);
          tm = tlm;
          break;
        case 'Tj':
          show(o, a.slice(0, 1), '');
          break;
        case 'TJ':
          show(o, a[0]?.t === 'arr' ? a[0].v : [], '');
          break;
        case "'":
          tlm = mul([1, 0, 0, 1, 0, -g.tl], tlm);
          tm = tlm;
          show(o, a.slice(0, 1), 'T* ');
          break;
        case '"':
          g.tw = n(0);
          g.tc = n(1);
          tlm = mul([1, 0, 0, 1, 0, -g.tl], tlm);
          tm = tlm;
          show(o, a.slice(2, 3), `${fmt(n(0))} Tw ${fmt(n(1))} Tc T* `);
          break;
        // ---- paths (redact mode)
        case 'm':
        case 'l':
        case 're':
        case 'c':
        case 'v':
        case 'y':
        case 'h':
          if (pathStart < 0) {
            pathStart = o.start;
            pathPts = [];
            clipped = false;
          }
          if (o.op === 're') {
            const [x, y, w, h] = [n(0), n(1), n(2), n(3)];
            for (const [px, py] of [
              [x, y],
              [x + w, y],
              [x + w, y + h],
              [x, y + h],
            ])
              pathPts.push(apply(g.ctm, px, py));
            cur = [x, y];
          } else if (o.op !== 'h') {
            for (let i = 0; i + 1 < a.length; i += 2) pathPts.push(apply(g.ctm, n(i), n(i + 1)));
            if (a.length >= 2) cur = [n(a.length - 2), n(a.length - 1)];
          }
          void cur;
          break;
        case 'W':
        case 'W*':
          clipped = true;
          break;
        case 'S':
        case 's':
        case 'f':
        case 'F':
        case 'f*':
        case 'B':
        case 'B*':
        case 'b':
        case 'b*':
        case 'n':
          if (this.opts.redact && pathStart >= 0 && !clipped && o.op !== 'n' && pathPts.length && pathPts.every(([x, y]) => inside(this.rects, x, y))) {
            edits.push({ start: pathStart, end: o.end, text: '' });
            this.result.paths++;
          }
          pathStart = -1;
          break;
        case 'BI': {
          if (!this.opts.redact) break;
          const [cx, cy] = apply(g.ctm, 0.5, 0.5);
          if (inside(this.rects, cx, cy)) {
            edits.push({ start: o.start, end: o.end, text: '' });
            this.result.images++;
          }
          break;
        }
        case 'Do': {
          const name = a[0]?.t === 'name' ? a[0].v : '';
          const ref = xres?.get(PDFName.of(name));
          const xo = ref instanceof PDFRef ? this.ctx.lookup(ref) : ref;
          if (!(xo instanceof PDFStream)) break;
          const subtype = xo.dict.lookup(PDFName.of('Subtype'))?.toString();
          if (subtype === '/Form' && depth < 12) {
            const fm = xo.dict.lookup(PDFName.of('Matrix'));
            const m: M = fm instanceof PDFArray && fm.size() === 6 ? (Array.from({ length: 6 }, (_, i) => num(fm.lookup(i)) ?? 0) as M) : IDENT;
            const fres = this.dict(xo.dict.lookup(PDFName.of('Resources'))) ?? res;
            const r = await this.run(streamBytes(xo), fres, mul(m, g.ctm), depth + 1);
            if (r.bytes || r.xobjects.size) {
              // a private copy: the form may be shared by other pages
              const dict = xo.dict.clone(this.ctx);
              dict.delete(PDFName.of('Filter'));
              dict.delete(PDFName.of('DecodeParms'));
              dict.delete(PDFName.of('Length'));
              // application data and XMP of the form may hold the original content
              for (const k of ['PieceInfo', 'Metadata', 'LastModified']) dict.delete(PDFName.of(k));
              if (r.xobjects.size) dict.set(PDFName.of('Resources'), withXObjects(this.ctx, fres, r.xobjects));
              const copy = this.ctx.flateStream(r.bytes ?? streamBytes(xo), {});
              for (const [key, val] of dict.entries()) copy.dict.set(key, val);
              xobjects.set(name, this.ctx.register(copy));
            }
          } else if (subtype === '/Image' && this.opts.redact && xo instanceof PDFRawStream) {
            const newRef = await this.image(xo, g.ctm, false);
            if (newRef) xobjects.set(name, newRef);
          }
          break;
        }
      }
    }
    if (!edits.length) return { bytes: null, xobjects };
    // apply edits (non-overlapping, in order)
    edits.sort((x, y) => x.start - y.start);
    const parts: Uint8Array[] = [];
    let pos = 0;
    for (const e of edits) {
      if (e.start < pos) continue;
      parts.push(bytes.subarray(pos, e.start), ascii(e.text));
      pos = e.end;
    }
    parts.push(bytes.subarray(pos));
    const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
    let o = 0;
    for (const p of parts) {
      out.set(p, o);
      o += p.length;
    }
    return { bytes: out, xobjects };
  }

  /**
   * Zero the samples of an image under the rectangles; returns the new image or null.
   * Its soft mask and stencil mask (the shape of what it shows) are redacted the same way.
   */
  private async image(img: PDFRawStream, ctm: M, isMaskOf: boolean): Promise<PDFRef | null> {
    // image space: unit square; pixel (0,0) at the top-left → (0,1) in unit space
    const d = img.dict;
    const w = num(d.lookup(PDFName.of('Width'))) ?? 0;
    const h = num(d.lookup(PDFName.of('Height'))) ?? 0;
    if (!w || !h) return null;
    const inv = invert(ctm);
    if (!inv) return null;
    const boxes: [number, number, number, number][] = [];
    for (const r of this.rects) {
      const pts = [apply(inv, r.x0, r.y0), apply(inv, r.x1, r.y0), apply(inv, r.x1, r.y1), apply(inv, r.x0, r.y1)];
      const ux0 = Math.max(0, Math.min(...pts.map((p) => p[0])));
      const ux1 = Math.min(1, Math.max(...pts.map((p) => p[0])));
      const uy0 = Math.max(0, Math.min(...pts.map((p) => p[1])));
      const uy1 = Math.min(1, Math.max(...pts.map((p) => p[1])));
      if (ux1 <= ux0 || uy1 <= uy0) continue;
      boxes.push([Math.floor(ux0 * w), Math.floor((1 - uy1) * h), Math.ceil(ux1 * w), Math.ceil((1 - uy0) * h)]);
    }
    if (!boxes.length) return null;
    if (this.imageDone.has(img)) return this.imageDone.get(img)!;
    const dec = await this.samplesOf(img, w, h);
    if (!dec) {
      this.warn('Некоторые изображения под областью скрытия не удалось изменить (формат сжатия): они закрашены, но их данные остались в файле.');
      return null;
    }
    const { comps, bpc } = dec;
    const samples = dec.samples.constructor === Uint8Array ? dec.samples : new Uint8Array(dec.samples.buffer, dec.samples.byteOffset, dec.samples.byteLength);
    const stride = Math.ceil((w * comps * bpc) / 8);
    if (samples.length < stride * h) return null;
    // the black box colour: zero samples (black in Gray/RGB; mask/Indexed/CMYK still lose their content)
    const zero = (buf: Uint8Array, n: number, bits: number, rowBytes: number): void => {
      for (const [x0, y0, x1, y1] of boxes)
        for (let y = Math.max(0, y0); y < Math.min(h, y1); y++)
          for (let x = Math.max(0, x0); x < Math.min(w, x1); x++) {
            const bit0 = x * n * bits;
            for (let b = bit0; b < bit0 + n * bits; b++) buf[y * rowBytes + (b >> 3)] &= ~(0x80 >> (b & 7));
          }
    };
    zero(samples, comps, bpc, stride);
    const dict = d.clone(this.ctx);
    // filters are undone; alternates, OPI proxies and XMP (thumbnails) would keep the original
    for (const k of ['Filter', 'DecodeParms', 'Length', 'Alternates', 'OPI', 'Metadata', 'SMaskInData']) dict.delete(PDFName.of(k));
    if (dec.colorSpace) dict.set(PDFName.of('ColorSpace'), dec.colorSpace);
    dict.set(PDFName.of('BitsPerComponent'), PDFNumber.of(bpc));
    if (dec.dropDecode) dict.delete(PDFName.of('Decode'));
    const stream = this.ctx.flateStream(samples, {});
    for (const [k, v] of dict.entries()) stream.dict.set(k, v);
    // the alpha channel of a JPX image becomes its soft mask
    if (dec.alpha) {
      zero(dec.alpha, 1, 8, w);
      const sm = this.ctx.flateStream(dec.alpha, { Type: 'XObject', Subtype: 'Image', Width: w, Height: h, ColorSpace: 'DeviceGray', BitsPerComponent: 8 });
      stream.dict.set(PDFName.of('SMask'), this.ctx.register(sm));
    }
    for (const key of ['SMask', 'Mask']) {
      const m = d.lookup(PDFName.of(key));
      if (m instanceof PDFRawStream) {
        const ref = await this.image(m, ctm, true);
        if (ref) stream.dict.set(PDFName.of(key), ref);
      }
    }
    const ref = this.ctx.register(stream);
    this.imageDone.set(img, ref);
    if (!isMaskOf) this.result.images++;
    return ref;
  }

  /** The samples of an image: its filters undone (pdf-lib's, PNG/TIFF predictors, then the image codecs). */
  private async samplesOf(
    img: PDFRawStream,
    w: number,
    h: number,
  ): Promise<{ samples: Uint8Array; comps: number; bpc: number; colorSpace?: PDFObject; dropDecode?: boolean; alpha?: Uint8Array } | null> {
    const d = img.dict;
    const filters = filterNames(d.lookup(PDFName.of('Filter')));
    const parmsRaw = d.lookup(PDFName.of('DecodeParms'));
    const parms = filters.map((_, i) => this.dict(parmsRaw instanceof PDFArray ? parmsRaw.lookup(i) : i === 0 ? parmsRaw : undefined));
    const isMask = d.lookup(PDFName.of('ImageMask'))?.toString() === 'true';
    const bpc = isMask ? 1 : (num(d.lookup(PDFName.of('BitsPerComponent'))) ?? 8);
    const csObj = d.get(PDFName.of('ColorSpace'));
    const cs = csObj instanceof PDFRef ? this.ctx.lookup(csObj) : csObj;
    const comps = isMask ? 1 : components(this.ctx, csObj);
    const indexed = cs instanceof PDFArray && cs.lookup(0)?.toString() === '/Indexed';
    const CODECS = ['DCTDecode', 'JPXDecode', 'JBIG2Decode', 'CCITTFaxDecode'] as const;
    type Codec = (typeof CODECS)[number];
    const last = filters[filters.length - 1];
    const codec = (CODECS as readonly string[]).includes(last) ? (last as Codec) : undefined;
    const generic = codec ? filters.slice(0, -1) : filters;
    if (generic.some((f) => (CODECS as readonly string[]).includes(f))) return null;
    let data: Uint8Array | null = img.getContents();
    for (let i = 0; i < generic.length && data; i++) {
      try {
        data = decodePDFRawStream(PDFRawStream.of(this.ctx.obj({ Filter: generic[i] }), data)).decode();
      } catch {
        return null;
      }
      // predictors belong to Flate/LZW; for the last filter they describe the image rows
      if (data && (generic[i] === 'FlateDecode' || generic[i] === 'LZWDecode')) data = unpredict(data, parms[i], comps, bpc, w);
    }
    if (!data) return null;
    if (!codec) return { samples: data, comps, bpc, colorSpace: csObj };
    const cparms = parms[filters.length - 1];
    const dec = this.opts.decodeImage;
    if (codec === 'DCTDecode') {
      if (dec) {
        const r = await dec('DCTDecode', data, { width: w, height: h, comps, indexed, colorTransform: num(cparms?.lookup(PDFName.of('ColorTransform'))) }).catch(() => null);
        if (r && r.comps === comps) return { samples: r.data, comps, bpc: 8, colorSpace: csObj };
        if (r) return { samples: r.data, comps: r.comps, bpc: 8, colorSpace: PDFName.of(r.comps === 1 ? 'DeviceGray' : r.comps === 4 ? 'DeviceCMYK' : 'DeviceRGB'), dropDecode: true };
      }
      if (!this.opts.decodeJpeg) return null;
      const r = await this.opts.decodeJpeg(data).catch(() => null);
      if (!r || r.width !== w || r.height !== h) return null;
      const rgb = new Uint8Array(w * h * 3);
      for (let p = 0; p < w * h; p++) rgb.set(r.rgba.subarray(p * 4, p * 4 + 3), p * 3);
      return { samples: rgb, comps: 3, bpc: 8, colorSpace: PDFName.of('DeviceRGB'), dropDecode: true };
    }
    if (!dec) return null;
    if (codec === 'JPXDecode') {
      // the colour space of the dictionary, or the one inside the JPEG 2000 data (then RGBA)
      const smaskInData = (num(d.lookup(PDFName.of('SMaskInData'))) ?? 0) >= 1;
      const r = await dec('JPXDecode', data, { width: w, height: h, comps: csObj && !smaskInData ? comps : 0, indexed, smaskInData }).catch(() => null);
      if (!r) return null;
      if (csObj && !smaskInData && r.comps === comps) return { samples: r.data, comps, bpc: 8, colorSpace: csObj, dropDecode: true };
      if (r.comps === 4 && (!csObj || smaskInData)) {
        const rgb = new Uint8Array(w * h * 3);
        const alpha = new Uint8Array(w * h);
        let opaque = true;
        for (let p = 0; p < w * h; p++) {
          rgb[p * 3] = r.data[p * 4];
          rgb[p * 3 + 1] = r.data[p * 4 + 1];
          rgb[p * 3 + 2] = r.data[p * 4 + 2];
          alpha[p] = r.data[p * 4 + 3];
          if (alpha[p] !== 255) opaque = false;
        }
        return { samples: rgb, comps: 3, bpc: 8, colorSpace: PDFName.of('DeviceRGB'), dropDecode: true, ...(opaque ? {} : { alpha }) };
      }
      if (r.comps === 1 || r.comps === 3) return { samples: r.data, comps: r.comps, bpc: 8, colorSpace: PDFName.of(r.comps === 1 ? 'DeviceGray' : 'DeviceRGB'), dropDecode: true };
      return null;
    }
    if (codec === 'JBIG2Decode') {
      const g = cparms?.lookup(PDFName.of('JBIG2Globals'));
      const globals = g instanceof PDFStream ? streamBytes(g) : undefined;
      const r = await dec('JBIG2Decode', data, { width: w, height: h, comps: 1, indexed: false, globals }).catch(() => null);
      return r ? { samples: r.data, comps: 1, bpc: 1, colorSpace: csObj } : null;
    }
    const params: Record<string, number | boolean> = {};
    for (const k of ['K', 'EndOfLine', 'EncodedByteAlign', 'Columns', 'Rows', 'EndOfBlock', 'BlackIs1']) {
      const v = cparms?.lookup(PDFName.of(k));
      if (v instanceof PDFNumber) params[k] = v.asNumber();
      else if (v instanceof PDFBool) params[k] = v.asBoolean();
    }
    const r = await dec('CCITTFaxDecode', data, { width: w, height: h, comps: 1, indexed: false, params }).catch(() => null);
    return r ? { samples: r.data, comps: 1, bpc: 1, colorSpace: csObj } : null;
  }
}

function invert(m: M): M | null {
  const det = m[0] * m[3] - m[1] * m[2];
  if (Math.abs(det) < 1e-12) return null;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}

function filterNames(f: PDFObject | undefined): string[] {
  if (f instanceof PDFName) return [f.asString().slice(1)];
  if (f instanceof PDFArray) return f.asArray().map((x) => x.toString().slice(1));
  return [];
}

function components(ctx: PDFContext, cs: PDFObject | undefined): number {
  const v = cs instanceof PDFRef ? ctx.lookup(cs) : cs;
  if (v instanceof PDFName) return v.asString() === '/DeviceCMYK' ? 4 : v.asString() === '/DeviceRGB' || v.asString() === '/CalRGB' ? 3 : 1;
  if (v instanceof PDFArray) {
    const kind = v.lookup(0)?.toString();
    if (kind === '/ICCBased') {
      const s = v.lookup(1);
      const st = s instanceof PDFRef ? ctx.lookup(s) : s;
      return st instanceof PDFStream ? (num(st.dict.lookup(PDFName.of('N'))) ?? 3) : 3;
    }
    if (kind === '/Indexed' || kind === '/Separation' || kind === '/CalGray') return 1;
    if (kind === '/DeviceN') return (v.lookup(1) as PDFArray | undefined)?.size?.() ?? 1;
    if (kind === '/Lab' || kind === '/CalRGB') return 3;
  }
  return 1;
}

/** Undo PNG/TIFF predictors of FlateDecode image data. */
function unpredict(data: Uint8Array, parms: PDFDict | undefined, comps: number, bpc: number, w: number): Uint8Array | null {
  const pred = num(parms?.lookup(PDFName.of('Predictor'))) ?? 1;
  const colors = num(parms?.lookup(PDFName.of('Colors'))) ?? comps;
  const bits = num(parms?.lookup(PDFName.of('BitsPerComponent'))) ?? bpc;
  const cols = num(parms?.lookup(PDFName.of('Columns'))) ?? w;
  if (pred === 2) {
    // TIFF: each sample is the difference to the one to its left (8-bit samples only)
    if (bits !== 8) return null;
    const stride = cols * colors;
    const out = data.slice();
    for (let r = 0; r * stride < out.length; r++) for (let i = r * stride + colors; i < Math.min(out.length, (r + 1) * stride); i++) out[i] = (out[i] + out[i - colors]) & 0xff;
    return out;
  }
  if (pred < 10) return pred === 1 ? data : null;
  const bpp = Math.max(1, Math.ceil((colors * bits) / 8));
  const stride = Math.ceil((cols * colors * bits) / 8);
  const rows = Math.floor(data.length / (stride + 1));
  const out = new Uint8Array(rows * stride);
  for (let r = 0; r < rows; r++) {
    const f = data[r * (stride + 1)];
    const line = data.subarray(r * (stride + 1) + 1, (r + 1) * (stride + 1));
    for (let i = 0; i < stride; i++) {
      const a = i >= bpp ? out[r * stride + i - bpp] : 0;
      const b = r ? out[(r - 1) * stride + i] : 0;
      const c = r && i >= bpp ? out[(r - 1) * stride + i - bpp] : 0;
      let v = line[i];
      if (f === 1) v += a;
      else if (f === 2) v += b;
      else if (f === 3) v += (a + b) >> 1;
      else if (f === 4) {
        const pa = Math.abs(b - c);
        const pb = Math.abs(a - c);
        const pc = Math.abs(a + b - 2 * c);
        v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      }
      out[r * stride + i] = v & 0xff;
    }
  }
  return out;
}

/** A copy of a resource dictionary with some XObjects replaced. */
function withXObjects(ctx: PDFContext, res: PDFDict | undefined, repl: Map<string, PDFRef>): PDFDict {
  const copy = res ? res.clone(ctx) : ctx.obj({});
  const xo = copy.lookup(PDFName.of('XObject'));
  const xdict = xo instanceof PDFDict ? xo.clone(ctx) : ctx.obj({});
  for (const [name, ref] of repl) xdict.set(PDFName.of(name), ref);
  copy.set(PDFName.of('XObject'), xdict);
  return copy;
}

/**
 * Remove content inside `rects` (PDF user space of the page) from a page of `doc`.
 * The page's content streams are rewritten into one stream; shared objects are never modified.
 */
export async function redactPage(doc: PDFDocument, page: PDFPage, rects: Rect[], opts: RedactOptions = {}): Promise<RedactResult> {
  const ctx = doc.context;
  const norm = rects.map((r) => ({ x0: Math.min(r.x0, r.x1), y0: Math.min(r.y0, r.y1), x1: Math.max(r.x0, r.x1), y1: Math.max(r.y0, r.y1) })).filter((r) => r.x1 > r.x0 && r.y1 > r.y0);
  const red = new Redactor(ctx, norm, opts);
  if (!norm.length) return red.result;
  const contents = page.node.Contents();
  const streams: PDFStream[] = [];
  if (contents instanceof PDFStream) streams.push(contents);
  else if (contents instanceof PDFArray)
    for (let i = 0; i < contents.size(); i++) {
      const s = contents.lookup(i);
      if (s instanceof PDFStream) streams.push(s);
    }
  const parts = streams.map(streamBytes);
  const all = new Uint8Array(parts.reduce((s, p) => s + p.length + 1, 0));
  let o = 0;
  for (const p of parts) {
    all.set(p, o);
    all[o + p.length] = 10;
    o += p.length + 1;
  }
  const res = page.node.Resources();
  const r = await red.run(all, res, IDENT);
  if (r.bytes) page.node.set(PDFName.of('Contents'), ctx.register(ctx.flateStream(r.bytes)));
  // a thumbnail, application data (PieceInfo) or XMP of the page would still show what was removed
  if (r.bytes || r.xobjects.size) for (const k of ['Thumb', 'PieceInfo', 'Metadata']) page.node.delete(PDFName.of(k));
  if (r.xobjects.size) page.node.set(PDFName.of('Resources'), withXObjects(ctx, res, r.xobjects));
  if (opts.redact) {
    // annotations (links, comments, form fields) overlapping the areas
    const annots = page.node.Annots();
    if (annots) {
      const keep: PDFObject[] = [];
      for (let i = 0; i < annots.size(); i++) {
        const raw = annots.get(i);
        const a = raw instanceof PDFRef ? ctx.lookup(raw) : raw;
        const rect = a instanceof PDFDict ? a.lookup(PDFName.of('Rect')) : undefined;
        const v = rect instanceof PDFArray ? Array.from({ length: 4 }, (_, k) => num(rect.lookup(k)) ?? 0) : null;
        const hit = v && norm.some((q) => Math.min(v[0], v[2]) < q.x1 && Math.max(v[0], v[2]) > q.x0 && Math.min(v[1], v[3]) < q.y1 && Math.max(v[1], v[3]) > q.y0);
        if (hit) red.result.annotations++;
        else keep.push(raw);
      }
      if (red.result.annotations) page.node.set(PDFName.of('Annots'), ctx.obj(keep));
    }
  }
  return red.result;
}

/** Exposed for tests. */
export const _internal = { tokenize, zlibSync };
