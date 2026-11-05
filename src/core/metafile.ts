/**
 * Windows metafiles (EMF, WMF) – the vector pictures Office documents carry (logos,
 * charts, equations). A GDI interpreter turns the records into a small display list
 * (paths, text, bitmaps) in output units; it is rendered as SVG for web formats and
 * drawn as vectors into PDFs. EMF+ "dual" files are drawn from their GDI records.
 * Clipping regions and raster operations other than copy are not applied.
 */
import { zlibSync } from 'fflate';
import { familyOf, type FontFamily } from './model';
import { bytesToBase64, escapeXml } from './util';

export type MetaCmd =
  | { k: 'path'; d: string; fill?: string; evenOdd: boolean; stroke?: string; width: number; dash?: number[] }
  | {
      k: 'text';
      x: number;
      y: number;
      text: string;
      size: number;
      family: FontFamily;
      bold: boolean;
      italic: boolean;
      underline: boolean;
      color: string;
      angle: number;
      anchor: 'start' | 'middle' | 'end';
      baseline: 'alphabetic' | 'top' | 'bottom';
    }
  | { k: 'image'; x: number; y: number; w: number; h: number; png: Uint8Array };

export interface Metafile {
  /** Size of the drawing in output units (the coordinate space of the commands). */
  width: number;
  height: number;
  /** Intended physical size in points. */
  widthPt: number;
  heightPt: number;
  cmds: MetaCmd[];
}

export const isEmf = (b: Uint8Array): boolean => b.length > 88 && b[0] === 1 && b[1] === 0 && b[2] === 0 && b[3] === 0 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46;
export const isWmf = (b: Uint8Array): boolean =>
  b.length > 22 && ((b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) || ((b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0 && (b[4] === 0 || b[4] === 1) && b[5] === 3));

// ---------------------------------------------------------------------------
// GDI state

type Xf = [number, number, number, number, number, number]; // m11 m12 m21 m22 dx dy
const IDENT: Xf = [1, 0, 0, 1, 0, 0];
const mulXf = (a: Xf, b: Xf): Xf => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];

interface Pen {
  style: number;
  width: number;
  color: string;
}
interface Brush {
  style: number;
  color: string;
}
interface Font {
  height: number;
  escapement: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  face: string;
  charset: number;
}
type GdiObj = { t: 'pen'; v: Pen } | { t: 'brush'; v: Brush } | { t: 'font'; v: Font } | { t: 'other' };

interface DC {
  mapMode: number;
  winOrg: [number, number];
  winExt: [number, number];
  vpOrg: [number, number];
  vpExt: [number, number];
  world: Xf;
  pen: Pen;
  brush: Brush;
  font: Font;
  textColor: string;
  bkColor: string;
  bkMode: number;
  textAlign: number;
  polyFill: number;
  cur: [number, number];
}

const color = (cr: number): string => [cr & 255, (cr >> 8) & 255, (cr >> 16) & 255].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
const STOCK: Record<number, GdiObj> = {
  0: { t: 'brush', v: { style: 0, color: 'FFFFFF' } },
  1: { t: 'brush', v: { style: 0, color: 'C0C0C0' } },
  2: { t: 'brush', v: { style: 0, color: '808080' } },
  3: { t: 'brush', v: { style: 0, color: '404040' } },
  4: { t: 'brush', v: { style: 0, color: '000000' } },
  5: { t: 'brush', v: { style: 1, color: '000000' } },
  6: { t: 'pen', v: { style: 0, width: 0, color: 'FFFFFF' } },
  7: { t: 'pen', v: { style: 0, width: 0, color: '000000' } },
  8: { t: 'pen', v: { style: 5, width: 0, color: '000000' } },
};
const DEFAULT_FONT: Font = { height: -12, escapement: 0, weight: 400, italic: false, underline: false, face: 'Arial', charset: 0 };

class Gdi {
  cmds: MetaCmd[] = [];
  dc: DC;
  private stack: DC[] = [];
  objects: (GdiObj | undefined)[] = [];
  /** Output = device units minus this origin, times `scale`. */
  origin: [number, number] = [0, 0];
  scale = 1;
  /** Device units per millimetre (for fixed map modes). */
  devPerMm: [number, number] = [96 / 25.4, 96 / 25.4];
  private path: string[] | null = null;
  private figureOpen = false;

  constructor() {
    this.dc = {
      mapMode: 1,
      winOrg: [0, 0],
      winExt: [1, 1],
      vpOrg: [0, 0],
      vpExt: [1, 1],
      world: [...IDENT],
      pen: { style: 0, width: 0, color: '000000' },
      brush: { style: 0, color: 'FFFFFF' },
      font: DEFAULT_FONT,
      textColor: '000000',
      bkColor: 'FFFFFF',
      bkMode: 2,
      textAlign: 0,
      polyFill: 1,
      cur: [0, 0],
    };
  }

  save(): void {
    this.stack.push({ ...this.dc, winOrg: [...this.dc.winOrg], winExt: [...this.dc.winExt], vpOrg: [...this.dc.vpOrg], vpExt: [...this.dc.vpExt], world: [...this.dc.world] as Xf, cur: [...this.dc.cur] });
  }

  restore(n: number): void {
    // n < 0: relative to the current level; n > 0: absolute level
    let target = n < 0 ? this.stack.length + n : n - 1;
    target = Math.max(0, Math.min(this.stack.length - 1, target));
    if (!this.stack.length) return;
    this.dc = this.stack[target];
    this.stack.length = target;
  }

  /** Logical → page → device → output. */
  private pageScale(): [number, number] {
    const d = this.dc;
    switch (d.mapMode) {
      case 1: // MM_TEXT
        return [1, 1];
      case 2: // LOMETRIC 0.1 mm
        return [0.1 * this.devPerMm[0], -0.1 * this.devPerMm[1]];
      case 3: // HIMETRIC 0.01 mm
        return [0.01 * this.devPerMm[0], -0.01 * this.devPerMm[1]];
      case 4: // LOENGLISH 0.01 in
        return [0.254 * this.devPerMm[0], -0.254 * this.devPerMm[1]];
      case 5: // HIENGLISH 0.001 in
        return [0.0254 * this.devPerMm[0], -0.0254 * this.devPerMm[1]];
      case 6: // TWIPS
        return [(25.4 / 1440) * this.devPerMm[0], -(25.4 / 1440) * this.devPerMm[1]];
      default: {
        // ISOTROPIC / ANISOTROPIC
        let sx = d.winExt[0] ? d.vpExt[0] / d.winExt[0] : 1;
        let sy = d.winExt[1] ? d.vpExt[1] / d.winExt[1] : 1;
        if (d.mapMode === 7) {
          const s = Math.min(Math.abs(sx), Math.abs(sy));
          sx = Math.sign(sx || 1) * s;
          sy = Math.sign(sy || 1) * s;
        }
        return [sx, sy];
      }
    }
  }

  pt(x: number, y: number): [number, number] {
    const w = this.dc.world;
    const px = x * w[0] + y * w[2] + w[4];
    const py = x * w[1] + y * w[3] + w[5];
    const [sx, sy] = this.pageScale();
    const dx = (px - this.dc.winOrg[0]) * sx + this.dc.vpOrg[0];
    const dy = (py - this.dc.winOrg[1]) * sy + this.dc.vpOrg[1];
    return [(dx - this.origin[0]) * this.scale, (dy - this.origin[1]) * this.scale];
  }

  /** Length in logical units → output units (average of both axes). */
  len(v: number): number {
    const a = this.pt(0, 0);
    const b = this.pt(v, 0);
    const c = this.pt(0, v);
    return (Math.hypot(b[0] - a[0], b[1] - a[1]) + Math.hypot(c[0] - a[0], c[1] - a[1])) / 2;
  }

  private f(n: number): string {
    return (Math.round(n * 100) / 100).toString();
  }

  private emit(d: string, fill: boolean, stroke: boolean): void {
    if (!d) return;
    if (this.path) {
      this.path.push(d);
      return;
    }
    this.paint(d, fill, stroke);
  }

  private paint(d: string, fill: boolean, stroke: boolean): void {
    const { pen, brush } = this.dc;
    const doFill = fill && brush.style !== 1;
    const doStroke = stroke && (pen.style & 0xf) !== 5;
    if (!doFill && !doStroke) return;
    const width = pen.width ? Math.max(0.3, this.len(pen.width)) : Math.max(0.3, this.scale);
    const st = pen.style & 0xf;
    const dash = st === 1 ? [width * 4, width * 2] : st === 2 ? [width, width * 2] : st === 3 || st === 4 ? [width * 4, width * 2, width, width * 2] : undefined;
    this.cmds.push({ k: 'path', d, fill: doFill ? brush.color : undefined, evenOdd: this.dc.polyFill === 1, stroke: doStroke ? pen.color : undefined, width, dash });
  }

  poly(points: [number, number][], close: boolean, fill: boolean): void {
    if (points.length < 2) return;
    const out = points.map(([x, y]) => this.pt(x, y));
    const d = 'M' + out.map(([x, y]) => `${this.f(x)} ${this.f(y)}`).join(' L') + (close ? ' Z' : '');
    if (this.path) {
      // inside a path bracket: continue the current figure when drawing "to" the current position
      this.path.push(d);
      this.figureOpen = !close;
      return;
    }
    this.emit(d, fill && close, true);
  }

  polyPoly(polys: [number, number][][], close: boolean, fill: boolean): void {
    const d = polys
      .filter((p) => p.length > 1)
      .map((p) => 'M' + p.map(([x, y]) => this.pt(x, y)).map(([x, y]) => `${this.f(x)} ${this.f(y)}`).join(' L') + (close ? ' Z' : ''))
      .join(' ');
    this.emit(d, fill && close, true);
  }

  /** Lines/curves continuing from the current position. */
  to(points: [number, number][], bezier: boolean): void {
    if (!points.length) return;
    const start = this.pt(...this.dc.cur);
    const out = points.map(([x, y]) => this.pt(x, y));
    let d = '';
    if (bezier) for (let i = 0; i + 2 < out.length; i += 3) d += ` C${out.slice(i, i + 3).map(([x, y]) => `${this.f(x)} ${this.f(y)}`).join(' ')}`;
    else d = ' L' + out.map(([x, y]) => `${this.f(x)} ${this.f(y)}`).join(' L');
    this.dc.cur = points[points.length - 1];
    if (this.path) {
      if (!this.figureOpen) this.path.push(`M${this.f(start[0])} ${this.f(start[1])}`);
      this.path.push(d);
      this.figureOpen = true;
      return;
    }
    this.emit(`M${this.f(start[0])} ${this.f(start[1])}${d}`, false, true);
  }

  bezier(points: [number, number][]): void {
    if (points.length < 4) return;
    const out = points.map(([x, y]) => this.pt(x, y));
    let d = `M${this.f(out[0][0])} ${this.f(out[0][1])}`;
    for (let i = 1; i + 2 < out.length; i += 3) d += ` C${out.slice(i, i + 3).map(([x, y]) => `${this.f(x)} ${this.f(y)}`).join(' ')}`;
    this.emit(d, false, true);
  }

  moveTo(x: number, y: number): void {
    this.dc.cur = [x, y];
    if (this.path) {
      const [px, py] = this.pt(x, y);
      this.path.push(`M${this.f(px)} ${this.f(py)}`);
      this.figureOpen = true;
    }
  }

  rect(l: number, t: number, r: number, b: number, rx = 0, ry = 0): void {
    if (rx || ry) {
      // rounded rectangle: corners as quarter ellipses (cubic approximation)
      const k = 0.5523;
      const ax = rx / 2;
      const ay = ry / 2;
      const P = (x: number, y: number): string => this.pt(x, y).map((v) => this.f(v)).join(' ');
      const d = `M${P(l + ax, t)} L${P(r - ax, t)} C${P(r - ax + ax * k, t)} ${P(r, t + ay - ay * k)} ${P(r, t + ay)} L${P(r, b - ay)} C${P(r, b - ay + ay * k)} ${P(r - ax + ax * k, b)} ${P(r - ax, b)} L${P(l + ax, b)} C${P(l + ax - ax * k, b)} ${P(l, b - ay + ay * k)} ${P(l, b - ay)} L${P(l, t + ay)} C${P(l, t + ay - ay * k)} ${P(l + ax - ax * k, t)} ${P(l + ax, t)} Z`;
      this.emit(d, true, true);
      return;
    }
    this.poly(
      [
        [l, t],
        [r, t],
        [r, b],
        [l, b],
      ],
      true,
      true,
    );
  }

  ellipse(l: number, t: number, r: number, b: number): void {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const rx = (r - l) / 2;
    const ry = (b - t) / 2;
    const k = 0.5523;
    const P = (x: number, y: number): string => this.pt(x, y).map((v) => this.f(v)).join(' ');
    const d = `M${P(cx + rx, cy)} C${P(cx + rx, cy + ry * k)} ${P(cx + rx * k, cy + ry)} ${P(cx, cy + ry)} C${P(cx - rx * k, cy + ry)} ${P(cx - rx, cy + ry * k)} ${P(cx - rx, cy)} C${P(cx - rx, cy - ry * k)} ${P(cx - rx * k, cy - ry)} ${P(cx, cy - ry)} C${P(cx + rx * k, cy - ry)} ${P(cx + rx, cy - ry * k)} ${P(cx + rx, cy)} Z`;
    this.emit(d, true, true);
  }

  /** ARC/CHORD/PIE: the ellipse in the box between two radial points (counter-clockwise). */
  arc(l: number, t: number, r: number, b: number, sx: number, sy: number, ex: number, ey: number, kind: 'arc' | 'chord' | 'pie'): void {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const rx = (r - l) / 2;
    const ry = (b - t) / 2;
    const a0 = Math.atan2(sy - cy, sx - cx);
    let a1 = Math.atan2(ey - cy, ex - cx);
    // GDI default direction is counter-clockwise in logical space with y down: angles decrease
    while (a1 >= a0) a1 -= Math.PI * 2;
    const steps = Math.max(8, Math.ceil(((a0 - a1) / (Math.PI * 2)) * 48));
    const pts: [number, number][] = [];
    for (let i = 0; i <= steps; i++) {
      const a = a0 + ((a1 - a0) * i) / steps;
      pts.push([cx + rx * Math.cos(a), cy + ry * Math.sin(a)]);
    }
    if (kind === 'pie') pts.push([cx, cy]);
    this.poly(pts, kind !== 'arc', kind !== 'arc');
  }

  beginPath(): void {
    this.path = [];
    this.figureOpen = false;
  }
  closeFigure(): void {
    if (this.path) {
      this.path.push('Z');
      this.figureOpen = false;
    }
  }
  endPath(): string {
    const d = (this.path ?? []).join(' ');
    this.path = null;
    this.pending = d;
    return d;
  }
  pending = '';
  usePath(fill: boolean, stroke: boolean): void {
    const d = this.pending;
    this.pending = '';
    if (d) this.paint(d, fill, stroke);
  }
  abortPath(): void {
    this.path = null;
    this.pending = '';
  }

  fillRect(l: number, t: number, r: number, b: number, colorHex?: string): void {
    const [x0, y0] = this.pt(l, t);
    const [x1, y1] = this.pt(r, b);
    this.cmds.push({ k: 'path', d: `M${this.f(x0)} ${this.f(y0)} L${this.f(x1)} ${this.f(y0)} L${this.f(x1)} ${this.f(y1)} L${this.f(x0)} ${this.f(y1)} Z`, fill: colorHex ?? this.dc.brush.color, evenOdd: false, width: 0 });
  }

  text(x: number, y: number, s: string, useCurrent: boolean): void {
    if (!s.trim()) return;
    const d = this.dc;
    if (useCurrent || d.textAlign & 1) [x, y] = d.cur;
    const [ox, oy] = this.pt(x, y);
    const f = d.font;
    // Negative height = character height (em); positive = cell height (≈ 1.2 em)
    const size = Math.max(1, this.len(Math.abs(f.height)) * (f.height > 0 ? 0.82 : 1));
    const h = d.textAlign & 6;
    const v = d.textAlign & 24;
    this.cmds.push({
      k: 'text',
      x: ox,
      y: oy,
      text: s,
      size,
      family: familyOf(f.face) ?? 'sans',
      bold: f.weight >= 600,
      italic: f.italic,
      underline: f.underline,
      color: d.textColor,
      angle: f.escapement / 10,
      anchor: h === 6 ? 'middle' : h === 2 ? 'end' : 'start',
      baseline: v === 24 ? 'alphabetic' : v === 8 ? 'bottom' : 'top',
    });
  }

  bitmap(x: number, y: number, w: number, h: number, dib: Uint8Array): void {
    const png = dibToPng(dib);
    if (!png) return;
    const [x0, y0] = this.pt(x, y);
    const [x1, y1] = this.pt(x + w, y + h);
    this.cmds.push({ k: 'image', x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0), png });
  }

  select(obj: GdiObj | undefined): void {
    if (!obj) return;
    if (obj.t === 'pen') this.dc.pen = obj.v;
    else if (obj.t === 'brush') this.dc.brush = obj.v;
    else if (obj.t === 'font') this.dc.font = obj.v;
  }
}

// ---------------------------------------------------------------------------
// Bitmaps (DIB → PNG)

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(b: Uint8Array): number {
  let c = 0xffffffff;
  for (const v of b) c = CRC_TABLE[(c ^ v) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export function encodePng(rgba: Uint8Array, w: number, h: number): Uint8Array {
  const raw = new Uint8Array((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) raw.set(rgba.subarray(y * w * 4, (y + 1) * w * 4), y * (w * 4 + 1) + 1);
  const chunk = (type: string, data: Uint8Array): Uint8Array => {
    const out = new Uint8Array(12 + data.length);
    const dv = new DataView(out.buffer);
    dv.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    dv.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
    return out;
  };
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, w);
  dv.setUint32(4, h);
  ihdr.set([8, 6, 0, 0, 0], 8);
  const parts = [Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlibSync(raw, { level: 6 })), chunk('IEND', new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** BITMAPINFO + bits (1/4/8/16/24/32 bpp, uncompressed or embedded PNG/JPEG) → PNG. */
function dibToPng(dib: Uint8Array): Uint8Array | null {
  if (dib.length < 40) return null;
  const dv = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const hs = dv.getUint32(0, true);
  if (hs < 12) return null;
  const core = hs === 12;
  const w = core ? dv.getUint16(4, true) : dv.getInt32(4, true);
  let h = core ? dv.getInt16(6, true) : dv.getInt32(8, true);
  const bpp = core ? dv.getUint16(10, true) : dv.getUint16(14, true);
  const comp = core ? 0 : dv.getUint32(16, true);
  const used = core ? 0 : dv.getUint32(32, true);
  const topDown = h < 0;
  h = Math.abs(h);
  if (w <= 0 || h <= 0 || w * h > 40_000_000) return null;
  if (comp === 5) return dib.subarray(hs); // BI_PNG
  if (comp !== 0 && comp !== 3) return null;
  const palN = bpp <= 8 ? used || 1 << bpp : 0;
  const palSize = core ? 3 : 4;
  let off = hs + (comp === 3 && !core ? 12 : 0);
  const pal: number[][] = [];
  for (let i = 0; i < palN; i++) pal.push([dib[off + i * palSize + 2], dib[off + i * palSize + 1], dib[off + i * palSize]]);
  off += palN * palSize;
  const stride = Math.ceil((w * bpp) / 32) * 4;
  if (off + stride * h > dib.length) return null;
  const rgba = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y++) {
    const row = off + (topDown ? y : h - 1 - y) * stride;
    for (let x = 0; x < w; x++) {
      let r = 0;
      let g = 0;
      let b = 0;
      if (bpp === 24 || bpp === 32) {
        const p = row + x * (bpp / 8);
        b = dib[p];
        g = dib[p + 1];
        r = dib[p + 2];
      } else if (bpp === 16) {
        const v = dib[row + x * 2] | (dib[row + x * 2 + 1] << 8);
        r = ((v >> 10) & 31) << 3;
        g = ((v >> 5) & 31) << 3;
        b = (v & 31) << 3;
      } else {
        const bit = x * bpp;
        const byte = dib[row + (bit >> 3)];
        const idx = (byte >> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
        [r, g, b] = pal[idx] ?? [0, 0, 0];
      }
      const o = (y * w + x) * 4;
      rgba[o] = r;
      rgba[o + 1] = g;
      rgba[o + 2] = b;
      rgba[o + 3] = 255;
    }
  }
  return encodePng(rgba, w, h);
}

// ---------------------------------------------------------------------------
// EMF

function utf16(d: Uint8Array, off: number, chars: number): string {
  let s = '';
  for (let i = 0; i < chars && off + i * 2 + 1 < d.length; i++) {
    const c = d[off + i * 2] | (d[off + i * 2 + 1] << 8);
    if (c === 0) break;
    s += String.fromCharCode(c);
  }
  return s;
}

const CHARSET_DECODER: Record<number, string> = { 204: 'windows-1251', 238: 'windows-1250', 161: 'windows-1253', 162: 'windows-1254', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257', 128: 'shift_jis', 134: 'gbk', 136: 'big5', 129: 'euc-kr' };
function ansi(bytes: Uint8Array, charset: number): string {
  try {
    return new TextDecoder(CHARSET_DECODER[charset] ?? 'windows-1252').decode(bytes);
  } catch {
    return String.fromCharCode(...bytes);
  }
}

export function parseEmf(data: Uint8Array): Metafile {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const bl = dv.getInt32(8, true);
  const bt = dv.getInt32(12, true);
  const br = dv.getInt32(16, true);
  const bb = dv.getInt32(20, true);
  const fl = dv.getInt32(24, true);
  const ft = dv.getInt32(28, true);
  const fr = dv.getInt32(32, true);
  const fb = dv.getInt32(36, true);
  const devW = dv.getInt32(72, true);
  const devH = dv.getInt32(76, true);
  const mmW = dv.getInt32(80, true);
  const mmH = dv.getInt32(84, true);
  const g = new Gdi();
  if (mmW > 0 && mmH > 0 && devW > 0 && devH > 0) g.devPerMm = [devW / mmW, devH / mmH];
  // Output space: the picture frame (0.01 mm) in device units
  const left = (fl / 100) * g.devPerMm[0];
  const top = (ft / 100) * g.devPerMm[1];
  let width = ((fr - fl) / 100) * g.devPerMm[0];
  let height = ((fb - ft) / 100) * g.devPerMm[1];
  g.origin = [left, top];
  if (!(width > 0 && height > 0)) {
    g.origin = [bl, bt];
    width = br - bl + 1;
    height = bb - bt + 1;
  }
  const widthPt = fr > fl ? ((fr - fl) / 100) * (72 / 25.4) : (width * 72) / 96;
  const heightPt = fb > ft ? ((fb - ft) / 100) * (72 / 25.4) : (height * 72) / 96;
  const i32 = (o: number): number => dv.getInt32(o, true);
  const u32 = (o: number): number => dv.getUint32(o, true);
  const f32 = (o: number): number => dv.getFloat32(o, true);
  const xform = (o: number): Xf => [f32(o), f32(o + 4), f32(o + 8), f32(o + 12), f32(o + 16), f32(o + 20)];
  let pos = 0;
  let guard = 0;
  while (pos + 8 <= data.length && guard++ < 1_000_000) {
    const type = u32(pos);
    const size = u32(pos + 4);
    if (size < 8 || pos + size > data.length) break;
    const p = pos + 8;
    const pts16 = (count: number, o: number): [number, number][] => Array.from({ length: count }, (_, i) => [dv.getInt16(o + i * 4, true), dv.getInt16(o + i * 4 + 2, true)] as [number, number]);
    const pts32 = (count: number, o: number): [number, number][] => Array.from({ length: count }, (_, i) => [i32(o + i * 8), i32(o + i * 8 + 4)] as [number, number]);
    switch (type) {
      case 14: // EOF
        pos = data.length;
        continue;
      case 9:
        g.dc.winExt = [i32(p), i32(p + 4)];
        break;
      case 10:
        g.dc.winOrg = [i32(p), i32(p + 4)];
        break;
      case 11:
        g.dc.vpExt = [i32(p), i32(p + 4)];
        break;
      case 12:
        g.dc.vpOrg = [i32(p), i32(p + 4)];
        break;
      case 17:
        g.dc.mapMode = u32(p);
        break;
      case 18:
        g.dc.bkMode = u32(p);
        break;
      case 19:
        g.dc.polyFill = u32(p);
        break;
      case 22:
        g.dc.textAlign = u32(p);
        break;
      case 24:
        g.dc.textColor = color(u32(p));
        break;
      case 25:
        g.dc.bkColor = color(u32(p));
        break;
      case 27:
        g.moveTo(i32(p), i32(p + 4));
        break;
      case 33:
        g.save();
        break;
      case 34:
        g.restore(i32(p));
        break;
      case 35:
        g.dc.world = xform(p);
        break;
      case 36: {
        const m = xform(p);
        const mode = u32(p + 24);
        if (mode === 1) g.dc.world = [...IDENT];
        else if (mode === 2) g.dc.world = mulXf(m, g.dc.world);
        else if (mode === 3) g.dc.world = mulXf(g.dc.world, m);
        else if (mode === 4) g.dc.world = m;
        break;
      }
      case 37: {
        const idx = u32(p);
        const stock = idx & 0x7fffffff;
        // stock objects 10..17 are fonts
        g.select(idx & 0x80000000 ? (STOCK[stock] ?? (stock >= 10 && stock <= 17 ? { t: 'font', v: DEFAULT_FONT } : undefined)) : g.objects[idx]);
        break;
      }
      case 38: // CREATEPEN
        g.objects[u32(p)] = { t: 'pen', v: { style: u32(p + 4), width: i32(p + 8), color: color(u32(p + 16)) } };
        break;
      case 95: // EXTCREATEPEN
        g.objects[u32(p)] = { t: 'pen', v: { style: u32(p + 20), width: u32(p + 24), color: color(u32(p + 32)) } };
        break;
      case 39: // CREATEBRUSHINDIRECT
        g.objects[u32(p)] = { t: 'brush', v: { style: u32(p + 4), color: color(u32(p + 8)) } };
        break;
      case 93:
      case 94: // pattern brushes: approximate with gray
        g.objects[u32(p)] = { t: 'brush', v: { style: 0, color: '808080' } };
        break;
      case 82: {
        // EXTCREATEFONTINDIRECTW
        g.objects[u32(p)] = { t: 'font', v: { height: i32(p + 4), escapement: i32(p + 12), weight: i32(p + 20), italic: data[p + 24] !== 0, underline: data[p + 25] !== 0, charset: data[p + 27], face: utf16(data, p + 32, 32) } };
        break;
      }
      case 40:
        g.objects[u32(p)] = undefined;
        break;
      case 43:
        g.rect(i32(p), i32(p + 4), i32(p + 8), i32(p + 12));
        break;
      case 44:
        g.rect(i32(p), i32(p + 4), i32(p + 8), i32(p + 12), i32(p + 16), i32(p + 20));
        break;
      case 42:
        g.ellipse(i32(p), i32(p + 4), i32(p + 8), i32(p + 12));
        break;
      case 45:
      case 46:
      case 47:
        g.arc(i32(p), i32(p + 4), i32(p + 8), i32(p + 12), i32(p + 16), i32(p + 20), i32(p + 24), i32(p + 28), type === 45 ? 'arc' : type === 46 ? 'chord' : 'pie');
        break;
      case 54: {
        const x = i32(p);
        const y = i32(p + 4);
        g.to([[x, y]], false);
        break;
      }
      case 3:
        g.poly(pts32(u32(p + 16), p + 20), true, true);
        break;
      case 4:
        g.poly(pts32(u32(p + 16), p + 20), false, false);
        break;
      case 2:
        g.bezier(pts32(u32(p + 16), p + 20));
        break;
      case 5:
        g.to(pts32(u32(p + 16), p + 20), true);
        break;
      case 6:
        g.to(pts32(u32(p + 16), p + 20), false);
        break;
      case 86:
        g.poly(pts16(u32(p + 16), p + 20), true, true);
        break;
      case 87:
        g.poly(pts16(u32(p + 16), p + 20), false, false);
        break;
      case 85:
        g.bezier(pts16(u32(p + 16), p + 20));
        break;
      case 88:
        g.to(pts16(u32(p + 16), p + 20), true);
        break;
      case 89:
        g.to(pts16(u32(p + 16), p + 20), false);
        break;
      case 7:
      case 8:
      case 90:
      case 91: {
        // POLYPOLYLINE / POLYPOLYGON (32 and 16 bit)
        const n = u32(p + 16);
        const counts = Array.from({ length: n }, (_, i) => u32(p + 24 + i * 4));
        let o = p + 24 + n * 4;
        const polys: [number, number][][] = [];
        const small = type >= 90;
        for (const c of counts) {
          polys.push(small ? pts16(c, o) : pts32(c, o));
          o += c * (small ? 4 : 8);
        }
        const closed = type === 8 || type === 91;
        g.polyPoly(polys, closed, closed);
        break;
      }
      case 59:
        g.beginPath();
        break;
      case 60:
        g.endPath();
        break;
      case 61:
        g.closeFigure();
        break;
      case 62:
        g.usePath(true, false);
        break;
      case 63:
        g.usePath(true, true);
        break;
      case 64:
        g.usePath(false, true);
        break;
      case 68:
        g.abortPath();
        break;
      case 83:
      case 84: {
        // EXTTEXTOUTA / EXTTEXTOUTW
        const e = p + 28; // EMRTEXT
        const x = i32(e);
        const y = i32(e + 4);
        const n = u32(e + 8);
        const off = u32(e + 12);
        const str = type === 84 ? utf16(data, pos + off, n) : ansi(data.subarray(pos + off, pos + off + n), g.dc.font.charset);
        g.text(x, y, str, false);
        break;
      }
      case 81: {
        // STRETCHDIBITS
        const xd = i32(p + 16);
        const yd = i32(p + 20);
        const offBmi = u32(p + 40);
        const cbBmi = u32(p + 44);
        const offBits = u32(p + 48);
        const cbBits = u32(p + 52);
        const wd = i32(p + 64);
        const hd = i32(p + 68);
        if (cbBmi) g.bitmap(xd, yd, wd, hd, joinDib(data, pos + offBmi, cbBmi, pos + offBits, cbBits));
        break;
      }
      case 76:
      case 77: {
        // BITBLT / STRETCHBLT
        const xd = i32(p + 16);
        const yd = i32(p + 20);
        const wd = i32(p + 24);
        const hd = i32(p + 28);
        const offBmi = u32(p + 76);
        const cbBmi = u32(p + 80);
        const offBits = u32(p + 84);
        const cbBits = u32(p + 88);
        if (cbBmi) {
          const sw = type === 77 ? i32(p + 92) : wd;
          const sh = type === 77 ? i32(p + 96) : hd;
          void sw;
          void sh;
          g.bitmap(xd, yd, wd, hd, joinDib(data, pos + offBmi, cbBmi, pos + offBits, cbBits));
        } else if (g.dc.brush.style !== 1) g.fillRect(xd, yd, xd + wd, yd + hd);
        break;
      }
      case 80: {
        // SETDIBITSTODEVICE
        const xd = i32(p + 16);
        const yd = i32(p + 20);
        const wd = i32(p + 32);
        const hd = i32(p + 36);
        const offBmi = u32(p + 40);
        const cbBmi = u32(p + 44);
        const offBits = u32(p + 48);
        const cbBits = u32(p + 52);
        if (cbBmi) g.bitmap(xd, yd, wd, hd, joinDib(data, pos + offBmi, cbBmi, pos + offBits, cbBits));
        break;
      }
      default:
        break;
    }
    pos += size;
  }
  return { width, height, widthPt, heightPt, cmds: g.cmds };
}

function joinDib(d: Uint8Array, bmiOff: number, bmiLen: number, bitsOff: number, bitsLen: number): Uint8Array {
  const out = new Uint8Array(bmiLen + bitsLen);
  out.set(d.subarray(bmiOff, bmiOff + bmiLen));
  out.set(d.subarray(bitsOff, bitsOff + bitsLen), bmiLen);
  return out;
}

// ---------------------------------------------------------------------------
// WMF

export function parseWmf(data: Uint8Array): Metafile {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let pos = 0;
  let bbox: [number, number, number, number] | null = null;
  let inch = 1440;
  if (dv.getUint32(0, true) === 0x9ac6cdd7) {
    bbox = [dv.getInt16(6, true), dv.getInt16(8, true), dv.getInt16(10, true), dv.getInt16(12, true)];
    inch = dv.getUint16(14, true) || 1440;
    pos = 22;
  }
  const headerWords = dv.getUint16(pos + 2, true);
  pos += headerWords * 2;
  const g = new Gdi();
  g.dc.mapMode = 8; // WMF logical units map through the window by default
  const recs: { fn: number; p: number; size: number }[] = [];
  let guard = 0;
  // First pass: the window extent of non-placeable files defines the picture
  let win: [number, number, number, number] | null = null;
  for (let q = pos; q + 6 <= data.length && guard++ < 1_000_000; ) {
    const size = dv.getUint32(q, true) * 2;
    const fn = dv.getUint16(q + 4, true);
    if (size < 6 || q + size > data.length) break;
    recs.push({ fn, p: q + 6, size });
    if (fn === 0x020b && !win) win = [dv.getInt16(q + 8, true), dv.getInt16(q + 6, true), 0, 0];
    if (fn === 0x020c && win) win = [win[0], win[1], dv.getInt16(q + 8, true), dv.getInt16(q + 6, true)];
    if (fn === 0) break;
    q += size;
  }
  const box = bbox ?? (win ? [win[0], win[1], win[0] + win[2], win[1] + win[3]] : [0, 0, 1000, 1000]);
  const lw = Math.abs(box[2] - box[0]) || 1;
  const lh = Math.abs(box[3] - box[1]) || 1;
  // Output in 1/96 in for placeable files, logical units otherwise
  const scale = bbox ? 96 / inch : 1;
  const widthPt = bbox ? (lw / inch) * 72 : (lw * 72) / 96;
  const heightPt = bbox ? (lh / inch) * 72 : (lh * 72) / 96;
  // Device space == logical space mapped through window→viewport set to the box
  g.dc.winOrg = [Math.min(box[0], box[2]), Math.min(box[1], box[3])];
  g.dc.winExt = [lw, lh];
  g.dc.vpOrg = [0, 0];
  g.dc.vpExt = [lw * scale, lh * scale];
  const i16 = (o: number): number => dv.getInt16(o, true);
  const u16 = (o: number): number => dv.getUint16(o, true);
  const u32 = (o: number): number => dv.getUint32(o, true);
  const addObj = (o: GdiObj): void => {
    let i = g.objects.findIndex((x) => x === undefined);
    if (i < 0) i = g.objects.length;
    g.objects[i] = o;
  };
  for (const { fn, p, size } of recs) {
    switch (fn) {
      case 0x020b: // SETWINDOWORG
        g.dc.winOrg = [i16(p + 2), i16(p)];
        break;
      case 0x020c: // SETWINDOWEXT
        g.dc.winExt = [i16(p + 2), i16(p)];
        g.dc.vpExt = [Math.abs(i16(p + 2)) * scale, Math.abs(i16(p)) * scale];
        break;
      case 0x0103:
        // map mode: placeable files are drawn into their bounding box regardless
        break;
      case 0x0102:
        g.dc.bkMode = u16(p);
        break;
      case 0x0106:
        g.dc.polyFill = u16(p);
        break;
      case 0x0209:
        g.dc.textColor = color(u32(p));
        break;
      case 0x0201:
        g.dc.bkColor = color(u32(p));
        break;
      case 0x012e:
        g.dc.textAlign = u16(p);
        break;
      case 0x001e:
        g.save();
        break;
      case 0x0127:
        g.restore(i16(p));
        break;
      case 0x012d:
        g.select(g.objects[u16(p)]);
        break;
      case 0x01f0:
        g.objects[u16(p)] = undefined;
        break;
      case 0x02fa:
        addObj({ t: 'pen', v: { style: u16(p), width: i16(p + 2), color: color(u32(p + 6)) } });
        break;
      case 0x02fc:
        addObj({ t: 'brush', v: { style: u16(p), color: color(u32(p + 2)) } });
        break;
      case 0x0142:
      case 0x01f9:
        addObj({ t: 'brush', v: { style: 0, color: '808080' } });
        break;
      case 0x02fb: {
        const faceBytes = data.subarray(p + 18, Math.min(p + 18 + 32, p + size - 6));
        const end = faceBytes.indexOf(0);
        addObj({ t: 'font', v: { height: i16(p), escapement: i16(p + 4), weight: i16(p + 8), italic: data[p + 10] !== 0, underline: data[p + 11] !== 0, charset: data[p + 13], face: ansi(end >= 0 ? faceBytes.subarray(0, end) : faceBytes, 0) } });
        break;
      }
      case 0x00f7:
      case 0x06ff:
        addObj({ t: 'other' });
        break;
      case 0x0214:
        g.moveTo(i16(p + 2), i16(p));
        break;
      case 0x0213:
        g.to([[i16(p + 2), i16(p)]], false);
        break;
      case 0x041b:
        g.rect(i16(p + 6), i16(p + 4), i16(p + 2), i16(p));
        break;
      case 0x061c:
        g.rect(i16(p + 10), i16(p + 8), i16(p + 6), i16(p + 4), i16(p + 2), i16(p));
        break;
      case 0x0418:
        g.ellipse(i16(p + 6), i16(p + 4), i16(p + 2), i16(p));
        break;
      case 0x0817:
      case 0x0830:
      case 0x081a:
        // ARC / CHORD / PIE: yEnd xEnd yStart xStart bottom right top left
        g.arc(i16(p + 14), i16(p + 12), i16(p + 10), i16(p + 8), i16(p + 6), i16(p + 4), i16(p + 2), i16(p), fn === 0x0817 ? 'arc' : fn === 0x0830 ? 'chord' : 'pie');
        break;
      case 0x0324:
      case 0x0325: {
        const n = u16(p);
        const pts = Array.from({ length: n }, (_, i) => [i16(p + 2 + i * 4), i16(p + 4 + i * 4)] as [number, number]);
        g.poly(pts, fn === 0x0324, fn === 0x0324);
        break;
      }
      case 0x0538: {
        const n = u16(p);
        const counts = Array.from({ length: n }, (_, i) => u16(p + 2 + i * 2));
        let o = p + 2 + n * 2;
        const polys: [number, number][][] = [];
        for (const c of counts) {
          polys.push(Array.from({ length: c }, (_, i) => [i16(o + i * 4), i16(o + i * 4 + 2)] as [number, number]));
          o += c * 4;
        }
        g.polyPoly(polys, true, true);
        break;
      }
      case 0x0521: {
        const n = u16(p);
        const str = ansi(data.subarray(p + 2, p + 2 + n), g.dc.font.charset);
        const o = p + 2 + n + (n & 1);
        g.text(i16(o + 2), i16(o), str, false);
        break;
      }
      case 0x0a32: {
        const y = i16(p);
        const x = i16(p + 2);
        const n = u16(p + 4);
        const opts = u16(p + 6);
        const o = p + 8 + (opts & 0x6 ? 8 : 0);
        g.text(x, y, ansi(data.subarray(o, o + n), g.dc.font.charset), false);
        break;
      }
      case 0x061d: {
        // PATBLT: rop(4) height width y x
        const hh = i16(p + 4);
        const ww = i16(p + 6);
        const y = i16(p + 8);
        const x = i16(p + 10);
        if (g.dc.brush.style !== 1) g.fillRect(x, y, x + ww, y + hh);
        break;
      }
      case 0x0f43: {
        // STRETCHDIB: rop(4) usage srcH srcW ySrc xSrc destH destW yDest xDest dib
        const hd = i16(p + 14);
        const wd = i16(p + 16);
        const yd = i16(p + 18);
        const xd = i16(p + 20);
        g.bitmap(xd, yd, wd, hd, data.subarray(p + 22, p + size - 6));
        break;
      }
      case 0x0b41: {
        // DIBSTRETCHBLT: rop(4) srcH srcW ySrc xSrc destH destW yDest xDest dib
        const hd = i16(p + 12);
        const wd = i16(p + 14);
        const yd = i16(p + 16);
        const xd = i16(p + 18);
        if (size > 6 + 22) g.bitmap(xd, yd, wd, hd, data.subarray(p + 20, p + size - 6));
        break;
      }
      case 0x0940: {
        // DIBBITBLT: rop(4) ySrc xSrc height width yDest xDest dib
        const hd = i16(p + 8);
        const wd = i16(p + 10);
        const yd = i16(p + 12);
        const xd = i16(p + 14);
        if (size > 6 + 18) g.bitmap(xd, yd, wd, hd, data.subarray(p + 16, p + size - 6));
        break;
      }
      default:
        break;
    }
  }
  return { width: lw * scale, height: lh * scale, widthPt, heightPt, cmds: g.cmds };
}

export function parseMetafile(data: Uint8Array): Metafile | null {
  try {
    if (isEmf(data)) return parseEmf(data);
    if (isWmf(data)) return parseWmf(data);
  } catch {
    return null;
  }
  return null;
}

// ---------------------------------------------------------------------------
// SVG output

const SVG_FAMILY: Record<FontFamily, string> = { sans: 'Arial, Liberation Sans, sans-serif', serif: 'Times New Roman, Liberation Serif, serif', mono: 'Courier New, Liberation Mono, monospace', calibri: 'Calibri, Carlito, sans-serif' };

export function metafileToSvg(m: Metafile): string {
  const f = (n: number): string => (Math.round(n * 100) / 100).toString();
  let body = '';
  for (const c of m.cmds) {
    if (c.k === 'path') {
      body += `<path d="${c.d}" fill="${c.fill ? '#' + c.fill : 'none'}"${c.evenOdd ? ' fill-rule="evenodd"' : ''}${c.stroke ? ` stroke="#${c.stroke}" stroke-width="${f(c.width)}"${c.dash ? ` stroke-dasharray="${c.dash.map(f).join(' ')}"` : ''}` : ''}/>`;
    } else if (c.k === 'text') {
      const rot = c.angle ? ` transform="rotate(${f(-c.angle)} ${f(c.x)} ${f(c.y)})"` : '';
      const base = c.baseline === 'top' ? ' dominant-baseline="text-before-edge"' : c.baseline === 'bottom' ? ' dominant-baseline="text-after-edge"' : '';
      body += `<text x="${f(c.x)}" y="${f(c.y)}" font-family="${SVG_FAMILY[c.family]}" font-size="${f(c.size)}"${c.bold ? ' font-weight="bold"' : ''}${c.italic ? ' font-style="italic"' : ''}${c.underline ? ' text-decoration="underline"' : ''} fill="#${c.color}"${c.anchor !== 'start' ? ` text-anchor="${c.anchor}"` : ''}${base}${rot} xml:space="preserve">${escapeXml(c.text)}</text>`;
    } else {
      body += `<image x="${f(c.x)}" y="${f(c.y)}" width="${f(c.w)}" height="${f(c.h)}" preserveAspectRatio="none" href="data:image/png;base64,${bytesToBase64(c.png)}"/>`;
    }
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${f(m.width)} ${f(m.height)}" width="${f(m.widthPt)}pt" height="${f(m.heightPt)}pt">${body}</svg>`;
}

/** EMF/WMF bytes → SVG bytes (null when the data is not a readable metafile). */
export function metafileToSvgBytes(data: Uint8Array): Uint8Array | null {
  const m = parseMetafile(data);
  return m && m.cmds.length ? new TextEncoder().encode(metafileToSvg(m)) : null;
}
