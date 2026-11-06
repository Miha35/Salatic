/**
 * EMF+ (GDI+) records → the vector display list. EMF+ records travel inside EMR_COMMENT records of an
 * EMF; "dual" files also carry GDI records, which parseEmf draws. Files that are EMF+ only (or whose
 * GDI part draws nothing) are drawn from these records instead.
 *
 * Supported: objects (solid, hatch and linear-gradient brushes; pens with width, dash styles and their
 * brush; paths; bitmaps, compressed or 24/32-bit pixels; fonts; string formats), Clear, Fill/Draw
 * Rects, Polygon, Lines, Ellipse, Pie, Arc, Path, Beziers, (Closed)Curve (cardinal splines),
 * DrawImage(Points), DrawString, DrawDriverString (Unicode), world and page transforms, Save/Restore,
 * containers, and rectangle/path clipping (replace, intersect, exclude).
 * Not drawn: regions, texture and path-gradient brushes (their colour is used), metafile images.
 */
import { familyOf, type Gradient } from './model';
import type { ClipPath, MetaCmd } from './metafile';
import { encodePng } from './png';

/** Row-vector affine matrix [m11 m12 m21 m22 dx dy]: p' = p · M. */
type Mx = [number, number, number, number, number, number];
const ID: Mx = [1, 0, 0, 1, 0, 0];
/** a then b */
const mul = (a: Mx, b: Mx): Mx => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
const app = (m: Mx, x: number, y: number): [number, number] => [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]];

const FAR = 'M-100000 -100000 L100000 -100000 L100000 100000 L-100000 100000 Z';

interface Brush {
  color: string;
  alpha: number;
  /** linear gradient: rectangle (world) and end colours */
  lin?: { rect: [number, number, number, number]; c0: string; c1: string; a0: number; a1: number };
}
interface Pen {
  width: number;
  /** width in world units (transformed) or device pixels */
  world: boolean;
  dash?: number[];
  brush: Brush;
}
interface PathObj {
  pts: [number, number][];
  types: number[];
}
interface FontObj {
  size: number;
  /** size in world units, else device pixels/points */
  unit: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  family: string;
}
interface ImageObj {
  png: Uint8Array;
  w: number;
  h: number;
}
interface FormatObj {
  align: number;
  lineAlign: number;
}
type Obj = Brush | Pen | PathObj | FontObj | ImageObj | FormatObj | null;

export interface EmfPlusTarget {
  /** EMF device units → output space */
  devToOut: (x: number, y: number) => [number, number];
  /** output units per EMF device unit */
  devScale: number;
  /** the picture in output space (Clear fills it) */
  frame: [number, number, number, number];
}

const argb = (v: number): { color: string; alpha: number } => ({
  color: (v & 0xffffff).toString(16).padStart(6, '0').toUpperCase(),
  alpha: ((v >>> 24) & 0xff) / 255,
});

/** Device pixels per page unit at the reference DPI (UnitWorld/UnitDisplay/UnitPixel count as pixels). */
function unitPx(unit: number, dpi: number): number {
  switch (unit) {
    case 3:
      return dpi / 72; // point
    case 4:
      return dpi; // inch
    case 5:
      return dpi / 300; // document
    case 6:
      return dpi / 25.4; // millimetre
    default:
      return 1;
  }
}

/**
 * Play EMF+ records (`blobs`: the contents of the EMF+ comments, in order). Returns the display list, or
 * null when there is no EMF+ header.
 */
export function playEmfPlus(blobs: Uint8Array[], t: EmfPlusTarget): { cmds: MetaCmd[]; dual: boolean } | null {
  const all = mergeObjects(blobs);
  const dv = new DataView(all.buffer, all.byteOffset, all.byteLength);
  const u16 = (p: number): number => dv.getUint16(p, true);
  const u32 = (p: number): number => dv.getUint32(p, true);
  const i32 = (p: number): number => dv.getInt32(p, true);
  const i16 = (p: number): number => dv.getInt16(p, true);
  const f32 = (p: number): number => dv.getFloat32(p, true);
  const cmds: MetaCmd[] = [];
  const objects: Obj[] = new Array(64).fill(null);
  let dpiX = 96;
  let dpiY = 96;
  let dual = false;
  let header = false;
  let world: Mx = ID;
  let pageUnit = 2;
  let pageScale = 1;
  let clip: ClipPath[] = [];
  const saved = new Map<number, { world: Mx; pageUnit: number; pageScale: number; clip: ClipPath[] }>();

  /** world → output */
  const toOut = (x: number, y: number): [number, number] => {
    const [wx, wy] = app(world, x, y);
    const k = pageScale;
    return t.devToOut(wx * k * unitPx(pageUnit, dpiX), wy * k * unitPx(pageUnit, dpiY));
  };
  /** output units per world unit (lengths) */
  const scaleW = (): number => Math.sqrt(Math.abs(world[0] * world[3] - world[1] * world[2])) * pageScale * unitPx(pageUnit, dpiX) * t.devScale;
  const fmt = (n: number): string => String(Math.round(n * 100) / 100);
  const P = (x: number, y: number): string => {
    const [a, b] = toOut(x, y);
    return `${fmt(a)} ${fmt(b)}`;
  };

  const readBrush = (p: number, end: number): Brush => {
    const type = p + 8 <= end ? u32(p + 4) : 0;
    const d = p + 8;
    if (type === 0 || d + 4 > end) return argb(d + 4 <= end ? u32(d) : 0xff000000);
    if (type === 1) return argb(u32(d + 4)); // hatch: its foreground
    if (type === 4 && d + 32 <= end) {
      const rect: [number, number, number, number] = [f32(d + 8), f32(d + 12), f32(d + 16), f32(d + 20)];
      const s = argb(u32(d + 24));
      const e = argb(u32(d + 28));
      const avg = [0, 2, 4].map((k) => Math.round((parseInt(s.color.slice(k, k + 2), 16) + parseInt(e.color.slice(k, k + 2), 16)) / 2).toString(16).padStart(2, '0')).join('').toUpperCase();
      return { color: avg, alpha: (s.alpha + e.alpha) / 2, lin: { rect, c0: s.color, c1: e.color, a0: s.alpha, a1: e.alpha } };
    }
    if (type === 3 && d + 12 <= end) return argb(u32(d + 4)); // path gradient: its centre colour
    return { color: '808080', alpha: 1 };
  };

  const readObject = (type: number, p: number, end: number): Obj => {
    switch (type) {
      case 1:
        return readBrush(p, end);
      case 2: {
        // pen: version, type, flags, unit, width, optional data, then its brush
        const flags = u32(p + 8);
        const unit = u32(p + 12);
        const width = f32(p + 16);
        let q = p + 20;
        let dash: number[] | undefined;
        let style = 0;
        if (flags & 0x1) q += 24;
        if (flags & 0x2) q += 4;
        if (flags & 0x4) q += 4;
        if (flags & 0x8) q += 4;
        if (flags & 0x10) q += 4;
        if (flags & 0x20) {
          style = i32(q);
          q += 4;
        }
        if (flags & 0x40) q += 4;
        if (flags & 0x80) q += 4;
        if (flags & 0x100) {
          const n = u32(q);
          dash = Array.from({ length: Math.min(n, 32) }, (_, k) => f32(q + 4 + k * 4));
          q += 4 + n * 4;
        }
        if (flags & 0x200) q += 4;
        if (flags & 0x400) q += 4 + u32(q) * 4;
        if (flags & 0x800) q += 4 + u32(q);
        if (flags & 0x1000) q += 4 + u32(q);
        if (!dash && style >= 1 && style <= 4) dash = [[3, 1], [1, 1], [3, 1, 1, 1], [3, 1, 1, 1, 1, 1]][style - 1];
        return { width, world: unit === 0, dash, brush: q < end ? readBrush(q, end) : { color: '000000', alpha: 1 } };
      }
      case 3: {
        // path: version, count, flags, points (int16 or float), types
        const count = Math.min(u32(p + 4), 1_000_000);
        const flags = u32(p + 8);
        if (flags & 0x1800) return null; // relative / run-length encoded points: not supported
        const compressed = !!(flags & 0x4000);
        let q = p + 12;
        const pts: [number, number][] = [];
        for (let k = 0; k < count && q + (compressed ? 4 : 8) <= end; k++) {
          pts.push(compressed ? [i16(q), i16(q + 2)] : [f32(q), f32(q + 4)]);
          q += compressed ? 4 : 8;
        }
        const types = Array.from({ length: pts.length }, (_, k) => (q + k < end ? all[q + k] : 1));
        return { pts, types };
      }
      case 5: {
        // image: bitmap (pixels or compressed file) only
        if (u32(p + 4) !== 1) return null;
        const w = i32(p + 8);
        const h = i32(p + 12);
        const stride = i32(p + 16);
        const format = u32(p + 20);
        const kind = u32(p + 24);
        const data = all.subarray(p + 28, end);
        if (kind === 1) {
          const png = data[0] === 0x89 && data[1] === 0x50;
          const jpg = data[0] === 0xff && data[1] === 0xd8;
          if (!png && !jpg) return null;
          // a PNG ends with its IEND chunk (records are padded after it)
          let n = data.length;
          if (png)
            for (let q = 8; q + 12 <= data.length; ) {
              const len = ((data[q] << 24) | (data[q + 1] << 16) | (data[q + 2] << 8) | data[q + 3]) >>> 0;
              const isEnd = data[q + 4] === 0x49 && data[q + 5] === 0x45 && data[q + 6] === 0x4e && data[q + 7] === 0x44;
              q += 12 + len;
              if (isEnd) {
                n = Math.min(q, data.length);
                break;
              }
            }
          return { png: data.slice(0, n), w, h };
        }
        const bpp = format === 0x0026200a || format === 0x000e200b || format === 0x00022009 ? 4 : format === 0x00021808 ? 3 : 0;
        if (!bpp || w <= 0 || h <= 0 || w * h > 40_000_000 || Math.abs(stride) * h > data.length) return null;
        const rgba = new Uint8Array(w * h * 4);
        for (let y = 0; y < h; y++)
          for (let x = 0; x < w; x++) {
            const s = (stride < 0 ? (h - 1 - y) * -stride : y * stride) + x * bpp;
            const d = (y * w + x) * 4;
            rgba[d] = data[s + 2];
            rgba[d + 1] = data[s + 1];
            rgba[d + 2] = data[s];
            rgba[d + 3] = bpp === 4 && format !== 0x00022009 ? data[s + 3] : 255;
          }
        return { png: encodePng(rgba, w, h, 4), w, h };
      }
      case 6: {
        const size = f32(p + 4);
        const unit = u32(p + 8);
        const style = i32(p + 12);
        const len = u32(p + 20);
        let family = '';
        for (let k = 0; k < len && p + 24 + k * 2 + 2 <= end; k++) family += String.fromCharCode(u16(p + 24 + k * 2));
        return { size, unit, bold: !!(style & 1), italic: !!(style & 2), underline: !!(style & 4), family };
      }
      case 7:
        return { align: u32(p + 12), lineAlign: u32(p + 16) };
      default:
        return null;
    }
  };

  const brushOf = (flags: number, id: number): Brush | null => {
    if (flags & 0x8000) return argb(id);
    const b = objects[id & 63];
    return b && 'color' in b ? b : null;
  };
  const penOf = (id: number): Pen | null => {
    const v = objects[id & 63];
    return v && 'brush' in v ? v : null;
  };

  const gradientOf = (b: Brush): Gradient | undefined => {
    if (!b.lin) return undefined;
    const [x, y, w, h] = b.lin.rect;
    const [x0, y0] = toOut(x, y + h / 2);
    const [x1, y1] = toOut(x + w, y + h / 2);
    return { kind: 'linear', coords: [x0, y0, x1, y1], matrix: [1, 0, 0, 1, 0, 0], stops: [{ pos: 0, color: b.lin.c0, ...(b.lin.a0 < 1 ? { alpha: b.lin.a0 } : {}) }, { pos: 1, color: b.lin.c1, ...(b.lin.a1 < 1 ? { alpha: b.lin.a1 } : {}) }] };
  };
  const fill = (d: string, b: Brush | null, evenOdd = true): void => {
    if (!b || !d || b.alpha <= 0) return;
    const grad = gradientOf(b);
    cmds.push({ k: 'path', d, fill: b.color, evenOdd, width: 0, ...(b.alpha < 1 ? { opacity: b.alpha } : {}), ...(grad ? { grad } : {}) });
  };
  const stroke = (d: string, pen: Pen | null): void => {
    if (!pen || !d || pen.brush.alpha <= 0) return;
    const w = pen.world ? pen.width * scaleW() : pen.width * t.devScale;
    const width = Math.max(w, 0.75 * t.devScale);
    cmds.push({ k: 'path', d, stroke: pen.brush.color, evenOdd: false, width, ...(pen.dash ? { dash: pen.dash.map((v) => v * width) } : {}), ...(pen.brush.alpha < 1 ? { strokeOpacity: pen.brush.alpha } : {}) });
  };

  const rectD = (x: number, y: number, w: number, h: number): string => `M${P(x, y)} L${P(x + w, y)} L${P(x + w, y + h)} L${P(x, y + h)} Z`;
  /** An elliptical arc as cubic Béziers (angles in degrees, clockwise from +x as GDI+ draws them). */
  const arcD = (x: number, y: number, w: number, h: number, start: number, sweep: number, move: boolean): string => {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    const n = Math.max(1, Math.ceil(Math.abs(sweep) / 90));
    const step = ((sweep / n) * Math.PI) / 180;
    let a = (start * Math.PI) / 180;
    const k = (4 / 3) * Math.tan(step / 4);
    let d = move ? `M${P(cx + rx * Math.cos(a), cy + ry * Math.sin(a))}` : `L${P(cx + rx * Math.cos(a), cy + ry * Math.sin(a))}`;
    for (let i = 0; i < n; i++) {
      const b = a + step;
      const p1: [number, number] = [cx + rx * (Math.cos(a) - k * Math.sin(a)), cy + ry * (Math.sin(a) + k * Math.cos(a))];
      const p2: [number, number] = [cx + rx * (Math.cos(b) + k * Math.sin(b)), cy + ry * (Math.sin(b) - k * Math.cos(b))];
      d += ` C${P(...p1)} ${P(...p2)} ${P(cx + rx * Math.cos(b), cy + ry * Math.sin(b))}`;
      a = b;
    }
    return d;
  };
  const ellipseD = (x: number, y: number, w: number, h: number): string => `${arcD(x, y, w, h, 0, 360, true)} Z`;
  const polyD = (pts: [number, number][], close: boolean): string => (pts.length ? `M${pts.map(([x, y]) => P(x, y)).join(' L')}${close ? ' Z' : ''}` : '');
  const pathD = (path: PathObj): string => {
    let d = '';
    for (let k = 0; k < path.pts.length; k++) {
      const ty = path.types[k] & 7;
      if (ty === 3 && k + 2 < path.pts.length) {
        d += ` C${P(...path.pts[k])} ${P(...path.pts[k + 1])} ${P(...path.pts[k + 2])}`;
        k += 2;
        if (path.types[k] & 0x80) d += ' Z';
        continue;
      }
      d += `${ty === 0 ? ' M' : ' L'}${P(...path.pts[k])}`;
      if (path.types[k] & 0x80) d += ' Z';
    }
    return d.trim();
  };
  /** Cardinal spline through the points (GDI+ curves), as Béziers. */
  const curveD = (pts: [number, number][], tension: number, closed: boolean): string => {
    if (pts.length < 2) return '';
    const n = pts.length;
    const at = (k: number): [number, number] => (closed ? pts[((k % n) + n) % n] : pts[Math.max(0, Math.min(n - 1, k))]);
    const segs = closed ? n : n - 1;
    const tt = tension / 3;
    let d = `M${P(...pts[0])}`;
    for (let k = 0; k < segs; k++) {
      const p0 = at(k - 1);
      const p1 = at(k);
      const p2 = at(k + 1);
      const p3 = at(k + 2);
      const c1: [number, number] = [p1[0] + (p2[0] - p0[0]) * tt, p1[1] + (p2[1] - p0[1]) * tt];
      const c2: [number, number] = [p2[0] - (p3[0] - p1[0]) * tt, p2[1] - (p3[1] - p1[1]) * tt];
      d += ` C${P(...c1)} ${P(...c2)} ${P(...p2)}`;
    }
    return closed ? `${d} Z` : d;
  };

  const readRects = (p: number, count: number, compressed: boolean): [number, number, number, number][] =>
    Array.from({ length: count }, (_, k) => (compressed ? [i16(p + k * 8), i16(p + k * 8 + 2), i16(p + k * 8 + 4), i16(p + k * 8 + 6)] : [f32(p + k * 16), f32(p + k * 16 + 4), f32(p + k * 16 + 8), f32(p + k * 16 + 12)]));
  const readPoints = (p: number, count: number, compressed: boolean): [number, number][] =>
    Array.from({ length: count }, (_, k) => (compressed ? [i16(p + k * 4), i16(p + k * 4 + 2)] : [f32(p + k * 8), f32(p + k * 8 + 4)]));
  const pointsFit = (p: number, count: number, compressed: boolean, end: number): boolean => count >= 0 && count < 1_000_000 && p + count * (compressed ? 4 : 8) <= end;

  const setClip = (d: string | null, mode: number, evenOdd = true): void => {
    if (d === null) clip = [];
    else if (mode === 0) clip = [{ d, evenOdd }];
    else if (mode === 1) clip = [...clip, { d, evenOdd }];
    else if (mode === 4) clip = [...clip, { d: `${FAR} ${d}`, evenOdd: true }];
    else if (mode === 2 && clip.length === 1) clip = [{ d: `${clip[0].d} ${d}`, evenOdd: false }];
    // union of several / xor / complement: kept as it was
    cmds.push({ k: 'clip', paths: [...clip] });
  };

  const text = (str: string, x: number, y: number, font: FontObj, b: Brush, anchor: 'start' | 'middle' | 'end', baseline: 'alphabetic' | 'top' | 'bottom'): void => {
    if (!str || b.alpha <= 0) return;
    const [ox, oy] = toOut(x, y);
    const [ax, ay] = app(world, 1, 0);
    const angle = (-Math.atan2(ay - world[5], ax - world[4]) * 180) / Math.PI;
    const size = font.unit === 0 ? font.size * scaleW() : font.size * unitPx(font.unit, dpiY) * t.devScale;
    cmds.push({ k: 'text', x: ox, y: oy, text: str, size, family: familyOf(font.family) ?? 'sans', bold: font.bold, italic: font.italic, underline: font.underline, color: b.color, angle: Math.abs(angle) < 0.01 ? 0 : angle, anchor, baseline });
  };

  let pos = 0;
  let guard = 0;
  while (pos + 12 <= all.length && guard++ < 2_000_000) {
    const type = u16(pos);
    const flags = u16(pos + 2);
    const size = u32(pos + 4);
    const dataSize = u32(pos + 8);
    if (size < 12 || pos + size > all.length) break;
    const p = pos + 12;
    const end = Math.min(p + dataSize, pos + size);
    const C = !!(flags & 0x4000);
    const id = flags & 0xff;
    switch (type) {
      case 0x4001: // header
        header = true;
        dual = !!(flags & 1);
        if (dataSize >= 16) {
          dpiX = u32(p + 8) || 96;
          dpiY = u32(p + 12) || 96;
        }
        break;
      case 0x4002: // end of file
        pos = all.length;
        continue;
      case 0x4008:
        // an object (records continuing one were joined by mergeObjects)
        objects[id & 63] = readObject((flags >> 8) & 0x7f, p, end);
        break;
      case 0x4009: {
        // clear: the whole picture
        const b = argb(u32(p));
        const [x0, y0, x1, y1] = t.frame;
        if (b.alpha > 0) cmds.push({ k: 'path', d: `M${fmt(x0)} ${fmt(y0)} L${fmt(x1)} ${fmt(y0)} L${fmt(x1)} ${fmt(y1)} L${fmt(x0)} ${fmt(y1)} Z`, fill: b.color, evenOdd: false, width: 0, ...(b.alpha < 1 ? { opacity: b.alpha } : {}) });
        break;
      }
      case 0x400a: {
        const b = brushOf(flags, u32(p));
        const n = u32(p + 4);
        if (p + 8 + n * (C ? 8 : 16) > end) break;
        fill(readRects(p + 8, n, C).map((r) => rectD(...r)).join(' '), b, false);
        break;
      }
      case 0x400b: {
        const n = u32(p);
        if (p + 4 + n * (C ? 8 : 16) > end) break;
        stroke(readRects(p + 4, n, C).map((r) => rectD(...r)).join(' '), penOf(id));
        break;
      }
      case 0x400c: {
        const n = u32(p + 4);
        if (flags & 0x0800 || !pointsFit(p + 8, n, C, end)) break;
        fill(polyD(readPoints(p + 8, n, C), true), brushOf(flags, u32(p)));
        break;
      }
      case 0x400d: {
        const n = u32(p);
        if (flags & 0x0800 || !pointsFit(p + 4, n, C, end)) break;
        stroke(polyD(readPoints(p + 4, n, C), !!(flags & 0x2000)), penOf(id));
        break;
      }
      case 0x400e: {
        const [r] = readRects(p + 4, 1, C);
        fill(ellipseD(...r), brushOf(flags, u32(p)));
        break;
      }
      case 0x400f: {
        const [r] = readRects(p, 1, C);
        stroke(ellipseD(...r), penOf(id));
        break;
      }
      case 0x4010: {
        const [r] = readRects(p + 12, 1, C);
        const cx = r[0] + r[2] / 2;
        const cy = r[1] + r[3] / 2;
        fill(`M${P(cx, cy)} ${arcD(...r, f32(p + 4), f32(p + 8), false)} Z`, brushOf(flags, u32(p)));
        break;
      }
      case 0x4011:
      case 0x4012: {
        const [r] = readRects(p + 8, 1, C);
        const cx = r[0] + r[2] / 2;
        const cy = r[1] + r[3] / 2;
        stroke(type === 0x4011 ? `M${P(cx, cy)} ${arcD(...r, f32(p), f32(p + 4), false)} Z` : arcD(...r, f32(p), f32(p + 4), true), penOf(id));
        break;
      }
      case 0x4014: {
        const path = objects[id & 63];
        if (path && 'pts' in path) fill(pathD(path), brushOf(flags, u32(p)));
        break;
      }
      case 0x4015: {
        const path = objects[id & 63];
        if (path && 'pts' in path) stroke(pathD(path), penOf(u32(p)));
        break;
      }
      case 0x4016: {
        // fill closed curve: brush, tension, count, points
        const n = u32(p + 8);
        if (flags & 0x0800 || !pointsFit(p + 12, n, C, end)) break;
        fill(curveD(readPoints(p + 12, n, C), f32(p + 4), true), brushOf(flags, u32(p)), !(flags & 0x2000));
        break;
      }
      case 0x4017: {
        const n = u32(p + 4);
        if (flags & 0x0800 || !pointsFit(p + 8, n, C, end)) break;
        stroke(curveD(readPoints(p + 8, n, C), f32(p), true), penOf(id));
        break;
      }
      case 0x4018: {
        // draw curve: tension, offset, segments, count, points
        const n = u32(p + 12);
        if (!pointsFit(p + 16, n, C, end)) break;
        const pts = readPoints(p + 16, n, C);
        const off = u32(p + 4);
        const segs = u32(p + 8);
        stroke(curveD(pts.slice(off, off + segs + 1), f32(p), false), penOf(id));
        break;
      }
      case 0x4019: {
        const n = u32(p);
        if (flags & 0x0800 || !pointsFit(p + 4, n, C, end) || n < 4) break;
        const pts = readPoints(p + 4, n, C);
        let d = `M${P(...pts[0])}`;
        for (let k = 1; k + 2 < pts.length; k += 3) d += ` C${P(...pts[k])} ${P(...pts[k + 1])} ${P(...pts[k + 2])}`;
        stroke(d, penOf(id));
        break;
      }
      case 0x401a:
      case 0x401b: {
        const img = objects[id & 63];
        if (!img || !('png' in img)) break;
        let corners: [number, number][];
        if (type === 0x401a) {
          const [r] = readRects(p + 24, 1, C);
          corners = [
            [r[0], r[1]],
            [r[0] + r[2], r[1]],
            [r[0], r[1] + r[3]],
          ];
        } else {
          if (flags & 0x0800 || u32(p + 24) !== 3 || !pointsFit(p + 28, 3, C, end)) break;
          corners = readPoints(p + 28, 3, C);
        }
        const [a, b, c] = corners.map(([x, y]) => toOut(x, y));
        // drawn upright: the box of the parallelogram (rotated pictures are rare in documents)
        const xs = [a[0], b[0], c[0], b[0] + c[0] - a[0]];
        const ys = [a[1], b[1], c[1], b[1] + c[1] - a[1]];
        const x0 = Math.min(...xs);
        const y0 = Math.min(...ys);
        cmds.push({ k: 'image', x: x0, y: y0, w: Math.max(...xs) - x0, h: Math.max(...ys) - y0, png: img.png });
        break;
      }
      case 0x401c: {
        // draw string: brush, format, length, layout rect, UTF-16 text
        const font = objects[id & 63];
        const b = brushOf(flags, u32(p));
        if (!font || !('family' in font) || !b) break;
        const format = objects[u32(p + 4) & 63];
        const len = u32(p + 8);
        const [x, y, w, h] = [f32(p + 12), f32(p + 16), f32(p + 20), f32(p + 24)];
        let str = '';
        for (let k = 0; k < len && p + 28 + k * 2 + 2 <= end; k++) str += String.fromCharCode(u16(p + 28 + k * 2));
        const al = format && 'align' in format ? format.align : 0;
        const la = format && 'align' in format ? format.lineAlign : 0;
        const lines = str.split(/\r\n|\n|\r/);
        const lh = font.size * 1.15;
        const blockH = lh * lines.length;
        const top = h > 0 ? (la === 1 ? y + (h - blockH) / 2 : la === 2 ? y + h - blockH : y) : y;
        lines.forEach((line, k) => {
          const lx = w > 0 ? (al === 1 ? x + w / 2 : al === 2 ? x + w : x) : x;
          text(line, lx, top + k * lh, font, b, al === 1 ? 'middle' : al === 2 ? 'end' : 'start', 'top');
        });
        break;
      }
      case 0x4036: {
        // draw driver string: brush, options, matrix present, glyph count, glyphs, positions
        const font = objects[id & 63];
        const b = brushOf(flags, u32(p));
        if (!font || !('family' in font) || !b) break;
        const opts = u32(p + 4);
        const n = u32(p + 12);
        if (!(opts & 1) || p + 16 + n * 10 > end) break; // glyph indices (not characters): cannot be mapped
        let str = '';
        for (let k = 0; k < n; k++) str += String.fromCharCode(u16(p + 16 + k * 2));
        const posAt = (k: number): [number, number] => [f32(p + 16 + n * 2 + k * 8), f32(p + 16 + n * 2 + k * 8 + 4)];
        if (opts & 4 || n === 1) text(str, ...posAt(0), font, b, 'start', 'alphabetic');
        else {
          // one run per stretch of evenly advancing glyphs would be ideal; per glyph keeps the positions exact
          for (let k = 0; k < n; k++) if (str[k].trim()) text(str[k], ...posAt(k), font, b, 'start', 'alphabetic');
        }
        break;
      }
      case 0x4025:
      case 0x4028: // save, begin container (no parameters)
        saved.set(u32(p), { world, pageUnit, pageScale, clip });
        break;
      case 0x4027: {
        // begin container: source rectangle mapped onto the destination
        const [dx, dy, dw, dh] = [f32(p), f32(p + 4), f32(p + 8), f32(p + 12)];
        const [sx, sy, sw, sh] = [f32(p + 16), f32(p + 20), f32(p + 24), f32(p + 28)];
        saved.set(u32(p + 32), { world, pageUnit, pageScale, clip });
        if (sw && sh) world = mul([dw / sw, 0, 0, dh / sh, dx - (sx * dw) / sw, dy - (sy * dh) / sh], world);
        break;
      }
      case 0x4026:
      case 0x4029: {
        const s = saved.get(u32(p));
        if (s) {
          ({ world, pageUnit, pageScale } = s);
          if (s.clip !== clip) {
            clip = s.clip;
            cmds.push({ k: 'clip', paths: [...clip] });
          }
        }
        break;
      }
      case 0x402a:
        world = [f32(p), f32(p + 4), f32(p + 8), f32(p + 12), f32(p + 16), f32(p + 20)];
        break;
      case 0x402b:
        world = ID;
        break;
      case 0x402c: {
        const m: Mx = [f32(p), f32(p + 4), f32(p + 8), f32(p + 12), f32(p + 16), f32(p + 20)];
        world = flags & 0x2000 ? mul(world, m) : mul(m, world);
        break;
      }
      case 0x402d: {
        const m: Mx = [1, 0, 0, 1, f32(p), f32(p + 4)];
        world = flags & 0x2000 ? mul(world, m) : mul(m, world);
        break;
      }
      case 0x402e: {
        const m: Mx = [f32(p), 0, 0, f32(p + 4), 0, 0];
        world = flags & 0x2000 ? mul(world, m) : mul(m, world);
        break;
      }
      case 0x402f: {
        const a = (f32(p) * Math.PI) / 180;
        const m: Mx = [Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0];
        world = flags & 0x2000 ? mul(world, m) : mul(m, world);
        break;
      }
      case 0x4030:
        pageUnit = flags & 0xff;
        pageScale = f32(p) || 1;
        break;
      case 0x4031:
        setClip(null, 0);
        break;
      case 0x4032: {
        const [r] = readRects(p, 1, false);
        setClip(rectD(...r), (flags >> 8) & 0xf, false);
        break;
      }
      case 0x4033: {
        const path = objects[id & 63];
        if (path && 'pts' in path) setClip(pathD(path), (flags >> 8) & 0xf);
        break;
      }
      default:
        break;
    }
    pos += size;
  }

  return header ? { cmds, dual } : null;
}

/**
 * The records of all comments in one buffer, an object continued over several records (flag 0x8000,
 * the first one giving the total size) joined into a single record.
 */
function mergeObjects(blobs: Uint8Array[]): Uint8Array {
  const total = blobs.reduce((n, b) => n + b.length, 0);
  const src = new Uint8Array(total);
  let o = 0;
  for (const b of blobs) {
    src.set(b, o);
    o += b.length;
  }
  const dv = new DataView(src.buffer, src.byteOffset, src.byteLength);
  const out: Uint8Array[] = [];
  let open: { flags: number; size: number; parts: Uint8Array[]; got: number } | null = null;
  const close = (): void => {
    if (!open) return;
    const len = open.parts.reduce((n, x) => n + x.length, 0);
    const rec = new Uint8Array(12 + len);
    const v = new DataView(rec.buffer);
    v.setUint16(0, 0x4008, true);
    v.setUint16(2, open.flags & 0x7fff, true);
    v.setUint32(4, 12 + len, true);
    v.setUint32(8, len, true);
    let q = 12;
    for (const x of open.parts) {
      rec.set(x, q);
      q += x.length;
    }
    out.push(rec);
    open = null;
  };
  for (let pos = 0; pos + 12 <= src.length; ) {
    const type = dv.getUint16(pos, true);
    const flags = dv.getUint16(pos + 2, true);
    const size = dv.getUint32(pos + 4, true);
    const dataSize = dv.getUint32(pos + 8, true);
    if (size < 12 || pos + size > src.length) break;
    const data = src.subarray(pos + 12, pos + 12 + Math.min(dataSize, size - 12));
    if (type === 0x4008 && flags & 0x8000) {
      // continued: the first piece starts with the object's total size
      if (!open || (open.flags & 0xff) !== (flags & 0xff)) {
        close();
        open = { flags, size: data.length >= 4 ? new DataView(data.buffer, data.byteOffset).getUint32(0, true) : 0, parts: [data.subarray(4)], got: data.length - 4 };
      } else {
        open.parts.push(data);
        open.got += data.length;
      }
      if (open.got >= open.size) close();
    } else if (type === 0x4008 && open && (open.flags & 0xff) === (flags & 0xff)) {
      // the last piece
      open.parts.push(data);
      close();
    } else {
      close();
      out.push(src.subarray(pos, pos + size));
    }
    pos += size;
  }
  close();
  const len = out.reduce((n, x) => n + x.length, 0);
  const merged = new Uint8Array(len);
  let q = 0;
  for (const x of out) {
    merged.set(x, q);
    q += x.length;
  }
  return merged;
}
