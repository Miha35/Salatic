/**
 * Vector display list → Enhanced Metafile (EMF), the vector picture format every Word version and
 * LibreOffice render natively. Used as the picture for SVG images and charts in DOCX and RTF.
 *
 * Mapping: MM_ANISOTROPIC with a window of `width·S × height·S` logical units onto a 96-dpi
 * viewport of the drawing's physical size, so coordinates keep 1/S precision.
 * Limits of plain EMF: no transparency (colours are blended with white) and no font embedding
 * (text uses the named system font).
 */
import type { Metafile, MetaCmd } from '../metafile';
import { decodePng } from '../png';

const S = 16;
const DPI = 96;
/** Reference device: 1920 px over 508 mm = 96 dpi in both directions. */
const DEV_PX = 1920;
const DEV_MM = 508;

const STOCK = { NULL_BRUSH: 0x80000005, BLACK_PEN: 0x80000007, NULL_PEN: 0x80000008, SYSTEM_FONT: 0x8000000d, WHITE_BRUSH: 0x80000000 };
const FACE: Record<string, string> = { sans: 'Arial', serif: 'Times New Roman', mono: 'Courier New', calibri: 'Calibri' };

class RecordWriter {
  private chunks: Uint8Array[] = [];
  bytes = 0;
  count = 0;

  rec(type: number, body: (dv: DataView, u8: Uint8Array) => void, size: number): void {
    const n = (size + 3) & ~3;
    const u8 = new Uint8Array(n);
    const dv = new DataView(u8.buffer);
    dv.setUint32(0, type, true);
    dv.setUint32(4, n, true);
    body(dv, u8);
    this.chunks.push(u8);
    this.bytes += n;
    this.count++;
  }

  /** A record whose body is a list of 32-bit integers. */
  ints(type: number, ...vals: number[]): void {
    this.rec(type, (dv) => vals.forEach((v, i) => dv.setInt32(8 + i * 4, v, true)), 8 + vals.length * 4);
  }

  concat(head: Uint8Array): Uint8Array {
    const out = new Uint8Array(head.length + this.bytes);
    out.set(head, 0);
    let o = head.length;
    for (const c of this.chunks) {
      out.set(c, o);
      o += c.length;
    }
    return out;
  }
}

const colorRef = (hex: string, alpha = 1): number => {
  let r = parseInt(hex.slice(0, 2), 16) || 0;
  let g = parseInt(hex.slice(2, 4), 16) || 0;
  let b = parseInt(hex.slice(4, 6), 16) || 0;
  if (alpha < 1) {
    const a = Math.max(0, alpha);
    r = Math.round(r * a + 255 * (1 - a));
    g = Math.round(g * a + 255 * (1 - a));
    b = Math.round(b * a + 255 * (1 - a));
  }
  return r | (g << 8) | (b << 16);
};

/** Absolute M/L/C/Z path data (the display-list format) → subpaths. */
function parseD(d: string): { op: 'M' | 'L' | 'C' | 'Z'; pts: number[] }[] {
  const out: { op: 'M' | 'L' | 'C' | 'Z'; pts: number[] }[] = [];
  const re = /([MLCZmlcz])|(-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)/g;
  let cur: { op: 'M' | 'L' | 'C' | 'Z'; pts: number[] } | null = null;
  for (let m; (m = re.exec(d)); ) {
    if (m[1]) {
      cur = { op: m[1].toUpperCase() as 'M' | 'L' | 'C' | 'Z', pts: [] };
      out.push(cur);
    } else if (cur) cur.pts.push(+m[2]);
  }
  return out;
}

export function metafileToEmf(mf: Metafile): Uint8Array {
  const w = new RecordWriter();
  const L = (v: number): number => Math.round(v * S);
  const widthPt = mf.widthPt > 0 ? mf.widthPt : (mf.width * 72) / 96;
  const heightPt = mf.heightPt > 0 ? mf.heightPt : (mf.height * 72) / 96;
  const devW = Math.max(1, Math.round((widthPt * DPI) / 72));
  const devH = Math.max(1, Math.round((heightPt * DPI) / 72));
  const winW = Math.max(1, L(mf.width));
  const winH = Math.max(1, L(mf.height));

  w.ints(17, 8); // SETMAPMODE MM_ANISOTROPIC
  w.ints(10, 0, 0); // SETWINDOWORGEX
  w.ints(9, winW, winH); // SETWINDOWEXTEX
  w.ints(12, 0, 0); // SETVIEWPORTORGEX
  w.ints(11, devW, devH); // SETVIEWPORTEXTEX
  w.ints(18, 1); // SETBKMODE TRANSPARENT

  // Current GDI objects (handles: 1 brush, 2 pen, 3 font)
  let brushKey = '';
  let penKey = '';
  let fontKey = '';
  let polyFill = -1;
  let textColor = -1;
  let textAlign = -1;
  const useBrush = (color: number): void => {
    const key = String(color);
    if (key === brushKey) return;
    if (brushKey) {
      w.ints(37, STOCK.WHITE_BRUSH);
      w.ints(40, 1);
    }
    w.ints(39, 1, 0, color, 0); // CREATEBRUSHINDIRECT BS_SOLID
    w.ints(37, 1);
    brushKey = key;
  };
  const usePen = (color: number, width: number, dash?: number[]): void => {
    const style = dash?.length ? dash.map((x) => Math.max(1, L(x))) : [];
    const key = `${color}/${width}/${style.join(',')}`;
    if (key === penKey) return;
    if (penKey) {
      w.ints(37, STOCK.BLACK_PEN);
      w.ints(40, 2);
    }
    // EXTCREATEPEN: PS_GEOMETRIC | PS_ENDCAP_FLAT | PS_JOIN_MITER | (PS_USERSTYLE or PS_SOLID)
    const penStyle = 0x10000 | 0x200 | 0x2000 | (style.length ? 7 : 0);
    w.ints(95, 2, 0, 0, 0, 0, penStyle, width, 0, color, 0, style.length, ...style);
    w.ints(37, 2);
    penKey = key;
  };
  const useFont = (c: Extract<MetaCmd, { k: 'text' }>): void => {
    const face = FACE[c.family] ?? 'Arial';
    const height = -Math.max(1, L(c.size));
    const esc = Math.round((c.angle || 0) * 10);
    const key = `${face}/${height}/${esc}/${+c.bold}${+c.italic}${+c.underline}`;
    if (key === fontKey) return;
    if (fontKey) {
      w.ints(37, STOCK.SYSTEM_FONT);
      w.ints(40, 3);
    }
    // EXTCREATEFONTINDIRECTW with a LogFontPanose (320 bytes), the form GDI itself writes
    w.rec(
      82,
      (dv, u8) => {
        dv.setUint32(8, 3, true);
        const o = 12;
        dv.setInt32(o, height, true);
        dv.setInt32(o + 8, esc, true);
        dv.setInt32(o + 12, esc, true);
        dv.setInt32(o + 16, c.bold ? 700 : 400, true);
        u8[o + 20] = c.italic ? 1 : 0;
        u8[o + 21] = c.underline ? 1 : 0;
        u8[o + 23] = 1; // DEFAULT_CHARSET
        u8[o + 26] = 4; // ANTIALIASED_QUALITY
        for (let i = 0; i < Math.min(31, face.length); i++) dv.setUint16(o + 28 + i * 2, face.charCodeAt(i), true);
      },
      12 + 320,
    );
    w.ints(37, 3);
    fontKey = key;
  };

  /** Emit a path's geometry inside BEGINPATH … ENDPATH. */
  const geometry = (d: string): boolean => {
    let any = false;
    w.ints(59); // BEGINPATH
    for (const seg of parseD(d)) {
      const p = seg.pts;
      if (seg.op === 'M') {
        for (let i = 0; i + 1 < p.length; i += 2) {
          w.ints(i === 0 ? 27 : 54, L(p[i]), L(p[i + 1])); // MOVETOEX, then implicit LINETO
          any ||= i > 0;
        }
      } else if (seg.op === 'L') {
        for (let i = 0; i + 1 < p.length; i += 2) w.ints(54, L(p[i]), L(p[i + 1]));
        any ||= p.length >= 2;
      } else if (seg.op === 'C') {
        const n = Math.floor(p.length / 6) * 3;
        if (!n) continue;
        const xs = p.slice(0, n * 2).map(L);
        // POLYBEZIERTO: bounds, count, points
        w.ints(5, Math.min(...xs.filter((_, i) => i % 2 === 0)), Math.min(...xs.filter((_, i) => i % 2 === 1)), Math.max(...xs.filter((_, i) => i % 2 === 0)), Math.max(...xs.filter((_, i) => i % 2 === 1)), n, ...xs);
        any = true;
      } else if (seg.op === 'Z') w.ints(61); // CLOSEFIGURE
    }
    w.ints(60); // ENDPATH
    return any;
  };

  const cmds = mf.cmds;
  for (let i = 0; i < cmds.length; i++) {
    const c = cmds[i];
    if (c.k === 'path') {
      const fill = c.fill !== undefined && c.fill !== 'none' && (c.opacity ?? 1) > 0.02;
      const stroke = c.stroke !== undefined && c.stroke !== 'none' && (c.strokeOpacity ?? 1) > 0.02;
      if (!fill && !stroke) continue;
      const mode = c.evenOdd ? 1 : 2;
      if (fill && mode !== polyFill) {
        w.ints(19, mode); // SETPOLYFILLMODE
        polyFill = mode;
      }
      if (fill) useBrush(colorRef(c.fill!, c.opacity ?? 1));
      if (stroke) usePen(colorRef(c.stroke!, c.strokeOpacity ?? 1), Math.max(1, L(c.width || 0.75)), c.dash);
      if (!geometry(c.d) && !fill) continue;
      // FILLPATH / STROKEPATH / STROKEANDFILLPATH with an (ignored) bounds rectangle
      w.ints(fill && stroke ? 63 : fill ? 62 : 64, 0, 0, winW, winH);
    } else if (c.k === 'text') {
      if (!c.text) continue;
      // A run followed by flowing runs (SVG tspans) is drawn at the current position, which GDI advances
      const chain = !c.flow && cmds[i + 1]?.k === 'text' && (cmds[i + 1] as Extract<MetaCmd, { k: 'text' }>).flow;
      const flowing = c.flow || chain;
      const h = c.anchor === 'middle' ? 6 : c.anchor === 'end' ? 2 : 0;
      const v = c.baseline === 'top' ? 0 : c.baseline === 'bottom' ? 8 : 24;
      const align = flowing ? 1 | v : h | v;
      useFont(c);
      const col = colorRef(c.color);
      if (col !== textColor) {
        w.ints(24, col);
        textColor = col;
      }
      if (align !== textAlign) {
        w.ints(22, align);
        textAlign = align;
      }
      let x = c.x;
      if (chain && h) {
        // estimate the chain's width to honour a centred/right anchor
        let wEm = 0;
        for (let j = i; j < cmds.length; j++) {
          const t = cmds[j];
          if (t.k !== 'text' || (j > i && !t.flow)) break;
          wEm += t.text.length * 0.52 * t.size;
        }
        x -= h === 6 ? wEm / 2 : wEm;
      }
      if (chain) w.ints(27, L(x), L(c.y)); // MOVETOEX to the chain start
      const units = [...c.text].flatMap((ch) => {
        const cp = ch.codePointAt(0)!;
        return cp > 0xffff ? [0xd800 + ((cp - 0x10000) >> 10), 0xdc00 + ((cp - 0x10000) & 0x3ff)] : [cp];
      });
      const xScale = (100 * 25.4 * devW) / DPI / winW;
      const yScale = (100 * 25.4 * devH) / DPI / winH;
      const strOff = 76;
      w.rec(
        84, // EXTTEXTOUTW
        (dv) => {
          dv.setUint32(24, 1, true); // GM_COMPATIBLE
          dv.setFloat32(28, xScale, true);
          dv.setFloat32(32, yScale, true);
          dv.setInt32(36, L(x), true);
          dv.setInt32(40, L(c.y), true);
          dv.setUint32(44, units.length, true);
          dv.setUint32(48, strOff, true);
          dv.setUint32(52, 0, true); // options
          dv.setUint32(72, 0, true); // offDx: let the player lay the glyphs out
          units.forEach((u, k) => dv.setUint16(strOff + k * 2, u, true));
        },
        strOff + units.length * 2,
      );
    } else if (c.k === 'image') {
      const img = decodePng(c.png);
      if (!img) continue;
      const { width, height, rgba } = img;
      const stride = (width * 3 + 3) & ~3;
      const bits = new Uint8Array(stride * height);
      for (let y = 0; y < height; y++) {
        const row = (height - 1 - y) * stride; // bottom-up DIB
        for (let x = 0; x < width; x++) {
          const s = (y * width + x) * 4;
          const a = rgba[s + 3] / 255;
          const o = row + x * 3;
          bits[o] = Math.round(rgba[s + 2] * a + 255 * (1 - a));
          bits[o + 1] = Math.round(rgba[s + 1] * a + 255 * (1 - a));
          bits[o + 2] = Math.round(rgba[s] * a + 255 * (1 - a));
        }
      }
      const offBmi = 80;
      const offBits = offBmi + 40;
      w.rec(
        81, // STRETCHDIBITS
        (dv, u8) => {
          const ints = [L(c.x), L(c.y), 0, 0, width, height, offBmi, 40, offBits, bits.length, 0, 0x00cc0020, L(c.w), L(c.h)];
          ints.forEach((v, k) => dv.setInt32(24 + k * 4, v, true));
          dv.setUint32(offBmi, 40, true);
          dv.setInt32(offBmi + 4, width, true);
          dv.setInt32(offBmi + 8, height, true);
          dv.setUint16(offBmi + 12, 1, true);
          dv.setUint16(offBmi + 14, 24, true);
          dv.setUint32(offBmi + 20, bits.length, true);
          u8.set(bits, offBits);
        },
        offBits + bits.length,
      );
    }
  }
  if (brushKey) {
    w.ints(37, STOCK.WHITE_BRUSH);
    w.ints(40, 1);
  }
  if (penKey) {
    w.ints(37, STOCK.BLACK_PEN);
    w.ints(40, 2);
  }
  if (fontKey) {
    w.ints(37, STOCK.SYSTEM_FONT);
    w.ints(40, 3);
  }
  w.ints(14, 0, 16, 20); // EOF: no palette, nSizeLast

  // Header
  const head = new Uint8Array(108);
  const dv = new DataView(head.buffer);
  const frameW = Math.round((widthPt * 2540) / 72);
  const frameH = Math.round((heightPt * 2540) / 72);
  dv.setUint32(0, 1, true);
  dv.setUint32(4, 108, true);
  [0, 0, devW - 1, devH - 1].forEach((v, i) => dv.setInt32(8 + i * 4, v, true));
  [0, 0, frameW, frameH].forEach((v, i) => dv.setInt32(24 + i * 4, v, true));
  dv.setUint32(40, 0x464d4520, true); // " EMF"
  dv.setUint32(44, 0x10000, true);
  dv.setUint32(48, 108 + w.bytes, true);
  dv.setUint32(52, w.count + 1, true);
  dv.setUint16(56, 4, true); // handles 0..3
  dv.setInt32(72, DEV_PX, true);
  dv.setInt32(76, DEV_PX, true);
  dv.setInt32(80, DEV_MM, true);
  dv.setInt32(84, DEV_MM, true);
  dv.setInt32(100, DEV_MM * 1000, true);
  dv.setInt32(104, DEV_MM * 1000, true);
  return w.concat(head);
}
