/** Small PNG encoder (8-bit gray/RGB/RGBA) built on fflate's zlib. Works in any JS runtime. */
import { unzlibSync, zlibSync } from 'fflate';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(data: Uint8Array, start: number, end: number): number {
  let c = 0xffffffff;
  for (let i = start; i < end; i++) c = CRC_TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  dv.setUint32(8 + data.length, crc32(out, 4, 8 + data.length));
  return out;
}

/**
 * Encode raw pixels as PNG.
 * @param channels 1 = gray, 3 = RGB, 4 = RGBA (8 bits per channel)
 */
export function encodePng(pixels: Uint8Array | Uint8ClampedArray, width: number, height: number, channels: 1 | 3 | 4): Uint8Array {
  const stride = width * channels;
  const raw = new Uint8Array((stride + 1) * height);
  // "Sub" filter on every row: cheap and compresses noticeably better than "None".
  for (let y = 0; y < height; y++) {
    const ro = y * (stride + 1);
    const so = y * stride;
    raw[ro] = 1;
    for (let x = 0; x < stride; x++) {
      const left = x >= channels ? pixels[so + x - channels] : 0;
      raw[ro + 1 + x] = (pixels[so + x] - left) & 0xff;
    }
  }
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, width);
  dv.setUint32(4, height);
  ihdr[8] = 8;
  ihdr[9] = channels === 1 ? 0 : channels === 3 ? 2 : 6;
  const sig = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const parts = [sig, chunk('IHDR', ihdr), chunk('IDAT', zlibSync(raw, { level: 6 })), chunk('IEND', new Uint8Array(0))];
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/**
 * Decode a PNG to 8-bit RGBA. Supports every colour type, bit depths 1–16, palettes with tRNS,
 * and Adam7 interlacing. Returns null for data that is not a valid PNG.
 */
export function decodePng(data: Uint8Array): { width: number; height: number; rgba: Uint8Array } | null {
  if (data.length < 33 || data[0] !== 0x89 || data[1] !== 0x50 || data[2] !== 0x4e || data[3] !== 0x47) return null;
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let width = 0;
  let height = 0;
  let depth = 8;
  let ctype = 6;
  let interlace = 0;
  let palette: Uint8Array | null = null;
  let trns: Uint8Array | null = null;
  const idat: Uint8Array[] = [];
  for (let p = 8; p + 12 <= data.length; ) {
    const len = dv.getUint32(p);
    const type = String.fromCharCode(data[p + 4], data[p + 5], data[p + 6], data[p + 7]);
    const body = data.subarray(p + 8, Math.min(data.length, p + 8 + len));
    if (type === 'IHDR') {
      width = dv.getUint32(p + 8);
      height = dv.getUint32(p + 12);
      depth = data[p + 16];
      ctype = data[p + 17];
      interlace = data[p + 20];
    } else if (type === 'PLTE') palette = body;
    else if (type === 'tRNS') trns = body;
    else if (type === 'IDAT') idat.push(body);
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  if (!width || !height || width * height > 50_000_000 || !idat.length) return null;
  const channels = ctype === 0 ? 1 : ctype === 2 ? 3 : ctype === 3 ? 1 : ctype === 4 ? 2 : ctype === 6 ? 4 : 0;
  if (!channels || ![1, 2, 4, 8, 16].includes(depth)) return null;
  let raw: Uint8Array;
  try {
    const all = new Uint8Array(idat.reduce((s, c) => s + c.length, 0));
    let o = 0;
    for (const c of idat) {
      all.set(c, o);
      o += c.length;
    }
    raw = unzlibSync(all);
  } catch {
    return null;
  }
  const bpp = Math.max(1, (channels * depth) >> 3);
  const rgba = new Uint8Array(width * height * 4);
  const max = (1 << Math.min(depth, 8)) - 1;
  const trnsGray = trns && ctype === 0 && trns.length >= 2 ? (trns[0] << 8) | trns[1] : -1;
  const trnsRgb = trns && ctype === 2 && trns.length >= 6 ? [(trns[0] << 8) | trns[1], (trns[2] << 8) | trns[3], (trns[4] << 8) | trns[5]] : null;
  let src = 0;
  const pass = (x0: number, y0: number, dx: number, dy: number): boolean => {
    const pw = Math.ceil((width - x0) / dx);
    const ph = Math.ceil((height - y0) / dy);
    if (pw <= 0 || ph <= 0) return true;
    const stride = Math.ceil((pw * channels * depth) / 8);
    let prev = new Uint8Array(stride);
    for (let r = 0; r < ph; r++) {
      if (src + 1 + stride > raw.length) return false;
      const filter = raw[src];
      const line = raw.slice(src + 1, src + 1 + stride);
      src += 1 + stride;
      for (let i = 0; i < stride; i++) {
        const a = i >= bpp ? line[i - bpp] : 0;
        const b = prev[i];
        const c = i >= bpp ? prev[i - bpp] : 0;
        let v = line[i];
        if (filter === 1) v += a;
        else if (filter === 2) v += b;
        else if (filter === 3) v += (a + b) >> 1;
        else if (filter === 4) {
          const pa = Math.abs(b - c);
          const pb = Math.abs(a - c);
          const pc = Math.abs(a + b - 2 * c);
          v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
        }
        line[i] = v & 0xff;
      }
      prev = line;
      const sample = (idx: number): number => {
        if (depth === 8) return line[idx];
        if (depth === 16) return (line[idx * 2] << 8) | line[idx * 2 + 1];
        const bit = idx * depth;
        return (line[bit >> 3] >> (8 - depth - (bit & 7))) & max;
      };
      const to8 = (v: number): number => (depth === 16 ? v >> 8 : depth === 8 ? v : Math.round((v * 255) / max));
      const y = y0 + r * dy;
      for (let k = 0; k < pw; k++) {
        const o = (y * width + x0 + k * dx) * 4;
        const s = k * channels;
        if (ctype === 3) {
          const idx = sample(s);
          rgba[o] = palette?.[idx * 3] ?? 0;
          rgba[o + 1] = palette?.[idx * 3 + 1] ?? 0;
          rgba[o + 2] = palette?.[idx * 3 + 2] ?? 0;
          rgba[o + 3] = trns && idx < trns.length ? trns[idx] : 255;
        } else if (channels <= 2) {
          const g = sample(s);
          rgba[o] = rgba[o + 1] = rgba[o + 2] = to8(g);
          rgba[o + 3] = channels === 2 ? to8(sample(s + 1)) : g === trnsGray ? 0 : 255;
        } else {
          const rv = sample(s);
          const gv = sample(s + 1);
          const bv = sample(s + 2);
          rgba[o] = to8(rv);
          rgba[o + 1] = to8(gv);
          rgba[o + 2] = to8(bv);
          rgba[o + 3] = channels === 4 ? to8(sample(s + 3)) : trnsRgb && rv === trnsRgb[0] && gv === trnsRgb[1] && bv === trnsRgb[2] ? 0 : 255;
        }
      }
    }
    return true;
  };
  const ok = interlace
    ? [
        [0, 0, 8, 8],
        [4, 0, 8, 8],
        [0, 4, 4, 8],
        [2, 0, 4, 4],
        [0, 2, 2, 4],
        [1, 0, 2, 2],
        [0, 1, 1, 2],
      ].every(([x, y, dx, dy]) => pass(x, y, dx, dy))
    : pass(0, 0, 1, 1);
  return ok ? { width, height, rgba } : null;
}
