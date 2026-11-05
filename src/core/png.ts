/** Small PNG encoder (8-bit gray/RGB/RGBA) built on fflate's zlib. Works in any JS runtime. */
import { zlibSync } from 'fflate';

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
