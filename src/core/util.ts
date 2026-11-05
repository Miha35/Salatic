/** Shared, environment-agnostic helpers (browser, Web Worker, Cloudflare Worker, Node). */

/** Error with a user-facing (Russian) message. */
export class ConvertError extends Error {
  readonly code: string;
  constructor(message: string, code = 'convert_failed') {
    super(message);
    this.name = 'ConvertError';
    this.code = code;
  }
}

const encoder = new TextEncoder();
export const utf8 = (s: string): Uint8Array => encoder.encode(s);

/** Decode text bytes: BOM-aware UTF-8/UTF-16, with a windows-1251 fallback for legacy Cyrillic files. */
export function decodeText(bytes: Uint8Array): string {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  }
  try {
    return new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(bytes);
  } catch {
    for (const enc of ['windows-1251', 'latin1']) {
      try {
        return new TextDecoder(enc).decode(bytes);
      } catch {
        /* encoding not supported by this runtime – try the next one */
      }
    }
    return new TextDecoder('utf-8').decode(bytes);
  }
}

const XML_ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' };
/** Escape text for XML; also strips characters that are illegal in XML 1.0. */
export function escapeXml(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '').replace(/[&<>"']/g, (c) => XML_ESC[c]);
}

const HTML_ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
export function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => HTML_ESC[c]);
}

export function bytesToBase64(bytes: Uint8Array): string {
  let bin = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK) as unknown as number[]);
  }
  return btoa(bin);
}

export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64.replace(/\s+/g, ''));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** Parse `data:` URI into bytes and mime type. Returns null for anything else. */
export function parseDataUri(uri: string): { mime: string; data: Uint8Array } | null {
  const m = /^data:([^;,]*)((?:;[^;,]*)*?)(;base64)?,(.*)$/s.exec(uri.trim());
  if (!m) return null;
  const mime = (m[1] || 'text/plain').toLowerCase();
  try {
    const data = m[3] ? base64ToBytes(m[4]) : utf8(decodeURIComponent(m[4]));
    return { mime, data };
  } catch {
    return null;
  }
}

export function stripExtension(name: string): string {
  const slash = Math.max(name.lastIndexOf('/'), name.lastIndexOf('\\'));
  const base = name.slice(slash + 1);
  const dot = base.lastIndexOf('.');
  return dot > 0 ? base.slice(0, dot) : base;
}

export function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot >= 0 ? name.slice(dot + 1).toLowerCase() : '';
}

/** Yield to the event loop so long jobs stay cancellable/responsive. */
export const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

/** Remove characters that cannot be rendered/stored meaningfully (keeps \n and \t). */
export function cleanText(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f­​‎‏﻿]/g, '');
}

export function concatBytes(parts: Uint8Array[]): Uint8Array {
  let len = 0;
  for (const p of parts) len += p.length;
  const out = new Uint8Array(len);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/** Max of a possibly huge array (spreading into Math.max overflows the call stack). */
export function maxOf(values: Iterable<number>, initial = -Infinity): number {
  let m = initial;
  for (const v of values) if (v > m) m = v;
  return m;
}

export function minOf(values: Iterable<number>, initial = Infinity): number {
  let m = initial;
  for (const v of values) if (v < m) m = v;
  return m;
}

const MAX_UNZIPPED = 400 * 1024 * 1024;
const MAX_RATIO = 250;

/**
 * Unzip an Office package defensively: entry sizes are checked from the central
 * directory before anything is inflated, so zip bombs are rejected cheaply.
 */
export function unzipOffice(bytes: Uint8Array, unzip: (b: Uint8Array, o?: { filter?: (f: { name: string; size: number; originalSize: number }) => boolean }) => Record<string, Uint8Array>): Record<string, Uint8Array> {
  let total = 0;
  unzip(bytes, {
    filter: (f) => {
      total += f.originalSize;
      if (f.originalSize > 1_000_000 && f.originalSize / Math.max(1, f.size) > MAX_RATIO) throw new ConvertError('Архив документа выглядит подозрительно (аномальное сжатие) и не будет распакован.', 'zip_bomb');
      return false;
    },
  });
  if (total > MAX_UNZIPPED) throw new ConvertError('Документ слишком большой после распаковки.', 'too_large');
  return unzip(bytes);
}

/** Image dimensions from PNG/JPEG/GIF/BMP/WebP headers without decoding. */
export function imageSize(data: Uint8Array): { width: number; height: number } | null {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (data.length > 24 && data[0] === 0x89 && data[1] === 0x50) {
    return { width: dv.getUint32(16), height: dv.getUint32(20) };
  }
  if (data.length > 10 && data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) {
    return { width: dv.getUint16(6, true), height: dv.getUint16(8, true) };
  }
  if (data.length > 26 && data[0] === 0x42 && data[1] === 0x4d) {
    return { width: Math.abs(dv.getInt32(18, true)), height: Math.abs(dv.getInt32(22, true)) };
  }
  if (data.length > 30 && data[0] === 0x52 && data[8] === 0x57 && data[9] === 0x45) {
    const chunk = String.fromCharCode(data[12], data[13], data[14], data[15]);
    if (chunk === 'VP8 ') return { width: dv.getUint16(26, true) & 0x3fff, height: dv.getUint16(28, true) & 0x3fff };
    if (chunk === 'VP8L') {
      const b = dv.getUint32(21, true);
      return { width: (b & 0x3fff) + 1, height: ((b >> 14) & 0x3fff) + 1 };
    }
    if (chunk === 'VP8X') {
      return {
        width: 1 + (data[24] | (data[25] << 8) | (data[26] << 16)),
        height: 1 + (data[27] | (data[28] << 8) | (data[29] << 16)),
      };
    }
  }
  if (data.length > 4 && data[0] === 0xff && data[1] === 0xd8) {
    let off = 2;
    while (off + 9 < data.length) {
      if (data[off] !== 0xff) {
        off++;
        continue;
      }
      const marker = data[off + 1];
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        off += 2;
        continue;
      }
      const len = dv.getUint16(off + 2);
      if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) {
        return { height: dv.getUint16(off + 5), width: dv.getUint16(off + 7) };
      }
      off += 2 + len;
    }
  }
  return null;
}

/** Sniff an image mime type from magic bytes. */
export function sniffImageMime(data: Uint8Array): string | null {
  if (data.length < 4) return null;
  if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return 'image/png';
  if (data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff) return 'image/jpeg';
  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) return 'image/gif';
  if (data[0] === 0x42 && data[1] === 0x4d) return 'image/bmp';
  if (data.length > 12 && data[0] === 0x52 && data[1] === 0x49 && data[8] === 0x57 && data[9] === 0x45) return 'image/webp';
  if ((data[0] === 0x49 && data[1] === 0x49 && data[2] === 0x2a && data[3] === 0) || (data[0] === 0x4d && data[1] === 0x4d && data[2] === 0 && data[3] === 0x2a)) return 'image/tiff';
  return null;
}
