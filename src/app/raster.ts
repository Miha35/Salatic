/** Image decoding/encoding with the browser's native codecs, and PDF page rasterisation. */
import { zipSync } from 'fflate';
import type { FormatId } from '../core/formats';
import { FORMATS } from '../core/formats';
import type { ImageInput } from '../core/pdf-tools';
import { imageSize } from '../core/util';
import { openPdf, pdfErrorMessage, renderPage } from './pdfjs';

export class RasterError extends Error {}

const MAX_PIXELS = 100_000_000;

function svgSize(text: string): { w: number; h: number } {
  const svg = /<svg\b[^>]*>/i.exec(text)?.[0] ?? '';
  const num = (attr: string): number => {
    const m = new RegExp(`\\b${attr}\\s*=\\s*["']([\\d.]+)(px)?["']`, 'i').exec(svg);
    return m ? parseFloat(m[1]) : 0;
  };
  let w = num('width');
  let h = num('height');
  const vb = /viewBox\s*=\s*["']\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)/i.exec(svg);
  if ((!w || !h) && vb) {
    const vw = parseFloat(vb[1]);
    const vh = parseFloat(vb[2]);
    if (w && !h) h = (w * vh) / vw;
    else if (h && !w) w = (h * vw) / vh;
    else {
      w = vw;
      h = vh;
    }
  }
  if (!w || !h) {
    w = 512;
    h = 512;
  }
  // Rasterise small vector images at a useful resolution
  const k = Math.max(1, 1024 / Math.max(w, h));
  return { w: Math.round(w * k), h: Math.round(h * k) };
}

function loadImageElement(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.decoding = 'async';
    img.onload = () => resolve(img);
    img.onerror = () => reject(new RasterError('Браузер не смог декодировать изображение.'));
    img.src = url;
  });
}

/** Decode any browser-supported image into a canvas at natural size. */
export async function decodeToCanvas(blob: Blob, format: FormatId): Promise<HTMLCanvasElement> {
  let source: CanvasImageSource;
  let w: number;
  let h: number;
  let cleanup = (): void => {};
  if (format === 'svg') {
    const text = await blob.text();
    const size = svgSize(text);
    const url = URL.createObjectURL(new Blob([text], { type: 'image/svg+xml' }));
    cleanup = () => URL.revokeObjectURL(url);
    const img = await loadImageElement(url);
    source = img;
    w = size.w;
    h = size.h;
  } else {
    try {
      const bmp = await createImageBitmap(blob);
      source = bmp;
      w = bmp.width;
      h = bmp.height;
      cleanup = () => bmp.close();
    } catch {
      const url = URL.createObjectURL(blob);
      try {
        const img = await loadImageElement(url);
        source = img;
        w = img.naturalWidth;
        h = img.naturalHeight;
      } finally {
        cleanup = () => URL.revokeObjectURL(url);
      }
    }
  }
  if (!w || !h) throw new RasterError('Изображение имеет нулевой размер.');
  if (w * h > MAX_PIXELS) throw new RasterError('Изображение слишком большое для обработки в браузере.');
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  canvas.getContext('2d')!.drawImage(source, 0, 0, w, h);
  cleanup();
  return canvas;
}

function canvasToBlob(canvas: HTMLCanvasElement, mime: string, quality?: number): Promise<Blob> {
  return new Promise((resolve, reject) =>
    canvas.toBlob((b) => (b ? resolve(b) : reject(new RasterError(`Браузер не поддерживает сохранение в ${mime}.`))), mime, quality),
  );
}

/** Flatten transparency onto white (for formats without alpha). */
function flatten(canvas: HTMLCanvasElement, background = '#ffffff'): HTMLCanvasElement {
  const out = document.createElement('canvas');
  out.width = canvas.width;
  out.height = canvas.height;
  const ctx = out.getContext('2d')!;
  ctx.fillStyle = background;
  ctx.fillRect(0, 0, out.width, out.height);
  ctx.drawImage(canvas, 0, 0);
  return out;
}

function encodeBmp(canvas: HTMLCanvasElement): Uint8Array {
  const { width: w, height: h } = canvas;
  const data = flatten(canvas).getContext('2d')!.getImageData(0, 0, w, h).data;
  const rowSize = (w * 3 + 3) & ~3;
  const size = 54 + rowSize * h;
  const buf = new Uint8Array(size);
  const dv = new DataView(buf.buffer);
  buf[0] = 0x42;
  buf[1] = 0x4d;
  dv.setUint32(2, size, true);
  dv.setUint32(10, 54, true);
  dv.setUint32(14, 40, true);
  dv.setInt32(18, w, true);
  dv.setInt32(22, h, true);
  dv.setUint16(26, 1, true);
  dv.setUint16(28, 24, true);
  dv.setUint32(34, rowSize * h, true);
  dv.setInt32(38, 2835, true);
  dv.setInt32(42, 2835, true);
  for (let y = 0; y < h; y++) {
    const row = 54 + (h - 1 - y) * rowSize;
    for (let x = 0; x < w; x++) {
      const s = (y * w + x) * 4;
      const d = row + x * 3;
      buf[d] = data[s + 2];
      buf[d + 1] = data[s + 1];
      buf[d + 2] = data[s];
    }
  }
  return buf;
}

async function encodeIco(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  const k = Math.min(1, 256 / Math.max(canvas.width, canvas.height));
  const w = Math.max(1, Math.round(canvas.width * k));
  const h = Math.max(1, Math.round(canvas.height * k));
  const c = document.createElement('canvas');
  c.width = w;
  c.height = h;
  const ctx = c.getContext('2d')!;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(canvas, 0, 0, w, h);
  const png = new Uint8Array(await (await canvasToBlob(c, 'image/png')).arrayBuffer());
  const out = new Uint8Array(22 + png.length);
  const dv = new DataView(out.buffer);
  dv.setUint16(2, 1, true); // type: icon
  dv.setUint16(4, 1, true); // one image
  out[6] = w >= 256 ? 0 : w;
  out[7] = h >= 256 ? 0 : h;
  dv.setUint16(10, 1, true); // colour planes
  dv.setUint16(12, 32, true); // bpp
  dv.setUint32(14, png.length, true);
  dv.setUint32(18, 22, true);
  out.set(png, 22);
  return out;
}

export async function encodeCanvas(canvas: HTMLCanvasElement, to: FormatId, quality = 0.9): Promise<Uint8Array> {
  switch (to) {
    case 'png':
      return new Uint8Array(await (await canvasToBlob(canvas, 'image/png')).arrayBuffer());
    case 'jpg':
      return new Uint8Array(await (await canvasToBlob(flatten(canvas), 'image/jpeg', quality)).arrayBuffer());
    case 'webp':
    case 'avif': {
      const blob = await canvasToBlob(canvas, FORMATS[to].mime, quality);
      // Browsers silently fall back to PNG for unsupported encoders
      if (blob.type !== FORMATS[to].mime) throw new RasterError(`Ваш браузер не умеет сохранять в формат ${FORMATS[to].label}.`);
      return new Uint8Array(await blob.arrayBuffer());
    }
    case 'bmp':
      return encodeBmp(canvas);
    case 'ico':
      return encodeIco(canvas);
    default:
      throw new RasterError(`Формат ${to} не является растровым.`);
  }
}

let avifSupport: Promise<boolean> | null = null;
export function supportsAvifEncoding(): Promise<boolean> {
  avifSupport ??= new Promise((resolve) => {
    const c = document.createElement('canvas');
    c.width = c.height = 1;
    c.toBlob((b) => resolve(b?.type === 'image/avif'), 'image/avif');
  });
  return avifSupport;
}

export async function convertImage(blob: Blob, from: FormatId, to: FormatId, quality: number): Promise<Uint8Array> {
  const canvas = await decodeToCanvas(blob, from);
  return encodeCanvas(canvas, to, quality);
}

/** Prepare images for PDF embedding: JPEG/PNG pass through untouched, others are re-encoded. */
export async function toPdfImage(blob: Blob, format: FormatId): Promise<ImageInput> {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  if (format === 'jpg' || format === 'png') {
    const size = imageSize(bytes);
    if (size) return { data: bytes, mime: format === 'jpg' ? 'image/jpeg' : 'image/png', width: size.width, height: size.height };
  }
  const canvas = await decodeToCanvas(blob, format);
  const data = await encodeCanvas(canvas, 'png');
  return { data, mime: 'image/png', width: canvas.width, height: canvas.height };
}

export interface PageImages {
  files: { name: string; data: Uint8Array }[];
}

/** Render PDF pages to PNG/JPEG. */
export async function pdfToImages(bytes: Uint8Array, to: 'png' | 'jpg', dpi: number, quality: number, base: string, onProgress?: (v: number) => void, signal?: AbortSignal): Promise<PageImages> {
  let pdf;
  try {
    pdf = await openPdf(bytes);
  } catch (e) {
    throw new RasterError(pdfErrorMessage(e));
  }
  const files: PageImages['files'] = [];
  try {
    const n = pdf.numPages;
    const pad = String(n).length;
    for (let i = 1; i <= n; i++) {
      if (signal?.aborted) throw new RasterError('Отменено');
      const page = await pdf.getPage(i);
      let scale = dpi / 72;
      const base1 = page.getViewport({ scale: 1 });
      // Keep within browser canvas limits (~16k px side, ~70 MP area)
      scale = Math.min(scale, 16000 / Math.max(base1.width, base1.height), Math.sqrt(70_000_000 / (base1.width * base1.height)));
      const canvas = document.createElement('canvas');
      await renderPage(page, canvas, scale);
      const data = await encodeCanvas(canvas, to, quality);
      files.push({ name: n === 1 ? `${base}.${to}` : `${base}-${String(i).padStart(pad, '0')}.${to}`, data });
      canvas.width = canvas.height = 0;
      page.cleanup();
      onProgress?.(i / n);
    }
  } finally {
    void pdf.loadingTask.destroy();
  }
  return { files };
}

export function zipFiles(files: { name: string; data: Uint8Array }[]): Uint8Array {
  const used = new Set<string>();
  const entries: Record<string, [Uint8Array, { level: 0 | 6 }]> = {};
  for (const f of files) {
    let name = f.name;
    let k = 2;
    while (used.has(name)) {
      const dot = f.name.lastIndexOf('.');
      name = dot > 0 ? `${f.name.slice(0, dot)} (${k})${f.name.slice(dot)}` : `${f.name} (${k})`;
      k++;
    }
    used.add(name);
    // Already-compressed formats are stored, text formats are deflated
    entries[name] = [f.data, { level: /\.(png|jpe?g|webp|avif|gif|zip|docx|xlsx|pdf)$/i.test(name) ? 0 : 6 }];
  }
  return zipSync(entries);
}
