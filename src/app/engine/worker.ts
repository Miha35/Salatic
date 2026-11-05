/**
 * Conversion engine running off the main thread. Heavy libraries (pdf.js, pdf-lib,
 * fonts) are loaded on first use only.
 */
import hbSubsetUrl from 'harfbuzzjs/dist/harfbuzz-subset.wasm?url';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { convert, readContent, writeContent, type ConvertEnv } from '../../core/convert';
import { asDoc } from '../../core/convert-lite';
import { htmlToDoc } from '../../core/readers/html';
import { blocksToHtml } from '../../core/writers/text';
import { encodePng } from '../../core/png';
import type { FontKey } from '../../core/pdf-fonts';
import { createSubsetter, type Subsetter } from '../../core/hb-subset';
import { FONT_URLS } from './fonts';
import type { EngineRequest, WorkerMessage } from './protocol';

interface WorkerScope {
  location: Location;
  postMessage(message: unknown, transfer: Transferable[]): void;
  addEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
}
declare const self: WorkerScope;

const abs = (u: string): string => new URL(u, self.location.href).href;

const fontCache = new Map<FontKey, Promise<Uint8Array>>();

let subsetterPromise: Promise<Subsetter | undefined> | null = null;
/** HarfBuzz subsetter (wasm, ~680 KB) – loaded once; PDFs still work without it, just larger. */
function loadSubsetter(): Promise<Subsetter | undefined> {
  subsetterPromise ??= fetch(abs(hbSubsetUrl))
    .then((r) => (r.ok ? r.arrayBuffer() : Promise.reject(new Error(String(r.status)))))
    .then((buf) => createSubsetter(buf))
    .catch(() => undefined);
  return subsetterPromise;
}

/** Fetch a font on first use only (most documents need one or two styles). */
function loadFont(key: FontKey): Promise<Uint8Array> {
  let p = fontCache.get(key);
  if (!p) {
    p = fetch(abs(FONT_URLS[key])).then(async (r) => {
      if (!r.ok) throw new Error(`Не удалось загрузить шрифт (${r.status})`);
      return new Uint8Array(await r.arrayBuffer());
    });
    p.catch(() => fontCache.delete(key));
    fontCache.set(key, p);
  }
  return p;
}

let pdfjsPromise: Promise<typeof import('pdfjs-dist')> | null = null;
function loadPdfjs(): Promise<typeof import('pdfjs-dist')> {
  pdfjsPromise ??= import('pdfjs-dist').then((m) => {
    m.GlobalWorkerOptions.workerSrc = abs(pdfWorkerUrl);
    return m;
  });
  return pdfjsPromise;
}

async function encodeImage(pixels: Uint8Array | Uint8ClampedArray, width: number, height: number, channels: 1 | 3 | 4): Promise<{ data: Uint8Array; mime: string }> {
  // Large opaque photos → JPEG (much smaller documents); everything else → lossless PNG
  if (typeof OffscreenCanvas !== 'undefined' && channels === 3 && width * height > 160_000) {
    try {
      const canvas = new OffscreenCanvas(width, height);
      const ctx = canvas.getContext('2d')!;
      const img = ctx.createImageData(width, height);
      for (let i = 0, j = 0; i < pixels.length; i += 3, j += 4) {
        img.data[j] = pixels[i];
        img.data[j + 1] = pixels[i + 1];
        img.data[j + 2] = pixels[i + 2];
        img.data[j + 3] = 255;
      }
      ctx.putImageData(img, 0, 0);
      const blob = await canvas.convertToBlob({ type: 'image/jpeg', quality: 0.88 });
      return { data: new Uint8Array(await blob.arrayBuffer()), mime: 'image/jpeg' };
    } catch {
      /* fall through to PNG */
    }
  }
  return { data: encodePng(pixels, width, height, channels), mime: 'image/png' };
}

async function transcodeImage(data: Uint8Array, mime: string): Promise<{ data: Uint8Array; mime: string } | null> {
  if (typeof OffscreenCanvas === 'undefined' || typeof createImageBitmap === 'undefined') return null;
  const bmp = await createImageBitmap(new Blob([data as BlobPart], { type: mime }));
  const canvas = new OffscreenCanvas(bmp.width, bmp.height);
  canvas.getContext('2d')!.drawImage(bmp, 0, 0);
  bmp.close();
  const blob = await canvas.convertToBlob({ type: 'image/png' });
  return { data: new Uint8Array(await blob.arrayBuffer()), mime: 'image/png' };
}

function env(id: number, subsetter?: Subsetter): ConvertEnv {
  return {
    subsetter,
    pdfjs: loadPdfjs,
    pdfDocParams: {
      cMapUrl: abs('/pdfjs/cmaps/'),
      cMapPacked: true,
      standardFontDataUrl: abs('/pdfjs/standard_fonts/'),
      wasmUrl: abs('/pdfjs/wasm/'),
      // Must be explicit inside a Web Worker: pdf.js would otherwise probe document.baseURI
      useWorkerFetch: true,
      // Inside our worker pdf.js parses in-thread ("fake worker"); silence the notice about it
      verbosity: 0,
    },
    fonts: loadFont,
    encodeImage,
    transcodeImage,
    onProgress: (value) => post({ id, type: 'progress', value }),
  };
}

function post(msg: WorkerMessage, transfer: Transferable[] = []): void {
  self.postMessage(msg, transfer);
}

async function handle(id: number, req: EngineRequest): Promise<{ result: unknown; transfer: Transferable[] }> {
  switch (req.op) {
    case 'convert': {
      const r = await convert(req.bytes, req.from, req.to, env(id, req.to === 'pdf' ? await loadSubsetter() : undefined), req.opts, req.title);
      return { result: r, transfer: [r.data.buffer] };
    }
    case 'toHtml': {
      const { content, warnings } = await readContent(req.bytes, req.from, env(id), req.title);
      const doc = asDoc(content, req.title);
      return { result: { html: blocksToHtml(doc.blocks), title: doc.title, warnings }, transfer: [] };
    }
    case 'fromHtml': {
      const doc = htmlToDoc(req.html);
      doc.title = req.title || doc.title;
      const r = await writeContent({ kind: 'doc', doc }, req.to, env(id, req.to === 'pdf' ? await loadSubsetter() : undefined), req.opts, req.title);
      return { result: r, transfer: [r.data.buffer] };
    }
    case 'imagesToPdf': {
      const { imagesToPdf } = await import('../../core/pdf-tools');
      const data = await imagesToPdf(req.images, req.opts);
      return { result: data, transfer: [data.buffer] };
    }
    case 'composePdf': {
      const { composePdf } = await import('../../core/pdf-tools');
      const needsFonts = req.pages.some((p) => p.annotations?.some((a) => a.type === 'text'));
      const data = await composePdf(req.sources, req.pages, loadFont, needsFonts ? await loadSubsetter() : undefined);
      return { result: data, transfer: [data.buffer] };
    }
  }
}

self.addEventListener('message', (e: MessageEvent<{ id: number; req: EngineRequest }>) => {
  const { id, req } = e.data;
  handle(id, req).then(
    ({ result, transfer }) => post({ id, type: 'done', result }, transfer),
    (err: unknown) => {
      const e2 = err as { message?: string; name?: string; code?: string };
      // Errors from our code carry Russian user-facing messages; others are unexpected
      const message = e2?.name === 'ConvertError' ? e2.message! : `Ошибка обработки: ${e2?.message ?? String(err)}`;
      post({ id, type: 'error', message, code: e2?.code });
    },
  );
});
