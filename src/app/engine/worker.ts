/**
 * Conversion engine running off the main thread. Heavy libraries (pdf.js, pdf-lib,
 * fonts) are loaded on first use only.
 */
import regularUrl from '@expo-google-fonts/pt-sans/400Regular/PTSans_400Regular.ttf?url';
import boldUrl from '@expo-google-fonts/pt-sans/700Bold/PTSans_700Bold.ttf?url';
import italicUrl from '@expo-google-fonts/pt-sans/400Regular_Italic/PTSans_400Regular_Italic.ttf?url';
import boldItalicUrl from '@expo-google-fonts/pt-sans/700Bold_Italic/PTSans_700Bold_Italic.ttf?url';
import monoUrl from '@expo-google-fonts/pt-mono/400Regular/PTMono_400Regular.ttf?url';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { convert, readContent, writeContent, type ConvertEnv } from '../../core/convert';
import { asDoc } from '../../core/convert-lite';
import { htmlToDoc } from '../../core/readers/html';
import { blocksToHtml } from '../../core/writers/text';
import { encodePng } from '../../core/png';
import type { FontSet } from '../../core/writers/pdf';
import type { EngineRequest, WorkerMessage } from './protocol';

interface WorkerScope {
  location: Location;
  postMessage(message: unknown, transfer: Transferable[]): void;
  addEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
}
declare const self: WorkerScope;

const abs = (u: string): string => new URL(u, self.location.href).href;

let fontsPromise: Promise<FontSet> | null = null;
function loadFonts(): Promise<FontSet> {
  const get = async (u: string): Promise<Uint8Array> => {
    const r = await fetch(abs(u));
    if (!r.ok) throw new Error(`Не удалось загрузить шрифт (${r.status})`);
    return new Uint8Array(await r.arrayBuffer());
  };
  fontsPromise ??= Promise.all([get(regularUrl), get(boldUrl), get(italicUrl), get(boldItalicUrl), get(monoUrl)])
    .then(([regular, bold, italic, boldItalic, mono]) => ({ regular, bold, italic, boldItalic, mono }))
    .catch((e) => {
      fontsPromise = null;
      throw e;
    });
  return fontsPromise;
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

function env(id: number): ConvertEnv {
  return {
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
    fonts: loadFonts,
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
      const r = await convert(req.bytes, req.from, req.to, env(id), req.opts, req.title);
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
      const r = await writeContent({ kind: 'doc', doc }, req.to, env(id), req.opts, req.title);
      return { result: r, transfer: [r.data.buffer] };
    }
    case 'imagesToPdf': {
      const { imagesToPdf } = await import('../../core/pdf-tools');
      const data = await imagesToPdf(req.images, req.opts);
      return { result: data, transfer: [data.buffer] };
    }
    case 'composePdf': {
      const { composePdf } = await import('../../core/pdf-tools');
      const data = await composePdf(req.sources, req.pages, loadFonts);
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
