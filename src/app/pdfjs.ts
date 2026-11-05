/** Main-thread pdf.js (rendering). Parsing runs in pdf.js' own worker. */
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

export type PdfJs = typeof import('pdfjs-dist');
export type PdfDoc = Awaited<ReturnType<PdfJs['getDocument']>['promise']>;

let promise: Promise<PdfJs> | null = null;

export function loadPdfjs(): Promise<PdfJs> {
  promise ??= import('pdfjs-dist').then((m) => {
    m.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
    return m;
  });
  return promise;
}

export async function openPdf(data: Uint8Array, password?: string): Promise<PdfDoc> {
  const pdfjs = await loadPdfjs();
  const origin = location.origin;
  return pdfjs.getDocument({
    // Font names and style flags let the PDF editor match the font of existing text
    fontExtraProperties: true,
    data: data.slice(),
    password,
    cMapUrl: `${origin}/pdfjs/cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${origin}/pdfjs/standard_fonts/`,
    wasmUrl: `${origin}/pdfjs/wasm/`,
    isEvalSupported: false,
  } as Parameters<PdfJs['getDocument']>[0]).promise;
}

export function pdfErrorMessage(e: unknown): string {
  const name = (e as Error)?.name;
  if (name === 'PasswordException') return 'PDF защищён паролем.';
  if (name === 'InvalidPDFException') return 'Файл PDF повреждён или имеет неверный формат.';
  return `Не удалось открыть PDF: ${(e as Error)?.message ?? e}`;
}

type PdfPage = Awaited<ReturnType<PdfDoc['getPage']>>;

/** Render a page into a canvas sized to the viewport, on a white background. */
export async function renderPage(page: PdfPage, canvas: HTMLCanvasElement, scale: number, rotation?: number): Promise<ReturnType<PdfPage['getViewport']>> {
  const viewport = page.getViewport({ scale, rotation: rotation ?? page.rotate });
  canvas.width = Math.max(1, Math.ceil(viewport.width));
  canvas.height = Math.max(1, Math.ceil(viewport.height));
  await page.render({ canvas, viewport, background: '#ffffff' } as Parameters<PdfPage['render']>[0]).promise;
  return viewport;
}
