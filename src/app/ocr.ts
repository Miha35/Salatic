/**
 * Text recognition (OCR) with tesseract.js. Tesseract runs in its own Web Worker;
 * all assets (worker, WebAssembly core, Russian + English models) are served from this
 * site – nothing is sent to third parties.
 */
import type { Block, TextRun } from '../core/model';
import { openPdf, renderPage, type PdfDoc } from './pdfjs';

type Tesseract = typeof import('tesseract.js');
type TWorker = Awaited<ReturnType<Tesseract['createWorker']>>;

export interface OcrProgress {
  (fraction: number, label: string): void;
}

const LANGS = ['rus', 'eng'];
let workerPromise: Promise<TWorker> | null = null;
let progressSink: OcrProgress | null = null;

function getWorker(): Promise<TWorker> {
  workerPromise ??= (async () => {
    const T = await import('tesseract.js');
    const origin = location.origin;
    const worker = await T.createWorker(LANGS, T.OEM.LSTM_ONLY, {
      workerPath: `${origin}/tesseract/worker.min.js`,
      corePath: `${origin}/tesseract/core`,
      langPath: `${origin}/tesseract/lang`,
      workerBlobURL: false,
      gzip: true,
      logger: (m: { status: string; progress: number }) => {
        if (!progressSink) return;
        if (m.status === 'recognizing text') progressSink(m.progress, 'Распознавание текста…');
        else if (/loading|initializ/i.test(m.status)) progressSink(m.progress, 'Загрузка модели распознавания (один раз)…');
      },
    });
    await worker.setParameters({ preserve_interword_spaces: '1', user_defined_dpi: '300' });
    return worker;
  })().catch((e) => {
    workerPromise = null;
    throw new Error(`Не удалось запустить распознавание текста: ${(e as Error)?.message ?? e}`);
  });
  return workerPromise;
}

/** Terminate the OCR worker (frees ~100 MB). */
export async function disposeOcr(): Promise<void> {
  const w = workerPromise;
  workerPromise = null;
  if (w) await (await w).terminate().catch(() => undefined);
}

export interface OcrResult {
  blocks: Block[];
  /** Text-only PDF page (invisible text) when requested. */
  pdf?: Uint8Array;
  confidence: number;
}

interface TLine {
  text: string;
  bbox: { x0: number; y0: number; x1: number; y1: number };
  rowAttributes?: { rowHeight: number };
  confidence: number;
}
interface TParagraph {
  lines: TLine[];
}
interface TBlock {
  paragraphs: TParagraph[];
}

/** Turn Tesseract's layout into document blocks. `pxPerPt` converts image pixels to points. */
export function layoutToBlocks(blocks: TBlock[] | null, pxPerPt: number): Block[] {
  const paras: { lines: TLine[]; height: number; x0: number; x1: number }[] = [];
  for (const b of blocks ?? []) {
    for (const p of b.paragraphs) {
      const lines = p.lines.filter((l) => l.text.trim() && l.confidence > 15);
      if (!lines.length) continue;
      const heights = lines.map((l) => l.rowAttributes?.rowHeight || l.bbox.y1 - l.bbox.y0).sort((a, b) => a - b);
      paras.push({ lines, height: heights[heights.length >> 1], x0: Math.min(...lines.map((l) => l.bbox.x0)), x1: Math.max(...lines.map((l) => l.bbox.x1)) });
    }
  }
  if (!paras.length) return [];
  // Body text height: the height carrying most lines
  const weights = new Map<number, number>();
  for (const p of paras) weights.set(Math.round(p.height), (weights.get(Math.round(p.height)) ?? 0) + p.lines.length);
  const body = [...weights].sort((a, b) => b[1] - a[1])[0][0];
  const out: Block[] = [];
  for (const p of paras) {
    let text = '';
    for (const l of p.lines) {
      const t = l.text.replace(/\s+$/, '').replace(/\s{2,}/g, ' ');
      if (!text) text = t;
      else if (/[\p{L}]-$/u.test(text) && /^\p{Ll}/u.test(t)) text = text.slice(0, -1) + t; // de-hyphenate
      else text += ' ' + t;
    }
    const runs: TextRun[] = [{ text: text.trim() }];
    const sizePt = Math.round((p.height / pxPerPt / 1.2) * 2) / 2;
    if (p.height > body * 1.35 && text.length < 200) {
      out.push({ type: 'heading', level: p.height > body * 1.8 ? 1 : 2, runs });
    } else {
      out.push({ type: 'paragraph', runs });
      if (Math.abs(sizePt - body / pxPerPt / 1.2) > 1.5) runs[0].size = sizePt;
    }
  }
  return out;
}

/** Recognise one image (canvas). */
export async function ocrCanvas(canvas: HTMLCanvasElement, opts: { pxPerPt: number; pdf?: boolean; onProgress?: OcrProgress }): Promise<OcrResult> {
  const worker = await getWorker();
  progressSink = opts.onProgress ?? null;
  try {
    const { data } = await worker.recognize(canvas, opts.pdf ? { pdfTitle: 'OCR', pdfTextOnly: true } : {}, { text: true, blocks: true, pdf: !!opts.pdf });
    return {
      blocks: layoutToBlocks(data.blocks as unknown as TBlock[], opts.pxPerPt),
      pdf: data.pdf ? new Uint8Array(data.pdf) : undefined,
      confidence: data.confidence,
    };
  } finally {
    progressSink = null;
  }
}

/** Upscale small images: Tesseract works best around 300 DPI (text height 20–40 px). */
export function prepareForOcr(source: HTMLCanvasElement): { canvas: HTMLCanvasElement; scale: number } {
  const longSide = Math.max(source.width, source.height);
  const scale = longSide < 1600 ? Math.min(3, 2400 / longSide) : longSide > 5000 ? 5000 / longSide : 1;
  if (scale === 1) return { canvas: source, scale };
  const c = document.createElement('canvas');
  c.width = Math.round(source.width * scale);
  c.height = Math.round(source.height * scale);
  const ctx = c.getContext('2d')!;
  ctx.imageSmoothingQuality = 'high';
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(source, 0, 0, c.width, c.height);
  return { canvas: c, scale };
}

export interface PdfPageOcr extends OcrResult {
  page: number;
  /** Maps canvas pixels (y down) to the page's PDF user space. */
  toUser: [number, number, number, number, number, number];
  width: number;
  height: number;
}

function invert(m: number[]): [number, number, number, number, number, number] {
  const [a, b, c, d, e, f] = m;
  const det = a * d - b * c;
  return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det];
}

/** Render and recognise selected pages of a PDF. */
export async function ocrPdfPages(data: Uint8Array, pages: number[], opts: { pdf?: boolean; onProgress?: OcrProgress; signal?: AbortSignal }): Promise<PdfPageOcr[]> {
  const doc: PdfDoc = await openPdf(data);
  const out: PdfPageOcr[] = [];
  try {
    for (let k = 0; k < pages.length; k++) {
      if (opts.signal?.aborted) throw new Error('Отменено');
      const page = await doc.getPage(pages[k] + 1);
      const base = page.getViewport({ scale: 1 });
      const scale = Math.min(300 / 72, 4200 / Math.max(base.width, base.height));
      const canvas = document.createElement('canvas');
      const vp = await renderPage(page, canvas, scale);
      const report: OcrProgress = (f, label) => opts.onProgress?.((k + f) / pages.length, `${label} стр. ${k + 1} из ${pages.length}`);
      report(0, 'Распознавание текста…');
      const r = await ocrCanvas(canvas, { pxPerPt: scale, pdf: opts.pdf, onProgress: report });
      out.push({ ...r, page: pages[k], toUser: invert(vp.transform), width: canvas.width, height: canvas.height });
      canvas.width = canvas.height = 0;
      page.cleanup();
    }
  } finally {
    void doc.loadingTask.destroy();
  }
  return out;
}

/** Pages of a PDF that have no extractable text (scanned). */
export async function pagesWithoutText(data: Uint8Array): Promise<{ total: number; missing: number[] }> {
  const doc = await openPdf(data);
  try {
    const missing: number[] = [];
    for (let i = 0; i < doc.numPages; i++) {
      const tc = await (await doc.getPage(i + 1)).getTextContent();
      if (!tc.items.some((it) => 'str' in it && it.str.trim())) missing.push(i);
    }
    return { total: doc.numPages, missing };
  } finally {
    void doc.loadingTask.destroy();
  }
}
