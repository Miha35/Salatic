/** Conversions that need text recognition: scanned PDFs and images. */
import type { ConvertOptions } from '../core/convert-lite';
import type { FormatId } from '../core/formats';
import type { Block, Doc } from '../core/model';
import type { TextLayer } from '../core/pdf-tools';
import { NO_TEXT_LAYER } from '../core/readers/pdf';
import { engine } from './engine/client';
import type { ConvertResult, ImagesPdfResult, PdfModelResult } from './engine/protocol';
import { ocrCanvas, ocrPdfPages, pagesWithoutText, prepareForOcr } from './ocr';
import { decodeToCanvas, toPdfImage } from './raster';

export type OcrMode = 'auto' | 'on' | 'off';
type Progress = (v: number, label?: string) => void;

const OCR_NOTE = 'Текст распознан автоматически (OCR) — проверьте имена, числа и редкие слова.';

/** PDF → DOCX/TXT/MD/HTML, recognising pages that have no text layer (or all pages). */
export async function pdfToDocWithOcr(bytes: Uint8Array, to: FormatId, opts: ConvertOptions, mode: OcrMode, title: string, onProgress: Progress, signal?: AbortSignal): Promise<ConvertResult> {
  const copy = bytes.slice();
  const model = await engine.run<PdfModelResult>({ op: 'pdfModel', bytes: copy, title }, { transfer: [copy.buffer], onProgress: (v) => onProgress(v * 0.3, 'Анализ PDF…'), signal });
  const pages = mode === 'on' ? model.pages.map((_, i) => i) : model.pages.flatMap((p, i) => (p.hasText ? [] : [i]));
  let doc: Doc = model.doc;
  let warnings = model.warnings;
  if (pages.length && mode !== 'off') {
    const ocr = await ocrPdfPages(bytes, pages, { onProgress: (v, label) => onProgress(0.3 + v * 0.6, label), signal });
    const byPage = new Map(ocr.map((r) => [r.page, r.blocks]));
    // Rebuild the document page by page: recognised pages replace their images
    const blocks: Block[] = [];
    model.pages.forEach((p, i) => {
      const recognised = byPage.get(i);
      if (recognised) {
        if (blocks.length && i > 0) blocks.push({ type: 'pagebreak' });
        blocks.push(...(recognised.length ? recognised : p.blocks.map((bi) => model.doc.blocks[bi])));
      } else blocks.push(...p.blocks.map((bi) => model.doc.blocks[bi]));
    });
    doc = { ...model.doc, blocks };
    warnings = [...warnings.filter((w) => !w.startsWith(NO_TEXT_LAYER)), OCR_NOTE];
    if (!blocks.length) throw new Error('Не удалось распознать текст: страницы пустые или изображение слишком нечёткое.');
  } else if (!doc.blocks.length) throw new Error('В PDF не найдено ни текста, ни изображений для извлечения.');
  const r = await engine.run<ConvertResult>({ op: 'writeDoc', doc, to, opts, title }, { onProgress: (v) => onProgress(0.9 + v * 0.1, 'Сохранение…'), signal });
  return { ...r, warnings: [...new Set([...warnings, ...r.warnings])] };
}

/** Image → DOCX/TXT: recognise the text of a photo or scan. */
export async function imageToDocWithOcr(file: Blob, format: FormatId, to: FormatId, title: string, onProgress: Progress): Promise<ConvertResult> {
  onProgress(0.05, 'Подготовка изображения…');
  const source = await decodeToCanvas(file, format);
  const { canvas } = prepareForOcr(source);
  // Unknown DPI: assume the long side of the picture is an A4 page (842 pt) to estimate font sizes
  const pxPerPt = Math.max(canvas.width, canvas.height) / 842;
  const r = await ocrCanvas(canvas, { pxPerPt, onProgress: (v, label) => onProgress(0.1 + v * 0.8, label) });
  if (!r.blocks.length) throw new Error('Текст на изображении не найден.');
  const doc: Doc = { title, blocks: r.blocks };
  const out = await engine.run<ConvertResult>({ op: 'writeDoc', doc, to, opts: {}, title });
  const warnings = [OCR_NOTE];
  if (r.confidence < 60) warnings.push(`Низкая уверенность распознавания (${Math.round(r.confidence)}%): изображение нечёткое или мелкое.`);
  return { ...out, warnings: [...warnings, ...out.warnings] };
}

/** PDF → searchable PDF: invisible recognised text over pages that lack a text layer. */
export async function searchablePdf(bytes: Uint8Array, onProgress: Progress, signal?: AbortSignal): Promise<{ data: Uint8Array; warnings: string[] }> {
  onProgress(0.02, 'Поиск страниц без текста…');
  const { total, missing } = await pagesWithoutText(bytes);
  if (!missing.length) return { data: bytes, warnings: ['В PDF уже есть текстовый слой на всех страницах — распознавание не требуется.'] };
  const ocr = await ocrPdfPages(bytes, missing, { pdf: true, onProgress: (v, label) => onProgress(0.05 + v * 0.85, label), signal });
  const layers: TextLayer[] = ocr.filter((r) => r.pdf).map((r) => ({ page: r.page, pdf: r.pdf!, toUser: r.toUser, width: r.width, height: r.height }));
  const base = bytes.slice();
  const data = await engine.run<Uint8Array>({ op: 'overlayText', base, layers }, { transfer: [base.buffer, ...layers.map((l) => l.pdf.buffer)], signal });
  return { data, warnings: [`Добавлен текстовый слой на ${missing.length} из ${total} стр.: теперь в PDF работают поиск и копирование текста.`, OCR_NOTE] };
}

/** Images → PDF with an invisible text layer (searchable scans). */
export async function imagesToSearchablePdf(files: { file: Blob; format: FormatId }[], pageSize: 'fit' | 'a4' | 'letter', onProgress: Progress): Promise<Uint8Array> {
  const inputs = [];
  for (const f of files) inputs.push(await toPdfImage(f.file, f.format));
  const { pdf, placements } = await engine.run<ImagesPdfResult>({ op: 'imagesToPdf', images: inputs, opts: { pageSize }, placements: true }, { transfer: inputs.map((i) => i.data.buffer) });
  const layers: TextLayer[] = [];
  for (let k = 0; k < files.length; k++) {
    const source = await decodeToCanvas(files[k].file, files[k].format);
    const { canvas } = prepareForOcr(source);
    const pl = placements[k];
    const r = await ocrCanvas(canvas, { pxPerPt: canvas.width / pl.w, pdf: true, onProgress: (v, label) => onProgress((k + v) / files.length, `${label} ${k + 1} из ${files.length}`) });
    if (!r.pdf) continue;
    // canvas pixels (y down) → page user space of the drawn image
    layers.push({ page: pl.page, pdf: r.pdf, width: canvas.width, height: canvas.height, toUser: [pl.w / canvas.width, 0, 0, -pl.h / canvas.height, pl.x, pl.y + pl.h] });
  }
  return engine.run<Uint8Array>({ op: 'overlayText', base: pdf, layers }, { transfer: [pdf.buffer, ...layers.map((l) => l.pdf.buffer)] });
}
