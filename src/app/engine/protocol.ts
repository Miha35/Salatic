import type { FormatId } from '../../core/formats';
import type { ConvertOptions } from '../../core/convert-lite';
import type { ImageInput, ImagePlacement, ImagesToPdfOptions, PageSpec, TextLayer } from '../../core/pdf-tools';
import type { PdfPageInfo } from '../../core/readers/pdf';
import type { Doc } from '../../core/model';

/** Document-level settings that the editor keeps alongside its HTML. */
export type DocMeta = Pick<Doc, 'font' | 'size' | 'page' | 'defaults'>;

export type EngineRequest =
  | { op: 'convert'; bytes: Uint8Array; from: FormatId; to: FormatId; opts: ConvertOptions; title?: string }
  | { op: 'toHtml'; bytes: Uint8Array; from: FormatId; title?: string }
  | { op: 'fromHtml'; html: string; to: FormatId; opts: ConvertOptions; title?: string; meta?: DocMeta }
  | { op: 'imagesToPdf'; images: ImageInput[]; opts: ImagesToPdfOptions; placements?: boolean }
  | { op: 'pdfModel'; bytes: Uint8Array; title?: string }
  | { op: 'writeDoc'; doc: Doc; to: FormatId; opts: ConvertOptions; title?: string }
  | { op: 'overlayText'; base: Uint8Array; layers: TextLayer[] }
  | { op: 'composePdf'; sources: Uint8Array[]; pages: PageSpec[] }
  | { op: 'decodeImage'; bytes: Uint8Array; format: 'tiff' | 'heic' }
  | { op: 'decryptPdf'; bytes: Uint8Array; password: string };

export interface ConvertResult {
  data: Uint8Array;
  mime: string;
  ext: string;
  warnings: string[];
}

export interface PdfModelResult {
  doc: Doc;
  pages: PdfPageInfo[];
  warnings: string[];
}

export interface ImagesPdfResult {
  pdf: Uint8Array;
  placements: ImagePlacement[];
}

export interface HtmlResult {
  html: string;
  title?: string;
  meta: DocMeta;
  warnings: string[];
}

export type WorkerMessage =
  | { id: number; type: 'progress'; value: number }
  | { id: number; type: 'done'; result: unknown }
  | { id: number; type: 'error'; message: string; code?: string };
