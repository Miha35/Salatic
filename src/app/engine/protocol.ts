import type { FormatId } from '../../core/formats';
import type { ConvertOptions } from '../../core/convert-lite';
import type { ImageInput, ImagesToPdfOptions, PageSpec } from '../../core/pdf-tools';

export type EngineRequest =
  | { op: 'convert'; bytes: Uint8Array; from: FormatId; to: FormatId; opts: ConvertOptions; title?: string }
  | { op: 'toHtml'; bytes: Uint8Array; from: FormatId; title?: string }
  | { op: 'fromHtml'; html: string; to: FormatId; opts: ConvertOptions; title?: string }
  | { op: 'imagesToPdf'; images: ImageInput[]; opts: ImagesToPdfOptions }
  | { op: 'composePdf'; sources: Uint8Array[]; pages: PageSpec[] };

export interface ConvertResult {
  data: Uint8Array;
  mime: string;
  ext: string;
  warnings: string[];
}

export interface HtmlResult {
  html: string;
  title?: string;
  warnings: string[];
}

export type WorkerMessage =
  | { id: number; type: 'progress'; value: number }
  | { id: number; type: 'done'; result: unknown }
  | { id: number; type: 'error'; message: string; code?: string };
