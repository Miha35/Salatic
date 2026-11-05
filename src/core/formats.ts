/** Format registry, detection and the conversion graph. */

import { unzipSync } from 'fflate';
import { extensionOf } from './util';
import { Cfb } from './cfb';

export type FormatId =
  | 'pdf'
  | 'docx'
  | 'doc'
  | 'txt'
  | 'md'
  | 'html'
  | 'csv'
  | 'tsv'
  | 'json'
  | 'xlsx'
  | 'xls'
  | 'odt'
  | 'rtf'
  | 'ods'
  | 'png'
  | 'jpg'
  | 'webp'
  | 'gif'
  | 'bmp'
  | 'svg'
  | 'ico'
  | 'avif';

export type FormatKind = 'document' | 'data' | 'image';

export interface FormatInfo {
  id: FormatId;
  label: string;
  ext: string;
  mime: string;
  kind: FormatKind;
  description: string;
}

export const FORMATS: Record<FormatId, FormatInfo> = {
  pdf: { id: 'pdf', label: 'PDF', ext: 'pdf', mime: 'application/pdf', kind: 'document', description: 'Документ PDF' },
  docx: {
    id: 'docx',
    label: 'Word',
    ext: 'docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    kind: 'document',
    description: 'Документ Microsoft Word',
  },
  doc: { id: 'doc', label: 'Word 97–2003', ext: 'doc', mime: 'application/msword', kind: 'document', description: 'Документ Word 97–2003' },
  txt: { id: 'txt', label: 'TXT', ext: 'txt', mime: 'text/plain;charset=utf-8', kind: 'document', description: 'Обычный текст' },
  md: { id: 'md', label: 'Markdown', ext: 'md', mime: 'text/markdown;charset=utf-8', kind: 'document', description: 'Разметка Markdown' },
  html: { id: 'html', label: 'HTML', ext: 'html', mime: 'text/html;charset=utf-8', kind: 'document', description: 'Веб-страница' },
  csv: { id: 'csv', label: 'CSV', ext: 'csv', mime: 'text/csv;charset=utf-8', kind: 'data', description: 'Таблица CSV' },
  tsv: {
    id: 'tsv',
    label: 'TSV',
    ext: 'tsv',
    mime: 'text/tab-separated-values;charset=utf-8',
    kind: 'data',
    description: 'Таблица с табуляцией',
  },
  json: { id: 'json', label: 'JSON', ext: 'json', mime: 'application/json', kind: 'data', description: 'Данные JSON' },
  xlsx: {
    id: 'xlsx',
    label: 'Excel',
    ext: 'xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    kind: 'data',
    description: 'Таблица Microsoft Excel',
  },
  xls: { id: 'xls', label: 'Excel 97–2003', ext: 'xls', mime: 'application/vnd.ms-excel', kind: 'data', description: 'Таблица Excel 97–2003' },
  odt: { id: 'odt', label: 'OpenDocument', ext: 'odt', mime: 'application/vnd.oasis.opendocument.text', kind: 'document', description: 'Текст OpenDocument (LibreOffice)' },
  rtf: { id: 'rtf', label: 'RTF', ext: 'rtf', mime: 'application/rtf', kind: 'document', description: 'Документ RTF (WordPad, Word)' },
  ods: { id: 'ods', label: 'OpenDocument Calc', ext: 'ods', mime: 'application/vnd.oasis.opendocument.spreadsheet', kind: 'data', description: 'Таблица OpenDocument (LibreOffice)' },
  png: { id: 'png', label: 'PNG', ext: 'png', mime: 'image/png', kind: 'image', description: 'Изображение PNG' },
  jpg: { id: 'jpg', label: 'JPG', ext: 'jpg', mime: 'image/jpeg', kind: 'image', description: 'Изображение JPEG' },
  webp: { id: 'webp', label: 'WebP', ext: 'webp', mime: 'image/webp', kind: 'image', description: 'Изображение WebP' },
  gif: { id: 'gif', label: 'GIF', ext: 'gif', mime: 'image/gif', kind: 'image', description: 'Изображение GIF' },
  bmp: { id: 'bmp', label: 'BMP', ext: 'bmp', mime: 'image/bmp', kind: 'image', description: 'Растровое изображение BMP' },
  svg: { id: 'svg', label: 'SVG', ext: 'svg', mime: 'image/svg+xml', kind: 'image', description: 'Векторное изображение SVG' },
  ico: { id: 'ico', label: 'ICO', ext: 'ico', mime: 'image/x-icon', kind: 'image', description: 'Иконка ICO' },
  avif: { id: 'avif', label: 'AVIF', ext: 'avif', mime: 'image/avif', kind: 'image', description: 'Изображение AVIF' },
};

/** Formats handled by the document/data engine (no canvas required). */
export const DOC_READABLE: FormatId[] = ['pdf', 'docx', 'doc', 'odt', 'rtf', 'txt', 'md', 'html', 'csv', 'tsv', 'json', 'xlsx', 'xls', 'ods'];
export const DOC_WRITABLE: FormatId[] = ['pdf', 'docx', 'odt', 'rtf', 'txt', 'md', 'html', 'csv', 'tsv', 'json', 'xlsx', 'ods'];
export const IMAGE_INPUTS: FormatId[] = ['png', 'jpg', 'webp', 'gif', 'bmp', 'svg', 'ico', 'avif'];
/** Image outputs the browser can always encode (avif is feature-detected at runtime). */
export const IMAGE_OUTPUTS: FormatId[] = ['png', 'jpg', 'webp', 'bmp', 'ico'];

const TARGETS: Partial<Record<FormatId, FormatId[]>> = {
  pdf: ['docx', 'odt', 'rtf', 'txt', 'md', 'html', 'pdf', 'png', 'jpg'],
  docx: ['pdf', 'odt', 'rtf', 'txt', 'md', 'html', 'xlsx', 'csv'],
  doc: ['docx', 'pdf', 'odt', 'rtf', 'txt', 'md', 'html', 'xlsx', 'csv'],
  odt: ['docx', 'pdf', 'rtf', 'txt', 'md', 'html', 'xlsx', 'ods', 'csv'],
  rtf: ['docx', 'pdf', 'odt', 'txt', 'md', 'html'],
  txt: ['pdf', 'docx', 'odt', 'rtf', 'md', 'html'],
  md: ['pdf', 'docx', 'odt', 'rtf', 'html', 'txt'],
  html: ['pdf', 'docx', 'odt', 'rtf', 'md', 'txt', 'xlsx', 'csv'],
  csv: ['xlsx', 'ods', 'json', 'tsv', 'html', 'pdf', 'docx', 'md'],
  tsv: ['xlsx', 'ods', 'json', 'csv', 'html', 'pdf', 'docx', 'md'],
  json: ['xlsx', 'ods', 'csv', 'tsv', 'html', 'pdf', 'docx', 'md'],
  xlsx: ['csv', 'ods', 'tsv', 'json', 'html', 'pdf', 'docx', 'md'],
  xls: ['xlsx', 'ods', 'csv', 'tsv', 'json', 'html', 'pdf', 'docx', 'md'],
  ods: ['xlsx', 'csv', 'tsv', 'json', 'html', 'pdf', 'docx', 'md'],
};

/** Available output formats for a given input format. */
export function targetsFor(from: FormatId, opts: { avif?: boolean } = {}): FormatId[] {
  if (FORMATS[from]?.kind === 'image') {
    const outs: FormatId[] = [...IMAGE_OUTPUTS, 'pdf', 'docx', 'txt'];
    if (opts.avif) outs.splice(3, 0, 'avif');
    return outs.filter((f) => f !== from);
  }
  return TARGETS[from] ?? [];
}

/** Pipeline that executes a conversion. */
export function pipelineFor(from: FormatId, to: FormatId): 'doc' | 'raster' | 'images-to-pdf' | 'ocr' | 'ocr-pdf' | null {
  if (!targetsFor(from, { avif: true }).includes(to)) return null;
  if (FORMATS[from].kind === 'image') return to === 'pdf' ? 'images-to-pdf' : to === 'docx' || to === 'txt' ? 'ocr' : 'raster';
  if (from === 'pdf' && (to === 'png' || to === 'jpg')) return 'raster';
  if (from === 'pdf' && to === 'pdf') return 'ocr-pdf';
  return 'doc';
}

/** Human label of a conversion target (OCR variants are spelled out). */
export function targetLabel(from: FormatId, to: FormatId): string {
  const base = `${FORMATS[to].label} (.${FORMATS[to].ext})`;
  const p = pipelineFor(from, to);
  if (p === 'ocr') return `${base} — распознать текст`;
  if (p === 'ocr-pdf') return 'PDF с распознанным текстом (OCR)';
  return base;
}

const EXT_MAP: Record<string, FormatId> = {
  pdf: 'pdf',
  docx: 'docx',
  doc: 'doc',
  dot: 'doc',
  docm: 'docx',
  dotx: 'docx',
  txt: 'txt',
  text: 'txt',
  log: 'txt',
  md: 'md',
  markdown: 'md',
  mdown: 'md',
  html: 'html',
  htm: 'html',
  xhtml: 'html',
  csv: 'csv',
  tsv: 'tsv',
  tab: 'tsv',
  json: 'json',
  xlsx: 'xlsx',
  xlsm: 'xlsx',
  xls: 'xls',
  xlt: 'xls',
  odt: 'odt',
  ott: 'odt',
  fodt: 'odt',
  ods: 'ods',
  ots: 'ods',
  rtf: 'rtf',
  png: 'png',
  jpg: 'jpg',
  jpeg: 'jpg',
  jfif: 'jpg',
  webp: 'webp',
  gif: 'gif',
  bmp: 'bmp',
  svg: 'svg',
  ico: 'ico',
  avif: 'avif',
};

const UNSUPPORTED_HINTS: Record<string, string> = {
  ppt: 'Презентации пока не поддерживаются.',
  pptx: 'Презентации пока не поддерживаются.',
  heic: 'Формат HEIC не поддерживается браузером. Сохраните фото как JPG.',
  tif: 'Формат TIFF не поддерживается браузером. Сохраните изображение как PNG.',
  tiff: 'Формат TIFF не поддерживается браузером. Сохраните изображение как PNG.',
};

export interface Detection {
  format: FormatId | null;
  /** Human-readable reason when the format is not supported. */
  error?: string;
}

function startsWith(bytes: Uint8Array, sig: number[], offset = 0): boolean {
  if (bytes.length < offset + sig.length) return false;
  for (let i = 0; i < sig.length; i++) if (bytes[offset + i] !== sig[i]) return false;
  return true;
}

function cfbKind(bytes: Uint8Array): FormatId | null {
  try {
    const cfb = new Cfb(bytes);
    if (cfb.has('WordDocument')) return 'doc';
    if (cfb.has('Workbook') || cfb.has('Book')) return 'xls';
    return null;
  } catch {
    return null;
  }
}

function zipKind(bytes: Uint8Array): FormatId | null {
  try {
    let kind: FormatId | null = null;
    // Entry names only; nothing but the tiny `mimetype` entry is inflated.
    const small = unzipSync(bytes, {
      filter: (f) => {
        if (f.name === 'word/document.xml') kind = 'docx';
        else if (f.name === 'xl/workbook.xml') kind = 'xlsx';
        return f.name === 'mimetype' && f.originalSize < 200;
      },
    });
    const mime = small['mimetype'] ? new TextDecoder().decode(small['mimetype']).trim() : '';
    if (mime === 'application/vnd.oasis.opendocument.text' || mime === 'application/vnd.oasis.opendocument.text-template') return 'odt';
    if (mime === 'application/vnd.oasis.opendocument.spreadsheet' || mime === 'application/vnd.oasis.opendocument.spreadsheet-template') return 'ods';
    return kind;
  } catch {
    return null;
  }
}

/** Detect format by magic bytes first, extension second. */
export function detectFormat(name: string, bytes: Uint8Array): Detection {
  const ext = extensionOf(name);
  if (startsWith(bytes, [0x25, 0x50, 0x44, 0x46])) return { format: 'pdf' };
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47])) return { format: 'png' };
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return { format: 'jpg' };
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return { format: 'gif' };
  if (startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) && startsWith(bytes, [0x57, 0x45, 0x42, 0x50], 8)) return { format: 'webp' };
  if (startsWith(bytes, [0x00, 0x00, 0x01, 0x00]) && (ext === 'ico' || ext === 'cur' || !ext)) return { format: 'ico' };
  if (startsWith(bytes, [0x66, 0x74, 0x79, 0x70], 4) && (startsWith(bytes, [0x61, 0x76, 0x69], 8) || ext === 'avif')) return { format: 'avif' };
  if (startsWith(bytes, [0x42, 0x4d]) && (ext === 'bmp' || ext === 'dib' || !ext)) return { format: 'bmp' };
  if (startsWith(bytes, [0x7b, 0x5c, 0x72, 0x74, 0x66])) return { format: 'rtf' };
  if (startsWith(bytes, [0xd0, 0xcf, 0x11, 0xe0])) {
    const kind = cfbKind(bytes);
    if (kind) return { format: kind };
    if (UNSUPPORTED_HINTS[ext]) return { format: null, error: UNSUPPORTED_HINTS[ext] };
    if (['doc', 'dot', 'xls', 'xlt'].includes(ext)) return { format: null, error: 'Файл повреждён: не удалось прочитать структуру документа Office.' };
    return { format: null, error: 'Этот файл Office старого формата не распознан (поддерживаются .doc и .xls).' };
  }
  if (startsWith(bytes, [0x50, 0x4b, 0x03, 0x04])) {
    const kind = zipKind(bytes);
    if (kind) return { format: kind };
    return { format: null, error: UNSUPPORTED_HINTS[ext] ?? 'Этот ZIP-архив не является документом Word, Excel или OpenDocument.' };
  }
  if (UNSUPPORTED_HINTS[ext]) return { format: null, error: UNSUPPORTED_HINTS[ext] };
  const byExt = EXT_MAP[ext];
  if (byExt && FORMATS[byExt].kind !== 'image') return { format: byExt };
  if (byExt === 'svg') return { format: 'svg' };
  // Sniff text content
  const head = new TextDecoder('utf-8').decode(bytes.subarray(0, 2048)).replace(/^﻿/, '').trimStart().toLowerCase();
  if (head.startsWith('<svg') || (head.startsWith('<?xml') && head.includes('<svg'))) return { format: 'svg' };
  if (head.startsWith('<!doctype html') || head.startsWith('<html')) return { format: 'html' };
  if (head.startsWith('{') || head.startsWith('[')) return { format: 'json' };
  if (byExt) return { format: byExt };
  if (isProbablyText(bytes)) return { format: 'txt' };
  return { format: null, error: 'Не удалось определить формат файла.' };
}

function isProbablyText(bytes: Uint8Array): boolean {
  const n = Math.min(bytes.length, 4096);
  let bad = 0;
  for (let i = 0; i < n; i++) {
    const b = bytes[i];
    if (b === 0) return false;
    if (b < 9 || (b > 13 && b < 32)) bad++;
  }
  return bad / Math.max(1, n) < 0.02;
}

/** `accept` attribute value (".a,.b,…") for file pickers that take the given formats. */
export function acceptFor(formats: FormatId[]): string {
  return Object.entries(EXT_MAP)
    .filter(([, f]) => formats.includes(f))
    .map(([ext]) => `.${ext}`)
    .join(',');
}

export function formatFromExtension(ext: string): FormatId | null {
  return EXT_MAP[ext.toLowerCase()] ?? null;
}
