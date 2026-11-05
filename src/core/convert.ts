/**
 * Full document/data conversion router (browser Web Worker, Node tests).
 * PDF support is loaded lazily so that light conversions stay light.
 */
import type { FormatId } from './formats';
import { FORMATS, targetsFor } from './formats';
import type { Content, Doc } from './model';
import type { FontLoader } from './writers/pdf';
import { asDoc, readLite, writeLite, type ConvertOptions, type ConvertOutput } from './convert-lite';
import { ConvertError } from './util';

export type { ConvertOptions, ConvertOutput };

type PdfJs = typeof import('pdfjs-dist');

export interface ConvertEnv {
  pdfjs?: () => Promise<PdfJs>;
  pdfDocParams?: Record<string, unknown>;
  fonts?: FontLoader;
  encodeImage?: (pixels: Uint8Array | Uint8ClampedArray, width: number, height: number, channels: 1 | 3 | 4) => Promise<{ data: Uint8Array; mime: string }>;
  transcodeImage?: (data: Uint8Array, mime: string) => Promise<{ data: Uint8Array; mime: string } | null>;
  onProgress?: (fraction: number) => void;
}

export async function readContent(bytes: Uint8Array, from: FormatId, env: ConvertEnv, title?: string): Promise<{ content: Content; warnings: string[] }> {
  if (from === 'pdf') {
    if (!env.pdfjs) throw new ConvertError('Чтение PDF недоступно в этой среде.', 'unsupported');
    const [{ pdfToDoc }, pdfjs] = await Promise.all([import('./readers/pdf'), env.pdfjs()]);
    const { doc, warnings } = await pdfToDoc(bytes, {
      pdfjs,
      docParams: env.pdfDocParams,
      encodeImage: env.encodeImage,
      onProgress: (f) => env.onProgress?.(f * 0.7),
    });
    doc.title ??= title;
    return { content: { kind: 'doc', doc }, warnings };
  }
  return readLite(bytes, from, title);
}

export async function writeContent(content: Content, to: FormatId, env: ConvertEnv, opts: ConvertOptions = {}, title?: string): Promise<ConvertOutput> {
  if (to === 'pdf') {
    if (!env.fonts) throw new ConvertError('Создание PDF недоступно в этой среде.', 'unsupported');
    const { docToPdf } = await import('./writers/pdf');
    const doc: Doc = asDoc(content, title);
    const r = await docToPdf({ ...doc, title: doc.title ?? title }, { fonts: env.fonts, transcodeImage: env.transcodeImage }, { pageSize: opts.pageSize });
    return { data: r.data, mime: FORMATS.pdf.mime, ext: 'pdf', warnings: r.warnings };
  }
  return writeLite(content, to, opts, title);
}

export async function convert(bytes: Uint8Array, from: FormatId, to: FormatId, env: ConvertEnv, opts: ConvertOptions = {}, title?: string): Promise<ConvertOutput> {
  if (!targetsFor(from).includes(to)) {
    throw new ConvertError(`Преобразование ${FORMATS[from]?.label ?? from} → ${FORMATS[to]?.label ?? to} не поддерживается.`, 'unsupported');
  }
  const { content, warnings } = await readContent(bytes, from, env, title);
  env.onProgress?.(0.75);
  const res = await writeContent(content, to, env, opts, title);
  env.onProgress?.(1);
  return { ...res, warnings: [...new Set([...warnings, ...res.warnings])] };
}
