/**
 * Conversion router for formats that need neither PDF libraries nor fonts.
 * Safe to bundle into the Cloudflare Worker (small, CPU-cheap).
 */
import { zipSync } from 'fflate';
import type { FormatId } from './formats';
import { FORMATS } from './formats';
import type { Content, Doc, Workbook } from './model';
import { bookToDoc, docToBook } from './model';
import { csvToBook, jsonToBook, xlsxToBook } from './readers/data';
import { docxToDoc } from './readers/docx';
import { htmlToDoc } from './readers/html';
import { markdownToDoc } from './readers/markdown';
import { textToDoc } from './readers/text';
import { bookToJson, bookToXlsx, sheetToCsv } from './writers/data';
import { docToDocx } from './writers/docx';
import { docToHtml, docToMarkdown, docToText } from './writers/text';
import { ConvertError, decodeText, utf8 } from './util';

export interface ConvertOptions {
  pageSize?: 'a4' | 'letter';
  /** CSV output delimiter. */
  csvDelimiter?: ',' | ';';
}

export interface ConvertOutput {
  data: Uint8Array;
  mime: string;
  ext: string;
  warnings: string[];
}

export const LITE_FORMATS: FormatId[] = ['docx', 'txt', 'md', 'html', 'csv', 'tsv', 'json', 'xlsx'];

export function readLite(bytes: Uint8Array, from: FormatId, title?: string): { content: Content; warnings: string[] } {
  switch (from) {
    case 'txt':
      return { content: { kind: 'doc', doc: textToDoc(decodeText(bytes), title) }, warnings: [] };
    case 'md':
      return { content: { kind: 'doc', doc: markdownToDoc(decodeText(bytes), title) }, warnings: [] };
    case 'html': {
      const doc = htmlToDoc(decodeText(bytes));
      doc.title ??= title;
      return { content: { kind: 'doc', doc }, warnings: [] };
    }
    case 'docx': {
      const { doc, warnings } = docxToDoc(bytes);
      doc.title ??= title;
      return { content: { kind: 'doc', doc }, warnings };
    }
    case 'csv':
      return { content: { kind: 'book', book: csvToBook(decodeText(bytes)) }, warnings: [] };
    case 'tsv':
      return { content: { kind: 'book', book: csvToBook(decodeText(bytes), '\t') }, warnings: [] };
    case 'json':
      return { content: { kind: 'book', book: jsonToBook(decodeText(bytes)) }, warnings: [] };
    case 'xlsx':
      return { content: { kind: 'book', book: xlsxToBook(bytes) }, warnings: [] };
    default:
      throw new ConvertError(`Чтение формата ${FORMATS[from]?.label ?? from} здесь не поддерживается.`, 'unsupported');
  }
}

export function asDoc(content: Content, title?: string): Doc {
  return content.kind === 'doc' ? content.doc : bookToDoc(content.book, title);
}

export function asBook(content: Content): Workbook {
  if (content.kind === 'book') return content.book;
  const book = docToBook(content.doc);
  if (!book.sheets.length) throw new ConvertError('В документе нет таблиц, которые можно сохранить как таблицу данных.', 'no_tables');
  return book;
}

const out = (data: Uint8Array, to: FormatId, warnings: string[] = []): ConvertOutput => ({ data, mime: FORMATS[to].mime, ext: FORMATS[to].ext, warnings });

export function writeLite(content: Content, to: FormatId, opts: ConvertOptions = {}, title?: string): ConvertOutput {
  switch (to) {
    case 'txt':
      return out(utf8(docToText(asDoc(content, title))), to);
    case 'md':
      return out(utf8(docToMarkdown(asDoc(content, title))), to);
    case 'html': {
      const doc = asDoc(content, title);
      return out(utf8(docToHtml({ ...doc, title: doc.title ?? title })), to);
    }
    case 'docx': {
      const doc = asDoc(content, title);
      const r = docToDocx({ ...doc, title: doc.title ?? title }, opts);
      return out(r.data, to, r.warnings);
    }
    case 'csv':
    case 'tsv': {
      const book = asBook(content);
      const delim = to === 'tsv' ? '\t' : (opts.csvDelimiter ?? ',');
      // UTF-8 BOM so that Excel opens Cyrillic text correctly
      const encode = (s: string): Uint8Array => utf8((to === 'csv' ? '﻿' : '') + s);
      if (book.sheets.length === 1) return out(encode(sheetToCsv(book.sheets[0], delim)), to);
      const files: Record<string, Uint8Array> = {};
      book.sheets.forEach((s, i) => {
        const safe = s.name.replace(/[\\/:*?"<>|]+/g, '_').trim() || `sheet${i + 1}`;
        files[`${String(i + 1).padStart(2, '0')}_${safe}.${to}`] = encode(sheetToCsv(s, delim));
      });
      return {
        data: zipSync(files),
        mime: 'application/zip',
        ext: 'zip',
        warnings: [`В книге ${book.sheets.length} листа(ов) — каждый сохранён отдельным файлом в ZIP-архиве.`],
      };
    }
    case 'json':
      return out(utf8(bookToJson(asBook(content))), to);
    case 'xlsx':
      return out(bookToXlsx(asBook(content)), to);
    default:
      throw new ConvertError(`Сохранение в формат ${FORMATS[to]?.label ?? to} здесь не поддерживается.`, 'unsupported');
  }
}
