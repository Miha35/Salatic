/** CSV/TSV, JSON and XLSX → Workbook. */
import { unzipSync } from 'fflate';
import type { CellValue, Sheet, Workbook } from '../model';
import { child, childrenNamed, find, findAll, parseXml, rootElement, textContent, type XNode } from '../xml';
import { ConvertError, unzipOffice } from '../util';

// ---------------------------------------------------------------------------
// CSV

/** Guess the delimiter from the first lines (comma, semicolon, tab, pipe). */
export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 20000).split(/\r?\n/).filter((l) => l.trim()).slice(0, 20);
  let best = ',';
  let bestScore = -1;
  for (const d of [',', ';', '\t', '|']) {
    const counts = sample.map((l) => countOutsideQuotes(l, d));
    if (!counts.length || counts[0] === 0) continue;
    const consistent = counts.filter((c) => c === counts[0]).length / counts.length;
    const score = consistent * 1000 + counts[0];
    if (score > bestScore) {
      best = d;
      bestScore = score;
    }
  }
  return best;
}

function countOutsideQuotes(line: string, d: string): number {
  let n = 0;
  let q = false;
  for (const ch of line) {
    if (ch === '"') q = !q;
    else if (ch === d && !q) n++;
  }
  return n;
}

/** RFC 4180 parser (quoted fields, escaped quotes, CRLF, newlines inside quotes). */
export function parseCsv(text: string, delimiter?: string): string[][] {
  text = text.replace(/^﻿/, '');
  const d = delimiter ?? detectDelimiter(text);
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  const n = text.length;
  let quoted = false;
  while (i < n) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      quoted = true;
      i++;
    } else if (ch === d) {
      row.push(field);
      field = '';
      i++;
    } else if (ch === '\n' || ch === '\r') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
    } else {
      field += ch;
      i++;
    }
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  // Drop trailing completely empty rows
  while (rows.length && rows[rows.length - 1].every((c) => c === '')) rows.pop();
  return rows;
}

export function csvToBook(text: string, delimiter?: string): Workbook {
  const rows = parseCsv(text, delimiter);
  const header = rows.length > 1 && rows[0].every((c) => c.trim() !== '' && !/^-?[\d\s.,]+$/.test(c));
  return { sheets: [{ name: 'Лист1', rows, header }] };
}

// ---------------------------------------------------------------------------
// JSON

function toCell(v: unknown): CellValue {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  return JSON.stringify(v);
}

function arrayToSheet(name: string, arr: unknown[]): Sheet {
  if (arr.every((x) => Array.isArray(x))) {
    return { name, rows: (arr as unknown[][]).map((r) => r.map(toCell)) };
  }
  if (arr.every((x) => x !== null && typeof x === 'object' && !Array.isArray(x))) {
    const keys: string[] = [];
    const seen = new Set<string>();
    for (const o of arr as Record<string, unknown>[]) {
      for (const k of Object.keys(o)) {
        if (!seen.has(k)) {
          seen.add(k);
          keys.push(k);
        }
      }
    }
    const rows: CellValue[][] = [keys];
    for (const o of arr as Record<string, unknown>[]) rows.push(keys.map((k) => toCell(o[k])));
    return { name, rows, header: true };
  }
  return { name, rows: [['value'], ...arr.map((v) => [toCell(v)])], header: true };
}

export function jsonToBook(text: string): Workbook {
  let data: unknown;
  try {
    data = JSON.parse(text.replace(/^﻿/, ''));
  } catch (e) {
    throw new ConvertError(`Некорректный JSON: ${(e as Error).message}`);
  }
  if (Array.isArray(data)) return { sheets: [arrayToSheet('Лист1', data)] };
  if (data && typeof data === 'object') {
    const entries = Object.entries(data as Record<string, unknown>);
    if (entries.length && entries.every(([, v]) => Array.isArray(v))) {
      return { sheets: entries.map(([k, v]) => arrayToSheet(k, v as unknown[])) };
    }
    return { sheets: [{ name: 'Лист1', rows: [['key', 'value'], ...entries.map(([k, v]) => [k, toCell(v)])], header: true }] };
  }
  return { sheets: [{ name: 'Лист1', rows: [['value'], [toCell(data)]], header: true }] };
}

// ---------------------------------------------------------------------------
// XLSX

const decoder = new TextDecoder('utf-8');
export const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);

function colIndex(ref: string): number {
  let n = 0;
  for (const ch of ref) {
    const c = ch.charCodeAt(0);
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

export function isDateFormat(code: string): boolean {
  const stripped = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  return /[dmyhs]/i.test(stripped) && !/^[#0.,%\s]*$/.test(stripped);
}

export function excelDate(serial: number, date1904: boolean): string {
  const ms = Math.round((serial - (date1904 ? -1462 : 25569)) * 86400000);
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return String(serial);
  const iso = d.toISOString();
  const timePart = iso.slice(11, 19);
  if (timePart === '00:00:00') return iso.slice(0, 10);
  if (serial < 1) return timePart;
  return `${iso.slice(0, 10)} ${timePart}`;
}

export function xlsxToBook(bytes: Uint8Array): Workbook {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipOffice(bytes, unzipSync);
  } catch (e) {
    if (e instanceof ConvertError) throw e;
    throw new ConvertError('Файл XLSX повреждён.');
  }
  const xml = (p: string): XNode | undefined => (files[p] ? parseXml(decoder.decode(files[p])) : undefined);
  const wb = xml('xl/workbook.xml');
  if (!wb) throw new ConvertError('В архиве нет книги Excel (xl/workbook.xml).');
  const date1904 = ['1', 'true'].includes(find(wb, 'workbookPr')?.attrs.date1904 ?? '');
  const rels = new Map<string, string>();
  for (const r of childrenNamed(rootElement(xml('xl/_rels/workbook.xml.rels') ?? { name: '', attrs: {}, children: [] }), 'Relationship')) {
    const t = r.attrs.Target;
    rels.set(r.attrs.Id, t.startsWith('/') ? t.slice(1) : `xl/${t}`);
  }
  // Shared strings
  const shared: string[] = [];
  const sst = xml('xl/sharedStrings.xml');
  for (const si of childrenNamed(rootElement(sst ?? { name: '', attrs: {}, children: [] }), 'si')) {
    const t = child(si, 't');
    shared.push(t ? textContent(t) : childrenNamed(si, 'r').map((r) => textContent(child(r, 't'))).join(''));
  }
  // Date styles
  const dateStyles = new Set<number>();
  const styles = xml('xl/styles.xml');
  if (styles) {
    const custom = new Map<number, string>();
    for (const nf of findAll(find(styles, 'numFmts'), 'numFmt')) custom.set(+nf.attrs.numFmtId, nf.attrs.formatCode ?? '');
    childrenNamed(find(styles, 'cellXfs'), 'xf').forEach((xf, idx) => {
      const id = +(xf.attrs.numFmtId ?? 0);
      if (BUILTIN_DATE_FORMATS.has(id) || (custom.has(id) && isDateFormat(custom.get(id)!))) dateStyles.add(idx);
    });
  }
  const sheets: Sheet[] = [];
  for (const s of findAll(find(wb, 'sheets'), 'sheet')) {
    const path = rels.get(s.attrs['r:id']);
    const sx = path ? xml(path) : undefined;
    if (!sx) continue;
    const rows: CellValue[][] = [];
    const data = find(sx, 'sheetData');
    let rowCursor = 0;
    for (const r of childrenNamed(data, 'row')) {
      const ri = r.attrs.r ? +r.attrs.r - 1 : rowCursor;
      rowCursor = ri + 1;
      const row: CellValue[] = [];
      let colCursor = 0;
      for (const c of childrenNamed(r, 'c')) {
        const ci = c.attrs.r ? colIndex(c.attrs.r) : colCursor;
        colCursor = ci + 1;
        const t = c.attrs.t ?? 'n';
        const v = child(c, 'v');
        let value: CellValue = null;
        if (t === 's') value = shared[+textContent(v)] ?? '';
        else if (t === 'inlineStr') value = findAll(child(c, 'is'), 't').map(textContent).join('');
        else if (t === 'str' || t === 'e') value = textContent(v);
        else if (t === 'b') value = textContent(v) === '1';
        else if (t === 'd') value = textContent(v);
        else if (v) {
          const num = +textContent(v);
          value = Number.isFinite(num) ? (dateStyles.has(+(c.attrs.s ?? -1)) ? excelDate(num, date1904) : num) : textContent(v);
        }
        if (ci < 16384) row[ci] = value;
      }
      for (let k = 0; k < row.length; k++) if (row[k] === undefined) row[k] = null;
      if (ri < 1048576) rows[ri] = row;
    }
    for (let k = 0; k < rows.length; k++) if (!rows[k]) rows[k] = [];
    while (rows.length && rows[rows.length - 1].every((v) => v === null || v === '')) rows.pop();
    sheets.push({ name: s.attrs.name ?? `Лист${sheets.length + 1}`, rows });
  }
  if (!sheets.length) throw new ConvertError('В книге Excel не найдено листов.');
  return { sheets };
}
