/** Workbook → CSV/TSV, JSON, XLSX. */
import { zipSync, type Zippable } from 'fflate';
import type { CellValue, Sheet, Workbook } from '../model';
import { cellToString } from '../model';
import { escapeXml, maxOf, utf8 } from '../util';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function sheetToCsv(sheet: Sheet, delimiter = ','): string {
  const needsQuote = new RegExp(`["\\r\\n${delimiter === '\t' ? '\\t' : delimiter.replace(/[|\\]/g, '\\$&')}]|^\\s|\\s$`);
  const width = sheet.rows.reduce((m, r) => Math.max(m, r.length), 0);
  return (
    sheet.rows
      .map((r) => {
        const cells: string[] = [];
        for (let i = 0; i < width; i++) {
          const s = cellToString(r[i] ?? null);
          cells.push(needsQuote.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
        }
        return cells.join(delimiter);
      })
      .join('\r\n') + '\r\n'
  );
}

function sheetToJsonValue(sheet: Sheet): unknown[] {
  const [head, ...rest] = sheet.rows;
  const keys = (head ?? []).map((k) => cellToString(k).trim());
  const usableHeader = sheet.header !== false && keys.length > 0 && keys.every((k) => k) && new Set(keys).size === keys.length && rest.length > 0;
  if (usableHeader) {
    return rest.map((r) => {
      const o: Record<string, CellValue> = {};
      keys.forEach((k, i) => (o[k] = r[i] ?? null));
      return o;
    });
  }
  return sheet.rows;
}

export function bookToJson(book: Workbook): string {
  if (book.sheets.length === 1) return JSON.stringify(sheetToJsonValue(book.sheets[0]), null, 2) + '\n';
  const o: Record<string, unknown[]> = {};
  for (const s of book.sheets) o[s.name] = sheetToJsonValue(s);
  return JSON.stringify(o, null, 2) + '\n';
}

// ---------------------------------------------------------------------------
// XLSX

function colName(i: number): string {
  let s = '';
  i++;
  while (i > 0) {
    const m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

const NUMERIC = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;

/** Strings that look like numbers become numbers (except IDs with leading zeros and very long digit strings). */
function coerce(v: CellValue): CellValue {
  if (typeof v !== 'string') return v;
  const t = v.trim();
  if (t !== v || !NUMERIC.test(t)) return v;
  if (/^[-+]?0\d/.test(t) || t.replace(/\D/g, '').length > 15) return v;
  const n = Number(t);
  return Number.isFinite(n) ? n : v;
}

function safeSheetNames(sheets: Sheet[]): string[] {
  const used = new Set<string>();
  return sheets.map((s, i) => {
    let base = (s.name || `Лист${i + 1}`).replace(/[[\]:*?/\\]/g, ' ').replace(/^'+|'+$/g, '').trim().slice(0, 31) || `Лист${i + 1}`;
    let name = base;
    let k = 2;
    while (used.has(name.toLowerCase())) {
      const suffix = ` (${k++})`;
      name = base.slice(0, 31 - suffix.length) + suffix;
    }
    used.add(name.toLowerCase());
    base = name;
    return name;
  });
}

function sheetXml(sheet: Sheet): string {
  const width = sheet.rows.reduce((m, r) => Math.max(m, r.length), 0);
  const widths = new Array(width).fill(8);
  let rowsXml = '';
  sheet.rows.forEach((r, ri) => {
    let cells = '';
    for (let ci = 0; ci < r.length; ci++) {
      const v = coerce(r[ci] ?? null);
      if (v === null || v === '') continue;
      const ref = `${colName(ci)}${ri + 1}`;
      const style = ri === 0 && sheet.header ? ' s="1"' : '';
      const len = typeof v === 'string' ? maxOf(v.split('\n').map((l) => l.length), 0) : String(v).length;
      widths[ci] = Math.min(60, Math.max(widths[ci], len + 2));
      if (typeof v === 'number') cells += `<c r="${ref}"${style}><v>${v}</v></c>`;
      else if (typeof v === 'boolean') cells += `<c r="${ref}"${style} t="b"><v>${v ? 1 : 0}</v></c>`;
      else cells += `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${escapeXml(v)}</t></is></c>`;
    }
    rowsXml += `<row r="${ri + 1}">${cells}</row>`;
  });
  const dim = width && sheet.rows.length ? `A1:${colName(width - 1)}${sheet.rows.length}` : 'A1';
  const cols = width ? `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '';
  const frozen = sheet.header && sheet.rows.length > 1 ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' : '';
  return (
    XML_HEAD +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<dimension ref="${dim}"/>${frozen}<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rowsXml}</sheetData></worksheet>`
  );
}

export function bookToXlsx(book: Workbook): Uint8Array {
  const sheets = book.sheets.length ? book.sheets : [{ name: 'Лист1', rows: [] }];
  const names = safeSheetNames(sheets);
  const files: Zippable = {
    '[Content_Types].xml': utf8(
      XML_HEAD +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': utf8(
      XML_HEAD +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>',
    ),
    'docProps/app.xml': utf8(XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Salatic</Application></Properties>'),
    'xl/workbook.xml': utf8(
      XML_HEAD +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView/></bookViews><sheets>' +
        names.map((n, i) => `<sheet name="${escapeXml(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
        '</sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': utf8(
      XML_HEAD +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
        `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        '</Relationships>',
    ),
    'xl/styles.xml': utf8(
      XML_HEAD +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>',
    ),
  };
  sheets.forEach((s, i) => (files[`xl/worksheets/sheet${i + 1}.xml`] = utf8(sheetXml(s))));
  return zipSync(files, { level: 6 });
}
