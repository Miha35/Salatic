import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { convert } from '../../src/core/convert';
import { DOC_READABLE, FORMATS, pipelineFor, targetsFor, type FormatId } from '../../src/core/formats';
import { docToDocx } from '../../src/core/writers/docx';
import { docToPdf } from '../../src/core/writers/pdf';
import { docToHtml, docToMarkdown, docToText } from '../../src/core/writers/text';
import { bookToJson, bookToXlsx, sheetToCsv } from '../../src/core/writers/data';
import { docxToDoc } from '../../src/core/readers/docx';
import { xlsxToBook } from '../../src/core/readers/data';
import { readLite, writeLite } from '../../src/core/convert-lite';
import { bookToOds, docToOdt } from '../../src/core/writers/odf';
import { odsToBook, odtToDoc } from '../../src/core/readers/odf';
import { docToRtf } from '../../src/core/writers/rtf';
import { rtfToDoc } from '../../src/core/readers/rtf';
import { docToEpub } from '../../src/core/writers/epub';
import { epubToDoc } from '../../src/core/readers/epub';
import type { Workbook } from '../../src/core/model';
import { bytes, env, hasCmd, loadFonts, pdfToText, soffice, text } from './helpers';
import { sample } from './sample';

const book: Workbook = {
  sheets: [
    {
      name: 'Продажи',
      header: true,
      rows: [
        ['Товар', 'Количество', 'Цена'],
        ['Яблоки', 10, 99.5],
        ['Груши, сорт «Лесная»', 3, 120],
      ],
    },
  ],
};

async function fixture(f: FormatId): Promise<Uint8Array> {
  switch (f) {
    case 'pdf':
      return (await docToPdf(sample, { fonts: loadFonts })).data;
    case 'docx':
      return docToDocx(sample).data;
    case 'txt':
      return bytes(docToText(sample));
    case 'md':
      return bytes(docToMarkdown(sample));
    case 'html':
      return bytes(docToHtml(sample));
    case 'csv':
      return bytes(sheetToCsv(book.sheets[0], ';'));
    case 'tsv':
      return bytes(sheetToCsv(book.sheets[0], '\t'));
    case 'json':
      return bytes(bookToJson(book));
    case 'xlsx':
      return bookToXlsx(book);
    case 'odt':
      return docToOdt(sample).data;
    case 'rtf':
      return docToRtf(sample).data;
    case 'ods':
      return bookToOds(book);
    case 'epub':
      return docToEpub(sample).data;
    // legacy binary formats are produced by LibreOffice (tests are skipped without it)
    case 'doc':
      return soffice(docToDocx(sample).data, 'docx', 'doc:MS Word 97');
    case 'xls':
      return soffice(bookToXlsx(book), 'xlsx', 'xls:MS Excel 97');
    default:
      throw new Error(f);
  }
}

function validate(to: FormatId, data: Uint8Array, mustContain: string): void {
  expect(data.length).toBeGreaterThan(0);
  switch (to) {
    case 'pdf':
      expect(text(data.subarray(0, 5))).toBe('%PDF-');
      if (hasCmd('pdftotext')) expect(pdfToText(data)).toContain(mustContain);
      break;
    case 'docx': {
      const { doc } = docxToDoc(data);
      expect(JSON.stringify(doc)).toContain(mustContain);
      break;
    }
    case 'xlsx':
      expect(JSON.stringify(xlsxToBook(data))).toContain(mustContain);
      break;
    case 'ods':
      expect(JSON.stringify(odsToBook(data).book)).toContain(mustContain);
      break;
    case 'odt':
      expect(JSON.stringify(odtToDoc(data).doc)).toContain(mustContain);
      break;
    case 'rtf':
      expect(JSON.stringify(rtfToDoc(data).doc)).toContain(mustContain);
      break;
    case 'epub':
      expect(JSON.stringify(epubToDoc(data).doc)).toContain(mustContain);
      break;
    case 'json':
      expect(JSON.stringify(JSON.parse(text(data)))).toContain(mustContain);
      break;
    default:
      expect(text(data)).toContain(mustContain);
  }
}

describe('conversion matrix', () => {
  for (const from of DOC_READABLE) {
    for (const to of targetsFor(from).filter((t) => pipelineFor(from, t) === 'doc')) {
      it.skipIf((from === 'doc' || from === 'xls') && !hasCmd('soffice'))(`${from} → ${to}`, async () => {
        const input = await fixture(from);
        const isData = FORMATS[from].kind === 'data';
        // docx/html → table formats export the table; data → anything keeps the cells
        const expectText = isData ? 'Яблоки' : FORMATS[to].kind === 'data' ? 'Москва' : 'Отчёт о работе';
        const out = await convert(input, from, to, env, {}, 'файл');
        expect(out.ext === FORMATS[to].ext || out.ext === 'zip').toBe(true);
        validate(to, out.data, expectText);
      });
    }
  }

  it('rejects unsupported pairs', async () => {
    await expect(convert(bytes('x'), 'txt', 'xlsx', env)).rejects.toThrow(/не поддерживается/);
  });

  it('explains when a document has no tables for a data format', () => {
    const { content } = readLite(bytes('<p>только текст</p>'), 'html');
    expect(() => writeLite(content, 'csv')).toThrow(/нет таблиц/);
  });

  it('writes multi-sheet workbooks to CSV as a ZIP', () => {
    const xlsx = bookToXlsx({ sheets: [book.sheets[0], { name: 'Второй', rows: [['a'], ['b']] }] });
    const { content } = readLite(xlsx, 'xlsx');
    const out = writeLite(content, 'csv');
    expect(out.ext).toBe('zip');
    const files = unzipSync(out.data);
    expect(Object.keys(files)).toEqual(['01_Продажи.csv', '02_Второй.csv']);
    expect(text(files['01_Продажи.csv'])).toContain('"Груши, сорт «Лесная»"');
  });

  it('uses the semicolon delimiter on request and adds a BOM for Excel', () => {
    const { content } = readLite(bookToXlsx(book), 'xlsx');
    const out = writeLite(content, 'csv', { csvDelimiter: ';' });
    expect([...out.data.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    const s = text(out.data);
    expect(s).toContain('Товар;Количество;Цена');
  });
});
