/** OpenDocument (.odt/.ods): our writer, our reader, and interoperability with LibreOffice. */
import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync } from 'fflate';
import { docToOdt, bookToOds } from '../../src/core/writers/odf';
import { odtToDoc, odsToBook, odfLength } from '../../src/core/readers/odf';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { bookToXlsx } from '../../src/core/writers/data';
import { docToMarkdown } from '../../src/core/writers/text';
import { detectFormat } from '../../src/core/formats';
import { convert } from '../../src/core/convert';
import { ConvertError } from '../../src/core/util';
import { blockText, type Doc } from '../../src/core/model';
import { env, hasCmd, pdfToText, soffice, text } from './helpers';
import { sample } from './sample';

const withExtras: Doc = {
  ...sample,
  header: [{ type: 'paragraph', runs: [{ text: 'Шапка документа' }] }],
  footer: [{ type: 'paragraph', align: 'center', runs: [{ text: 'Стр. ' }, { text: '1', field: 'page' }, { text: ' из ' }, { text: '1', field: 'pages' }] }],
  blocks: [
    ...sample.blocks,
    {
      type: 'paragraph',
      runs: [{ text: 'Сноска тут' }, { text: '', note: [{ type: 'paragraph', runs: [{ text: 'Текст сноски' }] }] }, { text: '  два  пробела', color: 'FF0000', highlight: 'FFFF00', size: 16, font: 'serif' }],
    },
  ],
};

describe('ODT writer', () => {
  const { data, warnings } = docToOdt(withExtras);
  const files = unzipSync(data);

  it('is a valid OpenDocument package', () => {
    // mimetype must be the first entry and stored uncompressed (ODF 1.3 §3.3)
    expect(text(data.subarray(30, 38))).toBe('mimetype');
    expect(data[8] | (data[9] << 8)).toBe(0);
    expect(text(files['mimetype'])).toBe('application/vnd.oasis.opendocument.text');
    for (const p of ['content.xml', 'styles.xml', 'meta.xml', 'META-INF/manifest.xml']) expect(files[p]).toBeDefined();
    expect(Object.keys(files).some((p) => p.startsWith('Pictures/'))).toBe(true);
    expect(warnings).toEqual([]);
    expect(detectFormat('x.zip', data).format).toBe('odt');
  });

  it('round-trips through our reader without losing structure', () => {
    expect(docToMarkdown(odtToDoc(docToOdt(sample).data).doc)).toBe(docToMarkdown(sample));
    const back = odtToDoc(data).doc;
    expect(back.header?.map(blockText)).toEqual(['Шапка документа']);
    const footer = back.footer?.[0];
    expect(footer?.type === 'paragraph' && footer.runs.filter((r) => r.field).map((r) => r.field)).toEqual(['page', 'pages']);
    const last = back.blocks[back.blocks.length - 1];
    expect(last.type).toBe('paragraph');
    if (last.type !== 'paragraph') return;
    const note = last.runs.find((r) => r.note);
    expect(note?.note?.map(blockText)).toEqual(['Текст сноски']);
    const styled = last.runs.find((r) => r.text.includes('два'));
    expect(styled).toMatchObject({ text: '  два  пробела', color: 'FF0000', highlight: 'FFFF00', size: 16, font: 'serif' });
  });

  it.skipIf(!hasCmd('soffice'))('opens in LibreOffice with headers, footers, notes and page numbers', () => {
    const pdf = soffice(data, 'odt', 'pdf');
    const t = pdfToText(pdf);
    expect(t).toContain('Шапка документа');
    expect(t).toMatch(/Стр\. 1 из 2/);
    expect(t).toMatch(/Стр\. 2 из 2/);
    expect(t).toContain('Текст сноски');
    const viaDocx = docxToDoc(soffice(data, 'odt', 'docx:MS Word 2007 XML')).doc;
    const types = viaDocx.blocks.map((b) => b.type);
    for (const t of ['heading', 'list', 'table', 'image', 'pagebreak']) expect(types).toContain(t);
    const list = viaDocx.blocks.find((b) => b.type === 'list');
    const nested = list?.type === 'list' ? list.items.flatMap((i) => i.blocks).find((b) => b.type === 'list') : undefined;
    expect(nested?.type === 'list' && nested.ordered).toBe(true);
  });
});

describe.skipIf(!hasCmd('soffice'))('ODT reader on LibreOffice files', () => {
  const odt = soffice(docToDocx(withExtras).data, 'docx', 'odt');
  const { doc } = odtToDoc(odt);

  it('reads structure, formatting and the document font', () => {
    const md = docToMarkdown(doc);
    expect(md).toContain('# Отчёт о работе');
    expect(md).toContain('## Список задач');
    expect(md).toContain('**жирный**');
    expect(md).toContain('[ссылка](https://example.com/)');
    expect(md).toMatch(/- Второй пункт\n {2}1\. Вложенный один\n {2}2\. Вложенный два/);
    expect(md).toContain('| Москва | 13 100 000 |');
    expect(md).toContain('[^1]: Текст сноски');
    expect(doc.font).toBe('sans'); // Arial, although LibreOffice labels it "roman"
    expect(doc.size).toBe(11);
    expect(doc.blocks.some((b) => b.type === 'image')).toBe(true);
    expect(doc.blocks.some((b) => b.type === 'pagebreak')).toBe(true);
  });

  it('reads the master page: header, footer fields and page geometry', () => {
    expect(doc.header?.map(blockText)).toEqual(['Шапка документа']);
    const f = doc.footer?.[0];
    expect(f?.type === 'paragraph' && f.runs.filter((r) => r.field).length).toBe(2);
    expect(doc.page?.width).toBeCloseTo(595.3, 0);
    expect(doc.page?.height).toBeCloseTo(841.9, 0);
    // the 2 cm body margin of the source is restored from margin + header extent
    expect(doc.page!.margins.top).toBeGreaterThan(50);
  });

  it('converts .odt to every document target', async () => {
    for (const to of ['docx', 'pdf', 'txt', 'md', 'html'] as const) {
      const r = await convert(odt, 'odt', to, env);
      expect(r.data.length).toBeGreaterThan(100);
    }
    const pdf = await convert(odt, 'odt', 'pdf', env);
    expect(pdfToText(pdf.data)).toContain('Съешь же ещё этих мягких');
  });
});

describe('ODS', () => {
  const book = {
    sheets: [
      { name: 'Продажи', header: true, rows: [['Товар', 'Цена', 'Дата', 'Флаг'], ['Яблоки', 99.5, '2024-03-15', true], ['Груши  «Лесные»', -3, '2024-03-15 10:30:00', false], ['007', '12', null, 'a\nb']] },
      { name: 'B', rows: [['x'], [1]] },
    ],
  };

  it('round-trips types, dates, multi-line text and spaces', () => {
    const ods = bookToOds(book);
    expect(detectFormat('t.ods', ods).format).toBe('ods');
    const back = odsToBook(ods).book;
    expect(back.sheets.map((s) => s.name)).toEqual(['Продажи', 'B']);
    expect(back.sheets[0].header).toBe(true);
    expect(back.sheets[0].rows).toEqual([
      ['Товар', 'Цена', 'Дата', 'Флаг'],
      ['Яблоки', 99.5, '2024-03-15', true],
      ['Груши  «Лесные»', -3, '2024-03-15 10:30:00', false],
      ['007', 12, null, 'a\nb'],
    ]);
  });

  it.skipIf(!hasCmd('soffice'))('is read correctly by LibreOffice', () => {
    const csv = text(soffice(bookToOds(book), 'ods', 'csv'));
    expect(csv).toContain('Яблоки,99.5,2024-03-15,TRUE');
    expect(csv).toContain('2024-03-15 10:30:00');
    expect(csv).toContain('"a\nb"');
  });

  it.skipIf(!hasCmd('soffice'))('reads LibreOffice spreadsheets with formulas and dates', () => {
    const csv = 'Дата,Сумма,Итог\n2024-03-15,10,=B2*2\n1999-12-31,2.5,=B3*2\n';
    const ods = soffice(new TextEncoder().encode(csv), 'csv', 'ods');
    const rows = odsToBook(ods).book.sheets[0].rows;
    expect(rows[1]).toEqual(['2024-03-15', 10, 20]);
    expect(rows[2]).toEqual(['1999-12-31', 2.5, 5]);
    const fromXlsx = odsToBook(soffice(bookToXlsx(book), 'xlsx', 'ods')).book;
    expect(fromXlsx.sheets[0].rows[1]).toEqual(['Яблоки', 99.5, '2024-03-15', true]);
  });

  it('does not expand huge runs of repeated empty cells', () => {
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><table:table table:name="S">' +
      '<table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16383"/></table:table-row>' +
      '<table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>' +
      '</table:table></office:spreadsheet></office:body></office:document-content>';
    const ods = zipSync({ mimetype: new TextEncoder().encode('application/vnd.oasis.opendocument.spreadsheet'), 'content.xml': new TextEncoder().encode(content) });
    const t0 = performance.now();
    const b = odsToBook(ods).book;
    expect(performance.now() - t0).toBeLessThan(500);
    expect(b.sheets[0].rows).toEqual([[1]]);
  });

  it('rejects encrypted and mismatched packages with readable errors', () => {
    const enc = zipSync({
      mimetype: new TextEncoder().encode('application/vnd.oasis.opendocument.text'),
      'content.xml': new Uint8Array([1, 2, 3]),
      'META-INF/manifest.xml': new TextEncoder().encode('<manifest:manifest><manifest:file-entry manifest:full-path="content.xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>'),
    });
    expect(() => odtToDoc(enc)).toThrow(/паролем/);
    expect(() => odtToDoc(bookToOds(book))).toThrow(ConvertError);
    expect(() => odsToBook(docToOdt(sample).data)).toThrow(/текстовый документ/);
    expect(() => odtToDoc(new Uint8Array([0x50, 0x4b, 3, 4, 0, 0]))).toThrow(ConvertError);
  });
});

describe('ODF lengths', () => {
  it('converts units to points', () => {
    expect(odfLength('2.54cm')).toBeCloseTo(72, 1);
    expect(odfLength('1in')).toBe(72);
    expect(odfLength('10mm')).toBeCloseTo(28.35, 1);
    expect(odfLength('12pt')).toBe(12);
    expect(odfLength('abc')).toBeUndefined();
  });
});

