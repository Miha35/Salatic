/** Legacy binary Office formats (.doc, .xls) – produced by LibreOffice, read by our parsers. */
import { describe, expect, it } from 'vitest';
import { docToDocx } from '../../src/core/writers/docx';
import { bookToXlsx } from '../../src/core/writers/data';
import { docBinaryToDoc } from '../../src/core/readers/doc';
import { xlsToBook } from '../../src/core/readers/xls';
import { Cfb, isCfb } from '../../src/core/cfb';
import { detectFormat } from '../../src/core/formats';
import { convert } from '../../src/core/convert';
import { blockText, forEachRun, type Block, type TextRun } from '../../src/core/model';
import { env, hasCmd, soffice, text } from './helpers';
import { sample } from './sample';
import { ConvertError } from '../../src/core/util';

const runsOf = (blocks: Block[]): TextRun[] => {
  const out: TextRun[] = [];
  forEachRun(blocks, (r) => out.push(r));
  return out;
};

describe.skipIf(!hasCmd('soffice'))('Word 97–2003 (.doc)', () => {
  const doc97 = soffice(docToDocx(sample).data, 'docx', 'doc:MS Word 97');
  const { doc } = docBinaryToDoc(doc97);

  it('is an OLE2 compound file detected as .doc', () => {
    expect(isCfb(doc97)).toBe(true);
    expect(new Cfb(doc97).has('WordDocument')).toBe(true);
    expect(detectFormat('report.bin', doc97).format).toBe('doc');
  });

  it('keeps the structure: headings, lists, table, image, page break', () => {
    const types = doc.blocks.map((b) => b.type);
    expect(types).toContain('heading');
    expect(types).toContain('list');
    expect(types).toContain('table');
    expect(types).toContain('image');
    expect(types).toContain('pagebreak');
    const h = doc.blocks.filter((b) => b.type === 'heading');
    expect(h.map((b) => blockText(b))).toEqual(expect.arrayContaining(['Отчёт о работе', 'Список задач', 'Вторая страница']));
    expect(h[0].type === 'heading' && h[0].level).toBe(1);
    const table = doc.blocks.find((b) => b.type === 'table');
    expect(table && table.type === 'table' && table.rows.map((r) => r.cells.map((c) => c.blocks.map(blockText).join(' ')))).toEqual([
      ['Город', 'Население'],
      ['Москва', '13 100 000'],
      ['Казань', '1 300 000'],
    ]);
    const img = doc.blocks.find((b) => b.type === 'image');
    expect(img?.type === 'image' && img.mime).toMatch(/^image\/(png|jpeg)$/);
  });

  it('keeps inline formatting, links and nested lists', () => {
    const runs = runsOf(doc.blocks);
    expect(runs.some((r) => r.bold && r.text.includes('жирный'))).toBe(true);
    expect(runs.some((r) => r.italic && r.text.includes('курсив'))).toBe(true);
    expect(runs.some((r) => r.underline && r.text.includes('подчёркнутый'))).toBe(true);
    expect(runs.some((r) => r.sup && r.text === '2')).toBe(true);
    expect(runs.some((r) => r.link === 'https://example.com/' && r.text === 'ссылка')).toBe(true);
    const list = doc.blocks.find((b) => b.type === 'list');
    expect(list?.type === 'list' && list.ordered).toBe(false);
    const nested = list?.type === 'list' ? list.items.flatMap((i) => i.blocks).find((b) => b.type === 'list') : undefined;
    expect(nested?.type === 'list' && nested.ordered).toBe(true);
    expect(nested?.type === 'list' && nested.items.map((i) => i.blocks.map(blockText).join(''))).toEqual(['Вложенный один', 'Вложенный два']);
  });

  it('converts .doc → .docx/.pdf/.md end to end', async () => {
    const md = text((await convert(doc97, 'doc', 'md', env)).data);
    expect(md).toContain('# Отчёт о работе');
    expect(md).toContain('| Москва | 13 100 000 |');
    const docx = await convert(doc97, 'doc', 'docx', env);
    const back = text(soffice(docx.data, 'docx', 'txt:Text (encoded):UTF8'));
    expect(back).toContain('Съешь же ещё этих мягких французских булок');
    const pdf = await convert(doc97, 'doc', 'pdf', env);
    expect(text(pdf.data.subarray(0, 5))).toBe('%PDF-');
  });

  it('rejects garbage with a readable error', () => {
    const broken = doc97.slice(0, 1024);
    expect(() => docBinaryToDoc(broken)).toThrow(ConvertError);
    expect(() => docBinaryToDoc(new Uint8Array(600))).toThrow(ConvertError);
  });
});

describe.skipIf(!hasCmd('soffice'))('Excel 97–2003 (.xls)', () => {
  const long = 'Очень длинная строка — '.repeat(800); // forces SST CONTINUE records
  const book = {
    sheets: [
      { name: 'Продажи', header: true, rows: [['Товар', 'Цена', 'Флаг', 'Big'], ['Яблоки', 99.5, true, 123456789], ['Груши «Лесные»', -3, false, 0.000123], ['Long', long, null, 1e15]] },
      { name: 'Второй', rows: [['x', 'y'], [1, 2]] },
    ],
  };
  const xls = soffice(bookToXlsx(book), 'xlsx', 'xls:MS Excel 97');

  it('is detected as .xls', () => {
    expect(detectFormat('data.bin', xls).format).toBe('xls');
  });

  it('reads sheets, numbers, Cyrillic, booleans and long shared strings', () => {
    const b = xlsToBook(xls);
    expect(b.sheets.map((s) => s.name)).toEqual(['Продажи', 'Второй']);
    const rows = b.sheets[0].rows;
    expect(rows[0]).toEqual(['Товар', 'Цена', 'Флаг', 'Big']);
    expect(rows[1]).toEqual(['Яблоки', 99.5, true, 123456789]);
    expect(rows[2]).toEqual(['Груши «Лесные»', -3, false, 0.000123]);
    expect(rows[3][1]).toBe(long);
    expect(rows[3][3]).toBe(1e15);
    expect(b.sheets[1].rows).toEqual([['x', 'y'], [1, 2]]);
  });

  it('reads dates and formula results', () => {
    const csv = 'Дата,Сумма,Итог\n2024-03-15,10,=B2*2\n1999-12-31,2.5,=B3*2\n';
    const x = soffice(new TextEncoder().encode(csv), 'csv', 'xls:MS Excel 97');
    const rows = xlsToBook(x).sheets[0].rows;
    expect(rows[1][0]).toBe('2024-03-15');
    expect(rows[2][0]).toBe('1999-12-31');
    expect(rows[1][2]).toBe(20);
  });

  it('converts .xls → .xlsx/.csv end to end', async () => {
    const single = soffice(bookToXlsx({ sheets: [book.sheets[0]] }), 'xlsx', 'xls:MS Excel 97');
    const csv = text((await convert(single, 'xls', 'csv', env)).data);
    expect(csv).toContain('Яблоки,99.5,true,123456789');
    const xlsx = await convert(xls, 'xls', 'xlsx', env);
    const back = text(soffice(xlsx.data, 'xlsx', 'csv:Text - txt - csv (StarCalc):44,34,76'));
    expect(back).toContain('Груши «Лесные»');
  });

  it('rejects a truncated workbook with a readable error', () => {
    expect(() => xlsToBook(xls.slice(0, 700))).toThrow(ConvertError);
  });
});
