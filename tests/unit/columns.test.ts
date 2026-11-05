/** PDF → document: reading order of multi-column pages. */
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { PDFDocument, type PDFFont, type PDFPage } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { pdfToDoc } from '../../src/core/readers/pdf';
import { docToDocx } from '../../src/core/writers/docx';
import { blockText, type Block, type Doc } from '../../src/core/model';
import { hasCmd, loadFonts, loadPdfjs, soffice } from './helpers';

const read = async (pdf: Uint8Array): Promise<Block[]> => (await pdfToDoc(pdf, { pdfjs: await loadPdfjs() })).doc.blocks;
const texts = (blocks: Block[]): string[] => blocks.map(blockText);

function wrap(font: PDFFont, words: string[], size: number, width: number): string[] {
  const lines: string[] = [];
  let cur = '';
  for (const w of words) {
    const t = cur ? `${cur} ${w}` : w;
    if (font.widthOfTextAtSize(t, size) > width && cur) {
      lines.push(cur);
      cur = w;
    } else cur = t;
  }
  if (cur) lines.push(cur);
  return lines;
}

const sentence = (tag: string, n: number): string[] => Array.from({ length: n }, (_, i) => `${tag}${i + 1}`);

describe('multi-column PDF pages', () => {
  it('reads interleaved columns column by column, between full-width title, footer and a table', async () => {
    const pdf = await PDFDocument.create();
    pdf.registerFontkit(fontkit);
    const font = await pdf.embedFont(await loadFonts('serif-regular'));
    const bold = await pdf.embedFont(await loadFonts('serif-bold'));
    const page: PDFPage = pdf.addPage([595, 842]);
    page.drawText('Заголовок статьи на всю ширину', { x: 60, y: 780, size: 18, font: bold });
    const left = wrap(font, ['Левая', 'колонка', ...sentence('лево', 90), 'конец.'], 10.5, 235);
    const right = wrap(font, ['Правая', 'колонка', ...sentence('право', 90), 'финал.'], 10.5, 235);
    const rows = Math.max(left.length, right.length);
    // drawn row by row across both columns, as some generators and OCR layers do
    for (let i = 0; i < rows; i++) {
      if (left[i]) page.drawText(left[i], { x: 60, y: 740 - i * 13, size: 10.5, font });
      if (right[i]) page.drawText(right[i], { x: 315, y: 740 - i * 13, size: 10.5, font });
    }
    const after = 740 - rows * 13 - 20;
    page.drawText('Подвал на всю ширину страницы после колонок.', { x: 60, y: after, size: 10.5, font });
    const table = [['Город', 'Население', 'Площадь'], ['Москва', '13 100 000', '2 561'], ['Казань', '1 300 000', '614'], ['Самара', '1 160 000', '541']];
    table.forEach((r, i) => r.forEach((c, j) => page.drawText(c, { x: 60 + j * 160, y: after - 40 - i * 14, size: 10.5, font: i ? font : bold })));
    const blocks = await read(await pdf.save());
    const t = texts(blocks);
    expect(blocks.map((b) => b.type)).toEqual(['heading', 'paragraph', 'paragraph', 'paragraph', 'table']);
    expect(t[1]).toMatch(/^Левая колонка лево1 лево2 .* лево90 конец\.$/);
    expect(t[2]).toMatch(/^Правая колонка право1 .* право90 финал\.$/);
    expect(t[3]).toBe('Подвал на всю ширину страницы после колонок.');
    const tb = blocks[4];
    expect(tb.type === 'table' && tb.rows.map((r) => r.cells.length)).toEqual([3, 3, 3, 3]);
  });

  it('joins a paragraph that continues from one column into the next', async () => {
    const pdf = await PDFDocument.create();
    pdf.registerFontkit(fontkit);
    const font = await pdf.embedFont(await loadFonts('sans-regular'));
    const page = pdf.addPage([595, 842]);
    const all = wrap(font, ['Начало', ...sentence('слово', 160), 'конец.'], 10, 235);
    const half = Math.ceil(all.length / 2);
    all.slice(0, half).forEach((l, i) => page.drawText(l, { x: 60, y: 780 - i * 12.5, size: 10, font }));
    all.slice(half).forEach((l, i) => page.drawText(l, { x: 315, y: 780 - i * 12.5, size: 10, font }));
    const blocks = await read(await pdf.save());
    expect(blocks.length).toBe(1);
    expect(texts(blocks)[0]).toMatch(/^Начало слово1 слово2 .* слово160 конец\.$/);
  });

  it('keeps label/value tables and single-column text intact', async () => {
    const pdf = await PDFDocument.create();
    pdf.registerFontkit(fontkit);
    const font = await pdf.embedFont(await loadFonts('sans-regular'));
    const page = pdf.addPage([595, 842]);
    const rows = ['Имя', 'Фамилия', 'Город', 'Телефон', 'Почта', 'Должность', 'Отдел', 'Стаж'];
    rows.forEach((k, i) => {
      page.drawText(k, { x: 60, y: 780 - i * 16, size: 11, font });
      page.drawText(`Значение поля номер ${i + 1}`, { x: 300, y: 780 - i * 16, size: 11, font });
    });
    const body = wrap(font, sentence('текст', 120), 11, 470);
    body.forEach((l, i) => page.drawText(l, { x: 60, y: 600 - i * 14, size: 11, font }));
    const blocks = await read(await pdf.save());
    const table = blocks.find((b) => b.type === 'table');
    expect(table?.type === 'table' && table.rows.length).toBe(8);
    expect(texts(blocks).find((t) => t.startsWith('текст1 '))).toMatch(/текст120$/);
  });

  it.skipIf(!hasCmd('soffice'))('reads a two-column Word section exported by LibreOffice in order', async () => {
    const paras = Array.from({ length: 8 }, (_, i) => `Абзац ${i + 1}. ` + sentence(`а${i + 1}x`, 45).join(' ') + '.');
    const doc: Doc = { blocks: [{ type: 'heading', level: 1, runs: [{ text: 'Две колонки' }] }, ...paras.map((t) => ({ type: 'paragraph' as const, runs: [{ text: t }] }))] };
    const files = unzipSync(docToDocx(doc).data);
    const xml = strFromU8(files['word/document.xml']);
    // The heading spans the page (its own continuous section); the body is set in two columns
    files['word/document.xml'] = strToU8(
      xml
        .replace(/(<w:p><w:pPr><w:pStyle w:val="Heading1"\/>.*?<\/w:p>)/, '$1<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:pPr></w:p>')
        .replace(/<w:cols w:space="708"\/><\/w:sectPr><\/w:body>/, '<w:type w:val="continuous"/><w:cols w:num="2" w:space="567"/></w:sectPr></w:body>'),
    );
    const pdf = soffice(zipSync(files), 'docx', 'pdf');
    const t = texts(await read(pdf)).filter((x) => x.trim());
    expect(t[0]).toBe('Две колонки');
    const order = t.slice(1).map((x) => /^Абзац (\d+)/.exec(x)?.[1]).filter(Boolean);
    expect(order).toEqual(['1', '2', '3', '4', '5', '6', '7', '8']);
    // every paragraph is complete (no column text glued into another paragraph)
    for (const x of t.slice(1)) expect(x).toMatch(/^Абзац \d+\. а\d+x1 .* а\d+x45\.$/);
  });
});
