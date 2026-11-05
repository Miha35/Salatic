import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToPdf } from '../../src/core/writers/pdf';
import { pdfToDoc } from '../../src/core/readers/pdf';
import { docToHtml, docToMarkdown, docToText } from '../../src/core/writers/text';
import { markdownToDoc } from '../../src/core/readers/markdown';
import { htmlToDoc } from '../../src/core/readers/html';
import { encodePng } from '../../src/core/png';
import { blockText, runsText, type Block, type Doc } from '../../src/core/model';
import { hasCmd, loadFonts, loadPdfjs, pdfToText, soffice, text } from './helpers';

const png = encodePng(
  new Uint8Array(120 * 80 * 3).map((_, i) => (i % 3 === 0 ? 220 : i % 3 === 1 ? 80 : (i >> 4) & 255)),
  120,
  80,
  3,
);

const longPara = 'Съешь же ещё этих мягких французских булок, да выпей чаю. '.repeat(12).trim();

export const sample: Doc = {
  title: 'Тестовый документ',
  blocks: [
    { type: 'heading', level: 1, runs: [{ text: 'Отчёт о работе' }] },
    {
      type: 'paragraph',
      runs: [
        { text: 'Обычный текст, ' },
        { text: 'жирный', bold: true },
        { text: ', ' },
        { text: 'курсив', italic: true },
        { text: ', ' },
        { text: 'подчёркнутый', underline: true },
        { text: ' и ' },
        { text: 'ссылка', link: 'https://example.com/' },
        { text: '. E=mc' },
        { text: '2', sup: true },
      ],
    },
    { type: 'heading', level: 2, runs: [{ text: 'Список задач' }] },
    {
      type: 'list',
      ordered: false,
      items: [
        { blocks: [{ type: 'paragraph', runs: [{ text: 'Первый пункт' }] }] },
        {
          blocks: [
            { type: 'paragraph', runs: [{ text: 'Второй пункт' }] },
            { type: 'list', ordered: true, items: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Вложенный один' }] }] }, { blocks: [{ type: 'paragraph', runs: [{ text: 'Вложенный два' }] }] }] },
          ],
        },
      ],
    },
    { type: 'paragraph', runs: [{ text: longPara }], align: 'justify' },
    {
      type: 'table',
      rows: [
        { cells: [{ header: true, blocks: [{ type: 'paragraph', runs: [{ text: 'Город' }] }] }, { header: true, blocks: [{ type: 'paragraph', runs: [{ text: 'Население' }] }] }] },
        { cells: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Москва' }] }] }, { blocks: [{ type: 'paragraph', runs: [{ text: '13 100 000' }] }] }] },
        { cells: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Казань' }] }] }, { blocks: [{ type: 'paragraph', runs: [{ text: '1 300 000' }] }] }] },
      ],
    },
    { type: 'image', data: png, mime: 'image/png', width: 120, height: 80, alt: 'Диаграмма' },
    { type: 'code', text: 'function hello() {\n  return "Привет";\n}' },
    { type: 'quote', blocks: [{ type: 'paragraph', runs: [{ text: 'Цитата великого человека.' }] }] },
    { type: 'pagebreak' },
    { type: 'heading', level: 2, runs: [{ text: 'Вторая страница' }], align: 'center' },
    { type: 'paragraph', runs: [{ text: 'Конец документа.' }] },
  ],
};

const flatText = (d: Doc): string => d.blocks.map(blockText).join('\n');

describe('DOCX writer/reader', () => {
  const { data, warnings } = docToDocx(sample);

  it('produces a well-formed OOXML package', () => {
    expect(warnings).toEqual([]);
    const files = unzipSync(data);
    for (const p of ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/_rels/document.xml.rels', 'word/media/image1.png']) {
      expect(files[p], p).toBeDefined();
    }
    const xml = text(files['word/document.xml']);
    expect(xml).toContain('Отчёт о работе');
    expect(xml).toContain('w:hyperlink');
    expect(xml).toContain('<w:br w:type="page"/>');
  });

  it('round-trips through our own reader', () => {
    const { doc } = docxToDoc(data);
    expect(doc.title).toBe('Тестовый документ');
    const types = doc.blocks.map((b) => b.type);
    expect(types).toEqual(['heading', 'paragraph', 'heading', 'list', 'paragraph', 'table', 'image', 'paragraph', 'quote', 'pagebreak', 'heading', 'paragraph']);
    const p = doc.blocks[1] as Extract<Block, { type: 'paragraph' }>;
    expect(runsText(p.runs)).toBe('Обычный текст, жирный, курсив, подчёркнутый и ссылка. E=mc2');
    expect(p.runs.find((r) => r.text === 'жирный')?.bold).toBe(true);
    expect(p.runs.find((r) => r.text === 'ссылка')?.link).toBe('https://example.com/');
    expect(p.runs.find((r) => r.text === '2')?.sup).toBe(true);
    const list = doc.blocks[3] as Extract<Block, { type: 'list' }>;
    expect(list.ordered).toBe(false);
    expect(list.items).toHaveLength(2);
    const nested = list.items[1].blocks[1] as Extract<Block, { type: 'list' }>;
    expect(nested.ordered).toBe(true);
    expect(nested.items.map((i) => blockText(i.blocks[0]))).toEqual(['Вложенный один', 'Вложенный два']);
    const table = doc.blocks[5] as Extract<Block, { type: 'table' }>;
    expect(table.rows[1].cells.map((c) => blockText(c.blocks[0]))).toEqual(['Москва', '13 100 000']);
    const img = doc.blocks[6] as Extract<Block, { type: 'image' }>;
    expect([img.width, img.height]).toEqual([120, 80]);
    expect((doc.blocks[7] as Extract<Block, { type: "paragraph" }>).runs[0].code).toBe(true);
    expect(blockText(doc.blocks[7])).toContain("  return");
  });

  it.skipIf(!hasCmd('soffice'))('opens in LibreOffice and keeps the text', () => {
    const txt = text(soffice(data, 'docx', 'txt:Text (encoded):UTF8'));
    for (const s of ['Отчёт о работе', 'Первый пункт', 'Вложенный два', 'Москва', 'Цитата великого человека.', 'Конец документа.', 'return "Привет";']) {
      expect(txt).toContain(s);
    }
  });

  it.skipIf(!hasCmd('soffice'))('reads a DOCX produced by LibreOffice', () => {
    const html = '<h1>Заголовок LO</h1><p>Абзац <b>жирный</b> <i>курсив</i></p><ul><li>раз</li><li>два</li></ul><ol><li>first</li></ol><table border="1"><tr><td>a1</td><td>b1</td></tr></table>';
    const docx = soffice(new TextEncoder().encode(`<html><body>${html}</body></html>`), 'html', 'docx:MS Word 2007 XML');
    const { doc } = docxToDoc(docx);
    const types = doc.blocks.filter((b) => !(b.type === 'paragraph' && !b.runs.length)).map((b) => b.type);
    expect(types).toEqual(['heading', 'paragraph', 'list', 'list', 'table']);
    const p = doc.blocks.find((b) => b.type === 'paragraph' && b.runs.length) as Extract<Block, { type: 'paragraph' }>;
    expect(p.runs.find((r) => r.text.includes('жирный'))?.bold).toBe(true);
    expect((doc.blocks.find((b) => b.type === 'list') as Extract<Block, { type: 'list' }>).ordered).toBe(false);
    expect(doc.blocks.filter((b) => b.type === 'list').map((b) => (b as Extract<Block, { type: 'list' }>).ordered)).toEqual([false, true]);
  });
});

describe('PDF writer', () => {
  it('lays out Cyrillic text across pages with fonts, tables and images', async () => {
    const { data, warnings } = await docToPdf(sample, { fonts: loadFonts });
    expect(text(data.subarray(0, 5))).toBe('%PDF-');
    expect(warnings).toEqual([]);
    if (hasCmd('pdftotext')) {
      const t = pdfToText(data);
      for (const s of ['Отчёт о работе', 'жирный', 'Первый пункт', 'Вложенный два', 'Москва', '13 100 000', 'Цитата великого человека.', 'Конец документа.']) {
        expect(t).toContain(s);
      }
      // the explicit page break must produce a second page
      expect(t.split('\f').filter((p) => p.trim()).length).toBeGreaterThanOrEqual(2);
    }
  });

  it('flows long documents over many pages and replaces unsupported characters', async () => {
    const blocks: Block[] = [];
    for (let i = 0; i < 120; i++) blocks.push({ type: 'paragraph', runs: [{ text: `Абзац ${i + 1}. ${longPara}` }] });
    blocks.push({ type: 'paragraph', runs: [{ text: 'Иероглифы 漢字 и стрелка →' }] });
    const { data, warnings } = await docToPdf({ blocks }, { fonts: loadFonts });
    expect(warnings.some((w) => w.includes('символы'))).toBe(true);
    if (hasCmd('pdftotext')) {
      const t = pdfToText(data);
      expect(t).toContain('Абзац 120.');
      expect(t).toContain('стрелка ->');
      expect(t.split('\f').length).toBeGreaterThan(10);
    }
  });

  it('splits a table that is taller than one page', async () => {
    const rows = Array.from({ length: 90 }, (_, i) => ({
      cells: [{ blocks: [{ type: 'paragraph', runs: [{ text: `Строка ${i + 1}` }] } as Block] }, { blocks: [{ type: 'paragraph', runs: [{ text: String(i * i) }] } as Block] }],
    }));
    const { data } = await docToPdf({ blocks: [{ type: 'table', rows }] }, { fonts: loadFonts });
    if (hasCmd('pdftotext')) {
      const t = pdfToText(data);
      expect(t).toContain('Строка 1 ');
      expect(t).toContain('Строка 90');
    }
  });
});

describe('PDF reader', () => {
  it('reconstructs structure from a PDF we generated', async () => {
    const { data } = await docToPdf(sample, { fonts: loadFonts });
    const { doc, warnings } = await pdfToDoc(data, { pdfjs: await loadPdfjs() });
    expect(warnings).toEqual([]);
    const all = flatText(doc);
    for (const s of ['Отчёт о работе', 'Первый пункт', 'Москва', 'Конец документа.']) expect(all).toContain(s);
    const headings = doc.blocks.filter((b): b is Extract<Block, { type: 'heading' }> => b.type === 'heading');
    expect(headings.map((h) => runsText(h.runs))).toEqual(['Отчёт о работе', 'Список задач', 'Вторая страница']);
    expect(headings[0].level).toBe(1);
    expect(headings[1].level).toBe(2);
    const para = doc.blocks.find((b) => b.type === 'paragraph' && runsText(b.runs).startsWith('Обычный текст')) as Extract<Block, { type: 'paragraph' }>;
    expect(para).toBeDefined();
    expect(para.runs.find((r) => r.text.includes('жирный'))?.bold).toBe(true);
    expect(para.runs.find((r) => r.text.includes('курсив'))?.italic).toBe(true);
    // The justified paragraph spanning many lines is rebuilt as a single paragraph
    expect(doc.blocks.some((b) => b.type === 'paragraph' && runsText(b.runs) === longPara)).toBe(true);
    const list = doc.blocks.find((b) => b.type === 'list') as Extract<Block, { type: 'list' }>;
    expect(list).toBeDefined();
    expect(list.ordered).toBe(false);
    expect(blockText(list.items[0].blocks[0])).toBe('Первый пункт');
    const table = doc.blocks.find((b) => b.type === 'table') as Extract<Block, { type: 'table' }>;
    expect(table).toBeDefined();
    expect(table.rows.map((r) => r.cells.map((c) => blockText(c.blocks[0])))).toEqual([
      ['Город', 'Население'],
      ['Москва', '13 100 000'],
      ['Казань', '1 300 000'],
    ]);
    const img = doc.blocks.find((b) => b.type === 'image') as Extract<Block, { type: 'image' }>;
    expect(img).toBeDefined();
    expect(img.mime).toBe('image/png');
    expect(img.width).toBe(120);
    const code = doc.blocks.find((b) => b.type === 'code') as Extract<Block, { type: 'code' }>;
    expect(code?.text).toBe('function hello() {\n  return "Привет";\n}');
    const nested = list.items[1].blocks[1] as Extract<Block, { type: 'list' }>;
    expect(nested).toMatchObject({ type: 'list', ordered: true });
    // page numbers are not part of the text
    expect(doc.blocks.some((b) => /^\d+$/.test(blockText(b).trim()))).toBe(false);
  });

  it.skipIf(!hasCmd('soffice'))('reads a PDF exported by LibreOffice', async () => {
    const md = '# Квартальный отчёт\n\nПервый абзац текста, достаточно длинный, чтобы занять строку.\n\n## Раздел 2\n\n- пункт А\n- пункт Б\n\nИтог.';
    const docx = docToDocx(markdownToDoc(md)).data;
    const pdf = soffice(docx, 'docx', 'pdf');
    const { doc } = await pdfToDoc(pdf, { pdfjs: await loadPdfjs() });
    const headings = doc.blocks.filter((b) => b.type === 'heading').map((b) => blockText(b));
    expect(headings).toEqual(['Квартальный отчёт', 'Раздел 2']);
    const list = doc.blocks.find((b) => b.type === 'list') as Extract<Block, { type: 'list' }>;
    expect(list?.items.map((i) => blockText(i.blocks[0]))).toEqual(['пункт А', 'пункт Б']);
    expect(flatText(doc)).toContain('Итог.');
  });

  it('rejects garbage with a readable error', async () => {
    await expect(pdfToDoc(new TextEncoder().encode('%PDF-1.4 garbage'), { pdfjs: await loadPdfjs() })).rejects.toThrow(/PDF/);
  });
});

describe('text writers', () => {
  it('Markdown output parses back to the same structure', () => {
    const md = docToMarkdown(sample);
    const back = markdownToDoc(md);
    expect(back.blocks.map((b) => b.type)).toEqual(['heading', 'paragraph', 'heading', 'list', 'paragraph', 'table', 'image', 'code', 'quote', 'hr', 'heading', 'paragraph']);
    expect(flatText(back)).toContain('Вложенный два');
    const p = back.blocks[1] as Extract<Block, { type: 'paragraph' }>;
    expect(p.runs.find((r) => r.text === 'жирный')?.bold).toBe(true);
    expect(p.runs.find((r) => r.text === 'ссылка')?.link).toBe('https://example.com/');
  });

  it('HTML output parses back to the same structure', () => {
    const html = docToHtml(sample);
    expect(html).toMatch(/^<!DOCTYPE html>/);
    const back = htmlToDoc(html);
    expect(back.title).toBe('Тестовый документ');
    expect(back.blocks.map((b) => b.type)).toEqual(sample.blocks.map((b) => b.type));
  });

  it('plain text keeps reading order and list markers', () => {
    const t = docToText(sample);
    expect(t).toContain('• Первый пункт');
    expect(t).toContain('1. Вложенный один');
    expect(t).toContain('Москва\t13 100 000');
    expect(t.indexOf('Отчёт о работе')).toBeLessThan(t.indexOf('Конец документа.'));
  });
});
