import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { PDFDocument } from 'pdf-lib';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToPdf } from '../../src/core/writers/pdf';
import { docToHtml, docToMarkdown, docToText } from '../../src/core/writers/text';
import { htmlToDoc } from '../../src/core/readers/html';
import { markdownToDoc } from '../../src/core/readers/markdown';
import { blockText, forEachRun, type Block, type Doc, type TextRun } from '../../src/core/model';
import { hasCmd, loadFonts, loadSubsetter, pdfToText, soffice, text, withTemp } from './helpers';
import { execFileSync } from 'node:child_process';
import { join } from 'node:path';
import { writeFileSync } from 'node:fs';

const note = (t: string): Block[] => [{ type: 'paragraph', runs: [{ text: t }] }];

/** A document exercising every typographic feature of the model. */
const rich: Doc = {
  title: 'Типографика',
  font: 'serif',
  size: 12,
  defaults: { spaceAfter: 6, lineHeight: 1.15 },
  page: { width: 595.28, height: 841.89, margins: { top: 56.7, right: 42.5, bottom: 56.7, left: 85 } },
  header: [{ type: 'paragraph', align: 'right', runs: [{ text: 'ООО «Ромашка» · Годовой отчёт' }] }],
  footer: [{ type: 'paragraph', align: 'center', runs: [{ text: 'Страница ' }, { text: '', field: 'page' }, { text: ' из ' }, { text: '', field: 'pages' }] }],
  blocks: [
    { type: 'heading', level: 1, runs: [{ text: 'Годовой отчёт' }] },
    {
      type: 'paragraph',
      firstLine: 35.4,
      align: 'justify',
      runs: [
        { text: 'Основной текст набран Times New Roman 12 пт с красной строкой' },
        { text: '', note: note('Первая сноска: пояснение к красной строке.') },
        { text: '. Фрагмент ' },
        { text: 'Calibri 14 пт', font: 'calibri', size: 14 },
        { text: ', фрагмент ' },
        { text: 'моноширинный', font: 'mono' },
        { text: ' и ' },
        { text: 'выделенный жёлтым', highlight: 'FFFF00' },
        { text: '.' },
      ],
    },
    { type: 'paragraph', indent: 36, spaceBefore: 12, runs: [{ text: 'Абзац с отступом слева 36 пт и интервалом перед ним 12 пт' }, { text: '', note: note('Вторая сноска.') }, { text: '.' }] },
    { type: 'pagebreak' },
    { type: 'paragraph', runs: [{ text: 'Вторая страница.' }] },
  ],
};

const runsOf = (d: Doc): TextRun[] => {
  const out: TextRun[] = [];
  forEachRun(d.blocks, (r) => out.push(r));
  return out;
};

describe('DOCX: fonts, sizes, paragraph format, page setup, header/footer, footnotes', () => {
  const data = docToDocx(rich).data;

  it('writes the parts Word needs', () => {
    const files = unzipSync(data);
    for (const p of ['word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml', 'word/_rels/footnotes.xml.rels']) expect(files[p], p).toBeDefined();
    const docXml = text(files['word/document.xml']);
    expect(docXml).toContain('<w:footnoteReference w:id="1"/>');
    expect(docXml).toContain('<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1700"');
    expect(text(files['word/footer1.xml'])).toContain('w:instr=" PAGE "');
    expect(text(files['word/styles.xml'])).toContain('w:ascii="Times New Roman"');
  });

  it('round-trips through our reader', () => {
    const { doc } = docxToDoc(data);
    expect(doc.font).toBe('serif');
    expect(doc.size).toBe(12);
    expect(doc.page?.margins.left).toBeCloseTo(85, 0);
    expect(blockText(doc.header![0])).toBe('ООО «Ромашка» · Годовой отчёт');
    expect(runsOf({ blocks: doc.footer! }).filter((r) => r.field).map((r) => r.field)).toEqual(['page', 'pages']);
    const runs = runsOf(doc);
    expect(runs.find((r) => r.text === 'Calibri 14 пт')).toMatchObject({ font: 'calibri', size: 14 });
    expect(runs.find((r) => r.text === 'моноширинный')?.font).toBe('mono');
    expect(runs.find((r) => r.text === 'выделенный жёлтым')?.highlight).toBe('FFFF00');
    const notes = runs.filter((r) => r.note).map((r) => blockText(r.note![0]));
    expect(notes).toEqual(['Первая сноска: пояснение к красной строке.', 'Вторая сноска.']);
    const p1 = doc.blocks.find((b) => b.type === 'paragraph' && blockText(b).startsWith('Основной')) as Extract<Block, { type: 'paragraph' }>;
    expect(p1.firstLine).toBeCloseTo(35.4, 0);
    expect(p1.align).toBe('justify');
    const p2 = doc.blocks.find((b) => b.type === 'paragraph' && blockText(b).startsWith('Абзац с отступом')) as Extract<Block, { type: 'paragraph' }>;
    expect(p2.indent).toBe(36);
    expect(p2.spaceBefore).toBe(12);
  });

  it.skipIf(!hasCmd('soffice'))('LibreOffice renders header, page fields and footnotes', () => {
    const pdf = soffice(data, 'docx', 'pdf');
    const t = pdfToText(pdf);
    expect(t).toContain('ООО «Ромашка» · Годовой отчёт');
    expect(t).toContain('Страница 1 из 2');
    expect(t).toContain('Страница 2 из 2');
    expect(t).toContain('Первая сноска: пояснение к красной строке.');
    const fonts = withTemp((dir) => {
      const f = join(dir, 'x.pdf');
      writeFileSync(f, pdf);
      return execFileSync('pdffonts', [f], { encoding: 'utf8' });
    });
    expect(fonts).toMatch(/LiberationSerif|TimesNewRoman/);
    expect(fonts).toMatch(/Carlito|Calibri/);
  });

  it.skipIf(!hasCmd('soffice'))('reads a DOCX re-saved by LibreOffice (theme fonts, style inheritance)', () => {
    const lo = soffice(data, 'docx', 'docx:MS Word 2007 XML');
    const { doc } = docxToDoc(lo);
    expect(doc.font).toBe('serif');
    expect(doc.size).toBe(12);
    expect(blockText(doc.header![0])).toContain('Ромашка');
    expect(runsOf({ blocks: doc.footer! }).some((r) => r.field === 'page')).toBe(true);
    const runs = runsOf(doc);
    expect(runs.filter((r) => r.note).length).toBe(2);
    expect(runs.find((r) => r.text.includes('Calibri 14'))).toMatchObject({ font: 'calibri', size: 14 });
  });
});

describe('PDF: header/footer fields, per-page footnotes, fonts', () => {
  it('places footnotes on the page of their reference and fills page fields', async () => {
    const { data, warnings } = await docToPdf(rich, { fonts: loadFonts, subsetter: await loadSubsetter() });
    expect(warnings).toEqual([]);
    const pdf = await PDFDocument.load(data);
    expect(pdf.getPageCount()).toBe(2);
    if (!hasCmd('pdftotext')) return;
    const pages = pdfToText(data).split('\f');
    expect(pages[0]).toContain('ООО «Ромашка» · Годовой отчёт');
    expect(pages[0]).toMatch(/Страница 1 из 2/);
    expect(pages[1]).toMatch(/Страница 2 из 2/);
    expect(pages[0]).toContain('1 Первая сноска: пояснение к красной строке.');
    expect(pages[0]).toContain('2 Вторая сноска.');
    expect(pages[1]).not.toContain('сноска');
    // reading order: body text precedes the notes on the page
    expect(pages[0].indexOf('Абзац с отступом')).toBeLessThan(pages[0].indexOf('Первая сноска'));
  });

  it('moves a footnote with its reference when the page fills up', async () => {
    const filler: Block[] = Array.from({ length: 34 }, (_, i) => ({ type: 'paragraph', runs: [{ text: `Строка-заполнитель номер ${i + 1}.` }] }) as Block);
    const doc: Doc = { blocks: [...filler, { type: 'paragraph', runs: [{ text: 'Ссылка у низа страницы' }, { text: '', note: note('Длинная сноска. '.repeat(30)) }] }, { type: 'paragraph', runs: [{ text: 'Хвост.' }] }] };
    const { data } = await docToPdf(doc, { fonts: loadFonts, subsetter: await loadSubsetter() });
    if (!hasCmd('pdftotext')) return;
    const pages = pdfToText(data).split('\f');
    const refPage = pages.findIndex((p) => p.includes('Ссылка у низа страницы'));
    expect(refPage).toBeGreaterThanOrEqual(0);
    expect(pages[refPage]).toContain('Длинная сноска.');
  });
});

describe('PDF → document keeps typography', () => {
  it('detects the body font, its size and differing fragments', async () => {
    const { pdfToDoc } = await import('../../src/core/readers/pdf');
    const { loadPdfjs } = await import('./helpers');
    const { data } = await docToPdf(rich, { fonts: loadFonts, subsetter: await loadSubsetter() });
    const { doc } = await pdfToDoc(data, { pdfjs: await loadPdfjs() });
    expect(doc.font).toBe('serif');
    expect(doc.size).toBe(12);
    const runs = runsOf(doc);
    expect(runs.find((r) => r.text.includes('Calibri 14'))).toMatchObject({ font: 'calibri', size: 14 });
    expect(runs.find((r) => r.text.includes('моноширинный'))?.code).toBe(true);
    // headers/footers repeated on every page are not body text
    expect(doc.blocks.some((b) => blockText(b).includes('Ромашка'))).toBe(false);
  });
});

describe('HTML and Markdown keep the new features', () => {
  it('HTML round-trip: styles, notes, header/footer, base font', () => {
    const html = docToHtml(rich);
    expect(html).toContain('class="footnotes"');
    expect(html).toContain('data-field="page"');
    const back = htmlToDoc(html);
    expect(back.font).toBe('serif');
    expect(back.size).toBe(12);
    expect(blockText(back.header![0])).toBe('ООО «Ромашка» · Годовой отчёт');
    expect(runsOf({ blocks: back.footer! }).map((r) => r.field).filter(Boolean)).toEqual(['page', 'pages']);
    const runs = runsOf(back);
    expect(runs.filter((r) => r.note).map((r) => blockText(r.note![0]))).toEqual(['Первая сноска: пояснение к красной строке.', 'Вторая сноска.']);
    expect(runs.find((r) => r.text === 'Calibri 14 пт')).toMatchObject({ font: 'calibri', size: 14 });
    expect(runs.find((r) => r.text === 'выделенный жёлтым')?.highlight).toBe('FFFF00');
    const p1 = back.blocks.find((b) => b.type === 'paragraph' && blockText(b).startsWith('Основной')) as Extract<Block, { type: 'paragraph' }>;
    expect(p1.firstLine).toBeCloseTo(35.4, 1);
    expect(back.blocks.some((b) => blockText(b).includes('сноска'))).toBe(false); // notes are not body text
  });

  it('reads pandoc/markdown-it style footnote HTML', () => {
    const html = '<p>Текст<sup id="fnref1"><a href="#fn1" class="footnote-ref">1</a></sup>.</p><section class="footnotes"><hr><ol><li id="fn1"><p>Примечание. <a href="#fnref1" class="footnote-back">↩︎</a></p></li></ol></section>';
    const d = htmlToDoc(html);
    expect(d.blocks).toHaveLength(1);
    const r = (d.blocks[0] as Extract<Block, { type: 'paragraph' }>).runs.find((x) => x.note);
    expect(blockText(r!.note![0])).toBe('Примечание.');
  });

  it('Markdown footnotes round-trip', () => {
    const md = docToMarkdown(rich);
    expect(md).toContain('[^1]');
    expect(md).toMatch(/\[\^2\]: Вторая сноска\./);
    const back = markdownToDoc(md);
    expect(runsOf(back).filter((r) => r.note).map((r) => blockText(r.note![0]))).toEqual(['Первая сноска: пояснение к красной строке.', 'Вторая сноска.']);
    const parsed = markdownToDoc('Текст[^a] и ещё[^b].\n\n[^a]: Первая.\n[^b]: Вторая,\n    продолжение.\n');
    expect(runsOf(parsed).filter((r) => r.note).map((r) => blockText(r.note![0]))).toEqual(['Первая.', 'Вторая,\nпродолжение.'.replace('\n', ' ')]);
  });

  it('plain text lists notes at the end', () => {
    const t = docToText(rich);
    expect(t).toContain('красной строкой[1]');
    expect(t).toMatch(/\[2\] Вторая сноска\.\n$/);
  });
});
