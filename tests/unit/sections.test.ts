/** Sections with columns: model round trips through every editable format, PDF layout, LibreOffice interop. */
import { describe, expect, it } from 'vitest';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToOdt } from '../../src/core/writers/odf';
import { odtToDoc } from '../../src/core/readers/odf';
import { docToRtf } from '../../src/core/writers/rtf';
import { rtfToDoc } from '../../src/core/readers/rtf';
import { docToHtml } from '../../src/core/writers/text';
import { htmlToDoc } from '../../src/core/readers/html';
import { docToPdf } from '../../src/core/writers/pdf';
import { blockText, type Block, type Doc } from '../../src/core/model';
import { hasCmd, loadFonts, pdfWords, soffice, type WordBox } from './helpers';

const para = (t: string): Block => ({ type: 'paragraph', runs: [{ text: t }] });
const WORDS = 'альфа бета гамма дельта эпсилон дзета эта тета йота каппа лямбда мю ню кси омикрон пи ро сигма тау ипсилон фи хи пси омега'.split(' ');
const text = (seed: number, n: number): string => Array.from({ length: n }, (_, i) => WORDS[(i * 7 + seed) % WORDS.length]).join(' ');

/** A heading and intro on the full width, two balanced columns, one column again, then three columns on a new page. */
const DOC: Doc = {
  blocks: [
    { type: 'heading', level: 1, runs: [{ text: 'Вёрстка в колонки' }] },
    para('Вступление на всю ширину страницы. ' + text(1, 40)),
    { type: 'section', columns: 2, gap: 24, start: 'continuous' },
    ...Array.from({ length: 5 }, (_, k) => para(`Абзац${k + 1}. ` + text(k, 45))),
    { type: 'section', columns: 1, gap: 36, start: 'continuous' },
    para('Снова одна колонка. ' + text(3, 30)),
    { type: 'section', columns: 3, gap: 18, separator: true, start: 'page' },
    ...Array.from({ length: 9 }, (_, k) => para(`Колонка${k + 1}. ` + text(k + 5, 50))),
  ],
};

type Sect = Extract<Block, { type: 'section' }>;
/** Section markers with the text that follows each (first 12 characters of the first paragraph). */
const shape = (d: Doc): string[] =>
  d.blocks.flatMap((b, i) => {
    if (b.type !== 'section') return [];
    const next = d.blocks.slice(i + 1).find((x) => x.type !== 'section');
    // the gap of a single column means nothing
    return [`${b.columns}|${b.columns > 1 ? Math.round(b.gap) : '-'}|${b.separator ? 'sep' : ''}|${b.start}|${next ? blockText(next).slice(0, 8) : ''}`];
  });
const EXPECTED = ['2|24||continuous|Абзац1. ', '1|-||continuous|Снова од', '3|18|sep|page|Колонка1'];

describe('sections round trip', () => {
  it('DOCX', () => {
    const back = docxToDoc(docToDocx(DOC).data).doc;
    expect(shape(back)).toEqual(EXPECTED);
    expect(back.blocks.filter((b) => b.type === 'paragraph').length).toBe(DOC.blocks.filter((b) => b.type === 'paragraph').length);
  });

  it('ODT (column sections are text:section elements)', async () => {
    const back = (await odtToDoc(docToOdt(DOC).data)).doc;
    // the page start travels as a break before the section's first paragraph
    expect(shape(back)).toEqual(EXPECTED);
  });

  it('HTML (CSS columns, used by the document editor)', () => {
    const html = docToHtml(DOC);
    expect(html).toContain('column-count:2');
    expect(html).toContain('column-rule:0.5pt solid');
    expect(shape(htmlToDoc(html))).toEqual(EXPECTED);
  });

  it('HTML from elsewhere: CSS column-count becomes a section', () => {
    const d = htmlToDoc('<p>До</p><div style="column-count:3;column-gap:12pt"><p>Один</p><p>Два</p></div><p>После</p>');
    expect(shape(d)).toEqual(['3|12||continuous|Один', '1|-||continuous|После']);
  });

  it('RTF', () => {
    const rtf = new TextDecoder().decode(docToRtf(DOC).data);
    expect(rtf).toMatch(/\\sect\\sectd\\sbknone[^\n]*\\cols2\\colsx480/);
    expect(rtf).toMatch(/\\sect\\sectd\\sbkpage[^\n]*\\cols3\\colsx360\\linebetcol/);
    expect(shape(rtfToDoc(docToRtf(DOC).data).doc)).toEqual(EXPECTED);
  });

  it('formats without sections (Markdown/TXT) simply ignore the markers', async () => {
    const { docToMarkdown, docToText } = await import('../../src/core/writers/text');
    expect(docToMarkdown(DOC)).toContain('Абзац1.');
    expect(docToText(DOC)).toContain('Колонка9.');
  });
});

/** Columns of a page: the left edges of its lines (words with no word just before them on the same line), clustered. */
function columnsOf(words: WordBox[]): number[][] {
  const starts = words.filter((w) => !words.some((o) => o !== w && o.page === w.page && Math.abs(o.y0 - w.y0) < 1 && o.x1 <= w.x0 + 0.5 && o.x1 > w.x0 - 14));
  const lefts = [...new Set(starts.map((w) => Math.round(w.x0)))].sort((a, b) => a - b);
  const groups: number[][] = [];
  for (const x of lefts) {
    const g = groups[groups.length - 1];
    if (g && x - g[g.length - 1] < 40) g.push(x);
    else groups.push([x]);
  }
  return groups;
}

describe.skipIf(!hasCmd('pdftotext'))('sections in PDF', () => {
  const lines = (pdf: Uint8Array) => pdfWords(pdf);

  it('our PDF: two balanced columns, full width again, three columns with separators on page 2', async () => {
    const pdf = (await docToPdf(DOC, { fonts: loadFonts })).data;
    const words = lines(pdf);
    const first = (t: string): WordBox => words.find((w) => w.text.startsWith(t))!;
    // page 1: Абзац1-3 on the left, Абзац4-5 on the right (balanced), then the full-width paragraph below both
    const mid = 595.28 / 2;
    for (const k of [1, 2, 3]) expect(first(`Абзац${k}.`).x0, `Абзац${k}`).toBeLessThan(mid);
    for (const k of [4, 5]) expect(first(`Абзац${k}.`).x0, `Абзац${k}`).toBeGreaterThan(mid);
    const colWords = words.filter((w) => w.page === 1 && w.y0 > first('Абзац1.').y0 - 1 && w.y1 < first('Снова').y0);
    const leftBottom = Math.max(...colWords.filter((w) => w.x0 < mid).map((w) => w.y1));
    const rightBottom = Math.max(...colWords.filter((w) => w.x0 > mid).map((w) => w.y1));
    expect(Math.abs(leftBottom - rightBottom), 'balanced columns').toBeLessThan(40);
    expect(first('Снова').y0).toBeGreaterThan(Math.max(leftBottom, rightBottom));
    expect(first('Снова').x0).toBeLessThan(80);
    // page 2: three columns
    expect(first('Колонка1.').page).toBe(2);
    expect(columnsOf(words.filter((w) => w.page === 2 && /\p{L}/u.test(w.text))).length).toBe(3); // body text, not the page number
  });

  it('footnotes referenced in columns get their room; headings stay with their paragraph', async () => {
    const blocks: Block[] = [{ type: 'section', columns: 2, gap: 20, start: 'continuous' }];
    for (let k = 0; k < 14; k++) {
      blocks.push({ type: 'paragraph', runs: [{ text: `Абзац${k + 1} ` + text(k, 40) }, { text: '', note: [para(`Сноска${k + 1} ` + text(k + 3, 25))] }, { text: ' конец.' }] });
      if (k % 4 === 3) blocks.push({ type: 'heading', level: 2, runs: [{ text: `Заголовок${k}` }] });
    }
    const words = lines((await docToPdf({ blocks }, { fonts: loadFonts })).data).filter((w) => /\p{L}/u.test(w.text));
    for (const page of new Set(words.map((w) => w.page))) {
      const pw = words.filter((w) => w.page === page);
      // footnote text is smaller than body text
      const body = pw.filter((w) => w.y1 - w.y0 > 11);
      const notes = pw.filter((w) => w.y1 - w.y0 < 10);
      if (notes.length) expect(Math.max(...body.map((w) => w.y1)), `page ${page}: body runs into the footnotes`).toBeLessThan(Math.min(...notes.map((w) => w.y0)));
      // every note sits on the page of its reference
      for (const n of notes.filter((w) => /^Сноска\d+$/.test(w.text))) {
        const k = n.text.slice(6);
        expect(pw.some((w) => w.text === `Абзац${k}`) || words.find((w) => w.text === `Абзац${k}`)!.page < page, `note ${k} on page ${page}`).toBe(true);
      }
      for (const h of pw.filter((w) => /^Заголовок/.test(w.text))) {
        const after = pw.filter((w) => Math.abs(w.x0 - h.x0) < 40 && w.y0 > h.y1 && w.y0 < h.y1 + 40);
        expect(after.length, `${h.text} is followed by text in its column`).toBeGreaterThan(0);
      }
    }
  });

  it.skipIf(!hasCmd('soffice'))('LibreOffice lays out our DOCX, ODT and RTF in the same columns', async () => {
    for (const [name, data, ext] of [
      ['docx', docToDocx(DOC).data, 'docx'],
      ['odt', docToOdt(DOC).data, 'odt'],
      ['rtf', docToRtf(DOC).data, 'rtf'],
    ] as const) {
      const words = lines(soffice(data, ext, 'pdf'));
      const p1 = words.filter((w) => w.page === 1 && /^Абзац\d\.$/.test(w.text));
      expect(columnsOf(p1).length, `${name}: page 1 columns`).toBe(2);
      const kol = words.filter((w) => /^Колонка\d\.$/.test(w.text));
      expect(new Set(kol.map((w) => w.page)).size >= 1 && kol[0].page > 1, `${name}: 3-column section starts a page`).toBe(true);
      expect(columnsOf(words.filter((w) => w.page === kol[0].page)).length, `${name}: 3 columns`).toBe(3);
    }
  });

  it.skipIf(!hasCmd('soffice'))('reads sections written by LibreOffice (DOCX and ODT)', async () => {
    const odtFromLo = soffice(docToDocx(DOC).data, 'docx', 'odt');
    const docxFromLo = soffice(odtFromLo, 'odt', 'docx:MS Word 2007 XML');
    const cols = (d: Doc): number[] => d.blocks.filter((b): b is Sect => b.type === 'section').map((b) => b.columns);
    expect(cols((await odtToDoc(odtFromLo)).doc)).toEqual(expect.arrayContaining([2, 3]));
    const fromDocx = docxToDoc(docxFromLo).doc;
    expect(cols(fromDocx)).toEqual(expect.arrayContaining([2, 3]));
    // and our PDF of it keeps the columns
    const words = lines((await docToPdf(fromDocx, { fonts: loadFonts })).data);
    expect(columnsOf(words.filter((w) => w.page === 1 && /^Абзац\d\.$/.test(w.text))).length).toBe(2);
  });
});
