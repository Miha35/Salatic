/** Sections with unequal column widths and sections with their own page size (landscape pages). */
import { describe, expect, it } from 'vitest';
import { PDFDocument } from 'pdf-lib';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToOdt } from '../../src/core/writers/odf';
import { odtToDoc } from '../../src/core/readers/odf';
import { docToRtf } from '../../src/core/writers/rtf';
import { rtfToDoc } from '../../src/core/readers/rtf';
import { docToPdf } from '../../src/core/writers/pdf';
import type { Block, Doc, PageSetup } from '../../src/core/model';
import { hasCmd, loadFonts, pdfWords, soffice, type WordBox } from './helpers';

const para = (t: string): Block => ({ type: 'paragraph', runs: [{ text: t }] });
const WORDS = 'альфа бета гамма дельта эпсилон дзета эта тета йота каппа лямбда мю ню кси омикрон пи ро сигма тау ипсилон фи хи пси омега'.split(' ');
const text = (seed: number, n: number): string => Array.from({ length: n }, (_, i) => WORDS[(i * 7 + seed) % WORDS.length]).join(' ');

const M = 56.7;
const PORTRAIT: PageSetup = { width: 595.28, height: 841.89, margins: { top: M, right: M, bottom: M, left: M } };
const LANDSCAPE: PageSetup = { width: 841.89, height: 595.28, margins: { top: 42.5, right: 42.5, bottom: 42.5, left: 42.5 } };
const GAP = 20;
const NARROW: string[] = Array.from({ length: 6 }, (_, k) => `Узкая${k + 1}. ` + text(k, 60));

/** A heading, two columns of 1:2 width, a landscape page, then portrait again. */
const DOC: Doc = {
  page: PORTRAIT,
  blocks: [
    { type: 'heading', level: 1, runs: [{ text: 'Неравные колонки' }] },
    { type: 'section', columns: 2, gap: GAP, start: 'continuous', widths: [1, 2], gaps: [GAP] },
    ...NARROW.map(para),
    { type: 'section', columns: 1, gap: 36, start: 'page', page: LANDSCAPE },
    para('Альбомная страница. ' + text(2, 120)),
    { type: 'section', columns: 1, gap: 36, start: 'page', page: PORTRAIT },
    para('Снова книжная. ' + text(4, 20)),
  ],
};

// column geometry of page 1: content 481.88 pt, gap 20 → 153.96 and 307.92 pt
const contentW = PORTRAIT.width - 2 * M;
const W1 = (contentW - GAP) / 3;
const X2 = M + W1 + GAP;

type Sect = Extract<Block, { type: 'section' }>;
/** Sections: columns, width ratio, start and own page (rounded points). */
const shape = (d: Doc): string[] =>
  d.blocks
    .filter((b): b is Sect => b.type === 'section')
    .map((b) => `${b.columns}|${b.widths ? `1:${(b.widths[1] / b.widths[0]).toFixed(2)}|${Math.round(b.gaps?.[0] ?? -1)}` : '|'}|${b.start}|${b.page ? `${Math.round(b.page.width)}x${Math.round(b.page.height)}` : ''}`);
const EXPECTED = ['2|1:2.00|20|continuous|', '1|||page|842x595', '1|||page|595x842'];

const pageSizes = async (pdf: Uint8Array): Promise<string[]> =>
  (await PDFDocument.load(pdf)).getPages().map((p) => {
    const { width, height } = p.getSize();
    return `${Math.round(width)}x${Math.round(height)}`;
  });

/** Lines of a set of words: clustered by baseline, each sorted left to right. */
function linesOf(words: WordBox[]): WordBox[][] {
  const sorted = [...words].sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const out: WordBox[][] = [];
  for (const w of sorted) {
    const line = out[out.length - 1];
    if (line && Math.abs(line[0].y0 - w.y0) < 2) line.push(w);
    else out.push([w]);
  }
  return out.map((l) => l.sort((a, b) => a.x0 - b.x0));
}

describe('unequal columns and per-section pages: round trips', () => {
  it('DOCX: w:cols with w:equalWidth="0" and a w:pgSz per section', () => {
    const out = docToDocx(DOC).data;
    const back = docxToDoc(out).doc;
    expect(shape(back)).toEqual(EXPECTED);
    expect(Math.round(back.page!.width)).toBe(595);
    expect(back.blocks.filter((b) => b.type === 'paragraph').length).toBe(8);
  });

  it('ODT: style:column rel-widths and a master page per page size', async () => {
    const back = (await odtToDoc(docToOdt(DOC).data)).doc;
    expect(shape(back)).toEqual(EXPECTED);
    expect(Math.round(back.page!.width)).toBe(595);
  });

  it('RTF: \\colno/\\colw/\\colsr and \\pgwsxn/\\pghsxn per section', () => {
    const out = docToRtf(DOC).data;
    const rtf = new TextDecoder().decode(out);
    expect(rtf).toMatch(/\\cols2\\colsx400\\colno1\\colw\d+\\colsr400\\colno2\\colw\d+/);
    expect(rtf).toMatch(/\\sbkpage\\pgwsxn16838\\pghsxn11906[^\n]*\\lndscpsxn/);
    const back = rtfToDoc(out).doc;
    expect(shape(back)).toEqual(EXPECTED);
  });

  it('HTML (the document editor) keeps them in data attributes', async () => {
    const { docToHtml } = await import('../../src/core/writers/text');
    const { htmlToDoc } = await import('../../src/core/readers/html');
    const html = docToHtml(DOC);
    expect(html).toContain('data-widths="1,2"');
    expect(html).toContain('data-page="841.89,595.28,42.5,42.5,42.5,42.5"');
    expect(shape(htmlToDoc(html))).toEqual(EXPECTED);
  });

  it('a page size of a merged section marker survives normalisation', async () => {
    const { normalizeSections } = await import('../../src/core/model');
    const out = normalizeSections([para('a'), { type: 'section', columns: 1, gap: 36, start: 'page', page: LANDSCAPE }, { type: 'section', columns: 2, gap: 20, start: 'continuous' }, para('b')]);
    expect(out[1]).toMatchObject({ type: 'section', columns: 2, start: 'page', page: LANDSCAPE });
  });
});

describe.skipIf(!hasCmd('pdftotext'))('unequal columns and per-section pages in PDF', () => {
  it('our PDF: 1:2 columns, every word once in reading order, then a landscape page and a portrait one', async () => {
    const pdf = (await docToPdf(DOC, { fonts: loadFonts })).data;
    expect(await pageSizes(pdf)).toEqual(['595x842', '842x595', '595x842']);
    const words = pdfWords(pdf);
    const heading = words.find((w) => w.text === 'колонки')!;
    const cols = words.filter((w) => w.page === 1 && w.y0 > heading.y1 && /\p{L}/u.test(w.text));
    const left = cols.filter((w) => w.x0 < X2 - GAP / 2);
    const right = cols.filter((w) => w.x0 >= X2 - GAP / 2);
    // the narrow column stays within its width, the wide one starts after the gap and uses its width
    expect(Math.min(...left.map((w) => w.x0))).toBeCloseTo(M, 0);
    expect(Math.max(...left.map((w) => w.x1))).toBeLessThanOrEqual(M + W1 + 0.5);
    expect(Math.max(...left.map((w) => w.x1))).toBeGreaterThan(M + W1 - 40);
    expect(Math.min(...right.map((w) => w.x0))).toBeCloseTo(X2, 0);
    expect(Math.max(...right.map((w) => w.x1))).toBeLessThanOrEqual(PORTRAIT.width - M + 0.5);
    expect(Math.max(...right.map((w) => w.x1))).toBeGreaterThan(PORTRAIT.width - M - 40);
    // reading order (left column, then right): exactly the words of the paragraphs, none lost or repeated
    const read = [...linesOf(left), ...linesOf(right)].flat().map((w) => w.text);
    expect(read).toEqual(NARROW.join(' ').split(' '));
    // a paragraph continued in the wide column is laid out again at that width:
    // the only short lines there are the last lines of paragraphs
    const rightLines = linesOf(right);
    const short = rightLines.filter((l) => l[l.length - 1].x1 - l[0].x0 < W1 + 10).length;
    const lastLines = right.filter((w) => /^Узкая\d\.$/.test(w.text)).length + 1;
    expect(short).toBeLessThanOrEqual(lastLines);
    // not balanced (the section ends with a page break, as in Word): the narrow column is filled to the bottom first
    const bottom = (ws: WordBox[]): number => Math.max(...ws.map((w) => w.y1));
    expect(bottom(left)).toBeGreaterThan(PORTRAIT.height - M - 30);
    expect(bottom(right)).toBeLessThan(bottom(left));
    // the landscape page uses its whole width and its own margin
    const land = words.filter((w) => w.page === 2);
    expect(land[0].text).toBe('Альбомная');
    expect(Math.min(...land.map((w) => w.x0))).toBeCloseTo(42.5, 0);
    expect(Math.max(...land.map((w) => w.x1))).toBeGreaterThan(LANDSCAPE.width - 42.5 - 40);
    expect(words.find((w) => w.text === 'Снова')!.page).toBe(3);
  });

  it('page numbers sit in the middle of each page, whatever its width', async () => {
    const pdf = (await docToPdf(DOC, { fonts: loadFonts })).data;
    const words = pdfWords(pdf);
    for (const [page, w] of [
      [1, PORTRAIT.width],
      [2, LANDSCAPE.width],
      [3, PORTRAIT.width],
    ] as const) {
      const num = words.find((x) => x.page === page && x.text === String(page) && x.y0 > 500)!;
      expect(num, `page number ${page}`).toBeDefined();
      expect((num.x0 + num.x1) / 2).toBeCloseTo(w / 2, -1);
    }
  });

  it('a header follows the width of a landscape page', async () => {
    const doc: Doc = { ...DOC, header: [{ type: 'paragraph', runs: [{ text: 'Колонтитул справа' }], align: 'right' }] };
    const words = pdfWords((await docToPdf(doc, { fonts: loadFonts })).data);
    const right = (page: number): number => words.find((w) => w.page === page && w.text === 'справа')!.x1;
    expect(right(1)).toBeCloseTo(PORTRAIT.width - M, 0);
    expect(right(2)).toBeCloseTo(LANDSCAPE.width - 42.5, 0);
  });

  it('a fixed paper size wins over the sections’ pages', async () => {
    const pdf = (await docToPdf(DOC, { fonts: loadFonts }, { pageSize: 'a4' })).data;
    expect(new Set(await pageSizes(pdf))).toEqual(new Set(['595x842']));
  });

  it.skipIf(!hasCmd('soffice'))('LibreOffice lays out our DOCX, ODT and RTF with the same columns and pages', async () => {
    for (const [name, data, ext] of [
      ['docx', docToDocx(DOC).data, 'docx'],
      ['odt', docToOdt(DOC).data, 'odt'],
      ['rtf', docToRtf(DOC).data, 'rtf'],
    ] as const) {
      const pdf = soffice(data, ext, 'pdf');
      expect(await pageSizes(pdf), `${name}: page sizes`).toEqual(['595x842', '842x595', '595x842']);
      const words = pdfWords(pdf).filter((w) => w.page === 1 && /^Узкая\d\.$/.test(w.text));
      const starts = [...new Set(words.map((w) => Math.round(w.x0)))].sort((a, b) => a - b);
      // paragraph starts in the narrow column at the margin, in the wide one after the narrow column and gap
      expect(starts[0], `${name}: narrow column`).toBeCloseTo(M, -1);
      expect(starts[starts.length - 1], `${name}: wide column`).toBeCloseTo(X2, -1);
    }
  });

  it.skipIf(!hasCmd('soffice'))('reads unequal columns and page sizes written by LibreOffice (ODT and DOCX)', async () => {
    const odtFromLo = soffice(docToDocx(DOC).data, 'docx', 'odt');
    const docxFromLo = soffice(odtFromLo, 'odt', 'docx:MS Word 2007 XML');
    for (const [name, doc] of [
      ['odt', (await odtToDoc(odtFromLo)).doc],
      ['docx', docxToDoc(docxFromLo).doc],
    ] as const) {
      expect(shape(doc), name).toEqual(EXPECTED);
      // and our PDF of it has the same pages and columns
      const pdf = (await docToPdf(doc, { fonts: loadFonts })).data;
      expect(await pageSizes(pdf), `${name}: pages`).toEqual(['595x842', '842x595', '595x842']);
      const starts = pdfWords(pdf).filter((w) => w.page === 1 && /^Узкая\d\.$/.test(w.text)).map((w) => w.x0);
      expect(Math.max(...starts), `${name}: wide column`).toBeCloseTo(X2, -1);
    }
  });
});
