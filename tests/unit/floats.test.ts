/**
 * Floating pictures (Word wp:anchor, ODF frames, CSS floats): text flowing around them in PDF, pictures
 * behind and in front of the text, round trips through DOCX/ODT/HTML, and LibreOffice reading ours.
 */
import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { PDFArray, PDFDocument, PDFName, PDFRawStream, decodePDFRawStream, type PDFRef } from 'pdf-lib';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToOdt } from '../../src/core/writers/odf';
import { odtToDoc } from '../../src/core/readers/odf';
import { docToHtml } from '../../src/core/writers/text';
import { htmlToDoc } from '../../src/core/readers/html';
import { docToPdf } from '../../src/core/writers/pdf';
import { encodePng } from '../../src/core/png';
import type { Block, Doc, ImageFloat } from '../../src/core/model';
import { hasCmd, loadFonts, pdfWords, renderPages, soffice, type WordBox } from './helpers';

const WORDS = 'альфа бета гамма дельта эпсилон дзета эта тета йота каппа лямбда мю ню кси омикрон пи ро сигма тау ипсилон фи хи пси омега'.split(' ');
const text = (seed: number, n: number): string => Array.from({ length: n }, (_, i) => WORDS[(i * 7 + seed) % WORDS.length]).join(' ');
const para = (t: string): Block => ({ type: 'paragraph', runs: [{ text: t }] });
/** A solid gray picture: 120 × 120 px = 90 × 90 pt; gray 128 is easy to find in a render. */
const gray = (v: number, w = 120, h = 120): Uint8Array => encodePng(new Uint8Array(w * h).fill(v), w, h, 1);
const img = (float: ImageFloat, v = 128, w = 120, h = 120): Block => ({ type: 'image', data: gray(v, w, h), mime: 'image/png', width: w, height: h, alt: 'Рисунок', float });

const M = 56.7;
const PAGE = { width: 595.28, height: 841.89, margins: { top: M, right: M, bottom: M, left: M } };
const LEFT: ImageFloat = { wrap: 'square', h: { from: 'column', align: 'left' }, v: { from: 'paragraph', offset: 0 }, dist: [0, 9, 0, 0] };
const RIGHT: ImageFloat = { wrap: 'square', h: { from: 'column', align: 'right' }, v: { from: 'paragraph', offset: 0 }, dist: [0, 0, 0, 9] };
const P1 = 'Первый ' + text(1, 150);
const P2 = 'Второй ' + text(5, 150);

// the right picture in another gray, so the two are told apart in renders
const DOC: Doc = { page: PAGE, blocks: [img(LEFT), para(P1), img(RIGHT, 100), para(P2)] };

/** Rows and columns (72 dpi, page pixels from the top-left) where a render has ≥ 60 pixels of exactly `v`. */
function blob(pgm: Buffer, v: number): { x0: number; x1: number; y0: number; y1: number } | null {
  const m = /^P5\s+(\d+)\s+(\d+)\s+255\s/.exec(new TextDecoder('latin1').decode(pgm.subarray(0, 40)))!;
  const W = +m[1];
  const H = +m[2];
  const at = (x: number, y: number): number => pgm[m[0].length + y * W + x];
  const rows: number[] = [];
  const cols: number[] = [];
  for (let y = 0; y < H; y++) {
    let n = 0;
    for (let x = 0; x < W; x++) if (at(x, y) === v) n++;
    if (n >= 60) rows.push(y);
  }
  for (let x = 0; x < W; x++) {
    let n = 0;
    for (let y = 0; y < H; y++) if (at(x, y) === v) n++;
    if (n >= 60) cols.push(x);
  }
  return rows.length && cols.length ? { x0: cols[0], x1: cols[cols.length - 1] + 1, y0: rows[0], y1: rows[rows.length - 1] + 1 } : null;
}

/** Words overlapping a box (page coordinates from the top). */
const overlapping = (words: WordBox[], b: { x0: number; x1: number; y0: number; y1: number }): WordBox[] => words.filter((w) => w.x1 > b.x0 + 0.5 && w.x0 < b.x1 - 0.5 && w.y1 > b.y0 + 0.5 && w.y0 < b.y1 - 0.5);

/** The content streams of a page, decoded, in order. */
async function streams(pdf: Uint8Array, page = 0): Promise<string[]> {
  const doc = await PDFDocument.load(pdf);
  const c = doc.getPage(page).node.Contents();
  const refs = c instanceof PDFArray ? c.asArray() : [c];
  return refs.map((r) => {
    const s = doc.context.lookup(r as PDFRef) as PDFRawStream;
    return new TextDecoder('latin1').decode(decodePDFRawStream(s).decode());
  });
}

describe.skipIf(!hasCmd('pdftotext') || !hasCmd('pdftoppm'))('floating pictures in PDF', () => {
  it('square wrap: the text flows beside a left and a right picture and never over them', async () => {
    const pdf = (await docToPdf(DOC, { fonts: loadFonts })).data;
    const words = pdfWords(pdf).filter((w) => w.page === 1);
    const pgm = renderPages(pdf, 72)[0];
    const pic = blob(pgm, 128)!;
    expect(pic).not.toBeNull();
    // the left picture: at the margin, at the top of its paragraph
    expect(pic.x0).toBeCloseTo(M, -1);
    expect(pic.x1 - pic.x0).toBeCloseTo(90, -1);
    expect(overlapping(words, pic)).toEqual([]);
    // lines beside it start right of it (with the 9 pt distance), lines below at the margin again
    const beside = words.filter((w) => w.y0 > pic.y0 && w.y1 < pic.y1);
    const besideStarts = beside.filter((w) => !beside.some((o) => Math.abs(o.y0 - w.y0) < 1 && o.x0 < w.x0));
    expect(besideStarts.length).toBeGreaterThan(3);
    for (const w of besideStarts) expect(w.x0).toBeGreaterThanOrEqual(M + 90 + 9 - 1);
    const p1Words = P1.split(' ');
    const below = words.filter((w) => w.y0 > pic.y1 + 2 && w.y1 < words.find((x) => x.text === 'Второй')!.y0);
    expect(Math.min(...below.map((w) => w.x0))).toBeCloseTo(M, 0);
    // every word of the paragraph once
    const firstP2 = words.findIndex((w) => w.text === 'Второй');
    expect(words.slice(0, firstP2).map((w) => w.text)).toEqual(p1Words);
  });

  it('the right picture keeps the text on its left', async () => {
    // only the right float, so the render has one gray block
    const pdf = (await docToPdf({ page: PAGE, blocks: [para('Начало ' + text(2, 20)), img(RIGHT), para(P2)] }, { fonts: loadFonts })).data;
    const words = pdfWords(pdf).filter((w) => w.page === 1);
    const pic = blob(renderPages(pdf, 72)[0], 128)!;
    expect(pic.x1).toBeCloseTo(PAGE.width - M, -1);
    expect(overlapping(words, pic)).toEqual([]);
    const beside = words.filter((w) => w.y0 > pic.y0 && w.y1 < pic.y1);
    expect(beside.length).toBeGreaterThan(10);
    expect(Math.max(...beside.map((w) => w.x1))).toBeLessThanOrEqual(PAGE.width - M - 90 - 9 + 1);
    expect(Math.min(...beside.map((w) => w.x0))).toBeCloseTo(M, 0);
  });

  it('a picture behind the text is painted first, one in front last; neither moves the text', async () => {
    const behind: ImageFloat = { wrap: 'none', behind: true, h: { from: 'margin', offset: 20 }, v: { from: 'paragraph', offset: 0 } };
    const front: ImageFloat = { wrap: 'none', h: { from: 'margin', offset: 300 }, v: { from: 'paragraph', offset: 0 } };
    const plain = (await docToPdf({ page: PAGE, blocks: [para(P1)] }, { fonts: loadFonts })).data;
    const pdf = (await docToPdf({ page: PAGE, blocks: [img(behind, 200), img(front, 60), para(P1)] }, { fonts: loadFonts })).data;
    // the text is where it is without the pictures
    const a = pdfWords(plain);
    const b = pdfWords(pdf);
    expect(b.map((w) => `${w.text}@${Math.round(w.x0)},${Math.round(w.y0)}`)).toEqual(a.map((w) => `${w.text}@${Math.round(w.x0)},${Math.round(w.y0)}`));
    const s = await streams(pdf);
    const doAt = (i: number): number => s.findIndex((x) => new RegExp(`/Image-?\\d+-${i} Do|Do`).test(x));
    const firstText = s.findIndex((x) => /Tj|TJ/.test(x));
    expect(doAt(0)).toBe(0); // the stream with the behind picture comes first
    expect(firstText).toBeGreaterThan(0);
    expect(/Do/.test(s[s.length - 1]) && !/Tj|TJ/.test(s[s.length - 1].split('Do').pop()!)).toBe(true); // in front: drawn after the text
    // the render shows text over the light (behind) picture: dark pixels inside its area
    const pgm = renderPages(pdf, 72)[0];
    const pic = blob(pgm, 200)!;
    expect(pic).not.toBeNull();
    const front60 = blob(pgm, 60)!;
    expect(front60.x0).toBeCloseTo(M + 300, -1);
  });

  it('top-and-bottom: nothing beside the picture, at its alignment', async () => {
    const tb: ImageFloat = { wrap: 'topBottom', h: { from: 'column', align: 'right' }, v: { from: 'paragraph', offset: 0 } };
    const pdf = (await docToPdf({ page: PAGE, blocks: [para('До ' + text(3, 30)), img(tb), para('После ' + text(4, 30))] }, { fonts: loadFonts })).data;
    const words = pdfWords(pdf);
    const pic = blob(renderPages(pdf, 72)[0], 128)!;
    expect(pic.x1).toBeCloseTo(PAGE.width - M, -1);
    expect(words.filter((w) => w.y1 > pic.y0 && w.y0 < pic.y1)).toEqual([]);
    expect(words.find((w) => w.text === 'После')!.y0).toBeGreaterThan(pic.y1);
  });

  it('a float near the page end moves to the next page with its paragraph; columns place floats in the column', async () => {
    const filler = Array.from({ length: 9 }, (_, k) => para(`Абзац${k} ` + text(k, 60)));
    const pdf = (await docToPdf({ page: PAGE, blocks: [...filler, img(LEFT), para(P1)] }, { fonts: loadFonts })).data;
    const pages = renderPages(pdf, 72);
    const words = pdfWords(pdf);
    const pageOfPic = pages.findIndex((p) => blob(p, 128));
    expect(pageOfPic).toBeGreaterThanOrEqual(0);
    expect(words.find((w) => w.text === 'Первый')!.page).toBe(pageOfPic + 1);
    for (const p of [1, 2]) expect(overlapping(words.filter((w) => w.page === p), blob(pages[p - 1], 128) ?? { x0: 0, x1: 0, y0: 0, y1: 0 })).toEqual([]);
    const cols = (await docToPdf({ page: PAGE, blocks: [{ type: 'section', columns: 2, gap: 20, start: 'continuous' }, img(LEFT), para(P1)] }, { fonts: loadFonts })).data;
    expect(blob(renderPages(cols, 72)[0], 128)).not.toBeNull();
  });

  it('a list item next to a float keeps its marker with its first line', async () => {
    const pdf = (await docToPdf({ page: PAGE, blocks: [img(LEFT), { type: 'list', ordered: true, items: [{ blocks: [para('Пункт ' + text(1, 30))] }, { blocks: [para('Пункт ' + text(2, 30))] }] }] }, { fonts: loadFonts })).data;
    const words = pdfWords(pdf);
    const pic = blob(renderPages(pdf, 72)[0], 128)!;
    expect(overlapping(words, pic)).toEqual([]);
    const marker = words.find((w) => w.text === '1.')!;
    const first = words.find((w) => w.text === 'Пункт')!;
    expect(marker.x0).toBeGreaterThan(pic.x1);
    expect(first.x0).toBeGreaterThan(marker.x1);
    expect(Math.abs(marker.y0 - first.y0)).toBeLessThan(1);
  });

  it('a header picture behind the text appears on every page (letterheads)', async () => {
    const logo: ImageFloat = { wrap: 'none', behind: true, h: { from: 'page', align: 'right' }, v: { from: 'page', offset: 10 } };
    const doc: Doc = { page: PAGE, header: [img(logo, 128, 80, 80), para('Шапка')], blocks: Array.from({ length: 12 }, (_, k) => para(`Абзац${k} ` + text(k, 60))) };
    const pdf = (await docToPdf(doc, { fonts: loadFonts })).data;
    const pages = renderPages(pdf, 72);
    expect(pages.length).toBeGreaterThan(1);
    for (const p of pages) {
      const b = blob(p, 128)!;
      expect(b).not.toBeNull();
      expect(b.x1).toBeCloseTo(PAGE.width, -1);
      expect(b.y0).toBeCloseTo(10, -1);
    }
  });
});

describe.skipIf(!hasCmd('pdftotext'))('paragraphs across pages (widows and orphans)', () => {
  /** Lines of a page: distinct word tops. */
  const lineTops = (ws: WordBox[]): number[] => [...new Set(ws.map((w) => Math.round(w.y0)))].sort((a, b) => a - b);

  it('a long paragraph at the page end is split, two lines at least on each page, no gap left', async () => {
    const filler = Array.from({ length: 8 }, (_, k) => para(`Абзац${k} ` + text(k, 64)));
    const long = 'Длинный ' + text(9, 260);
    const pdf = (await docToPdf({ page: PAGE, blocks: [...filler, para(long)] }, { fonts: loadFonts })).data;
    const words = pdfWords(pdf).filter((w) => /\p{L}/u.test(w.text));
    const start = words.find((w) => w.text === 'Длинный')!;
    expect(start.page).toBe(1);
    const p1 = words.filter((w) => w.page === 1 && w.y0 >= start.y0 - 1);
    const p2 = words.filter((w) => w.page === 2);
    expect(lineTops(p1).length).toBeGreaterThanOrEqual(2);
    expect(lineTops(p2).length).toBeGreaterThanOrEqual(2);
    // page 1 is filled down to its bottom margin (less than two lines left)
    expect(Math.max(...words.filter((w) => w.page === 1).map((w) => w.y1))).toBeGreaterThan(PAGE.height - M - 30);
    // and the words come in order, each once
    expect([...p1, ...p2].map((w) => w.text)).toEqual(long.split(' '));
  });

  it('one line of room: a four-line paragraph moves whole; three lines of room: a five-line one is split 3 + 2', async () => {
    // a small page (160 pt of text, about ten lines) and paragraphs of exact lines (line breaks)
    const small = { width: 300, height: 200, margins: { top: 20, right: 20, bottom: 20, left: 20 } };
    const lines = (tag: string, n: number): Block => ({ type: 'paragraph', runs: [{ text: Array.from({ length: n }, (_, k) => `${tag}${k + 1}`).join('\n') }] });
    const place = async (before: number, n: number): Promise<WordBox[]> => pdfWords((await docToPdf({ page: small, blocks: [lines('A', before), lines('B', n)] }, { fonts: loadFonts })).data);
    // 9 lines before: room for one more → B would leave a single line behind → B starts page 2
    let words = await place(9, 4);
    expect(words.find((w) => w.text === 'A9')!.page).toBe(1);
    expect(words.filter((w) => /^B\d$/.test(w.text)).map((w) => w.page)).toEqual([2, 2, 2, 2]);
    // 7 lines before: room for three → split 3 + 2
    words = await place(7, 5);
    expect(words.filter((w) => /^B\d$/.test(w.text)).map((w) => w.page)).toEqual([1, 1, 1, 2, 2]);
    // 7 lines before, six-line paragraph: 3 + 3
    words = await place(7, 6);
    expect(words.filter((w) => /^B\d$/.test(w.text)).map((w) => w.page)).toEqual([1, 1, 1, 2, 2, 2]);
    // 7 lines before, four-line paragraph: 3 + 1 would be a widow → 2 + 2
    words = await place(7, 4);
    expect(words.filter((w) => /^B\d$/.test(w.text)).map((w) => w.page)).toEqual([1, 1, 2, 2]);
  });
});

describe('floating pictures: round trips', () => {
  const shape = (d: Doc): string[] =>
    d.blocks.map((b) => {
      if (b.type !== 'image') return b.type === 'paragraph' ? `p:${b.runs.map((r) => r.text).join('').slice(0, 6)}` : b.type;
      const f = b.float;
      if (!f) return 'img';
      const d4 = (f.dist ?? []).map((v) => Math.round(v)).join(',');
      return `float:${f.wrap}${f.behind ? '/behind' : ''}|${f.h.from}:${f.h.align ?? Math.round(f.h.offset ?? 0)}|${f.v.from}:${f.v.align ?? Math.round(f.v.offset ?? 0)}|${d4}`;
    });
  const BEHIND: ImageFloat = { wrap: 'none', behind: true, h: { from: 'page', offset: 30 }, v: { from: 'page', offset: 40 }, dist: [0, 0, 0, 0] };
  const RT: Doc = { page: PAGE, blocks: [img(LEFT), para(P1), img(RIGHT), para(P2), img(BEHIND), para('Конец')] };
  const EXPECTED = ['float:square|column:left|paragraph:0|0,9,0,0', 'p:Первый', 'float:square|column:right|paragraph:0|0,0,0,9', 'p:Второй', 'float:none/behind|page:30|page:40|0,0,0,0', 'p:Конец'];

  it('DOCX: wp:anchor with wrapSquare / wrapNone + behindDoc, in the paragraph it belongs to', () => {
    const out = docToDocx(RT).data;
    const xml = new TextDecoder().decode(unzipSync(out)['word/document.xml']);
    expect((xml.match(/<wp:anchor /g) ?? []).length).toBe(3);
    expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>');
    expect(xml).toContain('behindDoc="1"');
    expect(xml).toContain('<wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH>');
    // the anchor sits in the paragraph that follows it, not in an empty one
    expect(xml).toMatch(/<w:p>(<w:pPr>.*?<\/w:pPr>)?<w:r><w:drawing><wp:anchor[^]*?<\/w:drawing><\/w:r><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>Первый/);
    expect(shape(docxToDoc(out).doc)).toEqual(EXPECTED);
  });

  it('ODT: frames anchored to the paragraph, wrap and position in their graphic style', async () => {
    const out = docToOdt(RT).data;
    const xml = new TextDecoder().decode(unzipSync(out)['content.xml']);
    expect(xml).toContain('style:wrap="parallel"');
    expect(xml).toContain('style:run-through="background"');
    expect(shape((await odtToDoc(out)).doc)).toEqual(EXPECTED);
  });

  it('HTML: CSS float for the document editor, data-float for the exact placement', () => {
    const html = docToHtml(RT);
    expect(html).toContain('float:left');
    expect(html).toContain('float:right');
    expect(shape(htmlToDoc(html))).toEqual(EXPECTED);
  });

  it('HTML from elsewhere: a CSS float inside a paragraph floats before it', () => {
    const src = `<p><img src="data:image/png;base64,${Buffer.from(gray(128)).toString('base64')}" width="120" height="120" style="float:right;margin:0 0 6pt 12pt">Текст рядом с картинкой.</p>`;
    const d = htmlToDoc(src);
    expect(d.blocks.map((b) => b.type)).toEqual(['image', 'paragraph']);
    const f = (d.blocks[0] as Extract<Block, { type: 'image' }>).float!;
    expect(f).toMatchObject({ wrap: 'square', h: { from: 'column', align: 'right' }, dist: [0, 0, 6, 12] });
  });

  it('a forged data-float is checked', () => {
    const src = `<p><img src="data:image/png;base64,${Buffer.from(gray(128)).toString('base64')}" width="10" height="10" data-float='{"wrap":"square","h":{"from":"evil","offset":1e99},"v":{"from":"page"},"dist":[1,2,"x",4]}'></p>`;
    const f = (htmlToDoc(src).blocks[0] as Extract<Block, { type: 'image' }>).float!;
    expect(f).toEqual({ wrap: 'square', h: { from: 'column' }, v: { from: 'page' } });
  });
});

describe.skipIf(!hasCmd('soffice') || !hasCmd('pdftotext'))('floating pictures and LibreOffice', () => {
  /** In LibreOffice's PDF: the words on the lines beside the picture (first 90 pt of the paragraph) start right of it. */
  function wrapsLeft(pdf: Uint8Array): void {
    const words = pdfWords(pdf).filter((w) => w.page === 1);
    const first = words.find((w) => w.text === 'Первый')!;
    const beside = words.filter((w) => w.y0 >= first.y0 - 1 && w.y1 < first.y0 + 85);
    const starts = beside.filter((w) => !beside.some((o) => Math.abs(o.y0 - w.y0) < 1 && o.x0 < w.x0));
    expect(starts.length).toBeGreaterThan(3);
    for (const w of starts) expect(w.x0).toBeGreaterThan(M + 90);
    const later = words.filter((w) => w.y0 > first.y0 + 120 && w.y0 < first.y0 + 200);
    expect(Math.min(...later.map((w) => w.x0))).toBeLessThan(M + 10);
  }

  it('LibreOffice lays out our DOCX and ODT with the text around the picture', () => {
    wrapsLeft(soffice(docToDocx(DOC).data, 'docx', 'pdf'));
    wrapsLeft(soffice(docToOdt(DOC).data, 'odt', 'pdf'));
  });

  it('reads the anchors LibreOffice writes (ODT → DOCX)', async () => {
    const odt = soffice(docToDocx(DOC).data, 'docx', 'odt');
    const docx = soffice(odt, 'odt', 'docx:MS Word 2007 XML');
    for (const d of [(await odtToDoc(odt)).doc, docxToDoc(docx).doc]) {
      const floats = d.blocks.filter((b): b is Extract<Block, { type: 'image' }> => b.type === 'image' && !!b.float);
      expect(floats.length).toBe(2);
      expect(floats[0].float!.wrap).toBe('square');
      // our PDF of it: still no words over the picture
      const pdf = (await docToPdf(d, { fonts: loadFonts })).data;
      const pic = blob(renderPages(pdf, 72)[0], 128);
      expect(pic).not.toBeNull();
      expect(overlapping(pdfWords(pdf).filter((w) => w.page === 1), pic!)).toEqual([]);
    }
  });
});

void PDFName;
