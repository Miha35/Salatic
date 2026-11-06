/** Redaction: content inside rectangles is removed from the content streams; everything else stays put. */
import { describe, expect, it } from 'vitest';
import fontkit from '@pdf-lib/fontkit';
import { unzlibSync } from 'fflate';
import { PDFDocument, PDFName, PDFRawStream, rgb, setTextRenderingMode, StandardFonts, TextRenderingMode, type PDFPage } from 'pdf-lib';
import { redactPage, type Rect } from '../../src/core/pdf-redact';
import { composePdf } from '../../src/core/pdf-tools';
import { encodePng } from '../../src/core/png';
import { docToDocx } from '../../src/core/writers/docx';
import { hasCmd, loadFonts, pdfToText, pdfWords, renderPages, soffice, type WordBox } from './helpers';
import { sample } from './sample';

const LINES = ['Получатель: Иванов Пётр Сергеевич', 'Паспорт 4510 123456 выдан ОВД', 'Счёт 40817810099910004312 в банке'];

/** A PDF with three lines in an embedded TrueType font (Type0/Identity-H, as pdf-lib and most tools write). */
async function cidPdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(await loadFonts('sans-regular'), { subset: true });
  const page = doc.addPage([400, 300]);
  LINES.forEach((t, i) => page.drawText(t, { x: 40, y: 240 - i * 40, size: 14, font }));
  return doc.save();
}

/** Bounding box (PDF user space, y up) of a word as poppler sees it on a page of height h. */
const rectOf = (w: WordBox, h: number, pad = 0.5): Rect => ({ x0: w.x0 - pad, y0: h - w.y1 - pad, x1: w.x1 + pad, y1: h - w.y0 + pad });

async function redactWords(pdf: Uint8Array, words: string[], redact = false): Promise<{ out: Uint8Array; glyphs: number; before: WordBox[]; after: WordBox[] }> {
  const before = pdfWords(pdf);
  const doc = await PDFDocument.load(pdf);
  const page = doc.getPage(0);
  const h = page.getHeight();
  const rects = before.filter((w) => w.page === 1 && words.includes(w.text)).map((w) => rectOf(w, h));
  expect(rects.length).toBe(words.length);
  const r = await redactPage(doc, page, rects, { redact });
  const out = await doc.save();
  return { out, glyphs: r.glyphs, before, after: pdfWords(out) };
}

/** Every word not redacted is still there, at the same place. */
function expectUnmoved(before: WordBox[], after: WordBox[], removed: string[]): void {
  const kept = before.filter((w) => !removed.includes(w.text));
  for (const w of kept) {
    const same = after.find((a) => a.page === w.page && a.text === w.text && Math.abs(a.x0 - w.x0) < 0.05 && Math.abs(a.y0 - w.y0) < 0.05);
    expect(same, `«${w.text}» moved or vanished`).toBeDefined();
  }
  for (const r of removed) expect(after.some((a) => a.text === r)).toBe(false);
}

describe.skipIf(!hasCmd('pdftotext'))('redaction of text', () => {
  it('embedded CID fonts: the words are gone, the rest of the line keeps its positions', async () => {
    const { out, glyphs, before, after } = await redactWords(await cidPdf(), ['4510', '123456', '40817810099910004312']);
    expect(glyphs).toBe(4 + 6 + 20);
    expectUnmoved(before, after, ['4510', '123456', '40817810099910004312']);
    const text = pdfToText(out);
    expect(text).not.toMatch(/4510|123456|4081781/);
    expect(text).toContain('выдан ОВД');
  });

  it('standard fonts without /Widths (Helvetica): advances from the AFM metrics', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([400, 200]);
    page.drawText('Account 1234-5678 balance: 99 USD', { x: 30, y: 150, size: 13, font });
    page.drawText('Wide WAVY kerning AVAWA test', { x: 30, y: 100, size: 13, font });
    const { glyphs, before, after } = await redactWords(await doc.save(), ['1234-5678', 'kerning']);
    expect(glyphs).toBe(9 + 7);
    expectUnmoved(before, after, ['1234-5678', 'kerning']);
  });

  it.skipIf(!hasCmd('soffice'))('a LibreOffice PDF (subset fonts, kerning in TJ)', async () => {
    const pdf = soffice(docToDocx(sample).data, 'docx', 'pdf');
    const words = pdfWords(pdf).filter((w) => w.page === 1 && /^[A-Za-zА-Яа-яЁё]{5,}$/.test(w.text));
    const targets = [...new Set(words.map((w) => w.text))].filter((t) => words.filter((w) => w.text === t).length === 1).slice(0, 4);
    expect(targets.length).toBe(4);
    const { glyphs, before, after } = await redactWords(pdf, targets);
    expect(glyphs).toBe(targets.join('').length);
    expectUnmoved(before.filter((w) => w.page === 1), after.filter((w) => w.page === 1), targets);
  });

  it('removes invisible text too (OCR layers)', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([300, 100]);
    page.pushOperators(setTextRenderingMode(TextRenderingMode.Invisible));
    page.drawText('SECRET visible-to-search', { x: 20, y: 50, size: 12, font });
    const pdf = await doc.save();
    expect(pdfToText(pdf)).toContain('SECRET');
    const { out, glyphs } = await redactWords(pdf, ['SECRET']);
    expect(glyphs).toBe(6);
    expect(pdfToText(out)).not.toContain('SECRET');
    expect(pdfToText(out)).toContain('visible-to-search');
  });

  it('text inside a form XObject is removed on a private copy (other pages keep it)', async () => {
    const src = await PDFDocument.load(await cidPdf());
    const doc = await PDFDocument.create();
    const form = await doc.embedPage(src.getPage(0));
    for (let i = 0; i < 2; i++) doc.addPage([400, 300]).drawPage(form, { x: 0, y: 0 });
    const pdf = await doc.save();
    const words = pdfWords(pdf);
    const target = words.find((w) => w.page === 1 && w.text === '123456')!;
    const d2 = await PDFDocument.load(pdf);
    const r = await redactPage(d2, d2.getPage(0), [rectOf(target, 300)]);
    expect(r.glyphs).toBe(6);
    const after = pdfWords(await d2.save());
    expect(after.filter((w) => w.text === '123456').map((w) => w.page)).toEqual([2]);
    expect(after.filter((w) => w.page === 1).length).toBe(words.filter((w) => w.page === 1).length - 1);
  });

  it('nothing in the area: the page is left untouched', async () => {
    const pdf = await cidPdf();
    const doc = await PDFDocument.load(pdf);
    const contents = doc.getPage(0).node.get(PDFName.of('Contents'));
    const r = await redactPage(doc, doc.getPage(0), [{ x0: 0, y0: 0, x1: 10, y1: 10 }]);
    expect(r.glyphs).toBe(0);
    expect(doc.getPage(0).node.get(PDFName.of('Contents'))).toBe(contents);
  });
});

/** Raw RGB samples of the first image XObject of a page. */
function imageSamples(page: PDFPage): { w: number; h: number; data: Uint8Array } {
  const xo = page.node.Resources()!.lookup(PDFName.of('XObject')) as unknown as { entries(): [PDFName, unknown][] };
  for (const [, ref] of xo.entries()) {
    const s = page.doc.context.lookup(ref as never);
    if (s instanceof PDFRawStream && s.dict.lookup(PDFName.of('Subtype'))?.toString() === '/Image') {
      const w = (s.dict.lookup(PDFName.of('Width')) as unknown as { asNumber(): number }).asNumber();
      const h = (s.dict.lookup(PDFName.of('Height')) as unknown as { asNumber(): number }).asNumber();
      return { w, h, data: unzlibSync(s.getContents()) };
    }
  }
  throw new Error('no image');
}

describe('redaction of graphics', () => {
  it('zeroes the image pixels under the area and nothing else', async () => {
    const doc = await PDFDocument.create();
    const px = new Uint8Array(40 * 20 * 3).fill(200);
    const img = await doc.embedPng(encodePng(px, 40, 20, 3));
    const page = doc.addPage([300, 200]);
    page.drawImage(img, { x: 100, y: 50, width: 160, height: 80 }); // 4 pt per pixel
    const d2 = await PDFDocument.load(await doc.save());
    // left quarter of the image, upper half
    const r = await redactPage(d2, d2.getPage(0), [{ x0: 90, y0: 90, x1: 140, y1: 140 }], { redact: true });
    expect(r.images).toBe(1);
    const out = await PDFDocument.load(await d2.save());
    const { w, h, data } = imageSamples(out.getPage(0));
    expect([w, h]).toEqual([40, 20]);
    const at = (x: number, y: number): number => data[(y * w + x) * 3];
    expect(at(0, 0)).toBe(0); // top-left: redacted
    expect(at(9, 9)).toBe(0);
    expect(at(10, 0)).toBe(200); // right of the area
    expect(at(0, 10)).toBe(200); // below it
    expect(at(39, 19)).toBe(200);
  });

  it('removes vector graphics lying inside the area, keeps those crossing it', async () => {
    const doc = await PDFDocument.create();
    const page = doc.addPage([300, 200]);
    page.drawRectangle({ x: 20, y: 20, width: 40, height: 30, color: rgb(1, 0, 0) }); // inside
    page.drawRectangle({ x: 150, y: 20, width: 120, height: 30, color: rgb(0, 0, 1) }); // crossing
    const d2 = await PDFDocument.load(await doc.save());
    const r = await redactPage(d2, d2.getPage(0), [{ x0: 10, y0: 10, x1: 200, y1: 100 }], { redact: true });
    expect(r.paths).toBe(1);
    if (!hasCmd('pdftoppm')) return;
    const [pgm] = renderPages(await d2.save(), 72);
    // PGM (P5): gray value at a point of the page
    const header = /^P5\s+(\d+)\s+(\d+)\s+255\s/.exec(new TextDecoder('latin1').decode(pgm.subarray(0, 40)))!;
    const W = +header[1];
    const gray = (x: number, y: number): number => pgm[header[0].length + (200 - y) * W + x];
    expect(gray(40, 35)).toBe(255); // the red box is gone
    expect(gray(250, 35)).toBeLessThan(100); // the blue one stays
  });

  it('drops annotations overlapping the area', async () => {
    const doc = await PDFDocument.create();
    const page = doc.addPage([300, 200]);
    const link = doc.context.register(doc.context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [20, 20, 80, 40], A: { S: 'URI', URI: 'https://example.com/secret' } }));
    const other = doc.context.register(doc.context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [200, 150, 260, 170], A: { S: 'URI', URI: 'https://example.com/ok' } }));
    page.node.set(PDFName.of('Annots'), doc.context.obj([link, other]));
    const r = await redactPage(doc, page, [{ x0: 10, y0: 10, x1: 100, y1: 50 }], { redact: true });
    expect(r.annotations).toBe(1);
    expect(page.node.Annots()!.size()).toBe(1);
  });
});

describe.skipIf(!hasCmd('pdftotext'))('editor operations (composePdf)', () => {
  it('redact boxes remove the data and paint the area; whiteout removes the text under it', async () => {
    const pdf = await cidPdf();
    const words = pdfWords(pdf);
    const box = (t: string): { x: number; y: number; w: number; h: number } => {
      const w = words.find((x) => x.text === t)!;
      return { x: w.x0 - 1, y: 300 - w.y1 - 1, w: w.x1 - w.x0 + 2, h: w.y1 - w.y0 + 2 };
    };
    const out = await composePdf([pdf], [{ src: 0, index: 0, rotate: 0, annotations: [{ type: 'redact', ...box('123456'), color: [0, 0, 0], rotate: 0 }, { type: 'rect', ...box('Иванов'), color: [1, 1, 1], opacity: 1, rotate: 0, erase: true }] }], loadFonts);
    const text = pdfToText(out);
    expect(text).not.toContain('123456');
    expect(text).not.toContain('Иванов');
    expect(text).toContain('4510');
    expect(text).toContain('Пётр');
  });

  it('editing existing text: the old words are removed, the new ones written', async () => {
    const pdf = await cidPdf();
    const w = pdfWords(pdf).filter((x) => x.text === 'Иванов' || x.text === 'Пётр' || x.text === 'Сергеевич');
    const x0 = Math.min(...w.map((x) => x.x0));
    const x1 = Math.max(...w.map((x) => x.x1));
    const top = Math.min(...w.map((x) => x.y0));
    const bottom = Math.max(...w.map((x) => x.y1));
    const out = await composePdf(
      [pdf],
      [{ src: 0, index: 0, rotate: 0, annotations: [{ type: 'erase', x: x0 - 1, y: 300 - bottom - 1, w: x1 - x0 + 2, h: bottom - top + 2, rotate: 0, fallback: [1, 1, 1] }, { type: 'text', x: x0, y: 240, size: 14, color: [0, 0, 0], text: 'Петров Иван', rotate: 0 }] }],
      loadFonts,
    );
    const text = pdfToText(out).replace(/\s+/g, ' ');
    expect(text).toContain('Получатель: Петров Иван');
    expect(text).not.toMatch(/Иванов|Сергеевич/);
  });
});
