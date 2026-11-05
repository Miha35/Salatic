/** Editing existing PDF text: line detection with font identification, and replacement through composePdf. */
import { describe, expect, it } from 'vitest';
import { buildSegments, familyFor } from '../../src/core/text-lines';
import { composePdf } from '../../src/core/pdf-tools';
import { docToPdf } from '../../src/core/writers/pdf';
import type { Doc } from '../../src/core/model';
import { hasCmd, loadFonts, loadPdfjs, loadSubsetter, pdfWords, renderPages } from './helpers';

const doc: Doc = {
  font: 'serif',
  size: 12,
  blocks: [
    { type: 'heading', level: 1, runs: [{ text: 'Договор поставки № 15' }] },
    { type: 'paragraph', runs: [{ text: 'Сумма договора: 120 000 рублей.' }] },
    { type: 'paragraph', runs: [{ text: 'Срок оплаты — ', italic: true }, { text: '10 дней', italic: true, bold: true }] },
    { type: 'paragraph', runs: [{ text: 'Код: ', font: 'mono' }, { text: 'ABC-123', font: 'mono' }] },
  ],
};

async function segmentsOf(pdf: Uint8Array) {
  const pdfjs = await loadPdfjs();
  const d = await pdfjs.getDocument({ data: pdf.slice(), verbosity: 0, fontExtraProperties: true }).promise;
  const page = await d.getPage(1);
  await page.getOperatorList(); // loads the fonts into commonObjs, as rendering does in the editor
  const tc = await page.getTextContent();
  const objs = page.commonObjs as unknown as { has(id: string): boolean; get(id: string): { name?: string; bold?: boolean; italic?: boolean; isSerifFont?: boolean; isMonospace?: boolean } };
  const info = (name: string) => {
    const f = objs.has(name) ? objs.get(name) : {};
    const n = f.name ?? '';
    return { family: familyFor(n, f.isSerifFont, f.isMonospace), bold: !!f.bold || /bold/i.test(n), italic: !!f.italic || /italic/i.test(n) };
  };
  const items = tc.items.filter((i): i is typeof i & { str: string; transform: number[]; width: number; fontName: string; hasEOL: boolean } => 'str' in i);
  return buildSegments(items, info);
}

describe('existing PDF text', () => {
  it('maps font names to families', () => {
    expect(familyFor('ABCDEF+TimesNewRomanPS-BoldMT')).toBe('serif');
    expect(familyFor('ArialMT')).toBe('sans');
    expect(familyFor('Calibri-Light')).toBe('calibri');
    expect(familyFor('CourierNewPSMT')).toBe('mono');
    expect(familyFor('Unknown', true)).toBe('serif');
  });

  it('finds lines with their size, family and style', async () => {
    const { data } = await docToPdf(doc, { fonts: loadFonts });
    const segs = await segmentsOf(data);
    const by = (t: string) => segs.find((s) => s.text === t)!;
    expect(segs.map((s) => s.text)).toEqual(['Договор поставки № 15', 'Сумма договора: 120 000 рублей.', 'Срок оплаты — 10 дней', 'Код: ABC-123']);
    expect(by('Договор поставки № 15')).toMatchObject({ font: 'serif', bold: true });
    expect(by('Договор поставки № 15').size).toBeCloseTo(24, 0);
    expect(by('Сумма договора: 120 000 рублей.')).toMatchObject({ font: 'serif', bold: false, italic: false, size: 12 });
    expect(by('Срок оплаты — 10 дней')).toMatchObject({ italic: true });
    expect(by('Код: ABC-123')).toMatchObject({ font: 'mono' });
  });

  it.skipIf(!hasCmd('pdftotext'))('replaces a line: old text covered, new text in place with the same font', async () => {
    const { data } = await docToPdf(doc, { fonts: loadFonts });
    const seg = (await segmentsOf(data)).find((s) => s.text.startsWith('Сумма'))!;
    const out = await composePdf(
      [data],
      [
        {
          src: 0,
          index: 0,
          rotate: 0,
          annotations: [
            { type: 'rect', x: seg.x - 1, y: seg.y - seg.size * 0.26, w: seg.w + 2, h: seg.size * 1.2, color: [1, 1, 1], opacity: 1, rotate: 0 },
            { type: 'text', x: seg.x, y: seg.y, size: seg.size, color: [0, 0, 0], text: 'Сумма договора: 150 000 рублей.', rotate: 0, font: seg.font, bold: seg.bold, italic: seg.italic },
          ],
        },
      ],
      loadFonts,
      await loadSubsetter(),
    );
    const words = pdfWords(out);
    const fresh = words.find((w) => w.text === '150')!;
    const old = words.find((w) => w.text === '120')!;
    // New text sits on the old baseline (same line box) and uses the serif face
    expect(Math.abs(fresh.y1 - old.y1)).toBeLessThan(1);
    // Visually the old digits are gone: the area of "120" is covered by the white box
    const [before] = renderPages(data, 72);
    const [after] = renderPages(out, 72);
    expect(before.equals(after)).toBe(false);
    const fonts = (await import('node:child_process')).execFileSync('pdffonts', ['-'], { input: out }).toString();
    expect(fonts).toMatch(/LiberationSerif/);
  });
});
