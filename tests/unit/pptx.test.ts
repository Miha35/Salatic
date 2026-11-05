/** PowerPoint: theme/master/layout inheritance, the slide model, PDF rendering vs LibreOffice, text export. */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { pptxToDeck, deckToDoc } from '../../src/core/readers/pptx';
import { deckToPdf } from '../../src/core/writers/slides-pdf';
import { docToMarkdown } from '../../src/core/writers/text';
import { detectFormat } from '../../src/core/formats';
import { convert } from '../../src/core/convert';
import type { SlideElement } from '../../src/core/slides';
import { env, hasCmd, loadFonts, loadSubsetter, pdfWords, renderPages, soffice, text } from './helpers';
import { makePptx } from './pptx-fixture';

const pptx = makePptx();
const { deck, warnings } = pptxToDeck(pptx);
const shapes = (i: number): Extract<SlideElement, { kind: 'shape' }>[] => deck.slides[i].elements.filter((e): e is Extract<SlideElement, { kind: 'shape' }> => e.kind === 'shape');
const textOf = (e: Extract<SlideElement, { kind: 'shape' }>): string => e.text?.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n') ?? '';

describe('PPTX reader', () => {
  it('reads the deck: size, title, slides, hidden slide', () => {
    expect(detectFormat('deck.zip', pptx).format).toBe('pptx');
    expect(warnings).toEqual([]);
    expect(deck.width).toBe(720);
    expect(deck.height).toBe(540);
    expect(deck.title).toBe('Тестовая презентация');
    expect(deck.slides.length).toBe(4);
    expect(deck.slides.map((s) => !!s.hidden)).toEqual([false, false, true, false]);
  });

  it('inherits placeholder geometry and text styles from layout and master', () => {
    const [title, sub] = shapes(0).filter((s) => s.role !== 'other');
    // ctrTitle: geometry and centred 60 pt from the layout, colour and font from the master/theme
    expect(title).toMatchObject({ role: 'title', x: 90, y: 88.4, w: 540, h: 188 });
    expect(title.text?.anchor).toBe('b');
    expect(title.text?.paragraphs[0]).toMatchObject({ align: 'center' });
    expect(title.text?.paragraphs[0].runs[0]).toMatchObject({ size: 60, color: '000000', font: 'calibri' });
    expect(sub).toMatchObject({ role: 'subtitle', x: 90, y: 283.6 });
    expect(sub.text?.paragraphs[0].bullet).toBeUndefined();
    expect(sub.text?.paragraphs[0].runs[0].size).toBe(24);
    // slide 2 title: no geometry on slide or layout → master
    const t2 = shapes(1).find((s) => s.role === 'title')!;
    expect(t2).toMatchObject({ x: 36, y: 28.8, w: 648, h: 104.4 });
    expect(t2.text?.paragraphs[0].runs[0].size).toBe(44);
  });

  it('resolves bullets, levels, numbering, autofit and links in the body', () => {
    const body = shapes(1).find((s) => s.role === 'body')!;
    const ps = body.text!.paragraphs;
    expect(body.text!.fontScale).toBeCloseTo(0.85);
    expect(ps.map((p) => p.bullet?.text)).toEqual(['•', '–', '1.', '2.', '•']);
    expect(ps.map((p) => p.level)).toEqual([0, 1, 0, 0, 0]);
    expect(ps[0].runs[0].size).toBeCloseTo(28 * 0.85, 1);
    expect(ps[1].runs[0].size).toBeCloseTo(24 * 0.85, 1);
    expect(ps[0]).toMatchObject({ marL: 18, indent: -18, spaceBefore: 10 });
    expect(ps[0].lineHeight).toBeCloseTo(0.9 * 0.9, 3);
    expect(ps[4].runs[0].link).toBe('https://example.com/deck');
  });

  it('applies group transforms, theme colour modifiers, cropping and connectors', () => {
    const els = deck.slides[1].elements;
    const g = shapes(1).find((s) => textOf(s) === 'В группе')!;
    // child (200,0,200,100) in a 400×200 child space mapped onto 200×100 at (400,300)
    expect(g).toMatchObject({ x: 500, y: 300, w: 100, h: 50, geom: 'roundRect' });
    // accent2 ED7D31 with lumMod 75% → C55A11 (PowerPoint's "Orange, Accent 2, Darker 25%")
    expect(g.fill).toBe('C55A11');
    expect(g.text?.paragraphs[0].runs[0]).toMatchObject({ color: 'FFFFFF', bold: true, size: 14 });
    const img = els.find((e) => e.kind === 'image');
    expect(img).toMatchObject({ x: 400, y: 150, w: 100, h: 100, mime: 'image/png', alt: 'Картинка', crop: { l: 0, t: 0, r: 0.5, b: 0 } });
    const line = els.find((e) => e.kind === 'line');
    expect(line).toMatchObject({ x1: 36, y1: 390, x2: 336, line: { color: 'FF0000', width: 2 } });
    // the master's decorative band is drawn on every slide, behind the content
    for (const s of deck.slides) expect(s.elements[0]).toMatchObject({ kind: 'shape', y: 530, h: 10, fill: '4472C4' });
  });

  it('reads tables with the default table style, spans and merged cells', () => {
    const t = deck.slides[1].elements.find((e) => e.kind === 'table');
    if (t?.kind !== 'table') throw new Error('no table');
    expect(t.cols).toEqual([150, 150]);
    expect(t.rows[0].cells.map((c) => c.fill)).toEqual(['4472C4', '4472C4']);
    expect(t.rows[0].cells[0].text.paragraphs[0].runs[0]).toMatchObject({ color: 'FFFFFF', bold: true });
    expect(t.rows[2].cells.map((c) => [c.span, c.merged])).toEqual([
      [2, false],
      [1, true],
    ]);
  });

  it('reads backgrounds, slide-number fields and speaker notes', () => {
    expect(deck.slides[0].background?.color).toBe('FFFFFF');
    expect(deck.slides[3].background?.color).toBe('1F3864');
    const num = shapes(3).find((s) => s.role === 'footer')!;
    expect(textOf(num)).toBe('4');
    expect(num).toMatchObject({ x: 516, y: 500.5 });
    const title = shapes(3).find((s) => s.role === 'title')!;
    expect(title.text?.paragraphs[0].runs[0].color).toBe('FFFFFF');
    expect(deck.slides[1].notes?.length).toBe(1);
  });

  it('flattens to a document for text formats', () => {
    const md = docToMarkdown(deckToDoc(deck));
    expect(md).toContain('## Заголовок презентации');
    expect(md).toContain('## Список и объекты');
    expect(md).toMatch(/- Первый пункт\n {2}- Вложенный пункт/);
    expect(md).toMatch(/1\. Нумерованный один\n2\. Нумерованный два/);
    expect(md).toContain('[Ссылка](https://example.com/deck)');
    expect(md).toContain('| Москва | 1 200 |');
    expect(md).toContain('Не забыть про сроки');
    expect(md).not.toContain('Скрытый слайд');
    expect(md).not.toMatch(/^4$/m); // slide-number placeholders are not content
  });

  it('rejects broken files', () => {
    expect(() => pptxToDeck(new Uint8Array([0x50, 0x4b, 3, 4]))).toThrow();
    expect(detectFormat('old.ppt', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])).error).toMatch(/\.pptx/);
  });
});

describe('PPTX → PDF', () => {
  let ours: Uint8Array;
  const render = async (): Promise<Uint8Array> => (ours ??= (await deckToPdf(deck, { fonts: loadFonts, subsetter: await loadSubsetter() })).data);

  it('renders one page per visible slide at the slide size', async () => {
    const pdf = await render();
    expect(text(pdf.subarray(0, 5))).toBe('%PDF-');
    if (!hasCmd('pdfinfo')) return;
    const pages = renderPages(pdf, 20);
    expect(pages.length).toBe(3);
  });

  it.skipIf(!hasCmd('pdftotext'))('places text where the inherited geometry puts it', async () => {
    const words = pdfWords(await render());
    const w = (page: number, t: string) => words.find((x) => x.page === page && x.text === t)!;
    // bottom-anchored centred title inside (90, 88.4, 540×188)
    const title = w(1, 'презентации');
    // with 90 % line spacing the descenders reach slightly below the box, as in PowerPoint
    expect(title.y1).toBeLessThanOrEqual(88.4 + 188 + 8);
    expect(title.y1).toBeGreaterThan(88.4 + 188 - 30);
    expect(Math.abs((title.x0 + title.x1) / 2 - 360)).toBeLessThan(2);
    // table cells start at the frame origin plus the cell inset
    expect(w(2, 'Москва').x0).toBeCloseTo(36 + 7.2, 0);
    // the slide number is right-aligned in its placeholder
    expect(w(3, '4').x1).toBeCloseTo(516 + 168 - 7.2, 0);
    expect(words.some((x) => x.text === 'Скрытый')).toBe(false);
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftotext'))('agrees with LibreOffice on the positions of shared-font text', async () => {
    const lo = pdfWords(soffice(pptx, 'pptx', 'pdf'));
    const mine = pdfWords(await render());
    // Table and group text use Calibri, which both renderers map to the metric-compatible Carlito
    for (const t of ['Город', 'Продажи', 'Москва', 'Широкая', 'группе']) {
      const a = lo.find((x) => x.text === t)!;
      const b = mine.find((x) => x.text === t)!;
      expect(Math.abs(a.x0 - b.x0), t).toBeLessThan(4);
      expect(Math.abs(a.y0 - b.y0), t).toBeLessThan(4);
    }
    // Titles: LibreOffice substitutes Calibri Light, so only the vertical placement is compared
    for (const [page, t] of [[1, 'Заголовок'], [1, 'Подзаголовок'], [3, 'Тёмный']] as const) {
      const a = lo.find((x) => x.page === page && x.text === t)!;
      const b = mine.find((x) => x.page === page && x.text === t)!;
      expect(Math.abs(a.y0 - b.y0), t).toBeLessThan(8);
    }
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftotext'))('renders a LibreOffice-made presentation', async () => {
    const tpl = readFileSync(join(__dirname, '../fixtures/deck.fodp.tpl'), 'utf8').replace('@PNG@', Buffer.from(makePptxPng()).toString('base64'));
    const loPptx = soffice(new TextEncoder().encode(tpl), 'fodp', 'pptx');
    const r = await convert(loPptx, 'pptx', 'pdf', env);
    const words = pdfWords(r.data).map((x) => x.text);
    for (const t of ['Годовой', 'Итоги', 'Выручка', 'Казань', 'Показатели', 'Москва', 'Рост', '中文']) expect(words).toContain(t);
    const md = text((await convert(loPptx, 'pptx', 'md', env)).data);
    expect(md).toMatch(/- Открыто 3 новых филиала\n {2}- Москва и Казань/);
    expect(md).toContain('| Москва | 1 200 |');
  });
});

function makePptxPng(): Uint8Array {
  // reuse the fixture's PNG generator through the deck's own picture
  const img = deck.slides[1].elements.find((e) => e.kind === 'image');
  return img?.kind === 'image' ? img.data : new Uint8Array();
}
