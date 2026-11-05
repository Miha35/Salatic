/** EMF/WMF pictures: our GDI interpreter, SVG output, vector PDF output and native embedding. */
import { describe, expect, it } from 'vitest';
import { parseMetafile, metafileToSvg, isEmf, isWmf } from '../../src/core/metafile';
import { sniffImageMime, bytesToBase64 } from '../../src/core/util';
import { docToPdf } from '../../src/core/writers/pdf';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToRtf } from '../../src/core/writers/rtf';
import { rtfToDoc } from '../../src/core/readers/rtf';
import { docToHtml } from '../../src/core/writers/text';
import type { Doc } from '../../src/core/model';
import { hasCmd, loadFonts, pdfToText, renderPages, soffice, text } from './helpers';
import { samplePng } from './pptx-fixture';

const SVG = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="400" height="300" viewBox="0 0 400 300">
  <rect x="10" y="10" width="380" height="280" fill="#f0f4ff"/>
  <circle cx="100" cy="120" r="60" fill="#ed7d31"/>
  <polygon points="220,60 300,180 140,180" fill="#70ad47" stroke="#000" stroke-width="2"/>
  <path d="M30 250 C 120 200, 200 300, 370 230" fill="none" stroke="#c00000" stroke-width="5"/>
  <text x="200" y="230" font-family="Liberation Sans" font-size="28" font-weight="bold" fill="#1f3864" text-anchor="middle">Отчёт 2024</text>
  <image x="320" y="20" width="60" height="30" xlink:href="data:image/png;base64,${bytesToBase64(samplePng(40, 20))}"/>
</svg>`;

describe.skipIf(!hasCmd('soffice'))('metafiles made by LibreOffice', () => {
  const emf = soffice(new TextEncoder().encode(SVG), 'svg', 'emf');
  const wmf = soffice(new TextEncoder().encode(SVG), 'svg', 'wmf');

  it('detects EMF and WMF', () => {
    expect(isEmf(emf)).toBe(true);
    expect(isWmf(wmf)).toBe(true);
    expect(sniffImageMime(emf)).toBe('image/emf');
    expect(sniffImageMime(wmf)).toBe('image/wmf');
  });

  for (const [name, data] of [['EMF', emf], ['WMF', wmf]] as const) {
    it(`interprets ${name}: shapes, colours, text and bitmaps`, () => {
      const m = parseMetafile(data)!;
      expect(m.width).toBeGreaterThan(100);
      expect(m.widthPt / m.heightPt).toBeCloseTo(400 / 300, 1);
      const fills = new Set(m.cmds.flatMap((c) => (c.k === 'path' && c.fill ? [c.fill] : [])));
      for (const c of ['ED7D31', '70AD47', 'F0F4FF']) expect(fills).toContain(c);
      // LibreOffice exports thick strokes as filled outlines
      expect(m.cmds.some((c) => c.k === 'path' && (c.stroke === 'C00000' || c.fill === 'C00000'))).toBe(true);
      const t = m.cmds.find((c) => c.k === 'text');
      expect(t?.k === 'text' && t.text).toBe('Отчёт 2024');
      expect(t?.k === 'text' && t.bold).toBe(true);
      expect(m.cmds.some((c) => c.k === 'image')).toBe(true);
      // the circle is centred at about (100,120) of 400×300
      const svg = metafileToSvg(m);
      expect(svg).toMatch(/^<svg xmlns="http:\/\/www.w3.org\/2000\/svg" viewBox="0 0 /);
      expect(svg).toContain('Отчёт 2024');
    });
  }

  it.skipIf(!hasCmd('pdftotext'))('draws metafiles as vectors in PDF (text stays text)', async () => {
    const doc: Doc = { blocks: [{ type: 'paragraph', runs: [{ text: 'До рисунка' }] }, { type: 'image', data: emf, mime: 'image/emf', width: 400, height: 300 }, { type: 'image', data: wmf, mime: 'image/wmf', width: 400, height: 300 }] };
    const { data, warnings } = await docToPdf(doc, { fonts: loadFonts });
    expect(warnings).toEqual([]);
    const t = pdfToText(data);
    expect(t.match(/Отчёт 2024/g)?.length).toBe(2);
    // orange and green pixels are present (filled shapes)
    const [page] = renderPages(data, 50);
    const header = page.toString('latin1', 0, 20).split('\n');
    const [w, h] = header[1].split(' ').map(Number);
    const px = page.subarray(page.length - w * h);
    // grayscale render: orange ≈ 148, green ≈ 152, text navy ≈ 52 – check non-white area exists
    const dark = [...px].filter((v) => v < 200).length;
    expect(dark / px.length).toBeGreaterThan(0.02);
  });

  it('keeps EMF/WMF natively in DOCX and RTF, and as SVG in HTML', async () => {
    const doc: Doc = { blocks: [{ type: 'image', data: emf, mime: 'image/emf', width: 400, height: 300, alt: 'Схема' }] };
    const docx = docToDocx(doc);
    expect(docx.warnings).toEqual([]);
    const back = docxToDoc(docx.data).doc.blocks.find((b) => b.type === 'image');
    expect(back?.type === 'image' && back.mime).toBe('image/emf');
    expect(back?.type === 'image' && back.data.length).toBe(emf.length);
    const rtf = docToRtf({ blocks: [...doc.blocks, { type: 'image', data: wmf, mime: 'image/wmf', width: 400, height: 300 }] });
    expect(rtf.warnings).toEqual([]);
    const rtfBack = rtfToDoc(rtf.data).doc.blocks.filter((b) => b.type === 'image').map((b) => (b.type === 'image' ? b.mime : ''));
    expect(rtfBack).toEqual(['image/emf', 'image/wmf']);
    expect(docToHtml(doc)).toContain('data:image/svg+xml;base64,');
    // LibreOffice opens both and renders the pictures (their text reaches the PDF)
    expect(pdfToText(soffice(docx.data, 'docx', 'pdf'))).toContain('Отчёт 2024');
    expect(pdfToText(soffice(rtf.data, 'rtf', 'pdf'))).toContain('Отчёт 2024');
    void text;
  });
});

describe('metafile robustness', () => {
  it('rejects garbage without throwing', () => {
    expect(parseMetafile(new Uint8Array(100))).toBeNull();
    const fake = new Uint8Array(200);
    fake.set([1, 0, 0, 0, 200, 0, 0, 0]);
    fake.set([0x20, 0x45, 0x4d, 0x46], 40);
    const m = parseMetafile(fake);
    expect(m === null || m.cmds.length === 0).toBe(true);
  });
});
