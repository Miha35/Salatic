/** Office charts (c:chartSpace) and SVG images: SVG rendering, the SVG → vector display list, and PDF output vs LibreOffice. */
import { describe, expect, it } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { chartToSvg, formatNumber } from '../../src/core/chart';
import { parseSvg } from '../../src/core/svg-render';
import { pptxToDeck } from '../../src/core/readers/pptx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToDocx } from '../../src/core/writers/docx';
import { convert } from '../../src/core/convert';
import type { Block } from '../../src/core/model';
import type { MetaCmd } from '../../src/core/metafile';
import { barChart, pieChart, sampleCharts, SAMPLE_CATS } from './chart-fixture';
import { makePptx } from './pptx-fixture';
import { sample } from './sample';
import { env, hasCmd, pdfToText, renderPages, soffice } from './helpers';

const norm = (s: string): string => s.replace(/\s+/g, ' ');
const texts = (svg: string): string[] => [...svg.matchAll(/<text[^>]*>([^<]*)<\/text>/g)].map((m) => norm(m[1]));
const count = (svg: string, tag: string): number => svg.split(`<${tag} `).length - 1;

describe('number formats', () => {
  it('formats values like Excel in a Russian locale', () => {
    expect(norm(formatNumber(1500000))).toBe('1 500 000');
    expect(formatNumber(0.25, '0%')).toBe('25%');
    expect(formatNumber(0.256, '0.0%')).toBe('25,6%');
    expect(norm(formatNumber(1234.5, '#,##0.00'))).toBe('1 234,50');
    expect(formatNumber(-3, '0')).toBe('−3');
    expect(formatNumber(2.5)).toBe('2,5');
    expect(norm(formatNumber(1200, '#,##0" ₽"'))).toBe('1 200 ₽');
  });
});

describe('chart → SVG', () => {
  const charts = sampleCharts();
  const svg = (k: string, w = 432, h = 252): string => chartToSvg(charts[k], { width: w, height: h })!;

  it('renders every sample chart type as well-formed SVG of the requested size', () => {
    for (const k of Object.keys(charts)) {
      const s = svg(k);
      expect(s, k).toMatch(/^<svg [^>]*viewBox="0 0 432 252"/);
      const mf = parseSvg(s);
      expect(mf, k).not.toBeNull();
      expect(mf!.widthPt).toBeCloseTo(432);
      expect(mf!.heightPt).toBeCloseTo(252);
      expect(mf!.cmds.length, k).toBeGreaterThan(5);
    }
  });

  it('clustered columns: one bar per value, legend, nice axis and data labels', () => {
    const s = svg('bar');
    // 3 series × 4 categories + 3 legend keys
    expect(count(s, 'rect')).toBe(12 + 3);
    const t = texts(s);
    expect(t[0]).toBe('Продажи, тыс. ₽');
    expect(t).toEqual(expect.arrayContaining(['План', 'Факт', 'Прогноз', ...SAMPLE_CATS, '0', '50', '100', '150', '200', '190']));
    expect(t).not.toContain('250');
  });

  it('bars are proportional to their values', () => {
    const s = chartToSvg(barChart(['a', 'b'], [{ name: 'S', values: [10, 40], color: '123456' }]), { width: 400, height: 300 })!;
    const bars = [...s.matchAll(/<rect x="[\d.]+" y="([\d.]+)" width="[\d.]+" height="([\d.]+)" fill="#123456"\/>/g)].map((m) => +m[2]);
    expect(bars.length).toBe(3); // the legend key, then two bars
    const [a, b] = bars.slice(1);
    expect(b / a).toBeCloseTo(4, 1);
  });

  it('horizontal stacked bars list the first category at the bottom', () => {
    const s = svg('hbar');
    const y = (label: string): number => +new RegExp(`<text[^>]*y="([\\d.]+)"[^>]*>${label}</text>`).exec(s)![1];
    expect(y('Москва')).toBeGreaterThan(y('Санкт-Петербург'));
    expect(y('Санкт-Петербург')).toBeGreaterThan(y('Казань'));
  });

  it('combo chart: secondary percentage axis and grouped thousands', () => {
    const t = texts(svg('combo'));
    expect(t).toEqual(expect.arrayContaining(['Выручка', 'Маржа', '500 000', '2 000 000', '0%', '30%']));
  });

  it('pie: slices as paths and percentage labels that sum to 100', () => {
    const s = svg('pie');
    expect(count(s, 'path')).toBeGreaterThanOrEqual(4);
    const pct = texts(s).filter((x) => /^\d+%$/.test(x));
    expect(pct).toEqual(['45%', '25%', '10%', '20%']);
  });

  it('uses theme accent colours for automatic series colours', () => {
    const s = chartToSvg(charts.bar, { width: 300, height: 200, colors: { accent1: 'AA0000', accent2: '00AA00', accent3: '0000AA' } })!;
    expect(s).toContain('fill="#AA0000"');
    expect(s).toContain('fill="#00AA00"');
    expect(s).toContain('fill="#0000AA"');
  });

  it('returns null for a chart without cached values', () => {
    const xml = pieChart(['a'], [1]).replace(/<c:numCache>.*?<\/c:numCache>/s, '');
    expect(chartToSvg(xml, { width: 100, height: 100 })).toBeNull();
    expect(chartToSvg('<c:chartSpace/>', { width: 100, height: 100 })).toBeNull();
  });
});

describe('SVG → vector display list', () => {
  it('parses shapes, paths with arcs, transforms, text runs and opacity', () => {
    const mf = parseSvg(
      '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 400 200">' +
        '<style>.r{fill:#ff0000;stroke:blue}</style>' +
        '<rect class="r" x="10" y="10" width="100" height="50"/>' +
        '<g transform="translate(200 0)"><circle cx="50" cy="50" r="40" fill="green" fill-opacity="0.5"/></g>' +
        '<path d="M10 150 a40 40 0 0 1 80 0 z" fill="none" stroke="#000" stroke-width="4"/>' +
        '<text x="10" y="190" style="font: bold 20px Arial">Привет <tspan fill="red">мир</tspan></text>' +
        '</svg>',
    )!;
    expect(mf.widthPt).toBe(150);
    expect(mf.heightPt).toBe(75);
    const paths = mf.cmds.filter((c): c is Extract<MetaCmd, { k: 'path' }> => c.k === 'path');
    expect(paths.length).toBe(3);
    expect(paths[0]).toMatchObject({ fill: 'FF0000', stroke: '0000FF' });
    expect(paths[1].opacity).toBeCloseTo(0.5);
    expect(paths[2].d).toMatch(/C/); // the arc became cubic Béziers
    expect(paths[2].fill).toBeUndefined();
    const textCmds = mf.cmds.filter((c): c is Extract<MetaCmd, { k: 'text' }> => c.k === 'text');
    expect(textCmds.map((c) => c.text.trim())).toEqual(['Привет', 'мир']);
    expect(textCmds[0]).toMatchObject({ bold: true, family: 'sans' });
    expect(textCmds[1]).toMatchObject({ flow: true, color: 'FF0000' });
    // viewBox 400×200 drawn into 200×100: the rect at x=10 lands at 5
    expect(paths[0].d).toMatch(/^M\s*5[ ,]/);
  });

  it('rejects input that is not SVG', () => {
    expect(parseSvg('<html><body/></html>')).toBeNull();
    expect(parseSvg('not xml at all')).toBeNull();
  });
});

/** A DOCX (our writer's sample) with one inline chart appended. */
function docxWithChart(chartXml: string): Uint8Array {
  const files = unzipSync(docToDocx(sample).data);
  const dec = (p: string): string => new TextDecoder().decode(files[p]);
  const drawing =
    '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="900" name="Chart 1"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdChart1"/>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  const doc = dec('word/document.xml');
  const at = doc.lastIndexOf('<w:sectPr');
  files['word/document.xml'] = strToU8(doc.slice(0, at) + drawing + doc.slice(at));
  files['word/_rels/document.xml.rels'] = strToU8(
    dec('word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>'),
  );
  files['[Content_Types].xml'] = strToU8(
    dec('[Content_Types].xml').replace('</Types>', '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'),
  );
  files['word/charts/chart1.xml'] = strToU8(chartXml);
  return zipSync(files);
}

describe('charts in Office documents', () => {
  const chartXml = sampleCharts().bar;

  it('PPTX: a chart frame becomes a vector image of the frame size', () => {
    const { deck, warnings } = pptxToDeck(makePptx(chartXml));
    expect(warnings).toEqual([]);
    expect(deck.slides.length).toBe(5);
    const img = deck.slides[4].elements.find((e) => e.kind === 'image');
    expect(img).toMatchObject({ kind: 'image', mime: 'image/svg+xml', x: 60, y: 140, w: 600, h: 350 });
  });

  it('PPTX → PDF draws the chart as text and vectors (no raster images)', async () => {
    const r = await convert(makePptx(chartXml), 'pptx', 'pdf', env);
    const pdf = new TextDecoder('latin1').decode(r.data);
    if (hasCmd('pdftotext')) {
      // the hidden slide 3 is not printed, so the chart slide is the 4th page
      const pages = pdfToText(r.data).split('\f');
      const page5 = norm(pages[3]);
      for (const s of ['Слайд с диаграммой', 'Продажи, тыс. ₽', 'Прогноз', 'Янв', 'Апр', '190']) expect(page5).toContain(s);
    }
    // the base deck already has one photo; the chart must not add raster images
    const images = (s: string): number => s.match(/\/Subtype\s*\/Image/g)?.length ?? 0;
    const base = new TextDecoder('latin1').decode((await convert(makePptx(), 'pptx', 'pdf', env)).data);
    expect(images(pdf)).toBe(images(base));
  });

  it('DOCX: an inline chart becomes an image block sized from wp:extent', () => {
    const { doc, warnings } = docxToDoc(docxWithChart(chartXml));
    expect(warnings).toEqual([]);
    const img = doc.blocks.find((b): b is Extract<Block, { type: 'image' }> => b.type === 'image' && b.mime === 'image/svg+xml');
    expect(img).toBeDefined();
    // 5486400 × 3200400 EMU = 432 × 252 pt = 576 × 336 px
    expect(img).toMatchObject({ width: 576, height: 336, alt: 'Диаграмма' });
  });

  it('DOCX → PDF and HTML keep the chart', async () => {
    const docx = docxWithChart(chartXml);
    const pdf = await convert(docx, 'docx', 'pdf', env);
    if (hasCmd('pdftotext')) {
      const t = norm(pdfToText(pdf.data));
      for (const s of ['Продажи, тыс. ₽', 'План', 'Факт', 'Мар', '155']) expect(t).toContain(s);
    }
    const html = new TextDecoder().decode((await convert(docx, 'docx', 'html', env)).data);
    expect(html).toContain('data:image/svg+xml;base64,');
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftotext'))('LibreOffice reads the same chart values from both fixtures', () => {
    for (const [file, ext] of [
      [makePptx(chartXml), 'pptx'],
      [docxWithChart(chartXml), 'docx'],
    ] as const) {
      const t = norm(pdfToText(soffice(file, ext, 'pdf')));
      for (const s of ['Продажи', 'План', 'Факт', 'Прогноз', 'Янв', 'Апр', '190']) expect(t, ext).toContain(s);
    }
  });

  it.skipIf(!hasCmd('pdftoppm'))('the chart page is not blank when rasterised by poppler', async () => {
    const r = await convert(makePptx(chartXml), 'pptx', 'pdf', env);
    const page = renderPages(r.data, 30)[3];
    // PGM: header then 8-bit pixels; count dark-ish pixels
    const px = page.subarray(page.length - Math.floor(page.length * 0.95));
    let dark = 0;
    for (const v of px) if (v < 200) dark++;
    expect(dark / px.length).toBeGreaterThan(0.03);
  });
});
