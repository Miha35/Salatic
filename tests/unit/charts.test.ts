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
import { barChart, moreCharts, pieChart, sampleCharts, SAMPLE_CATS } from './chart-fixture';
import { makePptx } from './pptx-fixture';
import { sample } from './sample';
import { docToHtmlFragment } from '../../src/core/writers/text';
import { htmlToDoc } from '../../src/core/readers/html';
import { env, hasCmd, pdfToText, renderPages, soffice, withTemp } from './helpers';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { metafileToEmf } from '../../src/core/writers/emf';
import { parseEmf, type Metafile } from '../../src/core/metafile';
import { decodePng, encodePng } from '../../src/core/png';
import { sniffImageMime } from '../../src/core/util';

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

const SVG_PICTURE =
  '<svg xmlns="http://www.w3.org/2000/svg" width="300" height="150" viewBox="0 0 300 150">' +
  '<rect x="5" y="5" width="290" height="140" rx="12" fill="#E8F0FE" stroke="#1A73E8" stroke-width="3"/>' +
  '<circle cx="60" cy="75" r="40" fill="#34A853" fill-opacity="0.6"/>' +
  '<text x="120" y="70" font-family="Arial" font-size="22" font-weight="bold" fill="#202124">Схема</text>' +
  '<text x="120" y="105" font-family="Arial" font-size="16" fill="#5F6368">вектор <tspan fill="#EA4335">SVG</tspan></text>' +
  '</svg>';

describe('EMF writer', () => {
  it('round-trips a chart through EMF and our EMF reader', () => {
    for (const [k, xml] of Object.entries(sampleCharts())) {
      const mf = parseSvg(chartToSvg(xml, { width: 432, height: 252 })!)!;
      const emf = metafileToEmf(mf);
      expect(sniffImageMime(emf), k).toBe('image/emf');
      const back = parseEmf(emf);
      expect(back.widthPt, k).toBeCloseTo(432, 0);
      expect(back.heightPt, k).toBeCloseTo(252, 0);
      // translucent fills come back as alpha-blended bitmaps, everything else as paths
      const translucent = mf.cmds.filter((c) => c.k === 'path' && c.fill && (c.opacity ?? 1) < 0.999).length;
      const count = (m: Metafile, t: string): number => m.cmds.filter((c) => c.k === t).length;
      expect([count(back, 'path') + count(back, 'image'), count(back, 'image'), count(back, 'text')], k).toEqual([count(mf, 'path'), translucent, count(mf, 'text')]);
      const txt = (m: Metafile): Extract<MetaCmd, { k: 'text' }>[] => m.cmds.filter((c): c is Extract<MetaCmd, { k: 'text' }> => c.k === 'text');
      const [a, b] = [txt(back), txt(mf)];
      expect(a.map((c) => c.text), k).toEqual(b.map((c) => c.text));
      // positions survive within the 1/16-unit EMF grid
      a.forEach((c, i) => expect(Math.hypot(c.x - b[i].x, c.y - b[i].y), `${k}: ${c.text}`).toBeLessThan(0.1));
    }
  });

  it('keeps tspan runs flowing, colours, bold and opacity (as real transparency)', () => {
    const back = parseEmf(metafileToEmf(parseSvg(SVG_PICTURE)!));
    const texts = back.cmds.filter((c): c is Extract<MetaCmd, { k: 'text' }> => c.k === 'text');
    expect(texts.map((t) => [t.text.trim(), t.bold, t.color, !!t.flow])).toEqual([
      ['Схема', true, '202124', false],
      ['вектор', false, '5F6368', false],
      ['SVG', false, 'EA4335', true],
    ]);
    // the 60 % green circle: an alpha bitmap with its own colour, not a colour mixed with white
    const img = back.cmds.find((c): c is Extract<MetaCmd, { k: 'image' }> => c.k === 'image')!;
    const dec = decodePng(img.png)!;
    const o = ((dec.height >> 1) * dec.width + (dec.width >> 1)) * 4;
    const [r, g, b, a] = dec.rgba.subarray(o, o + 4);
    expect(Math.abs(a - 153)).toBeLessThan(4);
    expect([r, g, b].every((v, i) => Math.abs(v - [0x34, 0xa8, 0x53][i]) < 6)).toBe(true);
  });

  it('embeds raster images as bitmaps', () => {
    const px = new Uint8Array([255, 0, 0, 255, 0, 0, 255, 128]);
    const mf: Metafile = { width: 20, height: 10, widthPt: 15, heightPt: 7.5, cmds: [{ k: 'image', x: 0, y: 0, w: 20, h: 10, png: encodePng(px, 2, 1, 4) }] };
    const back = parseEmf(metafileToEmf(mf));
    const img = back.cmds.find((c) => c.k === 'image');
    expect(img).toMatchObject({ x: 0, y: 0, w: 20, h: 10 });
    const dec = decodePng((img as Extract<MetaCmd, { k: 'image' }>).png)!;
    expect([...dec.rgba.subarray(0, 3)]).toEqual([255, 0, 0]);
    // the half-transparent blue pixel is blended with white
    expect([...dec.rgba.subarray(4, 7)]).toEqual([127, 127, 255]);
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftotext'))('LibreOffice draws the EMF with its text', () => {
    const emf = metafileToEmf(parseSvg(chartToSvg(sampleCharts().combo, { width: 432, height: 252 })!)!);
    const t = norm(pdfToText(soffice(emf, 'emf', 'pdf')));
    for (const s of ['Выручка и маржа', '1 500 000', '30%', 'Апр']) expect(t).toContain(s);
  });
});

describe('PNG decoder', () => {
  it('round-trips our encoder for gray, RGB and RGBA', () => {
    const w = 7;
    const h = 5;
    for (const ch of [1, 3, 4] as const) {
      const px = Uint8Array.from({ length: w * h * ch }, (_, i) => (i * 37 + ch) & 0xff);
      const d = decodePng(encodePng(px, w, h, ch))!;
      expect(d.width).toBe(w);
      for (let i = 0; i < w * h; i++) {
        const want = ch === 1 ? [px[i], px[i], px[i], 255] : ch === 3 ? [px[i * 3], px[i * 3 + 1], px[i * 3 + 2], 255] : [...px.subarray(i * 4, i * 4 + 4)];
        expect([...d.rgba.subarray(i * 4, i * 4 + 4)]).toEqual(want);
      }
    }
    expect(decodePng(new Uint8Array([1, 2, 3]))).toBeNull();
  });

  it.skipIf(!hasCmd('python3'))('decodes palette, 16-bit, low bit depth and interlaced PNGs made by Pillow', () => {
    const script = [
      'import sys',
      'from PIL import Image',
      'd = sys.argv[1]',
      'im = Image.new("RGBA", (13, 9))',
      'for y in range(9):',
      '    for x in range(13):',
      '        im.putpixel((x, y), ((x * 19) % 256, (y * 28) % 256, ((x + y) * 11) % 256, 255 if (x + y) % 3 else 0))',
      'im.save(d + "/rgba_interlaced.png", interlace=1)',
      'im.convert("RGB").quantize(16).save(d + "/palette.png")',
      'g = im.convert("L")',
      'g.save(d + "/gray.png")',
      'g.convert("1").save(d + "/mono.png")',
      'g16 = Image.new("I;16", g.size)',
      'g16.putdata([v * 257 for v in g.getdata()])',
      'g16.save(d + "/gray16.png")',
      'im.convert("RGB").save(d + "/ref.ppm")',
      'im.convert("RGB").quantize(16).convert("RGB").save(d + "/pal.ppm")',
      'g.save(d + "/gray.pgm")',
      'g.convert("1").convert("L").save(d + "/mono.pgm")',
    ].join('\n');
    withTemp((dir) => {
      try {
        execFileSync('python3', ['-I', '-c', script, dir], { stdio: 'ignore' });
      } catch {
        return; // Pillow not installed
      }
      const ppm = (name: string): Uint8Array => {
        const b = readFileSync(join(dir, name));
        // P5/P6 header: magic, width, height, maxval, then one whitespace byte
        let fields = 0;
        let i = 0;
        while (fields < 4) {
          while (/\s/.test(String.fromCharCode(b[i]))) i++;
          while (!/\s/.test(String.fromCharCode(b[i]))) i++;
          fields++;
        }
        return new Uint8Array(b.subarray(i + 1));
      };
      const rgb = (p: Uint8Array): number[] => [...p].filter((_, i) => i % 4 !== 3);
      const gray = (p: Uint8Array): number[] => [...p].filter((_, i) => i % 4 === 0);
      const png = (name: string) => decodePng(new Uint8Array(readFileSync(join(dir, name))))!;
      const inter = png('rgba_interlaced.png');
      expect(rgb(inter.rgba)).toEqual([...ppm('ref.ppm')]);
      expect([...inter.rgba].filter((_, i) => i % 4 === 3).slice(0, 4)).toEqual([0, 255, 255, 0]);
      expect(rgb(png('palette.png').rgba)).toEqual([...ppm('pal.ppm')]);
      expect(gray(png('gray.png').rgba)).toEqual([...ppm('gray.pgm')]);
      expect(gray(png('mono.png').rgba)).toEqual([...ppm('mono.pgm')]);
      // 16-bit samples keep their high byte
      expect(gray(png('gray16.png').rgba)).toEqual([...ppm('gray.pgm')]);
    });
  });
});

describe('charts and SVG in written documents', () => {
  const chartXml = sampleCharts().bar;
  const files = (zip: Uint8Array): Record<string, string> => Object.fromEntries(Object.entries(unzipSync(zip)).map(([k, v]) => [k, new TextDecoder().decode(v)]));

  it('DOCX → DOCX keeps a native chart part, without the embedded-workbook link, and the theme colours', () => {
    const withLink = chartXml.replace('</c:chartSpace>', '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>');
    const src = docxWithChart(withLink);
    const { doc } = docxToDoc(src);
    const out = docToDocx(doc).data;
    const f = files(out);
    expect(f['word/charts/chart1.xml']).toContain('<c:barChart>');
    expect(f['word/charts/chart1.xml']).not.toContain('externalData');
    expect(f['[Content_Types].xml']).toContain('/word/charts/chart1.xml');
    expect(f['word/_rels/document.xml.rels']).toMatch(/relationships\/chart" Target="charts\/chart1\.xml"/);
    expect(f['word/document.xml']).toContain('<c:chart xmlns:c=');
    expect(f['word/theme/theme1.xml']).toContain('<a:accent1><a:srgbClr val="4472C4"/>');
    // and reads back as the same chart
    const again = docxToDoc(out).doc.blocks.find((b): b is Extract<Block, { type: 'image' }> => b.type === 'image' && !!b.chart);
    expect(again?.chart?.xml).toContain('<c:barChart>');
    expect(again?.mime).toBe('image/svg+xml');
  });

  it('the document editor round trip (DOCX → HTML → DOCX) keeps the native chart', () => {
    const { doc } = docxToDoc(docxWithChart(chartXml));
    const html = docToHtmlFragment(doc, { sources: true });
    expect(html).toContain('data-chart=');
    expect(docToHtmlFragment(doc)).not.toContain('data-chart=');
    // the browser serialises the edited DOM; the worker reads it back
    const edited = htmlToDoc(html.replace('</p>', ' (правка)</p>'));
    const img = edited.blocks.find((b): b is Extract<Block, { type: 'image' }> => b.type === 'image' && !!b.chart);
    expect(img?.chart?.xml).toBe(chartXml);
    const original = doc.blocks.find((b): b is Extract<Block, { type: 'image' }> => b.type === 'image' && !!b.chart)!;
    expect(img?.chart?.colors).toEqual(original.chart!.colors);
    expect(files(docToDocx(edited).data)['word/charts/chart1.xml']).toContain('<c:barChart>');
    // forged attributes are ignored
    const forged = htmlToDoc(html.replace(/data-chart="[^"]*"/, 'data-chart="{&quot;xml&quot;:&quot;&lt;html/&gt;&quot;}"'));
    expect(forged.blocks.some((b) => b.type === 'image' && b.chart)).toBe(false);
  });

  it('PPTX → DOCX turns slide charts into Word charts with the deck theme', async () => {
    const out = (await convert(makePptx(chartXml), 'pptx', 'docx', env)).data;
    const f = files(out);
    expect(Object.keys(f).filter((k) => k.startsWith('word/charts/'))).toEqual(['word/charts/chart1.xml']);
    expect(f['word/theme/theme1.xml']).toMatch(/<a:accent1><a:srgbClr val="[0-9A-F]{6}"\/>/);
  });

  it('a chart that references other parts falls back to a vector picture', () => {
    const xml = chartXml.replace('</c:chartSpace>', '<c:userShapes r:id="rId9"/><c:printSettings><c:legacyDrawingHF r:id="rId3"/></c:printSettings></c:chartSpace>');
    const doc = { blocks: [{ type: 'image' as const, data: new TextEncoder().encode(chartToSvg(xml, { width: 300, height: 200 })!), mime: 'image/svg+xml', width: 400, height: 267, chart: { xml } }] };
    const f = files(docToDocx(doc).data);
    expect(Object.keys(f).some((k) => k.startsWith('word/charts/'))).toBe(false);
    expect(Object.keys(f)).toEqual(expect.arrayContaining(['word/media/image1.emf', 'word/media/image1.svg']));
  });

  it('SVG pictures go to DOCX as EMF with an SVG extension, and read back as SVG', () => {
    const doc = { blocks: [{ type: 'paragraph' as const, runs: [{ text: 'До рисунка' }] }, { type: 'image' as const, data: new TextEncoder().encode(SVG_PICTURE), mime: 'image/svg+xml', width: 300, height: 150, alt: 'Схема' }] };
    const out = docToDocx(doc).data;
    const f = files(out);
    expect(f['word/document.xml']).toMatch(/<a:blip r:embed="rId\d+"><a:extLst><a:ext uri="\{96DAC541-7B7A-43D3-8B79-37D633B846F1\}"><asvg:svgBlip [^>]*r:embed="rId\d+"\/>/);
    expect(f['[Content_Types].xml']).toContain('Extension="svg"');
    const img = docxToDoc(out).doc.blocks.find((b) => b.type === 'image');
    expect(img).toMatchObject({ mime: 'image/svg+xml', width: 300, height: 150, alt: 'Схема' });
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftotext'))('LibreOffice shows charts and SVG pictures from our DOCX and RTF', async () => {
    const chartDocx = docToDocx(docxToDoc(docxWithChart(chartXml)).doc).data;
    let t = norm(pdfToText(soffice(chartDocx, 'docx', 'pdf')));
    for (const s of ['Продажи', 'План', 'Факт', 'Прогноз', 'Янв', 'Апр', '190']) expect(t, 'docx chart').toContain(s);
    const doc = { blocks: [{ type: 'paragraph' as const, runs: [{ text: 'До рисунка' }] }, { type: 'image' as const, data: new TextEncoder().encode(SVG_PICTURE), mime: 'image/svg+xml', width: 300, height: 150 }] };
    t = norm(pdfToText(soffice(docToDocx(doc).data, 'docx', 'pdf')));
    for (const s of ['До рисунка', 'Схема', 'вектор', 'SVG']) expect(t, 'docx svg').toContain(s);
    const rtf = (await convert(docxWithChart(chartXml), 'docx', 'rtf', env)).data;
    t = norm(pdfToText(soffice(rtf, 'rtf', 'pdf')));
    for (const s of ['Продажи', 'Прогноз', 'Апр', '190']) expect(t, 'rtf chart').toContain(s);
  });
});

describe('more chart kinds', () => {
  const charts = moreCharts();
  const svgOf = (k: string): string => chartToSvg(charts[k], { width: 432, height: 252 })!;

  it('every kind renders to parseable SVG', () => {
    for (const k of Object.keys(charts)) {
      const s = svgOf(k);
      expect(s, k).toMatch(/^<svg /);
      expect(parseSvg(s)!.cmds.length, k).toBeGreaterThan(10);
    }
  });

  it('radar: one spoke per category, one polygon per series, category labels around', () => {
    const s = svgOf('radar');
    expect(s.match(/<line [^>]*stroke="#BFBFBF"/g)?.length).toBe(6);
    expect(s.match(/<path d="M[^"]+Z" fill="none" stroke="#(4472C4|ED7D31)"/g)?.length).toBe(2);
    expect(texts(s)).toEqual(expect.arrayContaining(['Скорость', 'Гарантия', 'Модель А']));
    expect(svgOf('radarFilled')).toMatch(/fill="#4472C4" fill-opacity/);
  });

  it('stock: high-low lines, white up bars and black down bars, only visible series in the legend', () => {
    const s = svgOf('stock');
    expect(s.match(/<line [^>]*stroke="#404040"/g)?.length).toBe(5);
    // closes below opens on days 1 and 4
    expect(s.match(/<rect [^>]*fill="#000000"/g)?.length).toBe(2);
    expect(s.match(/<rect [^>]*fill="#FFFFFF"/g)?.length).toBe(3);
    expect(texts(s)).toContain('Закрытие');
    expect(texts(s)).not.toContain('Открытие');
  });

  it('bubble: bubble area proportional to size', () => {
    const r = [...svgOf('bubble').matchAll(/<circle [^>]*r="([\d.]+)" fill="#4472C4"/g)].map((m) => +m[1]);
    expect(r.length).toBe(4);
    // sizes 40 and 10 of the same series: area ×4, radius ×2
    expect(Math.max(...r) / Math.min(...r.filter((x) => x > Math.min(...r)))).toBeCloseTo(2, 1);
  });

  it('3-D columns and bars: front, top and side faces; pie with walls', () => {
    const col = svgOf('column3d');
    expect(col.match(/<rect /g)!.length).toBeGreaterThanOrEqual(8 + 2); // 8 bars, 2 legend keys
    expect(col).toMatch(/fill="#5B85CF"|fill="#[0-9A-F]{6}" stroke="#[0-9A-F]{6}" stroke-width="0.4"/);
    expect(svgOf('cylinder3d')).toMatch(/<ellipse /);
    const pie = svgOf('pie3d');
    expect(pie.match(/<path [^>]*stroke="#FFFFFF"/g)?.length).toBe(4);
    expect(pie.match(/stroke-width="0.5"/g)!.length).toBeGreaterThanOrEqual(2); // walls of the front slices
  });

  it('surface: value bands in the legend, contour and 3-D views', () => {
    expect(texts(svgOf('contour'))).toEqual(expect.arrayContaining(['0–2', '8–10']));
    expect(svgOf('surface3d').match(/<line [^>]*stroke="#404040" stroke-width="0.35"/g)!.length).toBeGreaterThan(20);
  });

  it('pie of pie and bar of pie: the last points move to the second plot', () => {
    const pie = svgOf('pieOfPie');
    expect(pie.match(/<line [^>]*stroke="#BFBFBF"/g)?.length).toBe(2); // connectors
    expect(texts(pie)).toEqual(expect.arrayContaining(['40%', '8%', '7%', '5%']));
    const bar = svgOf('barOfPie');
    expect(bar.match(/<rect [^>]*stroke="#FFFFFF"/g)?.length).toBe(3);
  });

  it.skipIf(!hasCmd('pdftotext'))('PPTX → PDF keeps the labels of the new kinds', async () => {
    for (const k of ['radar', 'stock', 'bubble', 'column3d', 'pie3d', 'pieOfPie']) {
      const r = await convert(makePptx(charts[k]), 'pptx', 'pdf', env);
      const page = norm(pdfToText(r.data).split('\f')[3]);
      const title = /<a:t>([^<]+)<\/a:t>/.exec(charts[k])![1];
      expect(page, k).toContain(title);
    }
  });
});
