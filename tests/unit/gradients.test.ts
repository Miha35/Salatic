/** Gradient fills: DrawingML and SVG gradients → model → PDF shadings, checked against LibreOffice and by pixel. */
import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pptxToDeck } from '../../src/core/readers/pptx';
import { parseSvg } from '../../src/core/svg-render';
import { docToPdf } from '../../src/core/writers/pdf';
import { convert } from '../../src/core/convert';
import { gradientAverage, type Gradient } from '../../src/core/model';
import type { MetaCmd } from '../../src/core/metafile';
import type { SlideElement } from '../../src/core/slides';
import { env, hasCmd, loadFonts, soffice, withTemp } from './helpers';
import { PDFDocument, PDFRawStream } from 'pdf-lib';
import { GRADIENT_SLIDE, makePptx } from './pptx-fixture';

/** Rasterise one page at 72 dpi (1 px = 1 pt) and return an RGB sampler. */
function pixels(pdf: Uint8Array, page: number): (x: number, y: number) => [number, number, number] {
  const buf = withTemp((dir) => {
    writeFileSync(join(dir, 'in.pdf'), pdf);
    execFileSync('pdftoppm', ['-r', '72', '-f', String(page), '-l', String(page), join(dir, 'in.pdf'), join(dir, 'p')]);
    return readFileSync(join(dir, readdirSync(dir).find((n) => n.endsWith('.ppm'))!));
  });
  let i = 0;
  const fields: number[] = [];
  while (fields.length < 3) {
    while (/\s/.test(String.fromCharCode(buf[i]))) i++;
    if (buf[i] === 0x50) {
      i += 2;
      continue;
    }
    let n = '';
    while (!/\s/.test(String.fromCharCode(buf[i]))) n += String.fromCharCode(buf[i++]);
    fields.push(+n);
  }
  const [w] = fields;
  const data = buf.subarray(i + 1);
  return (x, y) => {
    const o = (Math.round(y) * w + Math.round(x)) * 3;
    return [data[o], data[o + 1], data[o + 2]];
  };
}

const hex = (c: string): [number, number, number] => [0, 2, 4].map((o) => parseInt(c.slice(o, o + 2), 16)) as [number, number, number];
const near = (a: number[], b: number[], tol: number): boolean => a.every((v, i) => Math.abs(v - b[i]) <= tol);
const lerp = (a: string, b: string, t: number): number[] => hex(a).map((v, i) => v + (hex(b)[i] - v) * t);
/** All indirect objects of a PDF as text (object streams decoded by pdf-lib). */
async function pdfObjects(b: Uint8Array): Promise<string> {
  const doc = await PDFDocument.load(b);
  return doc.context
    .enumerateIndirectObjects()
    .map(([, o]) => (o instanceof PDFRawStream ? o.dict.toString() : o.toString()))
    .join('\n');
}

describe('DrawingML gradients', () => {
  const { deck } = pptxToDeck(makePptx(undefined, GRADIENT_SLIDE));
  const slide = deck.slides[4];
  type Shape = Extract<SlideElement, { kind: 'shape' }>;
  /** The fixture's shapes by position: 1 linear rectangle, 2 radial ellipse, 3 theme-styled rounded rectangle. */
  const shape = (n: 1 | 2 | 3): Shape => {
    const at = { 1: [40, 150], 2: [380, 150], 3: [40, 340] }[n];
    return slide.elements.find((e): e is Shape => e.kind === 'shape' && e.x === at[0] && e.y === at[1])!;
  };

  it('reads the background gradient across the slide', () => {
    const g = slide.background?.gradient as Gradient;
    expect(g.kind).toBe('linear');
    // 90° (5400000): top to bottom over the full slide height
    expect(g.coords.map((v) => Math.round(v))).toEqual([360, 0, 360, 540]);
    expect(g.stops.map((s) => s.color)).toEqual(['FFFFFF', 'BDD7EE']);
    expect(slide.background?.color).toBe(gradientAverage(g));
  });

  it('a scaled 45° linear gradient runs corner to corner of the shape', () => {
    const g = shape(1).gradient!;
    expect(g.kind).toBe('linear');
    const [x1, y1, x2, y2] = g.coords;
    // the isolines run along the other diagonal: the start and end project onto (0,0) and (300,150)
    const d = [x2 - x1, y2 - y1];
    const t = (x: number, y: number): number => ((x - x1) * d[0] + (y - y1) * d[1]) / (d[0] ** 2 + d[1] ** 2);
    expect(t(0, 0)).toBeCloseTo(0, 5);
    expect(t(300, 150)).toBeCloseTo(1, 5);
    expect(t(300, 0)).toBeCloseTo(t(0, 150), 5);
    expect(g.stops.map((s) => [s.pos, s.color])).toEqual([
      [0, 'C00000'],
      [0.5, 'FFC000'],
      [1, '0070C0'],
    ]);
  });

  it('a circle path gradient is a circle from the focus to the farthest corner', () => {
    const g = shape(2).gradient!;
    expect(g.kind).toBe('radial');
    expect(g.coords[0]).toBe(150);
    expect(g.coords[1]).toBe(75);
    expect(g.coords[5]).toBeCloseTo(Math.hypot(150, 75), 5);
  });

  it('a theme gradient (fillRef) is tinted from the reference colour', () => {
    const g = shape(3).gradient!;
    expect(g).toBeDefined();
    // accent2 ED7D31 · lumMod 110 % · satMod 105 % · tint 67 % (in linear RGB)
    expect(near(hex(g.stops[0].color), hex('F6BDA4'), 3)).toBe(true);
  });
});

describe.skipIf(!hasCmd('pdftoppm'))('gradients in PDF', () => {
  it('slides: shadings, and colours that match LibreOffice', async () => {
    const pptx = makePptx(undefined, GRADIENT_SLIDE);
    const ours = (await convert(pptx, 'pptx', 'pdf', env)).data;
    const text = await pdfObjects(ours);
    expect(text).toMatch(/\/ShadingType 2/);
    expect(text).toMatch(/\/ShadingType 3/);
    expect(text).toMatch(/\/FunctionType 3/);
    const page = 4; // the hidden slide is not printed
    const px = pixels(ours, page);
    // linear rectangle at (40,150) 300×150: corners and centre
    expect(near(px(42, 152), hex('C00000'), 20)).toBe(true);
    expect(near(px(190, 225), hex('FFC000'), 12)).toBe(true);
    expect(near(px(338, 298), hex('0070C0'), 20)).toBe(true);
    // radial ellipse centred at (530,225): white centre, green towards the ends
    expect(near(px(530, 225), hex('FFFFFF'), 8)).toBe(true);
    expect(near(px(392, 225), lerp('FFFFFF', '00B050', 138 / Math.hypot(150, 75)), 10)).toBe(true);
    // background: white at the top, BDD7EE at the bottom
    expect(near(px(700, 3), hex('FFFFFF'), 6)).toBe(true);
    expect(near(px(700, 500), lerp('FFFFFF', 'BDD7EE', 500 / 540), 6)).toBe(true);
    if (!hasCmd('soffice')) return;
    const lo = pixels(soffice(pptx, 'pptx', 'pdf'), page);
    const points: [number, number][] = [
      [60, 170],
      [190, 225],
      [320, 280],
      [530, 225],
      [420, 225],
      [530, 165],
      [80, 360],
      [360, 390],
      [600, 430],
      [700, 100],
      [700, 450],
    ];
    for (const [x, y] of points) expect(near(px(x, y), lo(x, y), 10), `(${x},${y}) ours ${px(x, y)} vs LibreOffice ${lo(x, y)}`).toBe(true);
  });

  const SVG =
    '<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200" viewBox="0 0 400 200">' +
    '<defs><linearGradient id="a"><stop offset="0" stop-color="#ff0000"/><stop offset="1" stop-color="#0000ff"/></linearGradient>' +
    '<linearGradient id="v" href="#a" x2="0" y2="1"/>' +
    '<radialGradient id="r" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="#fff"/><stop offset="1" stop-color="#000"/></radialGradient>' +
    '<linearGradient id="fade" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="400" y2="0"><stop offset="0" stop-color="#000"/><stop offset="1" stop-color="#000" stop-opacity="0"/></linearGradient></defs>' +
    '<rect x="0" y="0" width="200" height="100" fill="url(#a)"/>' +
    '<rect x="200" y="0" width="100" height="100" fill="url(#v)"/>' +
    '<circle cx="350" cy="50" r="50" fill="url(#r)"/>' +
    '<rect x="0" y="120" width="400" height="80" fill="url(#fade)"/>' +
    '</svg>';

  it('SVG: bounding-box and user-space gradients, href templates, matrices', () => {
    const mf = parseSvg(SVG)!;
    const paths = mf.cmds.filter((c): c is Extract<MetaCmd, { k: 'path' }> => c.k === 'path');
    const [a, v, r, fade] = paths.map((p) => p.grad!);
    // objectBoundingBox: unit square mapped onto the rect
    expect(a.matrix).toEqual([200, 0, 0, 100, 0, 0]);
    expect(a.coords).toEqual([0, 0, 1, 0]);
    // the template's stops, its own direction
    expect(v.coords).toEqual([0, 0, 0, 1]);
    expect(v.stops.map((s) => s.color)).toEqual(['FF0000', '0000FF']);
    expect(r.kind).toBe('radial');
    expect(r.coords).toEqual([0.5, 0.5, 0, 0.5, 0.5, 0.5]);
    expect(fade.coords).toEqual([0, 0, 400, 0]);
    expect(fade.stops[1].alpha).toBe(0);
    expect(paths[0].fill).toBe(gradientAverage(a));
  });

  it('SVG: drawn with smooth colours and varying opacity (soft mask)', async () => {
    const pdf = (await docToPdf({ blocks: [{ type: 'image', data: new TextEncoder().encode(SVG), mime: 'image/svg+xml', width: 400, height: 200 }] }, { fonts: loadFonts })).data;
    expect(await pdfObjects(pdf)).toMatch(/\/S \/Luminosity/);
    const px = pixels(pdf, 1);
    // the picture's top-left corner on the page: find the first red pixel on a scan line
    let x0 = 0;
    let y0 = 0;
    search: for (let y = 0; y < 400; y++)
      for (let x = 0; x < 400; x++) {
        const [rr, gg, bb] = px(x, y);
        if (rr > 200 && gg < 60 && bb < 60) {
          x0 = x;
          y0 = y;
          break search;
        }
      }
    const k = 0.75; // 400 px = 300 pt
    const at = (x: number, y: number): [number, number, number] => px(x0 + x * k, y0 + y * k);
    expect(near(at(100, 50), lerp('FF0000', '0000FF', 0.5), 14)).toBe(true);
    expect(near(at(250, 90), lerp('FF0000', '0000FF', 0.9), 14)).toBe(true);
    expect(near(at(350, 50), [255, 255, 255], 12)).toBe(true);
    expect(near(at(350, 95), [26, 26, 26], 20)).toBe(true);
    // black fading to transparent over white: dark on the left, light on the right
    expect(near(at(20, 160), lerp('000000', 'FFFFFF', 0.05), 16)).toBe(true);
    expect(near(at(200, 160), lerp('000000', 'FFFFFF', 0.5), 16)).toBe(true);
    expect(near(at(380, 160), lerp('000000', 'FFFFFF', 0.95), 16)).toBe(true);
  });
});
