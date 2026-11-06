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
import { GRADIENT_SLIDE, makePptx, PATTERN_SLIDE, RECT_GRADIENT_SLIDE } from './pptx-fixture';
import { PATTERNS } from '../../src/core/drawingml';

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

  it('rectangular path gradients: concentric rectangles from the focus (ECMA-376), drawn as clipped axial shadings', async () => {
    const pptx = makePptx(undefined, RECT_GRADIENT_SLIDE);
    const deck = pptxToDeck(pptx).deck;
    const shapes = deck.slides[4].elements.filter((e): e is Extract<SlideElement, { kind: 'shape' }> => e.kind === 'shape' && !!e.gradient);
    expect(shapes.map((e) => e.gradient!.kind)).toEqual(['rect', 'rect', 'rect']);
    expect(shapes[1].gradient!.coords).toEqual([0, 0, 300, 150, 60, 30, 120, 60]);
    const px = pixels((await convert(pptx, 'pptx', 'pdf', env)).data, 4);
    const at = (x: number, y: number): number[] => px(x, y);
    // shape 1 at (40,150) 300×150, focus in the centre: white there, dark at the edges, halfway in between
    expect(near(at(190, 225), hex('FFFFFF'), 8)).toBe(true);
    expect(near(at(42, 225), hex('1F4E79'), 14)).toBe(true);
    expect(near(at(190, 152), hex('1F4E79'), 14)).toBe(true);
    expect(near(at(115, 225), lerp('FFFFFF', '1F4E79', 0.5), 12)).toBe(true); // half way to the left edge
    expect(near(at(190, 187.5), lerp('FFFFFF', '1F4E79', 0.5), 12)).toBe(true); // half way to the top edge
    // shape 2 at (380,150): the focus rectangle 60..120 × 30..60 is white throughout
    for (const [x, y] of [[445, 185], [495, 205]]) expect(near(at(x, y), hex('FFFFFF'), 8), `${x},${y}`).toBe(true);
    expect(near(at(678, 295), hex('1F4E79'), 14)).toBe(true);
  });

  it('pattern fills: the preset bitmaps as tiling patterns, transparent backgrounds let content through', async () => {
    const pptx = makePptx(undefined, PATTERN_SLIDE);
    const shapes = pptxToDeck(pptx).deck.slides[4].elements.filter((e): e is Extract<SlideElement, { kind: 'shape' }> => e.kind === 'shape' && !!e.pattern);
    expect(shapes.map((e) => e.pattern!.preset)).toEqual(['ltDnDiag', 'dkHorz', 'smCheck', 'cross', 'pct50']);
    expect(shapes[4].pattern!.bgClear).toBe(true);
    const pdf = (await convert(pptx, 'pptx', 'pdf', env)).data;
    expect(await pdfObjects(pdf)).toMatch(/\/PatternType 1/);
    // at 96 dpi one pattern pixel (0.75 pt) is one device pixel: count the foreground share
    const buf = withTemp((dir) => {
      writeFileSync(join(dir, 'in.pdf'), pdf);
      execFileSync('pdftoppm', ['-r', '96', '-f', '4', '-l', '4', join(dir, 'in.pdf'), join(dir, 'p')]);
      return readFileSync(join(dir, readdirSync(dir).find((n) => n.endsWith('.ppm'))!));
    });
    const head = /^P6\s+(\d+)\s+(\d+)\s+255\s/.exec(new TextDecoder('latin1').decode(buf.subarray(0, 30)))!;
    const w = +head[1];
    const px = (x: number, y: number): number[] => [...buf.subarray(head[0].length + (y * w + x) * 3, head[0].length + (y * w + x) * 3 + 3)];
    const share = (x0: number, y0: number, test: (p: number[]) => boolean): number => {
      let n = 0;
      // 64 × 64 device pixels = 8 × 8 cells
      for (let y = 0; y < 64; y++) for (let x = 0; x < 64; x++) if (test(px(Math.round(x0 * (96 / 72)) + x, Math.round(y0 * (96 / 72)) + y))) n++;
      return n / 4096;
    };
    const red = (p: number[]): boolean => p[0] > 150 && p[1] < 80 && p[2] < 80;
    const expected = (prst: string): number => PATTERNS[prst].reduce((n, row) => n + row.toString(2).split('1').length - 1, 0) / 64;
    ['ltDnDiag', 'dkHorz', 'smCheck', 'cross'].forEach((prst, i) => expect(share(50 + i * 165, 160, red), prst).toBeCloseTo(expected(prst), 1));
    // transparent background: blue dots on the yellow box underneath, never white
    expect(share(300, 360, (p) => p.every((v) => v > 235))).toBe(0);
    expect(share(300, 360, (p) => p[2] > p[0])).toBeCloseTo(0.5, 1);
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

  it('SVG: text filled with a gradient is clipped to the glyphs and painted', async () => {
    const svg =
      '<svg xmlns="http://www.w3.org/2000/svg" width="400" height="100" viewBox="0 0 400 100">' +
      '<defs><linearGradient id="t" gradientUnits="userSpaceOnUse" x1="20" y1="0" x2="380" y2="0"><stop offset="0" stop-color="#ff0000"/><stop offset="1" stop-color="#0000ff"/></linearGradient></defs>' +
      '<text x="20" y="80" font-family="Arial" font-size="80" font-weight="bold" fill="url(#t)">MMMMMM</text></svg>';
    const mf = parseSvg(svg)!;
    const t = mf.cmds.find((c): c is Extract<MetaCmd, { k: 'text' }> => c.k === 'text')!;
    expect(t.grad?.kind).toBe('linear');
    const pdf = (await docToPdf({ blocks: [{ type: 'image', data: new TextEncoder().encode(svg), mime: 'image/svg+xml', width: 400, height: 100 }] }, { fonts: loadFonts })).data;
    const px = pixels(pdf, 1);
    // the text: colored pixels, red on the left, blue on the right; between glyph strokes stays white
    const coloured: { x: number; c: number[] }[] = [];
    for (let y = 0; y < 300; y++)
      for (let x = 0; x < 595; x++) {
        const c = px(x, y);
        if (c[0] + c[1] + c[2] < 600 && Math.abs(c[0] - c[2]) > 60) coloured.push({ x, c });
      }
    expect(coloured.length).toBeGreaterThan(500);
    const xs = coloured.map((p) => p.x);
    const lo = Math.min(...xs);
    const hi = Math.max(...xs);
    const left = coloured.filter((p) => p.x < lo + (hi - lo) * 0.15);
    const right = coloured.filter((p) => p.x > hi - (hi - lo) * 0.15);
    expect(left.every((p) => p.c[0] > p.c[2]), 'red on the left').toBe(true);
    expect(right.every((p) => p.c[2] > p.c[0]), 'blue on the right').toBe(true);
    // clipped to the glyphs, not a filled box: white gaps between the strokes
    const ys: number[] = [];
    for (let y = 0; y < 300; y++) for (let x = lo; x <= hi; x++) if (coloured.length && px(x, y).every((v) => v < 250) && Math.abs(px(x, y)[0] - px(x, y)[2]) > 60) ys.push(y);
    const [y0, y1] = [Math.min(...ys), Math.max(...ys)];
    let white = 0;
    let all = 0;
    for (let y = y0; y <= y1; y += 2)
      for (let x = lo; x <= hi; x += 2) {
        all++;
        if (px(x, y).every((v) => v > 240)) white++;
      }
    expect(white / all).toBeGreaterThan(0.2);
  });

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
