/**
 * EMF+ (GDI+) pictures: files with EMF+ records only are drawn from them; dual files keep their GDI
 * records. Fixtures are assembled record by record (MS-EMFPLUS); LibreOffice, which has its own EMF+
 * renderer, draws the same files as an independent check.
 */
import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseEmf, type MetaCmd } from '../../src/core/metafile';
import { docToPdf } from '../../src/core/writers/pdf';
import { encodePng } from '../../src/core/png';
import { hasCmd, loadFonts, pdfToText, soffice, withTemp } from './helpers';

/** Little-endian byte writer. */
class Bytes {
  b: number[] = [];
  u16(v: number): this {
    this.b.push(v & 255, (v >>> 8) & 255);
    return this;
  }
  u32(v: number): this {
    return this.u16(v & 0xffff).u16((v >>> 16) & 0xffff);
  }
  i16(v: number): this {
    return this.u16(v < 0 ? v + 0x10000 : v);
  }
  f32(v: number): this {
    const a = new Uint8Array(new Float32Array([v]).buffer);
    this.b.push(...a);
    return this;
  }
  str16(s: string): this {
    for (const ch of s) this.u16(ch.charCodeAt(0));
    return this;
  }
  raw(a: Uint8Array): this {
    this.b.push(...a);
    return this;
  }
  pad(): this {
    while (this.b.length % 4) this.b.push(0);
    return this;
  }
}

/** One EMF+ record. */
const rec = (type: number, flags: number, data: (w: Bytes) => void): Uint8Array => {
  const d = new Bytes();
  data(d);
  d.pad();
  return Uint8Array.from(new Bytes().u16(type).u16(flags).u32(12 + d.b.length).u32(d.b.length).raw(Uint8Array.from(d.b)).b);
};
const ARGB = (a: number, r: number, g: number, b: number): number => ((a << 24) | (r << 16) | (g << 8) | b) >>> 0;
const RED = ARGB(255, 220, 0, 0);
const BLUE = ARGB(255, 0, 0, 220);
const GREEN = ARGB(255, 0, 160, 0);
const ORANGE = ARGB(255, 255, 140, 0);

const header = (dual: boolean): Uint8Array => rec(0x4001, dual ? 1 : 0, (w) => w.u32(0xdbc01002).u32(1).u32(96).u32(96));
const solidBrush = (id: number, color: number): Uint8Array => rec(0x4008, (1 << 8) | id, (w) => w.u32(0xdbc01002).u32(0).u32(color));
const pen = (id: number, width: number, color: number, dashStyle?: number): Uint8Array =>
  rec(0x4008, (2 << 8) | id, (w) => {
    w.u32(0xdbc01002).u32(0).u32(dashStyle ? 0x20 : 0).u32(0).f32(width);
    if (dashStyle) w.u32(dashStyle);
    w.u32(0xdbc01002).u32(0).u32(color);
  });
const pathObj = (id: number, pts: [number, number][], types: number[]): Uint8Array =>
  rec(0x4008, (3 << 8) | id, (w) => {
    w.u32(0xdbc01002).u32(pts.length).u32(0);
    for (const [x, y] of pts) w.f32(x).f32(y);
    for (const t of types) w.b.push(t);
  });
const font = (id: number, size: number, family: string): Uint8Array => rec(0x4008, (6 << 8) | id, (w) => w.u32(0xdbc01002).f32(size).u32(2).u32(0).u32(0).u32(family.length).str16(family));
const fillRectsColor = (color: number, r: [number, number, number, number]): Uint8Array => rec(0x400a, 0x8000, (w) => w.u32(color).u32(1).f32(r[0]).f32(r[1]).f32(r[2]).f32(r[3]));
const fillEllipse = (brush: number, r: [number, number, number, number]): Uint8Array => rec(0x400e, 0, (w) => w.u32(brush).f32(r[0]).f32(r[1]).f32(r[2]).f32(r[3]));
const drawLines = (penId: number, pts: [number, number][]): Uint8Array =>
  rec(0x400d, penId, (w) => {
    w.u32(pts.length);
    for (const [x, y] of pts) w.f32(x).f32(y);
  });
const fillPathColor = (pathId: number, color: number): Uint8Array => rec(0x4014, 0x8000 | pathId, (w) => w.u32(color));
const drawString = (fontId: number, color: number, s: string, x: number, y: number): Uint8Array => rec(0x401c, 0x8000 | fontId, (w) => w.u32(color).u32(0xff).u32(s.length).f32(x).f32(y).f32(0).f32(0).str16(s));
const eofPlus = (): Uint8Array => rec(0x4002, 0, () => undefined);

/** An EMF of 200 × 100 px (96 dpi) whose comments carry the EMF+ records; `gdi`: GDI records after them. */
function emfOf(plus: Uint8Array[], gdi: Uint8Array[] = []): Uint8Array {
  const body = new Bytes();
  for (const r of plus) body.raw(r);
  body.pad();
  const comment = Uint8Array.from(new Bytes().u32(70).u32(16 + body.b.length).u32(4 + body.b.length).u32(0x2b464d45).raw(Uint8Array.from(body.b)).b);
  const eof = Uint8Array.from(new Bytes().u32(14).u32(20).u32(0).u32(16).u32(20).b);
  const parts = [comment, ...gdi, eof];
  const len = 108 + parts.reduce((n, p) => n + p.length, 0);
  // header: bounds (px), frame (0.01 mm: 200 px at 96 dpi = 52.92 mm), signature, version, size, records,
  // handles, …, a 1920 × 1080 px device of 508 × 285.75 mm (96 dpi)
  const h = new Bytes().u32(1).u32(108).u32(0).u32(0).u32(199).u32(99).u32(0).u32(0).u32(5292).u32(2646);
  h.u32(0x464d4520).u32(0x10000).u32(len).u32(parts.length + 1).u16(1).u16(0).u32(0).u32(0).u32(0).u32(1920).u32(1080).u32(508).u32(286);
  h.u32(0).u32(0).u32(0).u32(508000).u32(285750);
  const out = new Uint8Array(len);
  out.set(Uint8Array.from(h.b), 0);
  let o = 108;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** A GDI rectangle in gray (to tell GDI and EMF+ drawing apart in dual files). */
const gdiGrayRect = (): Uint8Array[] => [
  Uint8Array.from(new Bytes().u32(39).u32(24).u32(1).u32(0).u32(0x808080).u32(0).b), // CREATEBRUSHINDIRECT
  Uint8Array.from(new Bytes().u32(37).u32(12).u32(1).b), // SELECTOBJECT
  Uint8Array.from(new Bytes().u32(43).u32(24).u32(0).u32(0).u32(200).u32(100).b), // RECTANGLE (device units in MM_TEXT)
];

const PICTURE = (): Uint8Array =>
  emfOf([
    header(false),
    fillRectsColor(RED, [10, 10, 80, 40]),
    solidBrush(0, BLUE),
    fillEllipse(0, [110, 10, 80, 40]),
    pen(1, 4, GREEN),
    drawLines(1, [
      [10, 62],
      [190, 62],
    ]),
    pathObj(2, [
      [100, 68],
      [130, 96],
      [70, 96],
    ], [0, 1, 0x81]),
    fillPathColor(2, ORANGE),
    font(3, 11, 'Arial'),
    drawString(3, ARGB(255, 0, 0, 0), 'EMFplus text', 140, 80),
    eofPlus(),
  ]);

/** RGB sampler of page 1 at 96 dpi, (0,0) at the picture's top-left = the page's (no margins). */
function raster(pdf: Uint8Array, dpi = 96): (x: number, y: number) => number[] {
  const buf = withTemp((dir) => {
    writeFileSync(join(dir, 'in.pdf'), pdf);
    execFileSync('pdftoppm', ['-r', String(dpi), '-f', '1', '-l', '1', join(dir, 'in.pdf'), join(dir, 'p')]);
    return readFileSync(join(dir, readdirSync(dir).find((n) => n.endsWith('.ppm'))!));
  });
  const head = /^P6\s+(\d+)\s+(\d+)\s+255\s/.exec(new TextDecoder('latin1').decode(buf.subarray(0, 30)))!;
  const w = +head[1];
  const data = buf.subarray(head[0].length);
  return (x, y) => [...data.subarray((Math.round(y) * w + Math.round(x)) * 3, (Math.round(y) * w + Math.round(x)) * 3 + 3)];
}
const near = (p: number[], q: number[], tol = 40): boolean => p.every((v, i) => Math.abs(v - q[i]) <= tol);

async function ourRaster(emf: Uint8Array): Promise<(x: number, y: number) => number[]> {
  const pdf = (await docToPdf({ page: { width: 150, height: 105, margins: { top: 0, right: 0, bottom: 0, left: 0 } }, blocks: [{ type: 'image', data: emf, mime: 'image/emf', width: 200, height: 100 }] }, { fonts: loadFonts })).data;
  return raster(pdf);
}

/** The colours the test picture has where (in its own pixels). */
const PROBES: [string, number, number, number[]][] = [
  ['red rectangle', 50, 30, [220, 0, 0]],
  ['blue ellipse', 150, 30, [0, 0, 220]],
  ['green line', 100, 62, [0, 160, 0]],
  ['orange triangle', 100, 88, [255, 140, 0]],
  ['white outside', 100, 5, [255, 255, 255]],
  ['white beside the ellipse', 112, 12, [255, 255, 255]],
];

describe('EMF+ records', () => {
  it('an EMF+-only file is drawn from its EMF+ records', () => {
    const mf = parseEmf(PICTURE());
    const paths = mf.cmds.filter((c): c is Extract<MetaCmd, { k: 'path' }> => c.k === 'path');
    expect(paths.map((p) => p.fill ?? `stroke:${p.stroke}`)).toEqual(['DC0000', '0000DC', 'stroke:00A000', 'FF8C00']);
    expect(paths[2].width).toBeCloseTo(4, 1);
    const text = mf.cmds.find((c) => c.k === 'text');
    expect(text).toMatchObject({ text: 'EMFplus text', family: 'sans' });
    expect(text && text.k === 'text' && text.size).toBeCloseTo(11, 1);
  });

  it('a dual file keeps its GDI records; a dual file whose GDI part draws nothing uses EMF+', () => {
    const dual = parseEmf(emfOf([header(true), fillRectsColor(RED, [10, 10, 80, 40]), eofPlus()], gdiGrayRect()));
    expect(dual.cmds.filter((c) => c.k === 'path').map((c) => c.k === 'path' && c.fill)).toEqual(['808080']);
    const empty = parseEmf(emfOf([header(true), fillRectsColor(RED, [10, 10, 80, 40]), eofPlus()]));
    expect(empty.cmds.filter((c) => c.k === 'path').map((c) => c.k === 'path' && c.fill)).toEqual(['DC0000']);
  });

  it('world and page transforms, save/restore, clipping', () => {
    const emf = emfOf([
      header(false),
      rec(0x4030, 6, (w) => w.f32(1)), // page unit: millimetres
      rec(0x4025, 0, (w) => w.u32(1)), // save
      rec(0x402d, 0, (w) => w.f32(10).f32(5)), // translate 10 mm, 5 mm
      rec(0x4032, 0, (w) => w.f32(0).f32(0).f32(5).f32(20)), // clip: replace with a 5 × 20 mm rectangle
      fillRectsColor(RED, [0, 0, 20, 10]),
      rec(0x4026, 0, (w) => w.u32(1)), // restore: no transform, no clip
      fillRectsColor(BLUE, [0, 20, 10, 5]),
      eofPlus(),
    ]);
    const mf = parseEmf(emf);
    const paths = mf.cmds.filter((c): c is Extract<MetaCmd, { k: 'path' }> => c.k === 'path');
    const px = 96 / 25.4;
    const nums = (d: string): number[] => (d.match(/-?[\d.]+/g) ?? []).map(Number);
    const red = nums(paths[0].d);
    expect(red[0]).toBeCloseTo(10 * px, 0);
    expect(red[1]).toBeCloseTo(5 * px, 0);
    expect(red[2]).toBeCloseTo(30 * px, 0);
    const clips = mf.cmds.filter((c) => c.k === 'clip');
    expect(clips.map((c) => c.k === 'clip' && c.paths.length)).toEqual([1, 0]);
    expect(nums(paths[1].d)[1]).toBeCloseTo(20 * px, 0);
  });

  it('bitmaps (32-bit pixels and PNG) are drawn where DrawImage puts them', () => {
    const rgba = new Uint8Array(4 * 4 * 4);
    for (let k = 0; k < 16; k++) rgba.set([0, 0, 255, 255], k * 4); // BGRA: red
    const bitmap = rec(0x4008, (5 << 8) | 4, (w) => w.u32(0xdbc01002).u32(1).u32(4).u32(4).u32(16).u32(0x0026200a).u32(0).raw(rgba));
    const png = encodePng(new Uint8Array(4 * 4 * 3).fill(90), 4, 4, 3);
    const pngImg = rec(0x4008, (5 << 8) | 5, (w) => w.u32(0xdbc01002).u32(1).u32(4).u32(4).u32(0).u32(0).u32(1).raw(png));
    const draw = (id: number, x: number): Uint8Array => rec(0x401a, id, (w) => w.u32(0xff).u32(2).f32(0).f32(0).f32(4).f32(4).f32(x).f32(10).f32(40).f32(40));
    const mf = parseEmf(emfOf([header(false), bitmap, pngImg, draw(4, 10), draw(5, 100), eofPlus()]));
    const imgs = mf.cmds.filter((c): c is Extract<MetaCmd, { k: 'image' }> => c.k === 'image');
    expect(imgs.map((i) => [Math.round(i.x), Math.round(i.y), Math.round(i.w), Math.round(i.h)])).toEqual([
      [10, 10, 40, 40],
      [100, 10, 40, 40],
    ]);
    expect(imgs[1].png).toEqual(png);
  });

  it('malformed records do not throw (truncated objects, huge counts)', () => {
    const bad = emfOf([header(false), rec(0x400a, 0x8000, (w) => w.u32(RED).u32(0x7fffffff)), rec(0x4008, (3 << 8) | 1, (w) => w.u32(0).u32(1_000_000).u32(0)), rec(0x4014, 0x8001, (w) => w.u32(RED)), eofPlus()]);
    expect(() => parseEmf(bad)).not.toThrow();
  });
});

describe.skipIf(!hasCmd('pdftoppm'))('EMF+ pictures in PDF', () => {
  it('our PDF shows the shapes and the text', async () => {
    const at = await ourRaster(PICTURE());
    for (const [name, x, y, rgb] of PROBES) expect(near(at(x, y), rgb), `${name}: ${at(x, y)}`).toBe(true);
    const pdf = (await docToPdf({ blocks: [{ type: 'image', data: PICTURE(), mime: 'image/emf', width: 200, height: 100 }] }, { fonts: loadFonts })).data;
    expect(pdfToText(pdf)).toContain('EMFplus text');
  });

  it.skipIf(!hasCmd('soffice'))('LibreOffice draws the same file the same way', () => {
    // its PDF of the picture: find the picture by the red rectangle's top-left corner (10,10)
    const at = raster(soffice(PICTURE(), 'emf', 'pdf'));
    let ox = -1;
    let oy = -1;
    outer: for (let y = 0; y < 1123; y++)
      for (let x = 0; x < 794; x++)
        if (near(at(x, y), [220, 0, 0])) {
          ox = x - 10;
          oy = y - 10;
          break outer;
        }
    expect(ox).toBeGreaterThanOrEqual(0);
    for (const [name, x, y, rgb] of PROBES) expect(near(at(ox + x, oy + y), rgb, 60), `LibreOffice ${name}: ${at(ox + x, oy + y)}`).toBe(true);
  });
});
