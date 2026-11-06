/** Clipping in vector pictures: EMF/WMF clip records, SVG clipPath → display list → PDF / SVG / EMF. */
import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseEmf, metafileToSvg, type Metafile, type MetaCmd } from '../../src/core/metafile';
import { parseSvg } from '../../src/core/svg-render';
import { metafileToEmf } from '../../src/core/writers/emf';
import { docToPdf } from '../../src/core/writers/pdf';
import { hasCmd, loadFonts, soffice, withTemp } from './helpers';
import { decodePng } from '../../src/core/png';

/** Rasterise page 1 at 72 dpi; RGB sampler with (0,0) at the picture's top-left (found by its red pixels). */
function raster(pdf: Uint8Array): (x: number, y: number) => number[] {
  const buf = withTemp((dir) => {
    writeFileSync(join(dir, 'in.pdf'), pdf);
    execFileSync('pdftoppm', ['-r', '72', '-f', '1', '-l', '1', join(dir, 'in.pdf'), join(dir, 'p')]);
    return readFileSync(join(dir, readdirSync(dir).find((n) => n.endsWith('.ppm'))!));
  });
  const head = /^P6\s+(\d+)\s+(\d+)\s+255\s/.exec(new TextDecoder('latin1').decode(buf.subarray(0, 30)))!;
  const w = +head[1];
  const data = buf.subarray(head[0].length);
  return (x, y) => [...data.subarray((Math.round(y) * w + Math.round(x)) * 3, (Math.round(y) * w + Math.round(x)) * 3 + 3)];
}
const isRed = (p: number[]): boolean => p[0] > 200 && p[1] < 60 && p[2] < 60;
const isBlue = (p: number[]): boolean => p[2] > 200 && p[0] < 60 && p[1] < 60;
const isWhite = (p: number[]): boolean => p.every((v) => v > 235);

/**
 * Sampler in picture units: the picture is found by the bounding box of its non-white pixels, whose
 * top-left corner is the known point `first` of the picture; `k` points per picture unit.
 */
function locate(px: (x: number, y: number) => number[], first: [number, number], k: number): (x: number, y: number) => number[] {
  let bx = Infinity;
  let by = Infinity;
  for (let y = 0; y < 842; y += 1)
    for (let x = 0; x < 595; x += 1)
      if (!isWhite(px(x, y))) {
        bx = Math.min(bx, x);
        by = Math.min(by, y);
      }
  return (x, y) => px(bx + (x - first[0]) * k, by + (y - first[1]) * k);
}

/** Our PDF of a picture of w × h points. */
async function pdfOf(data: Uint8Array, mime: string, w: number, h: number): Promise<(x: number, y: number) => number[]> {
  return raster((await docToPdf({ blocks: [{ type: 'image', data, mime, width: (w * 96) / 72, height: (h * 96) / 72 }] }, { fonts: loadFonts })).data);
}

/** A small EMF assembled record by record (logical units = output units × 16, as our writer maps them). */
function emf(records: [number, ...number[]][], w = 200, h = 100): Uint8Array {
  const base = metafileToEmf({ width: w, height: h, widthPt: w, heightPt: h, cmds: [] });
  const head = base.subarray(0, base.length - 20); // drop EOF
  const parts: Uint8Array[] = [head];
  for (const [type, ...vals] of records) {
    const r = new Uint8Array(8 + vals.length * 4);
    const dv = new DataView(r.buffer);
    dv.setUint32(0, type, true);
    dv.setUint32(4, r.length, true);
    vals.forEach((v, i) => dv.setInt32(8 + i * 4, v, true));
    parts.push(r);
  }
  const eof = new Uint8Array(20);
  new DataView(eof.buffer).setUint32(0, 14, true);
  new DataView(eof.buffer).setUint32(4, 20, true);
  new DataView(eof.buffer).setUint32(16, 20, true);
  parts.push(eof);
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  const dv = new DataView(out.buffer);
  dv.setUint32(48, out.length, true);
  dv.setUint32(52, dv.getUint32(52, true) - 1 + records.length + 1, true);
  return out;
}

const S = 16;
const RED = 0x0000ff; // COLORREF 0x00BBGGRR
const BLUE = 0xff0000;
/** Fill (0,0)-(200,100) red with a clip to x 50..150 inside SaveDC/RestoreDC, excluding x 90..110; then a blue bar unclipped. */
export const CLIP_EMF = (): Uint8Array =>
  emf([
    [39, 1, 0, RED, 0], // CREATEBRUSHINDIRECT red
    [37, 1], // SELECTOBJECT
    [37, 0x80000008], // NULL_PEN
    [33], // SAVEDC
    [30, 50 * S, 0, 150 * S, 100 * S], // INTERSECTCLIPRECT
    [29, 90 * S, 0, 110 * S, 100 * S], // EXCLUDECLIPRECT
    [43, 0, 0, 200 * S, 60 * S], // RECTANGLE (top 60 %)
    [34, -1], // RESTOREDC
    [39, 2, 0, BLUE, 0],
    [37, 2],
    [43, 0, 70 * S, 200 * S, 100 * S], // RECTANGLE, unclipped
  ]);

describe('EMF clipping', () => {
  it('reads clip records into the display list, restored with the DC', () => {
    const mf = parseEmf(CLIP_EMF());
    const clips = mf.cmds.filter((c): c is Extract<MetaCmd, { k: 'clip' }> => c.k === 'clip');
    expect(clips.map((c) => c.paths.length)).toEqual([1, 2, 0]);
    expect(clips[1].paths[1].evenOdd).toBe(true);
  });

  it.skipIf(!hasCmd('pdftoppm'))('our PDF clips like the spec says', async () => {
    // content covers the whole frame (the blue bar spans it, the red fill starts at the top)
    const at = locate(await pdfOf(CLIP_EMF(), 'image/emf', 200, 100), [0, 0], 1);
    expect(isWhite(at(25, 30)), 'left of the clip').toBe(true);
    expect(isRed(at(70, 30)), 'inside the clip').toBe(true);
    expect(isWhite(at(100, 30)), 'excluded band').toBe(true);
    expect(isRed(at(130, 30))).toBe(true);
    expect(isWhite(at(175, 30)), 'right of the clip').toBe(true);
    expect(isBlue(at(25, 85)) && isBlue(at(175, 85)), 'after RestoreDC: no clip').toBe(true);
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftoppm'))('LibreOffice clips the same EMF to the same rectangle', () => {
    const at = locate(raster(soffice(CLIP_EMF(), 'emf', 'pdf')), [0, 0], 1);
    expect(isWhite(at(25, 30))).toBe(true);
    expect(isRed(at(70, 30))).toBe(true);
    expect(isRed(at(130, 30))).toBe(true);
    expect(isWhite(at(175, 30))).toBe(true);
    expect(isBlue(at(25, 85)) && isBlue(at(175, 85))).toBe(true);
    // (LibreOffice ignores EXCLUDECLIPRECT here; per MS-EMF 2.3.2.1 the band 90..110 is excluded, as we do)
  });
});

const CLIP_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 200 100">' +
  '<defs><clipPath id="c"><circle cx="50" cy="50" r="40"/></clipPath><clipPath id="h"><rect x="0" y="0" width="200" height="50"/></clipPath></defs>' +
  '<rect x="0" y="0" width="100" height="100" fill="#ff0000" clip-path="url(#c)"/>' +
  '<g clip-path="url(#h)"><circle cx="150" cy="50" r="40" fill="#0000ff"/></g>' +
  '</svg>';

describe('SVG clipPath', () => {
  const mf = parseSvg(CLIP_SVG)!;

  it('becomes clip commands around the clipped content', () => {
    expect(mf.cmds.map((c) => c.k)).toEqual(['clip', 'path', 'clip', 'clip', 'path', 'clip']);
  });

  it.skipIf(!hasCmd('pdftoppm'))('is honoured in PDF', async () => {
    // content starts at (10, 10): the circle's left edge and the circles' tops; 200 px = 150 pt
    const at = locate(await pdfOf(new TextEncoder().encode(CLIP_SVG), 'image/svg+xml', 150, 75), [10, 10], 0.75);
    expect(isRed(at(50, 50))).toBe(true);
    expect(isWhite(at(15, 15)), 'rect corner outside the circle').toBe(true);
    expect(isBlue(at(150, 30))).toBe(true);
    expect(isWhite(at(150, 75)), 'lower half clipped').toBe(true);
  });

  it('survives the SVG and EMF writers', () => {
    const svg = metafileToSvg(mf);
    expect(svg).toContain('<clipPath id="c1">');
    expect(svg.match(/clip-path="url\(#c\d\)"/g)?.length).toBe(2);
    const back = parseEmf(metafileToEmf(mf));
    expect(back.cmds.filter((c) => c.k === 'clip').map((c) => (c as { paths: unknown[] }).paths.length)).toEqual([1, 0, 1, 0]);
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftoppm'))('LibreOffice honours the clip records our EMF writer emits', () => {
    const at = locate(raster(soffice(metafileToEmf(mf as Metafile), 'emf', 'pdf')), [10, 10], 0.75);
    expect(isWhite(at(15, 15))).toBe(true);
    expect(isRed(at(50, 50))).toBe(true);
    expect(isBlue(at(150, 30))).toBe(true);
    expect(isWhite(at(150, 75))).toBe(true);
  });
});

void readFileSync;

const ALPHA_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="300" height="120" viewBox="0 0 300 120">' +
  '<defs><linearGradient id="g"><stop offset="0" stop-color="#ff0000"/><stop offset="1" stop-color="#0000ff"/></linearGradient></defs>' +
  '<rect x="10" y="10" width="100" height="100" fill="#ff0000"/>' +
  '<circle cx="110" cy="60" r="50" fill="#0000ff" fill-opacity="0.5"/>' +
  '<rect x="180" y="10" width="110" height="100" fill="url(#g)"/>' +
  '</svg>';

describe('EMF writer: transparency and gradients', () => {
  const emf = metafileToEmf(parseSvg(ALPHA_SVG)!);

  it('writes them as alpha-blended bitmaps that our reader reads back', () => {
    const back = parseEmf(emf);
    expect(back.cmds.filter((c) => c.k === 'image').length).toBe(2);
    const img = back.cmds.find((c): c is Extract<MetaCmd, { k: 'image' }> => c.k === 'image')!;
    // the circle: 50 % blue at its centre, transparent at the corner of its box
    const dec = decodePng(img.png)!;
    const at = (x: number, y: number): number[] => [...dec.rgba.subarray((y * dec.width + x) * 4, (y * dec.width + x) * 4 + 4)];
    expect(at(dec.width >> 1, dec.height >> 1)[3]).toBeGreaterThan(120);
    expect(at(dec.width >> 1, dec.height >> 1)[3]).toBeLessThan(135);
    expect(at(1, 1)[3]).toBe(0);
  });

  it.skipIf(!hasCmd('pdftoppm'))('our PDF of the EMF: the overlap is purple, the gradient runs red → blue', async () => {
    // content starts at (10, 10); 300 px = 225 pt
    const at = locate(await pdfOf(emf, 'image/emf', 225, 90), [10, 10], 0.75);
    expect(isRed(at(40, 60))).toBe(true);
    const mix = at(85, 60); // inside both: blue at 50 % over red
    expect(mix[0]).toBeGreaterThan(100);
    expect(mix[2]).toBeGreaterThan(100);
    expect(mix[1]).toBeLessThan(40);
    const left = at(185, 60);
    const right = at(285, 60);
    expect(left[0] > left[2] && right[2] > right[0]).toBe(true);
  });

  it.skipIf(!hasCmd('soffice') || !hasCmd('pdftoppm'))('LibreOffice blends them with the content underneath too', () => {
    const at = locate(raster(soffice(emf, 'emf', 'pdf')), [10, 10], 0.75);
    expect(isRed(at(40, 60))).toBe(true);
    // red shows through the blue (blending with white would give light blue, with green ≈ 127).
    // LibreOffice takes the bitmap as straight alpha and darkens the blue (≈ 64 instead of 127);
    // GDI's AlphaBlend and MS-EMF define premultiplied values, which is what we write.
    const mix = at(85, 60);
    expect(mix[0]).toBeGreaterThan(100);
    expect(mix[2]).toBeGreaterThan(40);
    expect(mix[1]).toBeLessThan(40);
    const left = at(185, 60);
    const right = at(285, 60);
    expect(left[0] > left[2] && right[2] > right[0]).toBe(true);
  });
});
