/**
 * SVG pattern fills and masks → display list → PDF, compared pixel by pixel with Chromium's rendering
 * of the same SVG (an independent renderer).
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseSvg } from '../../src/core/svg-render';
import { metafileToEmf } from '../../src/core/writers/emf';
import { docToPdf } from '../../src/core/writers/pdf';
import { decodePng } from '../../src/core/png';
import { hasCmd, loadFonts, withTemp } from './helpers';

const W = 200;
const H = 100;
type Px = (x: number, y: number) => number[];

/** Our PDF of the SVG on a page of its size (1 px = 0.75 pt), rendered at 96 dpi: pixels match the SVG's. */
async function ours(svg: string): Promise<Px> {
  const pdf = (await docToPdf({ page: { width: (W * 3) / 4, height: (H * 3) / 4 + 30, margins: { top: 0, right: 0, bottom: 0, left: 0 } }, blocks: [{ type: 'image', data: new TextEncoder().encode(svg), mime: 'image/svg+xml', width: W, height: H }] }, { fonts: loadFonts })).data;
  const buf = withTemp((dir) => {
    writeFileSync(join(dir, 'in.pdf'), pdf);
    execFileSync('pdftoppm', ['-r', '96', '-f', '1', '-l', '1', join(dir, 'in.pdf'), join(dir, 'p')]);
    return readFileSync(join(dir, readdirSync(dir).find((n) => n.endsWith('.ppm'))!));
  });
  const head = /^P6\s+(\d+)\s+(\d+)\s+255\s/.exec(new TextDecoder('latin1').decode(buf.subarray(0, 30)))!;
  const w = +head[1];
  const data = buf.subarray(head[0].length);
  return (x, y) => [...data.subarray((y * w + x) * 3, (y * w + x) * 3 + 3)];
}

let browser: import('playwright').Browser | null = null;
const chromium = async (svg: string): Promise<Px> => {
  const page = await browser!.newPage({ viewport: { width: W, height: H } });
  await page.setContent(`<html><body style="margin:0;background:#fff">${svg}</body></html>`);
  const png = await page.screenshot({ clip: { x: 0, y: 0, width: W, height: H } });
  await page.close();
  const img = decodePng(new Uint8Array(png))!;
  const ch = img.rgba.length / (img.width * img.height);
  return (x, y) => [...img.rgba.subarray((y * img.width + x) * ch, (y * img.width + x) * ch + 3)];
};

/** Share of pixels where the two renderings differ clearly (anti-aliased edges aside). */
function differing(a: Px, b: Px): number {
  let n = 0;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) if (a(x, y).some((v, i) => Math.abs(v - b(x, y)[i]) > 80)) n++;
  return n / (W * H);
}
const svgOf = (body: string): string => `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${body}</svg>`;

const CASES: Record<string, string> = {
  'user-space checkerboard in a circle': svgOf(
    '<defs><pattern id="p" x="0" y="0" width="20" height="20" patternUnits="userSpaceOnUse"><rect width="10" height="10" fill="#d00"/><rect x="10" y="10" width="10" height="10" fill="#00c"/></pattern></defs>' +
      '<circle cx="100" cy="50" r="45" fill="url(#p)" stroke="#000" stroke-width="2"/>',
  ),
  'object-bounding-box tiles with rotated stripes': svgOf(
    '<defs><pattern id="s" width="0.1" height="0.25" patternTransform="rotate(30)"><rect width="6" height="40" fill="#080"/></pattern></defs>' + '<rect x="10" y="10" width="180" height="80" fill="url(#s)"/>',
  ),
  'a viewBox in the tile and an inherited pattern (href)': svgOf(
    '<defs><pattern id="base" width="25" height="25" patternUnits="userSpaceOnUse" viewBox="0 0 10 10"><circle cx="5" cy="5" r="4" fill="#f80"/></pattern>' +
      '<pattern id="child" href="#base" x="5" y="5"/></defs><rect x="0" y="0" width="200" height="100" fill="url(#child)"/>',
  ),
  'a mask: a white rectangle with a black hole': svgOf(
    '<defs><mask id="m"><rect x="0" y="0" width="200" height="100" fill="#fff"/><circle cx="100" cy="50" r="30" fill="#000"/></mask></defs>' + '<rect x="20" y="10" width="160" height="80" fill="#06c" mask="url(#m)"/>',
  ),
  'a mask of white shapes on nothing (with a black background first)': svgOf(
    '<defs><mask id="m2"><rect width="200" height="100" fill="black"/><circle cx="60" cy="50" r="35" fill="white"/><rect x="120" y="20" width="60" height="60" fill="white"/></mask></defs>' +
      '<g mask="url(#m2)"><rect width="200" height="100" fill="#c0c"/></g>',
  ),
};

describe.skipIf(!hasCmd('pdftoppm'))('SVG patterns and masks in our PDF match Chromium', () => {
  beforeAll(async () => {
    browser = await (await import('playwright')).chromium.launch();
  });
  afterAll(async () => {
    await browser?.close();
  });

  for (const [name, svg] of Object.entries(CASES))
    it(name, async () => {
      const share = differing(await ours(svg), await chromium(svg));
      expect(share).toBeLessThan(0.03);
    });

  it('the old behaviour would have failed: a pattern fill is not one flat colour', async () => {
    const px = await ours(CASES['user-space checkerboard in a circle']);
    // inside the circle: red and blue cells, and white between them
    const seen = new Set<string>();
    for (let y = 20; y < 80; y += 2) for (let x = 70; x < 130; x += 2) seen.add(px(x, y)[0] > 150 && px(x, y)[2] < 100 ? 'red' : px(x, y)[2] > 150 && px(x, y)[0] < 100 ? 'blue' : px(x, y).every((v) => v > 200) ? 'white' : 'other');
    expect([...seen]).toEqual(expect.arrayContaining(['red', 'blue', 'white']));
  });
});

describe('SVG patterns and masks: the display list', () => {
  it('a pattern fill becomes clip + tiles + restore; a mask becomes a clip with a hole', () => {
    const mf = parseSvg(CASES['user-space checkerboard in a circle'])!;
    const clips = mf.cmds.filter((c) => c.k === 'clip');
    expect(clips.length).toBeGreaterThan(20); // one per tile, then the restore
    expect(clips[clips.length - 1]).toEqual({ k: 'clip', paths: [] });
    const paths = mf.cmds.filter((c) => c.k === 'path');
    expect(paths.some((p) => p.k === 'path' && p.fill === 'DD0000')).toBe(true);
    expect(paths[paths.length - 1]).toMatchObject({ stroke: '000000', fill: undefined }); // the outline after the tiles
    const masked = parseSvg(CASES['a mask: a white rectangle with a black hole'])!;
    const c = masked.cmds.find((x) => x.k === 'clip')!;
    expect(c.k === 'clip' && c.paths.length).toBe(2);
    expect(c.k === 'clip' && c.paths[1].evenOdd).toBe(true);
  });

  it('a pattern referencing itself does not loop; an empty pattern paints nothing', () => {
    const loop = parseSvg(svgOf('<defs><pattern id="a" width="10" height="10" patternUnits="userSpaceOnUse"><rect width="5" height="5" fill="url(#a)"/></pattern></defs><rect width="200" height="100" fill="url(#a)"/>'))!;
    expect(loop.cmds.length).toBeLessThan(200_000);
    const empty = parseSvg(svgOf('<defs><pattern id="e" width="10" height="10"/></defs><rect width="200" height="100" fill="url(#e)"/>'))!;
    expect(empty.cmds.filter((c) => c.k === 'path')).toEqual([]);
  });

  it('the EMF writer keeps the tiles inside clip paths', () => {
    const emf = metafileToEmf(parseSvg(CASES['user-space checkerboard in a circle'])!);
    const view = new DataView(emf.buffer, emf.byteOffset, emf.byteLength);
    let clipRecords = 0;
    for (let o = 0; o + 8 <= emf.length; ) {
      const type = view.getUint32(o, true);
      const size = view.getUint32(o + 4, true);
      if (type === 67) clipRecords++; // EMR_SELECTCLIPPATH
      if (size < 8) break;
      o += size;
    }
    expect(clipRecords).toBeGreaterThan(20);
  });
});
