/**
 * Redaction, harder cases: vertical CJK text, images in every PDF codec (CCITT fax, JBIG2, JPEG 2000,
 * JPEG, filter chains, predictors), soft masks, and what a redacted file must not keep (thumbnails, title).
 */
import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import fontkit from '@pdf-lib/fontkit';
import { unzlibSync, zlibSync } from 'fflate';
import { concatTransformationMatrix, drawObject, PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream, PDFRef, popGraphicsState, pushGraphicsState, type PDFPage } from 'pdf-lib';
import { redactPage, type Rect } from '../../src/core/pdf-redact';
import { pdfjsImageDecoder } from '../../src/core/pdf-image-filters';
import { composePdf } from '../../src/core/pdf-tools';
import { encodePng } from '../../src/core/png';
import { hasCmd, loadFonts, loadSubsetter, pdfToText, renderPages } from './helpers';

const WASM = join(process.cwd(), 'node_modules/pdfjs-dist/wasm/');
const decodeImage = pdfjsImageDecoder({ wasmUrl: pathToFileURL(WASM).href, fetchBinary: async (f) => new Uint8Array(await readFile(join(WASM, f))) });

/** A grayscale render (PGM from pdftoppm at 72 dpi): value at a pixel, and the size. */
function render(pdf: Uint8Array): { w: number; h: number; at(x: number, y: number): number } {
  const [pgm] = renderPages(pdf, 72);
  const m = /^P5\s+(\d+)\s+(\d+)\s+255\s/.exec(new TextDecoder('latin1').decode(pgm.subarray(0, 40)))!;
  const w = +m[1];
  return { w, h: +m[2], at: (x, y) => pgm[m[0].length + y * w + x] };
}

// ---------------------------------------------------------------------------------------------
const VERT = '秘密情報機密番号東京';
const FS = 24;
const TOP = 360; // the column starts here (PDF y), at x = 150

/** Vertical text: a Type0 font with Identity-V; optional W2 entries for some characters. */
async function verticalPdf(w2?: (cidOf: (ch: string) => number) => (number | number[])[]): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const sub = await loadSubsetter();
  const font = await doc.embedFont(sub(await loadFonts('cjk'), [...VERT].map((c) => c.codePointAt(0)!)), { subset: false });
  doc.addPage([300, 400]).drawText(VERT, { x: 20, y: 20, size: 10, font });
  const hex = font.encodeText(VERT).toString();
  const d2 = await PDFDocument.load(await doc.save());
  const page = d2.getPage(0);
  const fonts = page.node.Resources()!.lookup(PDFName.of('Font'), PDFDict);
  const [name, ref] = fonts.entries()[0];
  const fd = d2.context.lookup(ref as PDFRef, PDFDict);
  fd.set(PDFName.of('Encoding'), PDFName.of('Identity-V'));
  if (w2) {
    const desc = d2.context.lookup((fd.lookup(PDFName.of('DescendantFonts'), PDFArray)).get(0) as PDFRef, PDFDict);
    const cidOf = (ch: string): number => parseInt(font.encodeText(ch).toString().slice(1, 5), 16);
    desc.set(PDFName.of('W2'), d2.context.obj(w2(cidOf)));
  }
  page.node.set(PDFName.of('Contents'), d2.context.register(d2.context.stream(`BT ${name.toString()} ${FS} Tf 1 0 0 1 150 ${TOP} Tm ${hex} Tj ET`)));
  return d2.save();
}

/** Rows (top-down, 72 dpi) with ink in the text column. */
const inkRows = (r: ReturnType<typeof render>): number[] => Array.from({ length: r.h }, (_, y) => y).filter((y) => Array.from({ length: 40 }, (_, k) => 130 + k).some((x) => r.at(x, y) < 128));

async function redactVertical(pdf: Uint8Array, rect: Rect): Promise<{ glyphs: number; out: Uint8Array }> {
  const doc = await PDFDocument.load(pdf);
  const r = await redactPage(doc, doc.getPage(0), [rect]);
  return { glyphs: r.glyphs, out: await doc.save() };
}

describe.skipIf(!hasCmd('pdftoppm') || !hasCmd('pdftotext'))('redaction of vertical CJK text', () => {
  it('Identity-V: the glyphs in the area are removed, the column below keeps its place', async () => {
    const pdf = await verticalPdf();
    const before = render(pdf);
    // glyph k hangs below the pen: PDF y from TOP - 24(k+1) to TOP - 24k; take glyphs 2..4 (情報機)
    const { glyphs, out } = await redactVertical(pdf, { x0: 130, y0: TOP - FS * 5 + 3, x1: 170, y1: TOP - FS * 2 - 3 });
    expect(glyphs).toBe(3);
    const text = pdfToText(out).replace(/\s+/g, '');
    expect(text).toContain('秘密');
    expect(text).toContain('密番号東京');
    for (const ch of '情報機') expect(text).not.toContain(ch);
    const after = render(out);
    const rowsBefore = inkRows(before);
    const rowsAfter = inkRows(after);
    // ink of the removed glyphs (page rows 400 - y) is gone, everything else is pixel-identical
    const gone = (y: number): boolean => y >= 400 - (TOP - FS * 2) && y < 400 - (TOP - FS * 5);
    expect(rowsAfter.filter(gone)).toEqual([]);
    expect(rowsBefore.filter(gone).length).toBeGreaterThan(30);
    for (let y = 0; y < 400; y++) if (!gone(y)) for (let x = 120; x < 180; x++) expect(after.at(x, y), `pixel ${x},${y}`).toBe(before.at(x, y));
  });

  it('W2 metrics (own vertical advance of a glyph) are honoured', async () => {
    // 秘 advances 1.5 em, so everything after it sits half an em lower
    const pdf = await verticalPdf((cidOf) => [cidOf('秘'), [-1500, 500, 880]]);
    const before = render(pdf);
    const shift = FS / 2;
    const { glyphs, out } = await redactVertical(pdf, { x0: 130, y0: TOP - shift - FS * 4 + 3, x1: 170, y1: TOP - shift - FS * 3 - 3 }); // 報 only
    expect(glyphs).toBe(1);
    expect(pdfToText(out)).not.toContain('報');
    expect(pdfToText(out).replace(/\s+/g, '')).toContain('機密番号東京');
    const after = render(out);
    const gone = (y: number): boolean => y >= 400 - (TOP - shift - FS * 3) && y < 400 - (TOP - shift - FS * 4);
    expect(inkRows(after).filter(gone)).toEqual([]);
    for (let y = 0; y < 400; y++) if (!gone(y)) for (let x = 120; x < 180; x++) expect(after.at(x, y), `pixel ${x},${y}`).toBe(before.at(x, y));
  });
});

// ---------------------------------------------------------------------------------------------
// Images in the codecs PDF knows. Fixtures are made with Pillow (libtiff for G4, OpenJPEG for JPX).

const PY = `
import base64, io, json, sys
from PIL import Image
n = 64
bw = Image.new('1', (n, n), 1)
for y in range(n):
    for x in range(n):
        if ((x // 8) + (y // 8)) % 2 == 0: bw.putpixel((x, y), 0)
buf = io.BytesIO(); bw.save(buf, 'TIFF', compression='group4')
t = Image.open(io.BytesIO(buf.getvalue()))
offs, counts = t.tag_v2[273], t.tag_v2[279]
assert len(offs) == 1
g4 = buf.getvalue()[offs[0]:offs[0] + counts[0]]
photometric = t.tag_v2[262]
rgb = Image.new('RGB', (32, 32))
for y in range(32):
    for x in range(32): rgb.putpixel((x, y), (x * 8, y * 8, 100))
j = io.BytesIO(); rgb.save(j, 'JPEG2000')
rgba = rgb.copy(); rgba.putalpha(Image.new('L', (32, 32), 180))
ja = io.BytesIO(); rgba.save(ja, 'JPEG2000')
jp = io.BytesIO(); rgb.save(jp, 'JPEG', quality=95)
cm = io.BytesIO(); rgb.convert('CMYK').save(cm, 'JPEG', quality=95)
print(json.dumps({'g4': base64.b64encode(g4).decode(), 'photometric': photometric, 'jpx': base64.b64encode(j.getvalue()).decode(), 'jpxa': base64.b64encode(ja.getvalue()).decode(), 'jpg': base64.b64encode(jp.getvalue()).decode(), 'cmyk': base64.b64encode(cm.getvalue()).decode()}))
`;
const hasPillow = ((): boolean => {
  try {
    execFileSync('python3', ['-c', 'import PIL'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
})();
const fx = hasPillow ? (JSON.parse(execFileSync('python3', ['-c', PY], { encoding: 'utf8' })) as Record<string, string | number>) : {};
const b64 = (k: string): Uint8Array => new Uint8Array(Buffer.from(fx[k] as string, 'base64'));

/** A JBIG2 embedded stream (as PDF wants it): page information + one immediate generic region coded with MMR (= G4). */
function jbig2Mmr(g4: Uint8Array, w: number, h: number): Uint8Array {
  const u32 = (v: number): number[] => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
  const segment = (no: number, type: number, data: number[]): number[] => [...u32(no), type, 0, 1, ...u32(data.length), ...data];
  const pageInfo = [...u32(w), ...u32(h), ...u32(0), ...u32(0), 0, 0, 0];
  const region = [...u32(w), ...u32(h), ...u32(0), ...u32(0), 0, /* MMR */ 1, ...g4];
  return Uint8Array.from([...segment(0, 48, pageInfo), ...segment(1, 38, region)]);
}

/** One page with one image XObject drawn at (50, 50) size 160 × 160. */
async function imagePdf(data: Uint8Array, dict: Record<string, unknown>, w: number, h: number): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([260, 260]);
  const stream = doc.context.stream(data, { Type: 'XObject', Subtype: 'Image', Width: w, Height: h, ...dict });
  const name = page.node.newXObject('Im', doc.context.register(stream));
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(160, 0, 0, 160, 50, 50), drawObject(name), popGraphicsState());
  return doc.save();
}

/** The image of page 1 after redaction: its dictionary and inflated samples. */
function imageOf(page: PDFPage): { dict: PDFDict; data: Uint8Array } {
  const xo = page.node.Resources()!.lookup(PDFName.of('XObject'), PDFDict);
  const s = page.doc.context.lookup(xo.values()[0] as PDFRef) as PDFRawStream;
  return { dict: s.dict, data: unzlibSync(s.getContents()) };
}

/** Redact the top-left quarter of the image (PDF 50..130 × 130..210). */
const QUARTER: Rect = { x0: 40, y0: 131, x1: 129, y1: 220 };

async function redactImage(pdf: Uint8Array): Promise<{ out: Uint8Array; images: number; warnings: string[] }> {
  const doc = await PDFDocument.load(pdf);
  const r = await redactPage(doc, doc.getPage(0), [QUARTER], { redact: true, decodeImage });
  return { out: await doc.save(), images: r.images, warnings: r.warnings };
}

/** Rendered: black over the quarter, unchanged elsewhere (both renders by poppler). */
function expectQuarterBlack(before: Uint8Array, after: Uint8Array): void {
  const a = render(before);
  const b = render(after);
  let changedOutside = 0;
  for (let y = 52; y < 208; y++)
    for (let x = 52; x < 208; x++) {
      // the zeroed pixels: image columns/rows 0..31 of 64 → page x 50..130, rows (from the top) 50..130
      const inQuarter = x < 131 && y < 131;
      if (x > 53 && y > 53 && x < 127 && y < 127) expect(b.at(x, y), `inside ${x},${y}`).toBeLessThan(40);
      else if (!inQuarter && Math.abs(a.at(x, y) - b.at(x, y)) > 8) changedOutside++;
    }
  expect(changedOutside).toBe(0);
}

describe.skipIf(!hasPillow || !hasCmd('pdftoppm'))('redaction of images in every codec', () => {
  const blackIs1 = (): boolean => fx.photometric === 1;

  it('CCITT fax (G4): decoded, the quarter zeroed, rewritten with Flate', async () => {
    const pdf = await imagePdf(b64('g4'), { ColorSpace: 'DeviceGray', BitsPerComponent: 1, Filter: 'CCITTFaxDecode', DecodeParms: { K: -1, Columns: 64, Rows: 64, BlackIs1: blackIs1() } }, 64, 64);
    const { out, images, warnings } = await redactImage(pdf);
    expect(warnings).toEqual([]);
    expect(images).toBe(1);
    const img = imageOf((await PDFDocument.load(out)).getPage(0));
    expect(img.dict.lookup(PDFName.of('Filter'))?.toString()).toBe('/FlateDecode');
    expect(img.dict.has(PDFName.of('DecodeParms'))).toBe(false);
    expect(img.data.length).toBe(8 * 64);
    expectQuarterBlack(pdf, out);
  });

  it('JBIG2 (generic region): decoded, the quarter zeroed', async () => {
    const pdf = await imagePdf(jbig2Mmr(b64('g4'), 64, 64), { ColorSpace: 'DeviceGray', BitsPerComponent: 1, Filter: 'JBIG2Decode' }, 64, 64);
    const { out, images, warnings } = await redactImage(pdf);
    expect(warnings).toEqual([]);
    expect(images).toBe(1);
    expectQuarterBlack(pdf, out);
  });

  it('JPEG 2000: exact samples, zero in the quarter only', async () => {
    const pdf = await imagePdf(b64('jpx'), { Filter: 'JPXDecode', ColorSpace: 'DeviceRGB', BitsPerComponent: 8 }, 32, 32);
    const { out, images } = await redactImage(pdf);
    expect(images).toBe(1);
    const { dict, data } = imageOf((await PDFDocument.load(out)).getPage(0));
    expect(dict.lookup(PDFName.of('ColorSpace'))?.toString()).toBe('/DeviceRGB');
    const px = (x: number, y: number): number[] => [...data.subarray((y * 32 + x) * 3, (y * 32 + x) * 3 + 3)];
    expect(px(0, 0)).toEqual([0, 0, 0]);
    expect(px(15, 15)).toEqual([0, 0, 0]);
    expect(px(16, 0)).toEqual([128, 0, 100]); // lossless: the original values
    expect(px(31, 31)).toEqual([248, 248, 100]);
    expect(px(0, 16)).toEqual([0, 128, 100]);
  });

  it('JPEG 2000 with its own alpha (SMaskInData): the alpha becomes a soft mask, zeroed in the area too', async () => {
    const pdf = await imagePdf(b64('jpxa'), { Filter: 'JPXDecode', SMaskInData: 1 }, 32, 32);
    const { out } = await redactImage(pdf);
    const page = (await PDFDocument.load(out)).getPage(0);
    const { dict, data } = imageOf(page);
    expect(data.length).toBe(32 * 32 * 3);
    expect(dict.has(PDFName.of('SMaskInData'))).toBe(false);
    const sm = page.doc.context.lookup(dict.get(PDFName.of('SMask')) as PDFRef) as PDFRawStream;
    const alpha = unzlibSync(sm.getContents());
    expect(alpha[0]).toBe(0);
    expect(alpha[31]).toBe(180);
  });

  it('JPEG (RGB and Adobe CMYK): the samples keep their colour space and Decode array', async () => {
    for (const [key, cs, comps] of [
      ['jpg', 'DeviceRGB', 3],
      ['cmyk', 'DeviceCMYK', 4],
    ] as const) {
      const pdf = await imagePdf(b64(key), { Filter: 'DCTDecode', ColorSpace: cs, BitsPerComponent: 8, ...(comps === 4 ? { Decode: [1, 0, 1, 0, 1, 0, 1, 0] } : {}) }, 32, 32);
      const { out, images } = await redactImage(pdf);
      expect(images, key).toBe(1);
      const { dict, data } = imageOf((await PDFDocument.load(out)).getPage(0));
      expect(dict.lookup(PDFName.of('ColorSpace'))?.toString(), key).toBe(`/${cs}`);
      expect(dict.has(PDFName.of('Decode')), key).toBe(comps === 4);
      expect(data.length).toBe(32 * 32 * comps);
      expect([...data.subarray(0, comps)]).toEqual(Array(comps).fill(0));
      expectQuarterBlack(pdf, out);
    }
  });
});

describe('redaction of images: filter chains, predictors, masks', () => {
  const W = 20;
  const gray = Uint8Array.from({ length: W * W }, (_, i) => 100 + (i % W)); // a horizontal ramp

  it('ASCIIHex + Flate with the TIFF predictor', async () => {
    // TIFF predictor 2: each sample the difference to its left neighbour
    const diff = gray.map((v, i) => (i % W ? (v - gray[i - 1]) & 255 : v));
    const hex = new TextEncoder().encode(Buffer.from(zlibSync(diff)).toString('hex') + '>');
    const pdf = await imagePdf(hex, { ColorSpace: 'DeviceGray', BitsPerComponent: 8, Filter: ['ASCIIHexDecode', 'FlateDecode'], DecodeParms: [null, { Predictor: 2, Colors: 1, BitsPerComponent: 8, Columns: W }] }, W, W);
    const { out, images, warnings } = await redactImage(pdf);
    expect(warnings).toEqual([]);
    expect(images).toBe(1);
    const { data } = imageOf((await PDFDocument.load(out)).getPage(0));
    expect(data[0]).toBe(0);
    expect(data[W - 1]).toBe(100 + W - 1);
    expect(data[(W - 1) * W]).toBe(100);
  });

  it('the soft mask of an image (its shape) is redacted too', async () => {
    const doc = await PDFDocument.create();
    const rgba = new Uint8Array(W * W * 4);
    for (let i = 0; i < W * W; i++) rgba.set([10, 20, 30, i % 3 ? 255 : 128], i * 4);
    const img = await doc.embedPng(encodePng(rgba, W, W, 4));
    doc.addPage([260, 260]).drawImage(img, { x: 50, y: 50, width: 160, height: 160 });
    const d2 = await PDFDocument.load(await doc.save());
    const r = await redactPage(d2, d2.getPage(0), [QUARTER], { redact: true });
    expect(r.images).toBe(1); // the mask is not counted as an image of its own
    const xo = d2.getPage(0).node.Resources()!.lookup(PDFName.of('XObject'), PDFDict);
    const im = d2.context.lookup(xo.values()[0] as PDFRef) as PDFRawStream;
    const sm = d2.context.lookup(im.dict.get(PDFName.of('SMask')) as PDFRef) as PDFRawStream;
    const alpha = unzlibSync(sm.getContents());
    expect(alpha[0]).toBe(0);
    expect(alpha[W - 1]).not.toBe(0);
    expect(alpha[W * W - 1]).not.toBe(0);
  });

  it('alternate images and the XMP of a redacted image are dropped', async () => {
    const doc = await PDFDocument.create();
    const alt = doc.context.register(doc.context.flateStream(gray, { Type: 'XObject', Subtype: 'Image', Width: W, Height: W, ColorSpace: 'DeviceGray', BitsPerComponent: 8 }));
    const meta = doc.context.register(doc.context.stream('<x:xmpmeta>secret</x:xmpmeta>', { Type: 'Metadata', Subtype: 'XML' }));
    const pdf = await imagePdf(zlibSync(gray), { ColorSpace: 'DeviceGray', BitsPerComponent: 8, Filter: 'FlateDecode', Alternates: [{ Image: alt }], Metadata: meta }, W, W);
    const { out } = await redactImage(pdf);
    const { dict } = imageOf((await PDFDocument.load(out)).getPage(0));
    expect(dict.has(PDFName.of('Alternates'))).toBe(false);
    expect(dict.has(PDFName.of('Metadata'))).toBe(false);
  });
});

describe('what a redacted file must not keep', () => {
  async function withThumb(): Promise<Uint8Array> {
    const doc = await PDFDocument.create();
    doc.setTitle('Договор с Ивановым П. С.');
    const font = await doc.embedFont('Helvetica');
    const page = doc.addPage([300, 200]);
    page.drawText('Secret 12345', { x: 30, y: 100, size: 14, font });
    const thumb = doc.context.register(doc.context.flateStream(new Uint8Array(30 * 20 * 3).fill(90), { Width: 30, Height: 20, ColorSpace: 'DeviceRGB', BitsPerComponent: 8 }));
    page.node.set(PDFName.of('Thumb'), thumb);
    page.node.set(PDFName.of('PieceInfo'), doc.context.obj({ App: { Private: 'Secret 12345', LastModified: '(D:2026)' } }));
    return doc.save();
  }

  it('the page thumbnail and private application data go; the title is not carried over', async () => {
    const out = await composePdf([await withThumb()], [{ src: 0, index: 0, rotate: 0, annotations: [{ type: 'redact', x: 20, y: 90, w: 140, h: 30, color: [0, 0, 0], rotate: 0 }] }]);
    const doc = await PDFDocument.load(out);
    expect(doc.getTitle()).toBeUndefined();
    const page = doc.getPage(0).node;
    expect(page.has(PDFName.of('Thumb'))).toBe(false);
    expect(page.has(PDFName.of('PieceInfo'))).toBe(false);
    if (hasCmd('pdftotext')) expect(pdfToText(out)).not.toContain('12345');
  });

  it('without redaction the title and the thumbnail stay', async () => {
    const out = await composePdf([await withThumb()], [{ src: 0, index: 0, rotate: 0 }]);
    const doc = await PDFDocument.load(out);
    expect(doc.getTitle()).toBe('Договор с Ивановым П. С.');
    expect(doc.getPage(0).node.get(PDFName.of('Thumb'))).toBeInstanceOf(PDFRef);
  });
});
