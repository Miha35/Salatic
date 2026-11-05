/** Image formats decoded by our own code: TIFF (all common compressions, multi-page) and HEIC. */
import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { decodeHeic, decodeTiff, encodeTiff, isHeic, isTiff } from '../../src/core/image-decode';
import { detectFormat } from '../../src/core/formats';
import { sniffImageMime } from '../../src/core/util';
import { hasCmd, withTemp } from './helpers';
import { makeHeic } from './heic-fixture';

const heifWasm = async (): Promise<Uint8Array> => new Uint8Array(readFileSync(join(__dirname, '../../node_modules/libheif-js/libheif-wasm/libheif.wasm')));

/** Make a TIFF with Pillow and return it with Pillow's own RGBA pixels of every page. */
function pilTiff(compression: string, mode: string, pages = 1): { tiff: Uint8Array; rgba: Uint8Array[]; w: number; h: number } {
  return withTemp((dir) => {
    const script = `
import sys
from PIL import Image, ImageDraw
w, h = 37, 23
frames = []
for k in range(${pages}):
    im = Image.new('RGB', (w, h))
    px = im.load()
    for y in range(h):
        for x in range(w):
            px[x, y] = ((x * 7 + k * 40) % 256, (y * 11) % 256, (x * y + k * 90) % 256)
    ImageDraw.Draw(im).rectangle([3 + k, 3, 12 + k, 9], fill=(250, 250, 250))
    frames.append(im.convert('${mode}'))
kw = {} if '${compression}' == 'raw' else {'compression': '${compression}'}
frames[0].save(sys.argv[1], save_all=True, append_images=frames[1:], **kw)
with open(sys.argv[2], 'wb') as f:
    for fr in frames:
        f.write(fr.convert('RGBA').tobytes())
`;
    const t = join(dir, 'a.tif');
    const r = join(dir, 'a.raw');
    execFileSync('python3', ['-I', '-c', script, t, r]);
    const raw = new Uint8Array(readFileSync(r));
    const size = 37 * 23 * 4;
    return { tiff: new Uint8Array(readFileSync(t)), rgba: Array.from({ length: pages }, (_, k) => raw.subarray(k * size, (k + 1) * size)), w: 37, h: 23 };
  });
}

const mae = (a: Uint8Array, b: Uint8Array): number => {
  let s = 0;
  let n = 0;
  for (let i = 0; i < a.length; i++) {
    if (i % 4 === 3) continue;
    s += Math.abs(a[i] - b[i]);
    n++;
  }
  return s / n;
};

describe.skipIf(!hasCmd('python3'))('TIFF', () => {
  const cases: [string, string, number][] = [
    ['raw', 'RGB', 0],
    ['tiff_lzw', 'RGB', 0],
    ['packbits', 'RGB', 0],
    ['tiff_adobe_deflate', 'RGB', 0],
    ['tiff_lzw', 'L', 0],
    ['tiff_lzw', 'RGBA', 0],
    ['jpeg', 'RGB', 6],
    ['group4', '1', 0],
  ];
  for (const [compression, mode, tolerance] of cases) {
    it(`decodes ${compression} ${mode}`, async () => {
      const { tiff, rgba, w, h } = pilTiff(compression, mode);
      expect(isTiff(tiff)).toBe(true);
      expect(detectFormat('scan.bin', tiff).format).toBe('tiff');
      const [page] = await decodeTiff(tiff);
      expect([page.width, page.height]).toEqual([w, h]);
      expect(mae(page.data, rgba[0])).toBeLessThanOrEqual(tolerance + 0.01);
    });
  }

  it('decodes every page of a multi-page TIFF', async () => {
    const { tiff, rgba } = pilTiff('tiff_lzw', 'RGB', 3);
    const pages = await decodeTiff(tiff);
    expect(pages.length).toBe(3);
    pages.forEach((p, k) => expect(mae(p.data, rgba[k])).toBe(0));
  });

  it('encodes TIFF that Pillow reads back', async () => {
    const data = new Uint8Array(4 * 3 * 4).map((_, i) => (i % 4 === 3 ? 255 : i * 5));
    const tiff = await encodeTiff({ width: 4, height: 3, data });
    const back = withTemp((dir) => {
      const f = join(dir, 'x.tif');
      require('node:fs').writeFileSync(f, tiff);
      return execFileSync('python3', ['-I', '-c', 'import sys;from PIL import Image;im=Image.open(sys.argv[1]).convert("RGBA");sys.stdout.buffer.write(im.tobytes())', f]);
    });
    expect([...back]).toEqual([...data]);
  });

  it('rejects broken files with a readable error', async () => {
    await expect(decodeTiff(new Uint8Array([0x49, 0x49, 0x2a, 0, 8, 0, 0, 0, 0, 0]))).rejects.toThrow(/TIFF/);
    expect(sniffImageMime(pilTiff('raw', 'RGB').tiff)).toBe('image/tiff');
  });
});

describe.skipIf(!hasCmd('ffmpeg'))('HEIC', () => {
  it('decodes an HEVC-coded HEIF still close to its source', async () => {
    const { heic, png } = makeHeic(64, 48);
    expect(isHeic(heic)).toBe(true);
    expect(detectFormat('IMG_0001.HEIC', heic).format).toBe('heic');
    expect(detectFormat('photo', heic).format).toBe('heic');
    const [img] = await decodeHeic(heic, heifWasm);
    expect([img.width, img.height]).toEqual([64, 48]);
    const ref = withTemp((dir) => {
      const f = join(dir, 's.png');
      require('node:fs').writeFileSync(f, png);
      return new Uint8Array(execFileSync('ffmpeg', ['-v', 'error', '-i', f, '-f', 'rawvideo', '-pix_fmt', 'rgba', '-']));
    });
    // lossy HEVC with 4:2:0 chroma: small average error on a colourful test pattern
    expect(mae(img.data, ref)).toBeLessThan(12);
  });

  it('does not mistake AVIF for HEIC and rejects garbage', async () => {
    const avif = new Uint8Array([0, 0, 0, 0x1c, 0x66, 0x74, 0x79, 0x70, 0x6d, 0x69, 0x66, 0x31, 0, 0, 0, 0, 0x6d, 0x69, 0x66, 0x31, 0x61, 0x76, 0x69, 0x66, 0x6d, 0x69, 0x61, 0x66]);
    expect(isHeic(avif)).toBe(false);
    await expect(decodeHeic(new Uint8Array([0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70, 0x68, 0x65, 0x69, 0x63, 0, 0, 0, 0, 0x6d, 0x69, 0x66, 0x31, 0x68, 0x65, 0x69, 0x63]), heifWasm)).rejects.toThrow(/HEIC/);
  });
});
