/**
 * Image formats browsers cannot decode: TIFF (every page; UTIF) and HEIC/HEIF
 * (libheif compiled to WebAssembly, loaded only when needed). Output is RGBA.
 */
import { ConvertError } from './util';
export { isHeic, isTiff } from './image-sniff';

export interface RgbaImage {
  width: number;
  height: number;
  data: Uint8Array;
}

const MAX_PIXELS = 100_000_000;

interface UtifIfd {
  [tag: string]: unknown;
  width: number;
  height: number;
  data: Uint8Array;
}
interface Utif {
  decode(buffer: ArrayBuffer): UtifIfd[];
  decodeImage(buffer: ArrayBuffer, ifd: UtifIfd, ifds?: UtifIfd[]): void;
  toRGBA8(ifd: UtifIfd): Uint8Array;
  encodeImage(rgba: Uint8Array, w: number, h: number, metadata?: Record<string, unknown>): ArrayBuffer;
}

async function utif(): Promise<Utif> {
  const mod = (await import('utif2')) as unknown as { default?: Utif } & Utif;
  return mod.default ?? mod;
}

/** Every page of a TIFF (thumbnails and reduced-resolution images are skipped). */
export async function decodeTiff(bytes: Uint8Array): Promise<RgbaImage[]> {
  const U = await utif();
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  let ifds: UtifIfd[];
  try {
    ifds = U.decode(buf);
  } catch {
    throw new ConvertError('Файл TIFF повреждён.');
  }
  const pages = ifds.filter((ifd) => {
    const sub = (ifd.t254 as number[] | undefined)?.[0] ?? 0;
    return ifd.t256 !== undefined && ifd.t257 !== undefined && (sub & 1) === 0;
  });
  const out: RgbaImage[] = [];
  for (const ifd of pages.length ? pages : ifds) {
    const w = (ifd.t256 as number[] | undefined)?.[0] ?? 0;
    const h = (ifd.t257 as number[] | undefined)?.[0] ?? 0;
    if (!w || !h) continue;
    if (w * h > MAX_PIXELS) throw new ConvertError('Изображение TIFF слишком большое для обработки в браузере.');
    try {
      U.decodeImage(buf, ifd, ifds);
      const rgba = U.toRGBA8(ifd);
      if (rgba.length >= w * h * 4) out.push({ width: ifd.width || w, height: ifd.height || h, data: rgba.subarray(0, w * h * 4) });
    } catch {
      /* an undecodable page (exotic compression) is skipped below */
    }
  }
  if (!out.length) throw new ConvertError('Не удалось декодировать TIFF: неподдерживаемое сжатие или повреждённый файл.');
  return out;
}

export async function encodeTiff(img: RgbaImage): Promise<Uint8Array> {
  const U = await utif();
  return new Uint8Array(U.encodeImage(img.data, img.width, img.height));
}

interface HeifImage {
  get_width(): number;
  get_height(): number;
  display(target: { data: Uint8ClampedArray; width: number; height: number }, cb: (r: { data: Uint8ClampedArray } | null) => void): void;
  free(): void;
}
interface Libheif {
  HeifDecoder: new () => { decode(data: Uint8Array): HeifImage[] };
}

let libheifPromise: Promise<Libheif> | null = null;

/** libheif needs its wasm binary; the caller provides it (fetched URL in browsers, a file in Node). */
export function loadLibheif(wasm: () => Promise<Uint8Array>): Promise<Libheif> {
  libheifPromise ??= (async () => {
    const [mod, wasmBinary] = await Promise.all([import('libheif-js/libheif-wasm/libheif.js'), wasm()]);
    const factory = ((mod as { default?: unknown }).default ?? mod) as (o: { wasmBinary: Uint8Array }) => Libheif | Promise<Libheif>;
    return await factory({ wasmBinary });
  })();
  libheifPromise.catch(() => (libheifPromise = null));
  return libheifPromise;
}

export async function decodeHeic(bytes: Uint8Array, wasm: () => Promise<Uint8Array>): Promise<RgbaImage[]> {
  const lib = await loadLibheif(wasm);
  const decoder = new lib.HeifDecoder();
  let images: HeifImage[];
  try {
    images = decoder.decode(bytes);
  } catch {
    images = [];
  }
  if (!images.length) throw new ConvertError('Не удалось прочитать HEIC: файл повреждён или это не изображение HEIF.');
  const out: RgbaImage[] = [];
  try {
    for (const img of images) {
      const width = img.get_width();
      const height = img.get_height();
      if (width * height > MAX_PIXELS) throw new ConvertError('Изображение HEIC слишком большое для обработки в браузере.');
      const data = await new Promise<Uint8ClampedArray>((resolve, reject) =>
        img.display({ data: new Uint8ClampedArray(width * height * 4), width, height }, (r) => (r ? resolve(r.data) : reject(new ConvertError('Не удалось декодировать HEIC.')))),
      );
      out.push({ width, height, data: new Uint8Array(data.buffer, data.byteOffset, data.byteLength) });
    }
  } finally {
    for (const img of images) img.free();
  }
  return out;
}
