/**
 * Decoders for the image filters of PDF that pdf-lib cannot undo (DCT, JPX, JBIG2, CCITT fax),
 * on top of pdf.js's stand-alone image decoders (OpenJPEG and jbig2 compiled to WebAssembly).
 * Used by redaction to rewrite the pixels of such images.
 */
import type { ImageFilterDecoder } from './pdf-redact';

type Decoders = typeof import('pdfjs-dist/image_decoders/pdf.image_decoders.mjs');

export interface ImageFilterOptions {
  /** Where openjpeg.wasm, jbig2.wasm and their JS fallbacks live (with a trailing slash). */
  wasmUrl: string;
  /** Fetches one of those files by name; without it they are fetched from `wasmUrl`. */
  fetchBinary?: (filename: string) => Promise<Uint8Array>;
}

/** A JBIG2 file (not the embedded organisation PDF uses) starts with an 8-byte id, flags and maybe a page count. */
function stripJbig2Header(b: Uint8Array): Uint8Array {
  const id = [0x97, 0x4a, 0x42, 0x32, 0x0d, 0x0a, 0x1a, 0x0a];
  if (b.length < 9 || id.some((v, i) => b[i] !== v)) return b;
  return b.subarray(b[8] & 2 ? 9 : 13);
}

/** The decoders hand out Uint8ClampedArrays at times; pdf-lib (and our sample code) want Uint8Array. */
const bytesOf = (a: ArrayBufferView): Uint8Array => (a instanceof Uint8Array && a.constructor === Uint8Array ? a : new Uint8Array(a.buffer, a.byteOffset, a.byteLength));

export function pdfjsImageDecoder(opts: ImageFilterOptions): ImageFilterDecoder {
  let mod: Promise<Decoders> | null = null;
  const load = (): Promise<Decoders> =>
    (mod ??= import('pdfjs-dist/image_decoders/pdf.image_decoders.mjs').then((m) => {
      const fetchBinary = opts.fetchBinary;
      // the options are static on their common base class; set once through either subclass
      m.JpxImage.setOptions(
        fetchBinary
          ? { useWasm: true, useWorkerFetch: false, wasmUrl: opts.wasmUrl, getHandler: () => ({ sendWithPromise: (_action, d) => fetchBinary(d.filename) }) }
          : { useWasm: true, useWorkerFetch: true, wasmUrl: opts.wasmUrl },
      );
      return m;
    }));

  return async (filter, data, info) => {
    const m = await load();
    const { width: w, height: h } = info;
    switch (filter) {
      case 'DCTDecode': {
        // the samples as stored (YCbCr/YCCK colour transforms undone, no Adobe inversion): the image keeps
        // its colour space and Decode array, so nothing but the redacted pixels changes
        const jpeg = new m.JpegImage({ colorTransform: info.colorTransform ?? -1 });
        jpeg.parse(data);
        if (jpeg.width !== w || jpeg.height !== h) return null;
        const out = jpeg.getData({ width: w, height: h, isSourcePDF: true });
        const comps = jpeg.numComponents;
        return out.length === w * h * comps ? { data: bytesOf(out), comps, bpc: 8 } : null;
      }
      case 'JPXDecode': {
        const out = bytesOf(await m.JpxImage.instance.decode(data, { numComponents: info.comps, isIndexedColormap: info.indexed, smaskInData: info.smaskInData, reducePower: 0 }));
        const comps = out.length / (w * h);
        return Number.isInteger(comps) && comps >= 1 && comps <= 4 ? { data: out, comps, bpc: 8 } : null;
      }
      case 'JBIG2Decode': {
        const out = bytesOf(await m.Jbig2Image.instance.decode(stripJbig2Header(data), w, h, info.globals ? stripJbig2Header(info.globals) : null));
        return out.length >= Math.ceil(w / 8) * h ? { data: out, comps: 1, bpc: 1 } : null;
      }
      case 'CCITTFaxDecode': {
        const p = info.params ?? {};
        const ccitt = {
          K: Number(p.K ?? 0) || 0,
          EndOfLine: !!p.EndOfLine,
          EncodedByteAlign: !!p.EncodedByteAlign,
          Columns: Number(p.Columns ?? 0) || 1728,
          Rows: Number(p.Rows ?? 0) || 0,
          EndOfBlock: p.EndOfBlock === undefined ? true : !!p.EndOfBlock,
          BlackIs1: !!p.BlackIs1,
        };
        const out = bytesOf(await m.Jbig2Image.instance.decode(data, w, h, null, ccitt));
        return out.length >= Math.ceil(w / 8) * h ? { data: out, comps: 1, bpc: 1 } : null;
      }
    }
    return null;
  };
}
