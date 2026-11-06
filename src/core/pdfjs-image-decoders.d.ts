/** The parts of pdf.js's stand-alone image decoders (pdfjs-dist/image_decoders) that we use. */
declare module 'pdfjs-dist/image_decoders/pdf.image_decoders.mjs' {
  interface WasmOptions {
    useWasm: boolean;
    useWorkerFetch: boolean;
    wasmUrl: string;
    getHandler?: () => { sendWithPromise(action: string, data: { kind: string; filename: string }): Promise<Uint8Array> };
  }
  export class JpxImage {
    static setOptions(o: WasmOptions): void;
    static get instance(): JpxImage;
    decode(bytes: Uint8Array, o?: { numComponents?: number; isIndexedColormap?: boolean; smaskInData?: boolean; reducePower?: number }): Promise<Uint8Array>;
  }
  export class Jbig2Image {
    static setOptions(o: WasmOptions): void;
    static get instance(): Jbig2Image;
    decode(
      bytes: Uint8Array,
      width: number,
      height: number,
      globals: Uint8Array | null,
      ccitt?: { K: number; EndOfLine: boolean; EncodedByteAlign: boolean; Columns: number; Rows: number; EndOfBlock: boolean; BlackIs1: boolean },
    ): Promise<Uint8Array>;
  }
  export class JpegImage {
    constructor(o?: { decodeTransform?: Int32Array | null; colorTransform?: number });
    width: number;
    height: number;
    numComponents: number;
    parse(data: Uint8Array): void;
    getData(o: { width: number; height: number; forceRGB?: boolean; forceRGBA?: boolean; isSourcePDF?: boolean }): Uint8Array | Uint8ClampedArray;
  }
}
