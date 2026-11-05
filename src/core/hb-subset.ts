/**
 * Font subsetting with HarfBuzz (hb-subset, compiled to WebAssembly by harfbuzzjs).
 *
 * pdf-lib's own subsetter (@pdf-lib/fontkit) corrupts many real-world fonts
 * (CFF/CID fonts such as Noto Sans CJK, TrueType fonts with odd-length glyphs such
 * as Carlito). We therefore subset with HarfBuzz – the reference implementation used
 * by browsers – and embed the already-small font into the PDF as is.
 */

/** Subset a font to the given code points. Returns a standalone OpenType font. */
export type Subsetter = (font: Uint8Array, codepoints: Iterable<number>) => Uint8Array;

interface HbExports {
  memory: WebAssembly.Memory;
  malloc(size: number): number;
  free(ptr: number): void;
  hb_blob_create(data: number, length: number, mode: number, userData: number, destroy: number): number;
  hb_blob_destroy(blob: number): void;
  hb_blob_get_data(blob: number, lengthPtr: number): number;
  hb_blob_get_length(blob: number): number;
  hb_face_create(blob: number, index: number): number;
  hb_face_destroy(face: number): void;
  hb_face_reference_blob(face: number): number;
  hb_set_add(set: number, codepoint: number): void;
  hb_subset_input_create_or_fail(): number;
  hb_subset_input_destroy(input: number): void;
  hb_subset_input_unicode_set(input: number): number;
  hb_subset_input_set_flags(input: number, flags: number): void;
  hb_subset_or_fail(face: number, input: number): number;
}

const HB_MEMORY_MODE_READONLY = 1;
const HB_SUBSET_FLAGS_NO_HINTING = 0x1;
const HB_SUBSET_FLAGS_NOTDEF_OUTLINE = 0x40;

export async function createSubsetter(wasm: BufferSource | WebAssembly.Module): Promise<Subsetter> {
  const instance = wasm instanceof WebAssembly.Module ? await WebAssembly.instantiate(wasm, {}) : (await WebAssembly.instantiate(wasm, {})).instance;
  const hb = instance.exports as unknown as HbExports;
  return (font, codepoints) => {
    const fontPtr = hb.malloc(font.byteLength);
    if (!fontPtr) throw new Error('hb-subset: out of memory');
    new Uint8Array(hb.memory.buffer).set(font, fontPtr);
    const blob = hb.hb_blob_create(fontPtr, font.byteLength, HB_MEMORY_MODE_READONLY, 0, 0);
    const face = hb.hb_face_create(blob, 0);
    hb.hb_blob_destroy(blob);
    const input = hb.hb_subset_input_create_or_fail();
    try {
      if (!input) throw new Error('hb-subset: cannot create input');
      // Hinting is useless inside PDFs; an outlined .notdef makes missing glyphs visible
      hb.hb_subset_input_set_flags(input, HB_SUBSET_FLAGS_NO_HINTING | HB_SUBSET_FLAGS_NOTDEF_OUTLINE);
      const set = hb.hb_subset_input_unicode_set(input);
      for (const cp of codepoints) hb.hb_set_add(set, cp);
      const subset = hb.hb_subset_or_fail(face, input);
      if (!subset) throw new Error('hb-subset: subsetting failed');
      const out = hb.hb_face_reference_blob(subset);
      const ptr = hb.hb_blob_get_data(out, 0);
      const len = hb.hb_blob_get_length(out);
      // Memory may have grown during subsetting: take a fresh view
      const result = new Uint8Array(hb.memory.buffer).slice(ptr, ptr + len);
      hb.hb_blob_destroy(out);
      hb.hb_face_destroy(subset);
      if (!len) throw new Error('hb-subset: empty result');
      return result;
    } finally {
      if (input) hb.hb_subset_input_destroy(input);
      hb.hb_face_destroy(face);
      hb.free(fontPtr);
    }
  };
}
