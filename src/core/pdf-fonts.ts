/**
 * Font handling for PDF output: loading on demand, metrics, per-character fallback
 * (Latin/Cyrillic → DejaVu → CJK → emoji) and HarfBuzz subsetting at embed time.
 */
import type { PDFDocument, PDFFont } from 'pdf-lib';
import type { FontFamily } from './model';
import type { Subsetter } from './hb-subset';

export type FontStyle = 'regular' | 'bold' | 'italic' | 'boldItalic';
export type FontKey = `${FontFamily}-${FontStyle}` | 'fallback-regular' | 'fallback-bold' | 'cjk' | 'emoji';

/** Loads the bytes of a font (TrueType/OpenType). */
export type FontLoader = (key: FontKey) => Promise<Uint8Array>;

interface FkGlyph {
  advanceWidth: number;
}
interface FkFont {
  unitsPerEm: number;
  ascent: number;
  descent: number;
  lineGap: number;
  'OS/2'?: { winAscent?: number; winDescent?: number; typoAscender?: number; typoDescender?: number };
  hasGlyphForCodePoint(cp: number): boolean;
  glyphForCodePoint(cp: number): FkGlyph;
}

export interface FontMetrics {
  /** Ascent and descent as fractions of the em (descent positive). */
  ascent: number;
  descent: number;
  /** Height of a single-spaced line as a fraction of the font size (Word semantics). */
  line: number;
}

/** Characters that are never drawn (layout handles them). */
const INVISIBLE = new Set([0x09, 0x0a, 0x0d]);

export const isCjk = (cp: number): boolean =>
  (cp >= 0x2e80 && cp <= 0x9fff) || (cp >= 0xa960 && cp <= 0xa97f) || (cp >= 0xac00 && cp <= 0xd7ff) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff00 && cp <= 0xffef) || (cp >= 0x20000 && cp <= 0x3ffff) || (cp >= 0x1100 && cp <= 0x11ff) || (cp >= 0x3000 && cp <= 0x303f);

export const isEmoji = (cp: number): boolean => (cp >= 0x1f000 && cp <= 0x1faff) || (cp >= 0x2600 && cp <= 0x27bf) || (cp >= 0x2b00 && cp <= 0x2bff) || cp === 0x200d || (cp >= 0xfe00 && cp <= 0xfe0f);

export function styleOf(bold: boolean, italic: boolean): FontStyle {
  return bold && italic ? 'boldItalic' : bold ? 'bold' : italic ? 'italic' : 'regular';
}

export class FontManager {
  private fk = new Map<FontKey, FkFont>();
  private bytes = new Map<FontKey, Uint8Array>();
  private advances = new Map<FontKey, Map<number, number>>();
  private used = new Map<FontKey, Set<number>>();
  private embedded = new Map<FontKey, PDFFont>();
  private failed = new Set<FontKey>();
  missing = false;

  constructor(
    private loader: FontLoader,
    private subsetter?: Subsetter,
  ) {}

  private async fontkit(): Promise<{ create(b: Uint8Array): FkFont }> {
    return (await import('@pdf-lib/fontkit')).default as unknown as { create(b: Uint8Array): FkFont };
  }

  async load(keys: Iterable<FontKey>): Promise<void> {
    const todo = [...new Set(keys)].filter((k) => !this.fk.has(k) && !this.failed.has(k));
    if (!todo.length) return;
    const fontkit = await this.fontkit();
    await Promise.all(
      todo.map(async (k) => {
        try {
          const data = await this.loader(k);
          this.bytes.set(k, data);
          this.fk.set(k, fontkit.create(data));
          this.advances.set(k, new Map());
        } catch {
          this.failed.add(k); // fallbacks are optional
        }
      }),
    );
  }

  loaded(key: FontKey): boolean {
    return this.fk.has(key);
  }

  has(key: FontKey, cp: number): boolean {
    return this.fk.get(key)?.hasGlyphForCodePoint(cp) ?? false;
  }

  /** Font that will draw `cp` for the given primary font (null if nothing can). */
  resolve(primary: FontKey, cp: number): FontKey | null {
    if (this.has(primary, cp)) return primary;
    const bold = primary.endsWith('bold') || primary.endsWith('boldItalic');
    for (const k of [bold ? 'fallback-bold' : 'fallback-regular', 'fallback-regular', 'cjk', 'emoji'] as FontKey[]) if (this.has(k, cp)) return k;
    return null;
  }

  /** Which fallback fonts are needed for this text (to load them before layout). */
  fallbacksFor(primary: FontKey, cp: number): FontKey[] {
    if (this.has(primary, cp) || INVISIBLE.has(cp)) return [];
    if (isCjk(cp)) return ['cjk'];
    if (isEmoji(cp)) return ['fallback-regular', 'emoji'];
    const bold = primary.endsWith('bold') || primary.endsWith('boldItalic');
    return [bold ? 'fallback-bold' : 'fallback-regular', 'cjk', 'emoji'];
  }

  /** Advance width in 1/1000 em. */
  advance(key: FontKey, cp: number): number {
    const cache = this.advances.get(key)!;
    let w = cache.get(cp);
    if (w === undefined) {
      const f = this.fk.get(key)!;
      w = (f.glyphForCodePoint(cp).advanceWidth * 1000) / f.unitsPerEm;
      cache.set(cp, w);
    }
    return w;
  }

  width(key: FontKey, text: string, size: number): number {
    let w = 0;
    for (const ch of text) w += this.advance(key, ch.codePointAt(0)!);
    return (w * size) / 1000;
  }

  metrics(key: FontKey): FontMetrics {
    const f = this.fk.get(key)!;
    const os2 = f['OS/2'];
    const upm = f.unitsPerEm;
    // Word lays out lines with the Windows metrics; fall back to hhea
    const asc = (os2?.winAscent || f.ascent) / upm;
    const desc = (os2?.winDescent || Math.abs(f.descent)) / upm;
    return { ascent: asc, descent: desc, line: asc + desc };
  }

  markUsed(key: FontKey, text: string): void {
    let set = this.used.get(key);
    if (!set) this.used.set(key, (set = new Set()));
    for (const ch of text) set.add(ch.codePointAt(0)!);
  }

  /** Embed every font that has been used, subset to the used characters. */
  async embedAll(pdf: PDFDocument, subset = true): Promise<void> {
    for (const [key, cps] of this.used) {
      if (this.embedded.has(key) || !this.bytes.has(key)) continue;
      const full = this.bytes.get(key)!;
      cps.add(0x20);
      let data = full;
      if (subset && this.subsetter) {
        try {
          data = this.subsetter(full, cps);
        } catch {
          data = full; // embedding the whole font is larger but always correct
        }
      }
      // liga off: widths are measured per character, ligatures would shift justified text
      this.embedded.set(key, await pdf.embedFont(data, { subset: false, features: { liga: false, dlig: false, clig: false } }));
    }
  }

  pdfFont(key: FontKey): PDFFont {
    const f = this.embedded.get(key);
    if (!f) throw new Error(`Font ${key} was not embedded`);
    return f;
  }
}
