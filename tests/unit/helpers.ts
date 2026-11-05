import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { ConvertEnv } from '../../src/core/convert';
import type { FontKey, FontLoader } from '../../src/core/pdf-fonts';
import { createSubsetter, type Subsetter } from '../../src/core/hb-subset';

const root = join(__dirname, '../..');
const nm = (p: string): Uint8Array => new Uint8Array(readFileSync(join(root, 'node_modules', p)));

const OFFICE = '@docx-editor.dev/fonts/assets/';
const STYLE_FILE = { regular: 'Regular', bold: 'Bold', italic: 'Italic', boldItalic: 'BoldItalic' } as const;
const FAMILY_FILE = { sans: 'LiberationSans', serif: 'LiberationSerif', mono: 'LiberationMono', calibri: 'Carlito' } as const;

export const FONT_FILES = {
  ...Object.fromEntries(
    Object.entries(FAMILY_FILE).flatMap(([fam, file]) => Object.entries(STYLE_FILE).map(([st, sf]) => [`${fam}-${st}`, `${OFFICE}${file}-${sf}.ttf`])),
  ),
  'fallback-regular': 'dejavu-fonts-ttf/ttf/DejaVuSans.ttf',
  'fallback-bold': 'dejavu-fonts-ttf/ttf/DejaVuSans-Bold.ttf',
  cjk: '@docx-editor.dev/fonts-cjk/assets/NotoSansCJKjp-Regular.otf',
  emoji: '@expo-google-fonts/noto-emoji/400Regular/NotoEmoji_400Regular.ttf',
} as Record<FontKey, string>;

export const loadFonts: FontLoader = async (key) => nm(FONT_FILES[key]);

let subsetterPromise: Promise<Subsetter> | null = null;
export const loadSubsetter = (): Promise<Subsetter> => (subsetterPromise ??= createSubsetter(readFileSync(join(root, 'node_modules/harfbuzzjs/dist/harfbuzz-subset.wasm'))));

export const loadPdfjs = async () => (await import('pdfjs-dist/legacy/build/pdf.mjs')) as unknown as typeof import('pdfjs-dist');

export const env: ConvertEnv = { pdfjs: loadPdfjs, fonts: loadFonts };
/** Same as the browser engine: HarfBuzz subsetting enabled. */
export const envWithSubsetter = async (): Promise<ConvertEnv> => ({ ...env, subsetter: await loadSubsetter() });

export const hasCmd = (cmd: string): boolean => {
  try {
    execFileSync('which', [cmd], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
};

/** Run a CLI on bytes written to a temp file; returns stdout or the produced file. */
export function withTemp<T>(fn: (dir: string) => T): T {
  const dir = mkdtempSync(join(tmpdir(), 'salatic-'));
  try {
    return fn(dir);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

export function pdfToText(pdf: Uint8Array): string {
  return withTemp((dir) => {
    const f = join(dir, 'in.pdf');
    writeFileSync(f, pdf);
    return execFileSync('pdftotext', ['-layout', f, '-'], { encoding: 'utf8' });
  });
}

export interface WordBox {
  page: number;
  text: string;
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

/** Word bounding boxes (points, y from the top) via `pdftotext -bbox`. */
export function pdfWords(pdf: Uint8Array): WordBox[] {
  const html = withTemp((dir) => {
    const f = join(dir, 'in.pdf');
    writeFileSync(f, pdf);
    return execFileSync('pdftotext', ['-bbox', f, '-'], { encoding: 'utf8' });
  });
  const out: WordBox[] = [];
  let page = 0;
  for (const line of html.split('\n')) {
    if (line.includes('<page ')) page++;
    const m = /<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)<\/word>/.exec(line);
    if (m) out.push({ page, text: m[5].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'), x0: +m[1], y0: +m[2], x1: +m[3], y1: +m[4] });
  }
  return out;
}

/** Convert with LibreOffice (headless) and return the produced file bytes. */
export function soffice(input: Uint8Array, ext: string, target: string): Uint8Array {
  return withTemp((dir) => {
    const f = join(dir, `in.${ext}`);
    writeFileSync(f, input);
    execFileSync('soffice', [`-env:UserInstallation=file://${dir}/profile`, '--headless', '--convert-to', target, '--outdir', dir, f], { stdio: 'ignore', timeout: 120_000 });
    const outExt = target.split(':')[0];
    const out = join(dir, `in.${outExt}`);
    if (!existsSync(out)) throw new Error(`LibreOffice did not produce ${outExt}`);
    return new Uint8Array(readFileSync(out));
  });
}

export const text = (b: Uint8Array): string => new TextDecoder().decode(b);
export const bytes = (s: string): Uint8Array => new TextEncoder().encode(s);

/** Rasterise every page with poppler (grayscale PGM) – an independent renderer for visual checks. */
export function renderPages(pdf: Uint8Array, dpi = 50): Buffer[] {
  return withTemp((dir) => {
    const f = join(dir, 'in.pdf');
    writeFileSync(f, pdf);
    execFileSync('pdftoppm', ['-r', String(dpi), '-gray', f, join(dir, 'p')]);
    return readdirSync(dir)
      .filter((n) => n.endsWith('.pgm'))
      .sort()
      .map((n) => readFileSync(join(dir, n)));
  });
}
