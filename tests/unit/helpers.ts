import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { ConvertEnv } from '../../src/core/convert';
import type { FontKey, FontLoader } from '../../src/core/writers/pdf';

const root = join(__dirname, '../..');
const nm = (p: string): Uint8Array => new Uint8Array(readFileSync(join(root, 'node_modules', p)));

export const FONT_FILES: Record<FontKey, string> = {
  regular: 'pdfjs-dist/standard_fonts/LiberationSans-Regular.ttf',
  bold: 'pdfjs-dist/standard_fonts/LiberationSans-Bold.ttf',
  italic: 'pdfjs-dist/standard_fonts/LiberationSans-Italic.ttf',
  boldItalic: 'pdfjs-dist/standard_fonts/LiberationSans-BoldItalic.ttf',
  mono: 'dejavu-fonts-ttf/ttf/DejaVuSansMono.ttf',
};

export const loadFonts: FontLoader = async (key) => nm(FONT_FILES[key]);

export const loadPdfjs = async () => (await import('pdfjs-dist/legacy/build/pdf.mjs')) as unknown as typeof import('pdfjs-dist');

export const env: ConvertEnv = { pdfjs: loadPdfjs, fonts: loadFonts };

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
