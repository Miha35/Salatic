/** Lines of existing PDF text (from pdf.js text items) that the PDF editor can replace in place. */
import type { FontFamily } from './model';

/** A line of existing PDF text (user space, unrotated) that can be edited in place. */
export interface TextSeg {
  id: string;
  x: number;
  y: number; // baseline
  w: number;
  size: number;
  text: string;
  bold: boolean;
  italic: boolean;
  font: FontFamily;
}

/** Map a PDF font (name and flags from pdf.js) to one of our families. */
export function familyFor(name: string, serif?: boolean, mono?: boolean): FontFamily {
  const n = name.replace(/^[A-Z]{6}\+/, '');
  if (mono || /courier|mono|consol|menlo/i.test(n)) return 'mono';
  if (/calibri|carlito/i.test(n)) return 'calibri';
  if (/arial|helvet|verdana|tahoma|segoe|sans|roboto|open ?sans|pt ?sans|dejavu ?sans|liberation ?sans/i.test(n)) return 'sans';
  if (serif || /times|serif|georgia|garamond|cambria|book|minion|palatino|caladea/i.test(n)) return 'serif';
  return 'sans';
}

/** Group pdf.js text items into editable lines. */
export function buildSegments(items: { str: string; transform: number[]; width: number; fontName: string; hasEOL?: boolean }[], fontInfo: (name: string) => { family: FontFamily; bold: boolean; italic: boolean }): TextSeg[] {
  const segs: TextSeg[] = [];
  let cur: (TextSeg & { x1: number; fontName: string }) | null = null;
  const flush = (): void => {
    if (cur && cur.text.trim()) segs.push({ id: `${segs.length}:${Math.round(cur.x)}:${Math.round(cur.y)}`, x: cur.x, y: cur.y, w: cur.x1 - cur.x, size: cur.size, text: cur.text.trim(), bold: cur.bold, italic: cur.italic, font: cur.font });
    cur = null;
  };
  for (const it of items) {
    const [a, b, c, d, e, f] = it.transform;
    const size = Math.hypot(c, d);
    // Only upright horizontal text can be replaced in place
    if (!it.str || Math.abs(b) > 0.01 * Math.abs(a) || Math.abs(c) > 0.01 * Math.abs(d) || a <= 0 || size < 1) {
      flush();
      continue;
    }
    if (!it.str.trim()) {
      if (cur) cur.text += ' ';
      if (it.hasEOL) flush();
      continue;
    }
    // A change of font (bold word, italic phrase) starts a new segment so that its style survives editing
    const sameLine = cur && cur.fontName === it.fontName && Math.abs(cur.y - f) < size * 0.25 && Math.abs(cur.size - size) < size * 0.3 && e >= cur.x1 - size * 0.5 && e - cur.x1 < size * 1.2;
    if (!sameLine) {
      flush();
      const fi = fontInfo(it.fontName);
      cur = { id: '', x: e, y: f, w: 0, x1: e + it.width, size, text: it.str, bold: fi.bold, italic: fi.italic, font: fi.family, fontName: it.fontName };
    } else if (cur) {
      if (e - cur.x1 > size * 0.15 && !cur.text.endsWith(' ') && !it.str.startsWith(' ')) cur.text += ' ';
      cur.text += it.str;
      cur.x1 = Math.max(cur.x1, e + it.width);
    }
    if (it.hasEOL) flush();
  }
  flush();
  return segs;
}
