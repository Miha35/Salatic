/**
 * Intermediate representations shared by every reader and writer.
 * Readers turn bytes into a `Doc` (flow document) or a `Workbook` (tabular data);
 * writers turn those back into bytes. N readers × M writers = N·M conversions.
 */

export type Align = 'left' | 'center' | 'right' | 'justify';

export interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  code?: boolean;
  sup?: boolean;
  sub?: boolean;
  /** Absolute URL (http/https/mailto). */
  link?: string;
  /** Hex colour without '#', e.g. "FF0000". */
  color?: string;
}

export interface ListItem {
  blocks: Block[];
}

export interface TableCell {
  blocks: Block[];
  header?: boolean;
  colspan?: number;
}

export interface TableRow {
  cells: TableCell[];
}

export type Block =
  | { type: 'paragraph'; runs: TextRun[]; align?: Align }
  | { type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; runs: TextRun[]; align?: Align }
  | { type: 'list'; ordered: boolean; items: ListItem[]; start?: number }
  | { type: 'table'; rows: TableRow[] }
  /** width/height are the intended display size in CSS pixels (96 dpi). */
  | { type: 'image'; data: Uint8Array; mime: string; width: number; height: number; alt?: string }
  | { type: 'code'; text: string }
  | { type: 'quote'; blocks: Block[] }
  | { type: 'hr' }
  | { type: 'pagebreak' };

export interface Doc {
  title?: string;
  blocks: Block[];
}

export type CellValue = string | number | boolean | null;

export interface Sheet {
  name: string;
  rows: CellValue[][];
  /** True when the first row holds column names. */
  header?: boolean;
}

export interface Workbook {
  sheets: Sheet[];
}

export type Content = { kind: 'doc'; doc: Doc } | { kind: 'book'; book: Workbook };

// ---------------------------------------------------------------------------
// Helpers

export const run = (text: string, style: Omit<TextRun, 'text'> = {}): TextRun => ({ text, ...style });

export function runsText(runs: TextRun[]): string {
  let s = '';
  for (const r of runs) s += r.text;
  return s;
}

/** Plain text of any block (used for tables → cells, search, etc.). */
export function blockText(b: Block): string {
  switch (b.type) {
    case 'paragraph':
    case 'heading':
      return runsText(b.runs);
    case 'list':
      return b.items.map((it) => it.blocks.map(blockText).join('\n')).join('\n');
    case 'table':
      return b.rows.map((r) => r.cells.map((c) => c.blocks.map(blockText).join('\n')).join('\t')).join('\n');
    case 'code':
      return b.text;
    case 'quote':
      return b.blocks.map(blockText).join('\n');
    case 'image':
      return b.alt ?? '';
    default:
      return '';
  }
}

const sameStyle = (a: TextRun, b: TextRun): boolean =>
  !!a.bold === !!b.bold &&
  !!a.italic === !!b.italic &&
  !!a.underline === !!b.underline &&
  !!a.strike === !!b.strike &&
  !!a.code === !!b.code &&
  !!a.sup === !!b.sup &&
  !!a.sub === !!b.sub &&
  a.link === b.link &&
  a.color === b.color;

/** Merge adjacent runs with identical formatting and drop empty ones. */
export function normalizeRuns(runs: TextRun[]): TextRun[] {
  const out: TextRun[] = [];
  for (const r of runs) {
    if (!r.text) continue;
    const prev = out[out.length - 1];
    if (prev && sameStyle(prev, r)) prev.text += r.text;
    else out.push({ ...r });
  }
  return out;
}

/** Trim leading/trailing whitespace of a run list (not inside). */
export function trimRuns(runs: TextRun[]): TextRun[] {
  const out = normalizeRuns(runs);
  while (out.length && !out[0].text.replace(/^[ \t\r\n]+/, '')) out.shift();
  while (out.length && !out[out.length - 1].text.replace(/[ \t\r\n]+$/, '')) out.pop();
  if (out.length) {
    out[0] = { ...out[0], text: out[0].text.replace(/^[ \t\r\n]+/, '') };
    const last = out.length - 1;
    out[last] = { ...out[last], text: out[last].text.replace(/[ \t\r\n]+$/, '') };
  }
  return out;
}

export function cellToString(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : '';
  return String(v);
}

/** Workbook → document: one heading + table per sheet. */
export function bookToDoc(book: Workbook, title?: string): Doc {
  const blocks: Block[] = [];
  for (const sheet of book.sheets) {
    if (book.sheets.length > 1) blocks.push({ type: 'heading', level: 2, runs: [run(sheet.name)] });
    const width = sheet.rows.reduce((m, r) => Math.max(m, r.length), 0);
    if (!width) {
      blocks.push({ type: 'paragraph', runs: [run('(пустой лист)', { italic: true })] });
      continue;
    }
    blocks.push({
      type: 'table',
      rows: sheet.rows.map((r, ri) => ({
        cells: Array.from({ length: width }, (_, ci) => ({
          header: ri === 0 && !!sheet.header,
          blocks: [{ type: 'paragraph', runs: [run(cellToString(r[ci] ?? null))] } as Block],
        })),
      })),
    });
  }
  return { title, blocks };
}

/** Document → workbook: every table becomes a sheet. */
export function docToBook(doc: Doc): Workbook {
  const tables: Extract<Block, { type: 'table' }>[] = [];
  const visit = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.type === 'table') tables.push(b);
      else if (b.type === 'quote') visit(b.blocks);
      else if (b.type === 'list') b.items.forEach((it) => visit(it.blocks));
    }
  };
  visit(doc.blocks);
  return {
    sheets: tables.map((t, i) => {
      const rows: CellValue[][] = t.rows.map((r) => {
        const cells: CellValue[] = [];
        for (const c of r.cells) {
          cells.push(c.blocks.map(blockText).join('\n'));
          for (let k = 1; k < (c.colspan ?? 1); k++) cells.push('');
        }
        return cells;
      });
      return { name: `Таблица ${i + 1}`, rows, header: t.rows[0]?.cells.every((c) => c.header) ?? false };
    }),
  };
}
