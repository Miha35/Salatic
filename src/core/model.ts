/**
 * Intermediate representations shared by every reader and writer.
 * Readers turn bytes into a `Doc` (flow document) or a `Workbook` (tabular data);
 * writers turn those back into bytes. N readers × M writers = N·M conversions.
 */

export type Align = 'left' | 'center' | 'right' | 'justify';

/**
 * Font families with metric-compatible free fonts:
 * sans ≈ Arial/Helvetica, serif ≈ Times New Roman, mono ≈ Courier New, calibri ≈ Calibri.
 */
export type FontFamily = 'sans' | 'serif' | 'mono' | 'calibri';

export interface GradientStop {
  /** 0…1 along the gradient */
  pos: number;
  color: string;
  alpha?: number;
}

/** A colour gradient. Coordinates are in gradient space; `matrix` maps it to the drawing space (y down). */
export interface Gradient {
  kind: 'linear' | 'radial' | 'rect';
  stops: GradientStop[];
  /**
   * linear: [x1, y1, x2, y2]; radial: [fx, fy, fr, cx, cy, r] (focal circle → outer circle);
   * rect: [x0, y0, x1, y1, fx0, fy0, fx1, fy1] (outer rectangle, focus rectangle): concentric rectangles
   */
  coords: number[];
  matrix?: [number, number, number, number, number, number];
}

/**
 * Tidy section markers: consecutive markers merge (the last one's columns, a page start wins),
 * a marker at the very end is dropped, and so is a leading one that changes nothing.
 */
export function normalizeSections(blocks: Block[]): Block[] {
  const out: Block[] = [];
  for (const b of blocks) {
    const prev = out[out.length - 1];
    if (b.type === 'section' && prev?.type === 'section') {
      // the later marker's columns; a page size set by either one is kept
      const page = b.page ?? prev.page;
      out[out.length - 1] = { ...b, ...(page ? { page } : {}), start: prev.start === 'page' || b.start === 'page' ? 'page' : 'continuous' };
    }
    else out.push(b);
  }
  while (out.length && out[out.length - 1].type === 'section') out.pop();
  const first = out[0];
  if (first?.type === 'section' && first.columns <= 1 && !first.page) out.shift();
  return out;
}

/** A two-colour preset pattern (DrawingML a:pattFill); `bgClear`: the background is transparent. */
export interface PatternFill {
  preset: string;
  fg: string;
  bg: string;
  bgClear?: boolean;
}

/** One colour standing in for a gradient where it cannot be drawn (stops weighted by the span they cover). */
export function gradientAverage(g: Gradient): string {
  const s = g.stops;
  if (!s.length) return '000000';
  const rgb = (c: string): number[] => [0, 2, 4].map((o) => parseInt(c.slice(o, o + 2), 16) || 0);
  const acc = [0, 0, 0];
  let total = 0;
  s.forEach((st, i) => {
    const lo = i ? (s[i - 1].pos + st.pos) / 2 : 0;
    const hi = i < s.length - 1 ? (st.pos + s[i + 1].pos) / 2 : 1;
    const w = Math.max(hi - lo, 1e-3);
    rgb(st.color).forEach((v, k) => (acc[k] += v * w));
    total += w;
  });
  return acc.map((v) => Math.round(v / total).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * A floating picture (Word's wp:anchor, an ODF frame, a CSS float). `square`: text flows beside it;
 * `topBottom`: nothing beside it; `none`: placed without moving the text, in front of it or behind it.
 * Positions are in points from the frame named by `from`; `paragraph`/`line` mean where it is anchored.
 */
export interface ImageFloat {
  wrap: 'square' | 'topBottom' | 'none';
  behind?: boolean;
  /** square: the sides text flows on (both/largest: the wider one) */
  side?: 'both' | 'left' | 'right' | 'largest';
  h: { from: 'page' | 'margin' | 'column' | 'character' | 'leftMargin' | 'rightMargin'; align?: 'left' | 'center' | 'right' | 'inside' | 'outside'; offset?: number };
  v: { from: 'page' | 'margin' | 'paragraph' | 'line' | 'topMargin' | 'bottomMargin'; align?: 'top' | 'center' | 'bottom' | 'inside' | 'outside'; offset?: number };
  /** Distance kept from the text: top, right, bottom, left (pt) */
  dist?: [number, number, number, number];
}

const FLOAT_H = ['page', 'margin', 'column', 'character', 'leftMargin', 'rightMargin'];
const FLOAT_V = ['page', 'margin', 'paragraph', 'line', 'topMargin', 'bottomMargin'];
const FLOAT_HA = ['left', 'center', 'right', 'inside', 'outside'];
const FLOAT_VA = ['top', 'center', 'bottom', 'inside', 'outside'];

/** A float description from untrusted data (an editor attribute), or undefined when it is not one. */
export function sanitizeFloat(v: unknown): ImageFloat | undefined {
  if (!v || typeof v !== 'object') return undefined;
  const o = v as Record<string, unknown>;
  const fin = (x: unknown): number | undefined => (typeof x === 'number' && Number.isFinite(x) && Math.abs(x) < 1e5 ? x : undefined);
  const axis = (a: unknown, froms: string[], aligns: string[], def: string): { from: string; align?: string; offset?: number } => {
    const r = a && typeof a === 'object' ? (a as Record<string, unknown>) : {};
    const from = typeof r.from === 'string' && froms.includes(r.from) ? r.from : def;
    const align = typeof r.align === 'string' && aligns.includes(r.align) ? r.align : undefined;
    const offset = fin(r.offset);
    return { from, ...(align ? { align } : {}), ...(offset !== undefined ? { offset } : {}) };
  };
  const wrap = o.wrap === 'square' || o.wrap === 'topBottom' || o.wrap === 'none' ? o.wrap : undefined;
  if (!wrap) return undefined;
  const side = o.side === 'both' || o.side === 'left' || o.side === 'right' || o.side === 'largest' ? o.side : undefined;
  const dist = Array.isArray(o.dist) && o.dist.length === 4 && o.dist.every((x) => fin(x) !== undefined && (x as number) >= 0) ? (o.dist as [number, number, number, number]) : undefined;
  return {
    wrap,
    ...(o.behind === true ? { behind: true } : {}),
    ...(side ? { side } : {}),
    h: axis(o.h, FLOAT_H, FLOAT_HA, 'column') as ImageFloat['h'],
    v: axis(o.v, FLOAT_V, FLOAT_VA, 'paragraph') as ImageFloat['v'],
    ...(dist ? { dist } : {}),
  };
}

export interface ChartSource {
  xml: string;
  /** Theme colour scheme the chart resolves scheme colours against (dk1, lt1, accent1…). */
  colors?: Record<string, string>;
}

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
  /** Background highlight, hex without '#'. */
  highlight?: string;
  /** Font family when it differs from the document default. */
  font?: FontFamily;
  /** Font size in points when it differs from the document default. */
  size?: number;
  /** Dynamic field rendered by page-aware writers (headers/footers). */
  field?: 'page' | 'pages';
  /** Footnote: this run is the reference mark (numbered by writers), `note` is its text. */
  note?: Block[];
}

/** Paragraph formatting (all lengths in points). */
export interface ParaFormat {
  align?: Align;
  /** Left indent. */
  indent?: number;
  /** First-line indent; negative values mean a hanging indent. */
  firstLine?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  /** Line spacing as a multiple of single spacing. */
  lineHeight?: number;
}

export interface PageSetup {
  /** Page size in points. */
  width: number;
  height: number;
  margins: { top: number; right: number; bottom: number; left: number };
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
  | ({ type: 'paragraph'; runs: TextRun[] } & ParaFormat)
  | ({ type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; runs: TextRun[] } & ParaFormat)
  | { type: 'list'; ordered: boolean; items: ListItem[]; start?: number }
  | { type: 'table'; rows: TableRow[] }
  /** width/height are the intended display size in CSS pixels (96 dpi). */
  | { type: 'image'; data: Uint8Array; mime: string; width: number; height: number; alt?: string; /** Source Office chart part (c:chartSpace XML) and its theme colours; `data` holds its SVG rendering. */ chart?: ChartSource; /** A floating picture: where it goes and how text flows around it */ float?: ImageFloat }
  | { type: 'code'; text: string }
  | { type: 'quote'; blocks: Block[] }
  | { type: 'hr' }
  | { type: 'pagebreak' }
  /**
   * Starts a section: the blocks that follow (up to the next section) flow in `columns` columns
   * `gap` points apart. `start: 'page'` begins it on a new page; 'continuous' on the same page
   * (the columns of the section before are then balanced).
   */
  | {
      type: 'section';
      columns: number;
      gap: number;
      separator?: boolean;
      start: 'continuous' | 'page';
      /** Unequal columns: relative widths, and the space after each column but the last (pt). */
      widths?: number[];
      gaps?: number[];
      /** Page size and margins of the section when they differ from the document's. */
      page?: PageSetup;
    };

export interface Doc {
  title?: string;
  blocks: Block[];
  /** Default font family and size (pt) of body text. */
  font?: FontFamily;
  size?: number;
  /** Paragraph formatting applied when a block does not specify its own. */
  defaults?: ParaFormat;
  /** Repeated on every page by page-aware writers. */
  header?: Block[];
  footer?: Block[];
  page?: PageSetup;
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

export type Content = { kind: 'doc'; doc: Doc } | { kind: 'book'; book: Workbook } | { kind: 'slides'; deck: import('./slides').Deck };

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
  !a.note &&
  !b.note &&
  !a.field &&
  !b.field &&
  a.font === b.font &&
  a.size === b.size &&
  a.highlight === b.highlight &&
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
    if (!r.text && !r.note && !r.field) continue;
    const prev = out[out.length - 1];
    if (prev && sameStyle(prev, r)) prev.text += r.text;
    else out.push({ ...r });
  }
  return out;
}

/** Trim leading/trailing whitespace of a run list (not inside). */
export function trimRuns(runs: TextRun[]): TextRun[] {
  const out = normalizeRuns(runs);
  const blank = (r: TextRun): boolean => !r.note && !r.field && !r.text.replace(/^[ \t\r\n]+/, '');
  while (out.length && blank(out[0])) out.shift();
  while (out.length && blank(out[out.length - 1])) out.pop();
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

/** Map a font name (Word, CSS, PDF) to the closest supported family. */
export function familyOf(name: string | undefined): FontFamily | undefined {
  if (!name) return undefined;
  const n = name.toLowerCase();
  if (/calibri|carlito/.test(n)) return 'calibri';
  if (/courier|consolas|mono|menlo|monaco|lucida console|source code|fira code/.test(n)) return 'mono';
  if (/sans/.test(n)) return 'sans';
  if (/times|georgia|garamond|cambria|caladea|book antiqua|palatino|serif|minion|baskerville|constantia|bookman|century schoolbook|didot|bodoni/.test(n)) return 'serif';
  return 'sans';
}

/** Visit every run in a list of blocks (including nested lists, tables, quotes and notes). */
export function forEachRun(blocks: Block[] | undefined, fn: (r: TextRun, b: Block) => void): void {
  for (const b of blocks ?? []) {
    if (b.type === 'paragraph' || b.type === 'heading') {
      for (const r of b.runs) {
        fn(r, b);
        if (r.note) forEachRun(r.note, fn);
      }
    } else if (b.type === 'list') b.items.forEach((i) => forEachRun(i.blocks, fn));
    else if (b.type === 'table') b.rows.forEach((row) => row.cells.forEach((c) => forEachRun(c.blocks, fn)));
    else if (b.type === 'quote') forEachRun(b.blocks, fn);
  }
}
