/**
 * Presentation model (slides with absolutely positioned elements). Produced by
 * the PPTX reader; rendered page-for-page to PDF, or flattened to a Doc for the
 * text formats. All lengths are in points, colours are hex RRGGBB.
 */
import type { Align, Block, ChartSource, TextRun } from './model';

export interface SlideBullet {
  /** Literal bullet character, or the formatted number for auto-numbered lists. */
  text: string;
  color?: string;
  font?: TextRun['font'];
  /** Bullet size relative to the first run (1 = same size). */
  scale: number;
}

export interface SlidePara {
  runs: TextRun[];
  align?: Align;
  level: number;
  /** Left margin of the text and first-line indent (negative = hanging bullet). */
  marL: number;
  indent: number;
  /** Line spacing as a multiple of single spacing, or exact points. */
  lineHeight?: number;
  linePts?: number;
  spaceBefore: number;
  spaceAfter: number;
  bullet?: SlideBullet;
  /** Size used for empty paragraphs (end-of-paragraph run properties). */
  size: number;
}

export interface SlideText {
  paragraphs: SlidePara[];
  anchor: 't' | 'ctr' | 'b';
  /** Insets: left, top, right, bottom. */
  insets: [number, number, number, number];
  wrap: boolean;
  /** normAutofit: font scale and line spacing reduction (fractions). */
  fontScale: number;
  lineReduction: number;
}

export type SlideRole = 'title' | 'subtitle' | 'body' | 'footer' | 'other';

export interface SlideLine {
  color: string;
  width: number;
  dash?: boolean;
  opacity?: number;
}

export type SlideElement =
  | {
      kind: 'shape';
      x: number;
      y: number;
      w: number;
      h: number;
      geom: string;
      rot?: number;
      fill?: string;
      fillOpacity?: number;
      line?: SlideLine;
      text?: SlideText;
      role: SlideRole;
    }
  | { kind: 'image'; x: number; y: number; w: number; h: number; data: Uint8Array; mime: string; crop?: { l: number; t: number; r: number; b: number }; rot?: number; alt?: string; chart?: ChartSource }
  | { kind: 'line'; x1: number; y1: number; x2: number; y2: number; line: SlideLine }
  | { kind: 'table'; x: number; y: number; w: number; h: number; cols: number[]; rows: SlideTableRow[] };

export interface SlideTableCell {
  text: SlideText;
  fill?: string;
  span: number;
  rowSpan: number;
  /** Covered by a span of another cell. */
  merged: boolean;
  borders: { l?: SlideLine; r?: SlideLine; t?: SlideLine; b?: SlideLine };
}

export interface SlideTableRow {
  h: number;
  cells: SlideTableCell[];
}

export interface Slide {
  background?: { color?: string; image?: { data: Uint8Array; mime: string } };
  elements: SlideElement[];
  notes?: Block[];
  hidden?: boolean;
}

export interface Deck {
  title?: string;
  width: number;
  height: number;
  slides: Slide[];
}
