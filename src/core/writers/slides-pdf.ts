/**
 * Deck → PDF, one page per slide at the slide's own size. Reuses the document
 * writer's font machinery (fallback scripts, HarfBuzz subsetting) and paragraph
 * layout; adds absolute positioning, shapes, rotation, cropping and tables.
 */
import { clip, concatTransformationMatrix, endPath, popGraphicsState, pushGraphicsState, rectangle, rgb, PDFDocument, type PDFPage } from 'pdf-lib';
import type { TextRun } from '../model';
import type { Deck, SlideElement, SlideLine, SlidePara, SlideText } from '../slides';
import type { FontKey } from '../pdf-fonts';
import { hexColor, PdfWriter, type Box, type Op, type PdfOptions, type PdfWriteEnv } from './pdf';

type Cmd = (page: PDFPage) => void;

interface TextLayout {
  ops: Op[];
  height: number;
}

/** An ellipse inscribed in a w × h box, as four cubic Béziers (y down). */
function ellipsePath(w: number, h: number): string {
  const k = 0.5523;
  const rx = w / 2;
  const ry = h / 2;
  const f = (n: number): string => n.toFixed(2);
  return `M${f(w)} ${f(ry)} C${f(w)} ${f(ry + k * ry)} ${f(rx + k * rx)} ${f(h)} ${f(rx)} ${f(h)} C${f(rx - k * rx)} ${f(h)} 0 ${f(ry + k * ry)} 0 ${f(ry)} C0 ${f(ry - k * ry)} ${f(rx - k * rx)} 0 ${f(rx)} 0 C${f(rx + k * rx)} 0 ${f(w)} ${f(ry - k * ry)} ${f(w)} ${f(ry)} Z`;
}

export class SlidePdfWriter extends PdfWriter {
  private cmds: Cmd[][] = [];

  constructor(env: PdfWriteEnv, opts: PdfOptions) {
    super(env, opts);
  }

  async render(deck: Deck): Promise<Uint8Array> {
    this.pdf = await PDFDocument.create();
    this.pdf.registerFontkit((await import('@pdf-lib/fontkit')).default);
    if (deck.title) this.pdf.setTitle(deck.title);
    this.pdf.setCreator('Salatic');
    this.pdf.setProducer('Salatic');
    this.pageW = deck.width;
    this.pageH = deck.height;
    const visible = deck.slides.filter((s) => !s.hidden);
    if (visible.length < deck.slides.length) this.warnings.add(`Скрытые слайды (${deck.slides.length - visible.length}) не включены в PDF.`);
    // Phase 1: measure text (marks the glyphs to embed) and embed images
    let done = 0;
    for (const slide of visible) {
      const cmds: Cmd[] = [];
      const W = deck.width;
      const H = deck.height;
      cmds.push((page) => page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: rgb(1, 1, 1) }));
      const bgGrad = slide.background?.gradient;
      const bgPatt = slide.background?.pattern;
      if (bgGrad) cmds.push((page) => this.fillGradient(page, `M0 0 H${W} V${H} H0 Z`, bgGrad, 0, H));
      else if (bgPatt) cmds.push((page) => this.fillPattern(page, `M0 0 H${W} V${H} H0 Z`, bgPatt, 0, H));
      else if (slide.background?.color) {
        const c = hexColor(slide.background.color);
        cmds.push((page) => page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: rgb(...c) }));
      }
      if (slide.background?.image) {
        const img = await this.image({ type: 'image', data: slide.background.image.data, mime: slide.background.image.mime, width: 1, height: 1 });
        if (img) cmds.push((page) => page.drawImage(img, { x: 0, y: 0, width: W, height: H }));
      }
      for (const e of slide.elements) await this.element(e, cmds);
      this.cmds.push(cmds);
      this.env.onProgress?.((++done / visible.length) * 0.8);
    }
    if (this.fm.missing) this.warnings.add('Некоторые символы отсутствуют во всех встроенных шрифтах и заменены на «?».');
    // Phase 2: fonts are final — draw
    await this.fm.embedAll(this.pdf, this.opts.subsetFonts !== false);
    for (const cmds of this.cmds) {
      const page = this.pdf.addPage([deck.width, deck.height]);
      for (const c of cmds) c(page);
    }
    return this.pdf.save({ useObjectStreams: true });
  }

  /** Draw ops whose y is measured from the top of the page, shifted down by `dy` (no clipping). */
  private drawAt(ops: Op[], dy: number, page: PDFPage): void {
    // draw() maps y to H - (offset + y - from); offset = from + dy keeps every op visible
    this.draw(ops, dy - 1e6, -1e6, 1e6, page);
  }

  // -- geometry helpers -------------------------------------------------------------

  /** Run `fn` with the page rotated by `deg` around the centre of the box (top-left coordinates). */
  private rotated(page: PDFPage, box: { x: number; y: number; w: number; h: number }, deg: number | undefined, fn: () => void): void {
    if (!deg) {
      fn();
      return;
    }
    const cx = box.x + box.w / 2;
    const cy = this.pageH - (box.y + box.h / 2);
    const a = (-deg * Math.PI) / 180; // PowerPoint rotates clockwise, PDF counter-clockwise
    const cos = Math.cos(a);
    const sin = Math.sin(a);
    page.pushOperators(pushGraphicsState(), concatTransformationMatrix(1, 0, 0, 1, cx, cy), concatTransformationMatrix(cos, sin, -sin, cos, 0, 0), concatTransformationMatrix(1, 0, 0, 1, -cx, -cy));
    fn();
    page.pushOperators(popGraphicsState());
  }

  private stroke(line: SlideLine | undefined): { borderColor?: ReturnType<typeof rgb>; borderWidth: number; borderDashArray?: number[]; borderOpacity?: number } {
    if (!line) return { borderWidth: 0 };
    return { borderColor: rgb(...hexColor(line.color)), borderWidth: line.width, borderDashArray: line.dash ? [line.width * 3, line.width * 2] : undefined, borderOpacity: line.opacity };
  }

  /** SVG path (y down, origin at the top-left of the box) for preset geometries. */
  private path(geom: string, w: number, h: number): string | null {
    const m = Math.min(w, h);
    const P = (pts: [number, number][]): string => 'M' + pts.map(([x, y]) => `${x.toFixed(2)},${y.toFixed(2)}`).join(' L') + ' Z';
    switch (geom) {
      case 'rect':
      case 'flowChartProcess':
      case 'wedgeRectCallout':
      case 'snip1Rect':
        return null;
      case 'roundRect':
      case 'flowChartAlternateProcess':
      case 'wedgeRoundRectCallout':
      case 'flowChartTerminator': {
        const r = geom === 'flowChartTerminator' ? m / 2 : m * 0.1667;
        return `M${r},0 L${w - r},0 Q${w},0 ${w},${r} L${w},${h - r} Q${w},${h} ${w - r},${h} L${r},${h} Q0,${h} 0,${h - r} L0,${r} Q0,0 ${r},0 Z`;
      }
      case 'triangle':
        return P([[w / 2, 0], [w, h], [0, h]]);
      case 'rtTriangle':
        return P([[0, 0], [w, h], [0, h]]);
      case 'diamond':
      case 'flowChartDecision':
        return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
      case 'parallelogram':
      case 'flowChartInputOutput':
        return P([[m * 0.25, 0], [w, 0], [w - m * 0.25, h], [0, h]]);
      case 'trapezoid':
        return P([[m * 0.25, 0], [w - m * 0.25, 0], [w, h], [0, h]]);
      case 'pentagon':
        return P([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]]);
      case 'hexagon':
        return P([[m * 0.25, 0], [w - m * 0.25, 0], [w, h / 2], [w - m * 0.25, h], [m * 0.25, h], [0, h / 2]]);
      case 'octagon': {
        const k = m * 0.29;
        return P([[k, 0], [w - k, 0], [w, k], [w, h - k], [w - k, h], [k, h], [0, h - k], [0, k]]);
      }
      case 'homePlate':
        return P([[0, 0], [w - h / 2, 0], [w, h / 2], [w - h / 2, h], [0, h]]);
      case 'chevron':
        return P([[0, 0], [w - h / 2, 0], [w, h / 2], [w - h / 2, h], [0, h], [h / 2, h / 2]]);
      case 'rightArrow':
        return P([[0, h * 0.25], [w - h / 2, h * 0.25], [w - h / 2, 0], [w, h / 2], [w - h / 2, h], [w - h / 2, h * 0.75], [0, h * 0.75]]);
      case 'leftArrow':
        return P([[w, h * 0.25], [h / 2, h * 0.25], [h / 2, 0], [0, h / 2], [h / 2, h], [h / 2, h * 0.75], [w, h * 0.75]]);
      case 'upArrow':
        return P([[w * 0.25, h], [w * 0.25, w / 2], [0, w / 2], [w / 2, 0], [w, w / 2], [w * 0.75, w / 2], [w * 0.75, h]]);
      case 'downArrow':
        return P([[w * 0.25, 0], [w * 0.75, 0], [w * 0.75, h - w / 2], [w, h - w / 2], [w / 2, h], [0, h - w / 2], [w * 0.25, h - w / 2]]);
      case 'plus':
      case 'mathPlus': {
        const a = m * 0.25;
        return P([[w / 2 - a, 0], [w / 2 + a, 0], [w / 2 + a, h / 2 - a], [w, h / 2 - a], [w, h / 2 + a], [w / 2 + a, h / 2 + a], [w / 2 + a, h], [w / 2 - a, h], [w / 2 - a, h / 2 + a], [0, h / 2 + a], [0, h / 2 - a], [w / 2 - a, h / 2 - a]]);
      }
      case 'star5': {
        const pts: [number, number][] = [];
        for (let i = 0; i < 10; i++) {
          const ang = -Math.PI / 2 + (i * Math.PI) / 5;
          const r = i % 2 ? 0.38 : 0.5;
          pts.push([w / 2 + Math.cos(ang) * r * w, h / 2 + Math.sin(ang) * r * h * 1.05 + h * 0.03]);
        }
        return P(pts);
      }
      default:
        return null;
    }
  }

  // -- elements -------------------------------------------------------------------------

  private async element(e: SlideElement, cmds: Cmd[]): Promise<void> {
    const H = this.pageH;
    switch (e.kind) {
      case 'line': {
        const c = hexColor(e.line.color);
        cmds.push((page) => page.drawLine({ start: { x: e.x1, y: H - e.y1 }, end: { x: e.x2, y: H - e.y2 }, thickness: e.line.width, color: rgb(...c), dashArray: e.line.dash ? [e.line.width * 3, e.line.width * 2] : undefined, opacity: e.line.opacity }));
        return;
      }
      case 'image': {
        if (/emf|wmf|svg/.test(e.mime)) {
          const op = await this.metaOp(e.data, e.x, e.y, e.w, e.h, e.mime);
          if (op) cmds.push((page) => this.rotated(page, e, e.rot, () => this.drawMeta(page, op, e.x, H - e.y)));
          return;
        }
        const img = await this.image({ type: 'image', data: e.data, mime: e.mime, width: 1, height: 1 });
        if (!img) return;
        cmds.push((page) =>
          this.rotated(page, e, e.rot, () => {
            const crop = e.crop;
            if (!crop) {
              page.drawImage(img, { x: e.x, y: H - e.y - e.h, width: e.w, height: e.h });
              return;
            }
            const fw = e.w / Math.max(0.01, 1 - crop.l - crop.r);
            const fh = e.h / Math.max(0.01, 1 - crop.t - crop.b);
            page.pushOperators(pushGraphicsState(), rectangle(e.x, H - e.y - e.h, e.w, e.h), clip(), endPath());
            page.drawImage(img, { x: e.x - crop.l * fw, y: H - e.y - e.h - crop.b * fh, width: fw, height: fh });
            page.pushOperators(popGraphicsState());
          }),
        );
        return;
      }
      case 'shape': {
        const fill = e.fill ? hexColor(e.fill) : undefined;
        const border = this.stroke(e.line);
        if (fill || e.line) {
          const path = this.path(e.geom, e.w, e.h);
          const grad = e.gradient;
          const patt = e.pattern;
          cmds.push((page) =>
            this.rotated(page, e, e.rot, () => {
              if (patt) {
                const round = e.geom === 'ellipse' || e.geom === 'flowChartConnector' || e.geom === 'donut';
                const d = round ? ellipsePath(e.w, e.h) : (path ?? `M0 0 H${e.w} V${e.h} H0 Z`);
                this.fillPattern(page, d, patt, e.x, H - e.y, e.fillOpacity ?? 1);
                if (e.line) page.drawSvgPath(d, { x: e.x, y: H - e.y, ...border });
                return;
              }
              if (grad) {
                // the shading fills the outline; the border is stroked on top
                const round = e.geom === 'ellipse' || e.geom === 'flowChartConnector' || e.geom === 'donut';
                const d = round ? ellipsePath(e.w, e.h) : (path ?? `M0 0 H${e.w} V${e.h} H0 Z`);
                this.fillGradient(page, d, grad, e.x, H - e.y);
                if (e.line) page.drawSvgPath(d, { x: e.x, y: H - e.y, ...border });
                return;
              }
              const common = { color: fill ? rgb(...fill) : undefined, opacity: e.fillOpacity, ...border };
              if (e.geom === 'ellipse' || e.geom === 'flowChartConnector' || e.geom === 'donut') page.drawEllipse({ x: e.x + e.w / 2, y: H - e.y - e.h / 2, xScale: e.w / 2, yScale: e.h / 2, ...common });
              else if (path) page.drawSvgPath(path, { x: e.x, y: H - e.y, ...common });
              else page.drawRectangle({ x: e.x, y: H - e.y - e.h, width: e.w, height: e.h, ...common });
            }),
          );
        }
        if (e.text) {
          const layout = await this.layoutText(e.text, e.x, e.y, e.w, e.h);
          cmds.push((page) => this.rotated(page, e, e.rot, () => this.drawAt(layout.ops, 0, page)));
        }
        return;
      }
      case 'table':
        await this.slideTable(e, cmds);
        return;
    }
  }

  private async slideTable(e: Extract<SlideElement, { kind: 'table' }>, cmds: Cmd[]): Promise<void> {
    const H = this.pageH;
    const total = e.cols.reduce((a, b) => a + b, 0) || e.w;
    const scale = total > 0 ? Math.min(1.5, Math.max(0.5, e.w / total)) : 1;
    const cols = e.cols.map((c) => c * scale);
    const xs = [e.x];
    for (const c of cols) xs.push(xs[xs.length - 1] + c);
    // Row heights grow with their content
    const layouts: (TextLayout | null)[][] = [];
    const heights: number[] = [];
    for (const row of e.rows) {
      const ls: (TextLayout | null)[] = [];
      let h = row.h;
      for (let ci = 0; ci < row.cells.length; ci++) {
        const cell = row.cells[ci];
        if (cell.merged || ci >= cols.length) {
          ls.push(null);
          continue;
        }
        const w = xs[Math.min(xs.length - 1, ci + cell.span)] - xs[ci];
        const l = await this.layoutText(cell.text, xs[ci], 0, w, 0, true);
        ls.push(l);
        if (cell.rowSpan === 1) h = Math.max(h, l.height + cell.text.insets[1] + cell.text.insets[3]);
      }
      layouts.push(ls);
      heights.push(h);
    }
    let y = e.y;
    e.rows.forEach((row, ri) => {
      const rh = heights[ri];
      row.cells.forEach((cell, ci) => {
        const l = layouts[ri][ci];
        if (!l || cell.merged) return;
        const x = xs[ci];
        const w = xs[Math.min(xs.length - 1, ci + cell.span)] - x;
        const h = heights.slice(ri, ri + cell.rowSpan).reduce((a, b) => a + b, 0);
        const top = y;
        const [, it, , ib] = cell.text.insets;
        const free = h - it - ib - l.height;
        const dy = top + it + (cell.text.anchor === 'ctr' ? free / 2 : cell.text.anchor === 'b' ? free : 0);
        const fill = cell.fill ? hexColor(cell.fill) : undefined;
        const b = cell.borders;
        cmds.push((page) => {
          const rect = `M0 0 H${w} V${h} H0 Z`;
          if (cell.pattern) this.fillPattern(page, rect, cell.pattern, x, H - top);
          else if (cell.gradient && cell.box) this.fillGradient(page, rect, { ...cell.gradient, matrix: [w / cell.box[0], 0, 0, h / cell.box[1], 0, 0] }, x, H - top);
          else if (fill) page.drawRectangle({ x, y: H - top - h, width: w, height: h, color: rgb(...fill) });
          const edge = (ln: SlideLine | undefined, x1: number, y1: number, x2: number, y2: number): void => {
            if (ln) page.drawLine({ start: { x: x1, y: H - y1 }, end: { x: x2, y: H - y2 }, thickness: ln.width, color: rgb(...hexColor(ln.color)) });
          };
          edge(b.t, x, top, x + w, top);
          edge(b.b, x, top + h, x + w, top + h);
          edge(b.l, x, top, x, top + h);
          edge(b.r, x + w, top, x + w, top + h);
          this.drawAt(l.ops, dy, page);
        });
      });
      y += rh;
    });
  }

  /** Lay out a text body inside a box; ops are absolute (page coordinates, y from the top). */
  private async layoutText(t: SlideText, x: number, y: number, w: number, h: number, relative = false): Promise<TextLayout> {
    const [il, it, ir, ib] = t.insets;
    // Unwrapped boxes are normally sized to their text by PowerPoint, so the box width works for both
    const width = Math.max(t.wrap ? 4 : 40, w - il - ir);
    const boxes: { box: Box; para: SlidePara }[] = [];
    for (const p of t.paragraphs) boxes.push({ box: await this.paragraph(p, x + il, width), para: p });
    let total = 0;
    boxes.forEach(({ box }, i) => (total += (i ? box.spaceBefore : 0) + box.height + (i < boxes.length - 1 ? box.spaceAfter : 0)));
    const free = h - it - ib - total;
    let cy = relative ? 0 : y + it + (t.anchor === 'ctr' ? free / 2 : t.anchor === 'b' ? free : 0);
    const ops: Op[] = [];
    boxes.forEach(({ box }, i) => {
      if (i) cy += box.spaceBefore;
      for (const o of box.ops) ops.push(shiftOp(o, cy));
      cy += box.height + box.spaceAfter;
    });
    return { ops, height: total };
  }

  private async paragraph(p: SlidePara, x: number, width: number): Promise<Box> {
    const runs: TextRun[] = p.runs.length ? p.runs : [{ text: '', size: p.size }];
    const ctx = { family: runs[0].font ?? 'sans', size: p.size };
    let lineHeight = p.lineHeight ?? 1;
    if (p.linePts) {
      const key: FontKey = this.keyFor(runs[0], ctx);
      await this.fm.load([key]);
      lineHeight = p.linePts / Math.max(1, this.fm.metrics(key).line * (runs[0].size ?? p.size));
    }
    const left = Math.max(0, p.marL);
    const box = await this.textBox(runs.some((r) => r.text) ? runs : [{ ...runs[0], text: ' ' }], x, width, ctx, { align: p.align, indent: left, firstLine: p.bullet ? 0 : p.indent, spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter, lineHeight });
    if (p.bullet) {
      const first = box.ops.find((o): o is Extract<Op, { k: 'text' }> => o.k === 'text');
      const size = (runs[0].size ?? p.size) * p.bullet.scale;
      const family = p.bullet.font ?? runs[0].font ?? 'sans';
      const primary: FontKey = `${family}-regular`;
      await this.fm.load([primary]);
      const cp = p.bullet.text.codePointAt(0) ?? 0x2022;
      const need = this.fm.fallbacksFor(primary, cp);
      if (need.length) await this.fm.load(need);
      const key = this.fm.resolve(primary, cp) ?? primary;
      this.fm.markUsed(key, p.bullet.text);
      const color = hexColor(p.bullet.color ?? runs[0].color ?? '000000');
      const bx = x + Math.max(0, p.marL + p.indent);
      const baseline = first ? first.y : size;
      box.ops.push({ k: 'text', x: bx, y: baseline, top: first?.top ?? 0, bottom: first?.bottom ?? size, text: p.bullet.text, font: key, size, color });
    }
    return box;
  }
}

function shiftOp(o: Op, dy: number): Op {
  switch (o.k) {
    case 'text':
      return { ...o, y: o.y + dy, top: o.top + dy, bottom: o.bottom + dy };
    case 'line':
      return { ...o, y1: o.y1 + dy, y2: o.y2 + dy, top: o.top + dy, bottom: o.bottom + dy };
    default:
      return { ...o, y: o.y + dy, top: o.top + dy, bottom: o.bottom + dy };
  }
}

export async function deckToPdf(deck: Deck, env: PdfWriteEnv, opts: PdfOptions = {}): Promise<{ data: Uint8Array; warnings: string[] }> {
  const w = new SlidePdfWriter(env, opts);
  const data = await w.render(deck);
  return { data, warnings: [...w.warnings] };
}
