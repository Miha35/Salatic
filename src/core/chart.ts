/**
 * Office charts (DrawingML chartSpace, as in word/charts/*.xml and ppt/charts/*.xml)
 * → SVG drawn from the cached values stored in the file. Covers column/bar
 * (clustered, stacked, 100 %), line, area, pie, doughnut, pie/bar of pie, scatter,
 * bubble, radar, stock (high-low lines, up/down bars) and surface (contour) charts,
 * their 3-D variants (oblique projection with depth, shaded faces), combinations
 * with a secondary value axis, titles, legends, gridlines, axis number formats,
 * per-point colours, marker symbols and value/percentage data labels.
 */
import { colorOf, type Theme } from './drawingml';
import { elements, parseXml, textContent, type XNode } from './xml';
import { escapeXml } from './util';

const OFFICE_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const AXIS = 'BFBFBF';
const GRID = 'D9D9D9';
const LABEL = '595959';
const FONT = "Calibri, Carlito, 'Liberation Sans', sans-serif";

const local = (n: string): string => n.replace(/^[a-z0-9]+:/i, '');
const kid = (n: XNode | undefined, name: string): XNode | undefined => elements(n).find((c) => local(c.name) === name);
const kids = (n: XNode | undefined, name: string): XNode[] => elements(n).filter((c) => local(c.name) === name);
const deep = (n: XNode | undefined, name: string): XNode | undefined => {
  if (!n) return undefined;
  for (const c of elements(n)) {
    if (local(c.name) === name) return c;
    const r = deep(c, name);
    if (r) return r;
  }
  return undefined;
};
const val = (n: XNode | undefined): string | undefined => n?.attrs.val;
const on = (n: XNode | undefined, def = false): boolean => (n ? n.attrs.val !== '0' && n.attrs.val !== 'false' : def);

interface Series {
  name: string;
  cats: string[];
  vals: (number | null)[];
  xs?: (number | null)[];
  /** bubble sizes */
  sizes?: (number | null)[];
  /** marker symbol: circle, square, diamond, triangle, x, star, plus, dash, dot */
  symbol?: string;
  color?: string;
  lineColor?: string;
  lineWidth?: number;
  noLine?: boolean;
  pointColors: Map<number, string>;
  marker: boolean;
  smooth: boolean;
  format?: string;
}

interface Group {
  type: 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'bubble' | 'radar' | 'stock' | 'surface' | 'ofPie';
  /** a 3-D variant (bar3DChart, pie3DChart …) */
  threeD: boolean;
  /** 3-D bar shape: box, cylinder, cone, pyramid (and *ToMax) */
  shape: string;
  radarStyle: string;
  hiLow: boolean;
  upDown?: { gap: number; up?: string; down?: string };
  bubbleScale: number;
  sizeArea: boolean;
  wireframe: boolean;
  /** pie of pie / bar of pie: the last `splitPos` points go to the second plot */
  ofPie?: { bar: boolean; splitPos: number; secondSize: number };
  horizontal: boolean;
  grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  series: Series[];
  gap: number;
  /** Bar overlap within a cluster, −100…100 (stacked groupings imply 100). */
  overlap: number;
  hole: number;
  firstAngle: number;
  secondary: boolean;
  vary: boolean;
  showVal: boolean;
  showPercent: boolean;
  showCat: boolean;
}

interface Axis {
  min?: number;
  max?: number;
  format?: string;
  deleted: boolean;
  grid: boolean;
  reversed: boolean;
  /** Category positions at tick marks ("midCat", areas) instead of between them. */
  midCat?: boolean;
}

export interface ChartOptions {
  /** Theme colours (accent1…6, dk1, lt1 …) for scheme references. */
  colors?: Record<string, string>;
  width: number;
  height: number;
}

function strings(ref: XNode | undefined): string[] {
  const cache = deep(ref, 'strCache') ?? deep(ref, 'numCache') ?? deep(ref, 'multiLvlStrCache');
  if (!cache) {
    const lit = deep(ref, 'strLit') ?? deep(ref, 'numLit');
    if (!lit) return [];
    return ptList(lit).map((v) => v ?? '');
  }
  if (local(cache.name) === 'multiLvlStrCache') {
    // innermost level gives the labels
    const lvl = kids(cache, 'lvl')[0];
    return ptList(lvl).map((v) => v ?? '');
  }
  return ptList(cache).map((v) => v ?? '');
}

function ptList(n: XNode | undefined): (string | null)[] {
  const count = +(val(kid(n, 'ptCount')) ?? 0);
  const out: (string | null)[] = Array.from({ length: count }, () => null);
  for (const pt of kids(n, 'pt')) {
    const idx = +(pt.attrs.idx ?? 0);
    if (idx > 100_000) continue;
    out[idx] = textContent(kid(pt, 'v'));
  }
  return out;
}

function numbers(ref: XNode | undefined): { vals: (number | null)[]; format?: string } {
  const cache = deep(ref, 'numCache') ?? deep(ref, 'numLit');
  const vals = ptList(cache).map((v) => (v === null || v === '' || !Number.isFinite(+v) ? null : +v));
  return { vals, format: textContent(kid(cache, 'formatCode')) || undefined };
}

function richText(n: XNode | undefined): string {
  if (!n) return '';
  const rich = deep(n, 'rich');
  if (rich) return kids(rich, 'p').map((p) => elements(p).filter((r) => local(r.name) === 'r' || local(r.name) === 'fld').map((r) => textContent(kid(r, 't'))).join('')).join('\n');
  const ref = deep(n, 'strRef');
  if (ref) return strings(ref).join(' ');
  return textContent(deep(n, 'v'));
}

/** Format a value the way Excel would for common number formats. */
export function formatNumber(v: number, code?: string): string {
  const group = (s: string): string => s.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  const fixed = (x: number, d: number): string => {
    const [i, f] = Math.abs(x).toFixed(d).split('.');
    return (x < 0 ? '−' : '') + group(i) + (f ? ',' + f : '');
  };
  if (code && code !== 'General') {
    const main = code.split(';')[0];
    const pct = main.includes('%');
    const dec = /\.(0+)/.exec(main)?.[1].length ?? 0;
    const thousands = main.includes(',') || main.includes('#,##');
    const x = pct ? v * 100 : v;
    let s = thousands ? fixed(x, dec) : (x < 0 ? '−' : '') + Math.abs(x).toFixed(dec).replace('.', ',');
    if (pct) s += '%';
    const cur = /"([^"]+)"|\[\$([^\]-]+)/.exec(main);
    if (cur) s = main.trim().startsWith('"') || main.trim().startsWith('[$') ? `${cur[1] ?? cur[2]}${s}` : `${s} ${cur[1] ?? cur[2]}`;
    return s;
  }
  if (Number.isInteger(v)) return fixed(v, 0);
  const abs = Math.abs(v);
  const d = abs >= 100 ? 1 : abs >= 1 ? 2 : 3;
  return fixed(v, d).replace(/,?0+$/, '');
}

/** Excel starts value axes at zero unless the data sit in a narrow band far from it. */
function autoMin(min: number, max: number): number {
  if (min >= 0 && max > 0) return min / max > 5 / 6 ? min : 0;
  if (max <= 0 && min < 0) return min;
  return min;
}

/** Rounded axis bounds and a 1/2/2.5/5 × 10ⁿ step giving about `target` intervals. */
function niceScale(min: number, max: number, fixedMin?: number, fixedMax?: number, target = 5): { min: number; max: number; step: number } {
  if (fixedMin !== undefined) min = fixedMin;
  if (fixedMax !== undefined) max = fixedMax;
  if (min === max) {
    max = min + 1;
    if (min > 0) min = 0;
  }
  const span = max - min;
  const raw = span / Math.max(2, target);
  const mag = 10 ** Math.floor(Math.log10(raw));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  return { min: fixedMin ?? Math.floor(min / step) * step, max: fixedMax ?? Math.ceil(max / step) * step, step };
}

interface View3D {
  rotX: number;
  rotY: number;
  depthPercent: number;
  rAngAx: boolean;
}

function parseChart(xml: XNode, theme: Theme): { title?: string; titleSize: number; groups: Group[]; catAxis: Axis; valAxis: Axis; val2Axis?: Axis; xAxis?: Axis; legend?: string; bg?: string; plotBg?: string; view3D?: View3D } {
  const chart = deep(xml, 'chart');
  const plot = kid(chart, 'plotArea');
  const map = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
  const color = (sp: XNode | undefined, kind: 'fill' | 'line'): string | undefined => {
    if (!sp) return undefined;
    const holder = kind === 'fill' ? sp : kid(sp, 'ln');
    if (!holder || kid(holder, 'noFill')) return undefined;
    const solid = kid(holder, 'solidFill') ?? kid(holder, 'gradFill') ?? kid(holder, 'pattFill');
    if (!solid) return undefined;
    const target = local(solid.name) === 'gradFill' ? deep(solid, 'gs') : local(solid.name) === 'pattFill' ? kid(solid, 'fgClr') : solid;
    return colorOf(target, theme, map)?.hex;
  };
  const groups: Group[] = [];
  const axIds = new Map<string, XNode>();
  for (const ax of elements(plot)) if (/Ax$/.test(local(ax.name))) axIds.set(val(kid(ax, 'axId')) ?? '', ax);
  for (const g of elements(plot)) {
    const name = local(g.name);
    const type: Group['type'] | undefined = /^bar/.test(name)
      ? 'bar'
      : /^line/.test(name)
        ? 'line'
        : name === 'radarChart'
          ? 'radar'
          : name === 'stockChart'
            ? 'stock'
            : /^area/.test(name)
              ? 'area'
              : name === 'ofPieChart'
                ? 'ofPie'
                : /^pie/.test(name)
                  ? 'pie'
                  : name === 'doughnutChart'
                    ? 'doughnut'
                    : name === 'scatterChart'
                      ? 'scatter'
                      : name === 'bubbleChart'
                        ? 'bubble'
                        : /^surface/.test(name)
                          ? 'surface'
                          : undefined;
    if (!type) continue;
    const dl = kid(g, 'dLbls');
    const ud = kid(g, 'upDownBars');
    const grp: Group = {
      type,
      threeD: /3DChart$/.test(name),
      shape: val(kid(g, 'shape')) ?? val(deep(g, 'shape')) ?? 'box',
      radarStyle: val(kid(g, 'radarStyle')) ?? 'marker',
      hiLow: !!kid(g, 'hiLowLines'),
      upDown: ud ? { gap: +(val(kid(ud, 'gapWidth')) ?? 150), up: color(kid(kid(ud, 'upBars'), 'spPr'), 'fill'), down: color(kid(kid(ud, 'downBars'), 'spPr'), 'fill') } : undefined,
      bubbleScale: +(val(kid(g, 'bubbleScale')) ?? 100),
      sizeArea: val(kid(g, 'sizeRepresents')) !== 'w',
      wireframe: on(kid(g, 'wireframe')),
      ofPie: name === 'ofPieChart' ? { bar: val(kid(g, 'ofPieType')) === 'bar', splitPos: Math.max(1, +(val(kid(g, 'splitPos')) ?? 3)), secondSize: +(val(kid(g, 'secondPieSize')) ?? 75) } : undefined,
      horizontal: val(kid(g, 'barDir')) === 'bar',
      grouping: (val(kid(g, 'grouping')) as Group['grouping']) ?? (type === 'bar' ? 'clustered' : 'standard'),
      series: [],
      gap: +(val(kid(g, 'gapWidth')) ?? 150),
      overlap: +(val(kid(g, 'overlap')) ?? 0),
      hole: +(val(kid(g, 'holeSize')) ?? 50),
      firstAngle: +(val(kid(g, 'firstSliceAng')) ?? 0),
      secondary: false,
      vary: on(kid(g, 'varyColors'), type === 'pie' || type === 'doughnut'),
      showVal: on(kid(dl, 'showVal')),
      showPercent: on(kid(dl, 'showPercent')),
      showCat: on(kid(dl, 'showCatName')),
    };
    const ids = kids(g, 'axId').map((a) => val(a) ?? '');
    const valAx = ids.map((id) => axIds.get(id)).find((a) => a && local(a.name) === 'valAx');
    if (valAx && val(kid(valAx, 'axPos')) === 'r') grp.secondary = true;
    for (const s of kids(g, 'ser')) {
      const sp = kid(s, 'spPr');
      const sdl = kid(s, 'dLbls');
      if (sdl) {
        grp.showVal ||= on(kid(sdl, 'showVal'));
        grp.showPercent ||= on(kid(sdl, 'showPercent'));
      }
      const num = numbers(kid(s, 'val') ?? kid(s, 'yVal'));
      const ser: Series = {
        name: richText(kid(s, 'tx')),
        cats: strings(kid(s, 'cat') ?? kid(s, 'xVal')),
        vals: num.vals,
        format: num.format,
        color: color(sp, 'fill'),
        lineColor: color(sp, 'line'),
        lineWidth: kid(sp, 'ln')?.attrs.w ? +kid(sp, 'ln')!.attrs.w / 12700 : undefined,
        noLine: !!kid(kid(sp, 'ln'), 'noFill'),
        pointColors: new Map(),
        marker: (() => {
          const m = kid(s, 'marker');
          if (m) return val(kid(m, 'symbol')) !== 'none';
          if (type === 'radar') return grp.radarStyle === 'marker';
          return type === 'line' ? val(kid(g, 'marker')) !== '0' : type === 'scatter';
        })(),
        symbol: val(kid(kid(s, 'marker'), 'symbol')),
        smooth: on(kid(s, 'smooth')),
      };
      if (type === 'bubble') {
        const xs = numbers(kid(s, 'xVal')).vals;
        if (xs.some((x) => x !== null)) ser.xs = xs;
        ser.sizes = numbers(kid(s, 'bubbleSize')).vals;
        ser.marker = false;
        ser.noLine = true;
      }
      if (type === 'scatter') {
        const xs = numbers(kid(s, 'xVal')).vals;
        if (xs.some((x) => x !== null)) ser.xs = xs;
        const style = val(kid(g, 'scatterStyle')) ?? 'marker';
        if (/marker/i.test(style) && !/line/i.test(style)) ser.noLine ||= !kid(sp, 'ln') || !!kid(kid(sp, 'ln'), 'noFill');
      }
      for (const dp of kids(s, 'dPt')) {
        const c = color(kid(dp, 'spPr'), 'fill');
        if (c) ser.pointColors.set(+(val(kid(dp, 'idx')) ?? 0), c);
      }
      grp.series.push(ser);
    }
    if (grp.series.length) groups.push(grp);
  }
  const axisOf = (ax: XNode): Axis => {
    const sc = kid(ax, 'scaling');
    const min = val(kid(sc, 'min'));
    const max = val(kid(sc, 'max'));
    return {
      min: min !== undefined ? +min : undefined,
      max: max !== undefined ? +max : undefined,
      format: kid(ax, 'numFmt')?.attrs.formatCode,
      deleted: on(kid(ax, 'delete')),
      grid: !!kid(ax, 'majorGridlines'),
      reversed: val(kid(sc, 'orientation')) === 'maxMin',
      midCat: val(kid(ax, 'crossBetween')) === 'midCat',
    };
  };
  const axis = (kind: string, secondary = false): Axis | undefined => {
    const list = elements(plot).filter((a) => local(a.name) === kind);
    const ax = kind === 'valAx' ? (list.find((a) => (val(kid(a, 'axPos')) === 'r') === secondary) ?? (secondary ? undefined : list[0])) : list[0];
    return ax ? axisOf(ax) : undefined;
  };
  const catAxis = axis('catAx') ?? axis('dateAx') ?? { deleted: false, grid: false, reversed: false };
  // For scatter charts the vertical value axis is the one on the left or right
  const vertical = elements(plot).filter((a) => local(a.name) === 'valAx').find((a) => ['l', 'r'].includes(val(kid(a, 'axPos')) ?? ''));
  const valAxis = (vertical ? axisOf(vertical) : axis('valAx')) ?? { deleted: false, grid: true, reversed: false };
  const scatterX = groups.some((g) => g.type === 'scatter' || g.type === 'bubble') ? elements(plot).filter((a) => local(a.name) === 'valAx').find((a) => ['b', 't'].includes(val(kid(a, 'axPos')) ?? '')) : undefined;
  let title: string | undefined;
  const t = kid(chart, 'title');
  if (t) title = richText(kid(t, 'tx')) || (groups.length === 1 && groups[0].series.length === 1 ? groups[0].series[0].name : undefined);
  else if (!on(kid(chart, 'autoTitleDeleted')) && groups.length === 1 && groups[0].series.length === 1 && groups[0].type !== 'scatter') title = groups[0].series[0].name || undefined;
  const sz = deep(t, 'defRPr')?.attrs.sz ?? deep(t, 'rPr')?.attrs.sz;
  const legend = kid(chart, 'legend');
  const v3 = kid(chart, 'view3D');
  return {
    view3D: v3 || groups.some((g) => g.threeD)
      ? { rotX: +(val(kid(v3, 'rotX')) ?? (groups.some((g) => g.type === 'pie') ? 30 : 15)), rotY: +(val(kid(v3, 'rotY')) ?? 20), depthPercent: +(val(kid(v3, 'depthPercent')) ?? 100), rAngAx: on(kid(v3, 'rAngAx'), true) }
      : undefined,
    title,
    titleSize: sz ? +sz / 100 : 14,
    groups,
    catAxis,
    valAxis,
    val2Axis: groups.some((g) => g.secondary) ? axis('valAx', true) : undefined,
    xAxis: scatterX ? axisOf(scatterX) : undefined,
    legend: legend ? (val(kid(legend, 'legendPos')) ?? 'r') : undefined,
    bg: color(kid(xml.children.find((c): c is XNode => typeof c !== 'string' && local(c.name) === 'chartSpace') ?? xml, 'spPr'), 'fill'),
    plotBg: color(kid(plot, 'spPr'), 'fill'),
  };
}

const f = (n: number): string => (Math.round(n * 100) / 100).toString();

function textWidth(s: string, size: number): number {
  // Calibri-like average advance; good enough for layout of labels
  let w = 0;
  for (const ch of s) w += /[ilI.,:;'|!]/.test(ch) ? 0.27 : /[mwMW]/.test(ch) ? 0.82 : /[A-ZА-ЯЁ0-9]/.test(ch) ? 0.58 : 0.5;
  return w * size;
}

/** Render a chart part to SVG (width/height in points). Returns null when the chart has no data. */
export function chartToSvg(xmlText: string, opts: ChartOptions): string | null {
  const xml = parseXml(xmlText);
  const theme: Theme = { colors: { ...Object.fromEntries(OFFICE_ACCENTS.map((c, i) => [`accent${i + 1}`, c])), dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', ...(opts.colors ?? {}) }, major: '', minor: '', fills: [] };
  const c = parseChart(xml, theme);
  if (!c.groups.length || c.groups.every((g) => g.series.every((s) => s.vals.every((v) => v === null)))) return null;
  const W = opts.width;
  const H = opts.height;
  const palette = [1, 2, 3, 4, 5, 6].map((i) => theme.colors[`accent${i}`] ?? OFFICE_ACCENTS[i - 1]);
  const shade = (hex: string, k: number): string =>
    [0, 2, 4]
      .map((o) => Math.round(parseInt(hex.slice(o, o + 2), 16) * k).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  const auto = (i: number): string => (i < 6 ? palette[i] : shade(palette[i % 6], i < 12 ? 0.6 : 0.8));
  let out = '';
  const rect = (x: number, y: number, w: number, h: number, fill: string, stroke?: string): void => {
    out += `<rect x="${f(x)}" y="${f(y)}" width="${f(Math.max(0, w))}" height="${f(Math.max(0, h))}" fill="#${fill}"${stroke ? ` stroke="#${stroke}" stroke-width="0.75"` : ''}/>`;
  };
  const line = (x1: number, y1: number, x2: number, y2: number, color: string, w = 0.75): void => {
    out += `<line x1="${f(x1)}" y1="${f(y1)}" x2="${f(x2)}" y2="${f(y2)}" stroke="#${color}" stroke-width="${w}"/>`;
  };
  const text = (x: number, y: number, s: string, size: number, anchor = 'start', color = LABEL, bold = false): void => {
    out += `<text x="${f(x)}" y="${f(y)}" font-family="${FONT}" font-size="${f(size)}"${bold ? ' font-weight="bold"' : ''} fill="#${color}"${anchor !== 'start' ? ` text-anchor="${anchor}"` : ''}>${escapeXml(s)}</text>`;
  };
  if (c.bg) rect(0, 0, W, H, c.bg);
  const fs = Math.max(7, Math.min(10, Math.min(W, H) / 28));
  let top = 6;
  let bottom = H - 6;
  let left = 6;
  let right = W - 6;
  if (c.title) {
    const ts = Math.min(c.titleSize, H / 10);
    const lines = c.title.split('\n');
    lines.forEach((l, i) => text(W / 2, top + ts * (i + 1), l, ts, 'middle', '404040'));
    top += ts * lines.length + ts * 0.6;
  }
  const pieLike = c.groups.every((g) => g.type === 'pie' || g.type === 'doughnut' || g.type === 'ofPie');
  const surface = c.groups[0].type === 'surface';
  const radar = c.groups[0].type === 'radar';
  // Legend entries
  const entries: { name: string; color: string; line: boolean }[] = [];
  if (c.legend) {
    if (surface) {
      // value bands, not series
      const b = surfaceBands(c);
      const fmt = (v: number): string => formatNumber(v, c.valAxis.format ?? c.groups[0].series[0]?.format);
      for (let k = 0; k + 1 < b.length; k++) entries.push({ name: `${fmt(b[k])}–${fmt(b[k + 1])}`, color: auto(k), line: false });
    } else if (pieLike) {
      const s = c.groups[0].series[0];
      s.vals.forEach((_, i) => entries.push({ name: s.cats[i] ?? String(i + 1), color: s.pointColors.get(i) ?? auto(i), line: false }));
    } else {
      let k = 0;
      for (const g of c.groups)
        for (const s of g.series) {
          const lineLike = g.type === 'line' || g.type === 'scatter' || g.type === 'stock' || (g.type === 'radar' && g.radarStyle !== 'filled');
          if (g.type === 'stock' && s.noLine && !s.marker) {
            k++;
            continue;
          }
          entries.push({ name: s.name || `Ряд ${k + 1}`, color: (lineLike ? s.lineColor ?? s.color : s.color) ?? auto(k), line: lineLike });
          k++;
        }
    }
  }
  const legendItemW = (e: { name: string }): number => 14 + textWidth(e.name, fs) + 10;
  if (entries.length) {
    if (c.legend === 'r' || c.legend === 'l') {
      const lw = Math.min(W * 0.35, Math.max(...entries.map(legendItemW)) + 6);
      const lx = c.legend === 'r' ? right - lw : left;
      const total = entries.length * fs * 1.6;
      let ly = top + Math.max(0, (bottom - top - total) / 2) + fs;
      for (const e of entries) {
        if (e.line) line(lx + 2, ly - fs * 0.35, lx + 12, ly - fs * 0.35, e.color, 2);
        else rect(lx + 2, ly - fs * 0.8, 8, 8, e.color);
        text(lx + 15, ly, e.name, fs);
        ly += fs * 1.6;
      }
      if (c.legend === 'r') right -= lw + 4;
      else left += lw + 4;
    } else {
      // top/bottom: centred rows
      const rows: { name: string; color: string; line: boolean }[][] = [[]];
      let rowW = 0;
      for (const e of entries) {
        if (rowW + legendItemW(e) > right - left && rows[rows.length - 1].length) {
          rows.push([]);
          rowW = 0;
        }
        rows[rows.length - 1].push(e);
        rowW += legendItemW(e);
      }
      const blockH = rows.length * fs * 1.6;
      let ly = c.legend === 't' ? top + fs : bottom - blockH + fs;
      for (const row of rows) {
        const rw = row.reduce((sum, e) => sum + legendItemW(e), 0);
        let lx = (W - rw) / 2;
        for (const e of row) {
          if (e.line) line(lx, ly - fs * 0.35, lx + 10, ly - fs * 0.35, e.color, 2);
          else rect(lx, ly - fs * 0.8, 8, 8, e.color);
          text(lx + 13, ly, e.name, fs);
          lx += legendItemW(e);
        }
        ly += fs * 1.6;
      }
      if (c.legend === 't') top += blockH + 4;
      else bottom -= blockH + 4;
    }
  }
  if (radar) drawRadar(c, left, top, right, bottom, fs, auto, (s) => (out += s));
  else if (surface) drawSurface(c, left, top, right, bottom, fs, auto, (s) => (out += s));
  else if (pieLike) {
    drawPies(c.groups, left, top, right, bottom, fs, auto, (s) => (out += s), c.view3D);
  } else {
    drawAxes(c, left, top, right, bottom, fs, auto, (s) => (out += s));
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${f(W)} ${f(H)}" width="${f(W)}pt" height="${f(H)}pt">${out}</svg>`;
}

function drawPies(groups: Group[], left: number, top: number, right: number, bottom: number, fs: number, auto: (i: number) => string, emit: (s: string) => void, view?: View3D): void {
  const g = groups[0];
  if (g.type === 'ofPie' && g.ofPie) return drawOfPie(g, left, top, right, bottom, fs, auto, emit);
  if (g.threeD && g.type === 'pie') return drawPie3D(g, left, top, right, bottom, fs, auto, emit, view);
  const s = g.series[0];
  const vals = s.vals.map((v) => Math.max(0, v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (!total) return;
  const showLabels = g.showPercent || g.showVal || g.showCat;
  const pad = showLabels ? fs * 2.2 : 4;
  const r = Math.max(10, Math.min(right - left, bottom - top) / 2 - pad);
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2;
  const inner = g.type === 'doughnut' ? (r * Math.min(90, Math.max(10, g.hole))) / 100 : 0;
  let a = ((g.firstAngle - 90) * Math.PI) / 180;
  vals.forEach((v, i) => {
    if (!v) return;
    const sweep = (v / total) * Math.PI * 2;
    const a2 = a + sweep;
    const large = sweep > Math.PI ? 1 : 0;
    const p = (rad: number, ang: number): string => `${f(cx + rad * Math.cos(ang))} ${f(cy + rad * Math.sin(ang))}`;
    const color = s.pointColors.get(i) ?? (g.vary ? auto(i) : (s.color ?? auto(0)));
    let d: string;
    if (sweep >= Math.PI * 2 - 1e-6) {
      d = `M${p(r, a)}A${f(r)} ${f(r)} 0 1 1 ${p(r, a + Math.PI)}A${f(r)} ${f(r)} 0 1 1 ${p(r, a)}Z`;
      if (inner) d += `M${p(inner, a)}A${f(inner)} ${f(inner)} 0 1 0 ${p(inner, a + Math.PI)}A${f(inner)} ${f(inner)} 0 1 0 ${p(inner, a)}Z`;
    } else if (inner) d = `M${p(r, a)}A${f(r)} ${f(r)} 0 ${large} 1 ${p(r, a2)}L${p(inner, a2)}A${f(inner)} ${f(inner)} 0 ${large} 0 ${p(inner, a)}Z`;
    else d = `M${f(cx)} ${f(cy)}L${p(r, a)}A${f(r)} ${f(r)} 0 ${large} 1 ${p(r, a2)}Z`;
    emit(`<path d="${d}" fill="#${color}" fill-rule="evenodd" stroke="#FFFFFF" stroke-width="1"/>`);
    if (showLabels) {
      const mid = a + sweep / 2;
      const lr = inner ? (r + inner) / 2 : r * 0.68;
      const parts: string[] = [];
      if (g.showCat && s.cats[i]) parts.push(s.cats[i]);
      if (g.showVal) parts.push(formatNumber(v, s.format));
      if (g.showPercent) parts.push(`${Math.round((v / total) * 100)}%`);
      const tx = cx + lr * Math.cos(mid);
      const ty = cy + lr * Math.sin(mid) + fs * 0.35;
      emit(`<text x="${f(tx)}" y="${f(ty)}" font-family="${FONT}" font-size="${f(fs)}" font-weight="bold" fill="#FFFFFF" text-anchor="middle">${escapeXml(parts.join(' '))}</text>`);
    }
    a = a2;
  });
}

function drawAxes(c: ReturnType<typeof parseChart>, left: number, top: number, right: number, bottom: number, fs: number, auto: (i: number) => string, emit: (s: string) => void): void {
  const groups = c.groups.filter((g) => g.type !== 'pie' && g.type !== 'doughnut' && g.type !== 'ofPie' && g.type !== 'radar' && g.type !== 'surface');
  const scatter = groups.some((g) => g.type === 'scatter' || g.type === 'bubble');
  const bubbles = groups.some((g) => g.type === 'bubble');
  const horizontal = groups.some((g) => g.type === 'bar' && g.horizontal);
  const cats = groups.flatMap((g) => g.series).reduce<string[]>((best, s) => (s.cats.length > best.length ? s.cats : best), []);
  const nCats = Math.max(1, cats.length, ...groups.flatMap((g) => g.series.map((s) => s.vals.length)));
  // Value ranges (stacked groups sum per category)
  const range = (secondary: boolean): { min: number; max: number } => {
    let lo = 0;
    let hi = 0;
    let any = false;
    for (const g of groups.filter((x) => x.secondary === secondary)) {
      if (g.grouping === 'percentStacked') {
        hi = Math.max(hi, 1);
        any = true;
        continue;
      }
      if (g.grouping === 'stacked') {
        for (let i = 0; i < nCats; i++) {
          let pos = 0;
          let neg = 0;
          for (const s of g.series) {
            const v = s.vals[i] ?? 0;
            if (v >= 0) pos += v;
            else neg += v;
          }
          hi = Math.max(hi, pos);
          lo = Math.min(lo, neg);
          any = true;
        }
      } else
        for (const s of g.series)
          for (const v of s.vals)
            if (v !== null) {
              hi = Math.max(hi, v);
              lo = Math.min(lo, v);
              any = true;
            }
    }
    return any ? { min: lo, max: hi } : { min: 0, max: 1 };
  };
  // Longer axes get more gridlines (as in Office): about one interval per 4.5 em, never narrower than a label
  const vLen = horizontal ? (right - left) * 0.75 : bottom - top - fs * 2;
  const xLen = (right - left) * 0.8;
  const intervals = (len: number, along: boolean, sample: number): number => Math.min(10, Math.max(3, Math.round(len / Math.max(fs * 4.5, along ? textWidth(formatNumber(sample), fs) * 1.6 : 0))));
  const r1 = range(false);
  if (scatter) r1.min = autoMin(Math.min(...groups.flatMap((g) => g.series.flatMap((s) => s.vals.filter((v): v is number => v !== null)))), r1.max);

  const vs = niceScale(r1.min, r1.max, c.valAxis.min, c.valAxis.max, intervals(vLen, horizontal, Math.max(Math.abs(r1.min), Math.abs(r1.max))));
  const hasSecondary = groups.some((g) => g.secondary);
  const r2 = hasSecondary ? range(true) : null;
  const vs2 = r2 ? niceScale(r2.min, r2.max, c.val2Axis?.min, c.val2Axis?.max, intervals(vLen, horizontal, Math.max(Math.abs(r2.min), Math.abs(r2.max)))) : null;
  const pct = groups.some((g) => g.grouping === 'percentStacked' && !g.secondary);
  const valFmt = (v: number, fmt?: string): string => (pct ? `${Math.round(v * 100)}%` : formatNumber(v, fmt ?? c.valAxis.format ?? groups[0]?.series[0]?.format));
  // X scale for scatter
  let xs: { min: number; max: number; step: number } | null = null;
  if (scatter) {
    const all = groups.flatMap((g) => g.series.flatMap((s) => (s.xs ?? s.vals.map((_, i) => i + 1)).filter((x): x is number => x !== null)));
    const xr = { min: autoMin(Math.min(...all), Math.max(...all)), max: Math.max(...all) };
    xs = niceScale(xr.min, xr.max, c.xAxis?.min, c.xAxis?.max, intervals(xLen, true, Math.max(...all.map(Math.abs))));
  }
  // Space for labels
  const ticks = (sc: { min: number; max: number; step: number }): number[] => {
    const out: number[] = [];
    for (let v = sc.min, k = 0; v <= sc.max + sc.step * 1e-6 && k < 50; v += sc.step, k++) out.push(Math.round(v / sc.step) * sc.step);
    return out;
  };
  const vTicks = ticks(vs);
  const vLabelW = c.valAxis.deleted ? 0 : Math.max(...vTicks.map((t) => textWidth(valFmt(t), fs))) + 6;
  const v2Ticks = vs2 ? ticks(vs2) : [];
  const v2LabelW = vs2 && !c.val2Axis?.deleted ? Math.max(...v2Ticks.map((t) => textWidth(formatNumber(t, c.val2Axis?.format), fs))) + 6 : 0;
  const catLabels = scatter ? ticks(xs!).map((t) => formatNumber(t, c.xAxis?.format)) : cats;
  let plotL: number;
  let plotR: number;
  let rotate = false;
  let plotT = top + fs * 0.6;
  let plotB: number;
  if (horizontal) {
    const catW = c.catAxis.deleted ? 0 : Math.min((right - left) * 0.35, Math.max(0, ...cats.map((t) => textWidth(t, fs))) + 6);
    plotL = left + catW;
    plotR = right - 4;
    plotB = bottom - (c.valAxis.deleted ? 2 : fs * 1.6);
  } else {
    plotL = left + vLabelW;
    plotR = right - (v2LabelW || 6);
    const slot = (plotR - plotL) / Math.max(1, catLabels.length);
    const longest = Math.max(0, ...catLabels.map((t) => textWidth(t, fs)));
    rotate = !c.catAxis.deleted && longest > slot * 0.95;
    const catH = c.catAxis.deleted ? 2 : rotate ? Math.min((bottom - top) * 0.4, longest * 0.71 + fs) : fs * 1.6;
    plotB = bottom - catH;
    plotT = Math.min(plotT, plotB - 10);
  }
  // 3-D: the back wall is offset by the depth vector
  const d3 = c.view3D && groups.some((g) => g.threeD && (g.type === 'bar' || g.type === 'line' || g.type === 'area')) ? depthOf(c.view3D, plotR - plotL, plotB - plotT) : null;
  // a series axis (standard 3-D lines/areas) puts every series in its own row: more depth
  const rowsInDepth = Math.max(1, ...groups.filter((g) => g.threeD && (g.type === 'line' || g.type === 'area') && g.grouping === 'standard').map((g) => g.series.length));
  if (d3 && rowsInDepth > 1) {
    const k = Math.min(3.5, 1.5 + 0.8 * (rowsInDepth - 1));
    d3.dx *= k;
    d3.dy *= k;
  }
  if (d3) {
    if (d3.dx > 0) plotR -= d3.dx;
    else plotL -= d3.dx;
    if (d3.dy < 0) plotT -= d3.dy;
    else plotB -= d3.dy;
  }
  const plotW = plotR - plotL;
  const plotH = plotB - plotT;
  if (plotW < 10 || plotH < 10) return;
  if (d3) {
    const { dx, dy } = d3;
    const back: [number, number][] = [
      [plotL + dx, plotT + dy],
      [plotR + dx, plotT + dy],
      [plotR + dx, plotB + dy],
      [plotL + dx, plotB + dy],
    ];
    emit(`<path d="${poly(back)}" fill="none" stroke="#${GRID}" stroke-width="0.75"/>`);
    emit(`<path d="${poly([[plotL, plotT], [plotL + dx, plotT + dy], [plotL + dx, plotB + dy], [plotL, plotB]])}" fill="none" stroke="#${GRID}" stroke-width="0.75"/>`);
    emit(`<path d="${poly([[plotL, plotB], [plotR, plotB], [plotR + dx, plotB + dy], [plotL + dx, plotB + dy]])}" fill="none" stroke="#${AXIS}" stroke-width="0.75"/>`);
  }
  if (c.plotBg) emit(`<rect x="${f(plotL)}" y="${f(plotT)}" width="${f(plotW)}" height="${f(plotH)}" fill="#${c.plotBg}"/>`);
  // value → coordinate
  const vPos = (v: number, sc = vs): number => (horizontal ? plotL + ((v - sc.min) / (sc.max - sc.min)) * plotW : plotB - ((v - sc.min) / (sc.max - sc.min)) * plotH);
  const slot = (horizontal ? plotH : plotW) / nCats;
  // Bar charts list the first category at the bottom, like Office; midCat puts points on the edges
  const midCat = !!c.valAxis.midCat && !horizontal && nCats > 1 && groups.every((g) => g.type !== 'bar');
  const catPos = (i: number): number => {
    const k = c.catAxis.reversed !== horizontal ? nCats - 1 - i : i;
    if (midCat) return plotL + (k * plotW) / (nCats - 1);
    return horizontal ? plotT + (k + 0.5) * slot : plotL + (k + 0.5) * slot;
  };
  const xPos = (x: number): number => plotL + ((x - xs!.min) / (xs!.max - xs!.min)) * plotW;
  // Gridlines and value labels
  for (const t of vTicks) {
    const p = vPos(t);
    if (c.valAxis.grid && d3) {
      // on the back wall and across the side wall (columns) or the floor (bars)
      const { dx, dy } = d3;
      emit(
        horizontal
          ? `<path d="M${f(p)} ${f(plotB)}L${f(p + dx)} ${f(plotB + dy)}L${f(p + dx)} ${f(plotT + dy)}" fill="none" stroke="#${GRID}" stroke-width="0.75"/>`
          : `<path d="M${f(plotL)} ${f(p)}L${f(plotL + dx)} ${f(p + dy)}L${f(plotR + dx)} ${f(p + dy)}" fill="none" stroke="#${GRID}" stroke-width="0.75"/>`,
      );
    } else if (c.valAxis.grid) emit(horizontal ? `<line x1="${f(p)}" y1="${f(plotT)}" x2="${f(p)}" y2="${f(plotB)}" stroke="#${GRID}" stroke-width="0.75"/>` : `<line x1="${f(plotL)}" y1="${f(p)}" x2="${f(plotR)}" y2="${f(p)}" stroke="#${GRID}" stroke-width="0.75"/>`);
    if (!c.valAxis.deleted) {
      const label = escapeXml(valFmt(t));
      emit(horizontal ? `<text x="${f(p)}" y="${f(plotB + fs * 1.25)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="middle">${label}</text>` : `<text x="${f(plotL - 4)}" y="${f(p + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="end">${label}</text>`);
    }
  }
  if (vs2 && !c.val2Axis?.deleted)
    for (const t of v2Ticks) emit(`<text x="${f(plotR + 4)}" y="${f(vPos(t, vs2) + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}">${escapeXml(formatNumber(t, c.val2Axis?.format))}</text>`);
  // Category labels / scatter x labels
  if (scatter && xs) {
    for (const t of ticks(xs)) {
      const p = xPos(t);
      if (c.xAxis?.grid) emit(`<line x1="${f(p)}" y1="${f(plotT)}" x2="${f(p)}" y2="${f(plotB)}" stroke="#${GRID}" stroke-width="0.75"/>`);
      if (!c.xAxis?.deleted) emit(`<text x="${f(p)}" y="${f(plotB + fs * 1.25)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="middle">${escapeXml(formatNumber(t, c.xAxis?.format))}</text>`);
    }
  } else if (!c.catAxis.deleted) {
    cats.forEach((t, i) => {
      const p = catPos(i);
      if (horizontal) emit(`<text x="${f(plotL - 4)}" y="${f(p + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="end">${escapeXml(t)}</text>`);
      else if (rotate) emit(`<text x="${f(p)}" y="${f(plotB + fs)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="end" transform="rotate(-45 ${f(p)} ${f(plotB + fs)})">${escapeXml(t)}</text>`);
      else emit(`<text x="${f(p)}" y="${f(plotB + fs * 1.25)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="middle">${escapeXml(t)}</text>`);
    });
  }
  // Axis line at zero / bottom
  const zero = vPos(Math.min(Math.max(0, vs.min), vs.max));
  if (!d3) emit(horizontal ? `<line x1="${f(zero)}" y1="${f(plotT)}" x2="${f(zero)}" y2="${f(plotB)}" stroke="#${AXIS}" stroke-width="0.75"/>` : `<line x1="${f(plotL)}" y1="${f(zero)}" x2="${f(plotR)}" y2="${f(zero)}" stroke="#${AXIS}" stroke-width="0.75"/>`);
  // Series
  let seriesIndex = 0;
  const labels: string[] = [];
  const prisms: { x: number; y: number; svg: string }[] = [];
  /** A 3-D bar: front rectangle with top and side faces (or a cylinder, cone or pyramid). */
  const prism = (x: number, y: number, w: number, h: number, col: string, shape: string, depth: number): void => {
    const dx = d3!.dx * depth;
    const dy = d3!.dy * depth;
    const top = tone(col, 1.25);
    const side = tone(col, 0.72);
    const edge = ` stroke="#${tone(col, 0.6)}" stroke-width="0.4"`;
    let svg: string;
    if (/^cylinder/.test(shape)) {
      const rx = w / 2;
      const ry = Math.max(1.5, Math.abs(dy) / 2);
      const cxm = x + w / 2 + dx / 2;
      svg =
        `<path d="M${f(cxm - rx)} ${f(y + dy / 2)}L${f(cxm - rx)} ${f(y + h + dy / 2)}A${f(rx)} ${f(ry)} 0 0 0 ${f(cxm + rx)} ${f(y + h + dy / 2)}L${f(cxm + rx)} ${f(y + dy / 2)}Z" fill="#${col}"${edge}/>` +
        `<ellipse cx="${f(cxm)}" cy="${f(y + dy / 2)}" rx="${f(rx)}" ry="${f(ry)}" fill="#${top}"${edge}/>`;
    } else if (/^(cone|pyramid)/.test(shape)) {
      const apex: [number, number] = [x + w / 2 + dx / 2, y + dy / 2];
      if (/^cone/.test(shape)) {
        const rx = w / 2;
        const ry = Math.max(1.5, Math.abs(dy) / 2);
        const cxm = x + w / 2 + dx / 2;
        svg = `<path d="M${f(cxm - rx)} ${f(y + h + dy / 2)}A${f(rx)} ${f(ry)} 0 0 0 ${f(cxm + rx)} ${f(y + h + dy / 2)}L${f(apex[0])} ${f(apex[1])}Z" fill="#${col}"${edge}/>`;
      } else
        svg =
          `<path d="${poly([[x, y + h], [x + w, y + h], apex])}" fill="#${col}"${edge}/>` +
          `<path d="${poly([[x + w, y + h], [x + w + dx, y + h + dy], apex])}" fill="#${side}"${edge}/>`;
    } else
      svg =
        `<path d="${poly([[x, y], [x + dx, y + dy], [x + w + dx, y + dy], [x + w, y]])}" fill="#${top}"${edge}/>` +
        `<path d="${poly([[x + w, y], [x + w + dx, y + dy], [x + w + dx, y + h + dy], [x + w, y + h]])}" fill="#${side}"${edge}/>` +
        `<rect x="${f(x)}" y="${f(y)}" width="${f(w)}" height="${f(h)}" fill="#${col}"${edge}/>`;
    prisms.push({ x, y: y + h, svg });
  };
  for (const g of groups) {
    const sc = g.secondary && vs2 ? vs2 : vs;
    const n = g.series.length;
    const color = (s: Series, k: number): string => s.color ?? auto(k);
    if (g.type === 'bar') {
      // Office layout: a slot holds n bars that overlap by `ov` of a bar, plus a gap of gapWidth % of a bar
      const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
      const ov = stacked ? 1 : Math.max(-1, Math.min(1, g.overlap / 100));
      const span = n - (n - 1) * ov;
      const barW = slot / (span + g.gap / 100);
      const groupW = barW * span;
      const pos = new Array(nCats).fill(0);
      const neg = new Array(nCats).fill(0);
      const totals = Array.from({ length: nCats }, (_, i) => g.series.reduce((sum, s) => sum + Math.abs(s.vals[i] ?? 0), 0) || 1);
      g.series.forEach((s, si) => {
        const k = seriesIndex + si;
        for (let i = 0; i < nCats; i++) {
          let v = s.vals[i];
          if (v === null || v === undefined) continue;
          if (g.grouping === 'percentStacked') v /= totals[i];
          let a0: number;
          let a1: number;
          if (stacked) {
            if (v >= 0) {
              a0 = pos[i];
              pos[i] += v;
              a1 = pos[i];
            } else {
              a0 = neg[i];
              neg[i] += v;
              a1 = neg[i];
            }
          } else {
            a0 = Math.min(Math.max(0, sc.min), sc.max);
            a1 = v;
          }
          const fill = s.pointColors.get(i) ?? (g.vary && n === 1 ? auto(i) : color(s, k));
          const off = catPos(i) - groupW / 2 + si * barW * (1 - ov);
          const p0 = vPos(a0, sc);
          const p1 = vPos(a1, sc);
          if (g.threeD && d3) {
            if (horizontal) prism(Math.min(p0, p1), off, Math.abs(p1 - p0), barW, fill, 'box', 0.7);
            else prism(off, Math.min(p0, p1), barW, Math.abs(p1 - p0), fill, g.shape, 0.7);
          } else if (horizontal) emit(`<rect x="${f(Math.min(p0, p1))}" y="${f(off)}" width="${f(Math.abs(p1 - p0))}" height="${f(barW)}" fill="#${fill}"/>`);
          else emit(`<rect x="${f(off)}" y="${f(Math.min(p0, p1))}" width="${f(barW)}" height="${f(Math.abs(p1 - p0))}" fill="#${fill}"/>`);
          if (g.showVal) {
            const t = escapeXml(g.grouping === 'percentStacked' ? `${Math.round(v * 100)}%` : formatNumber(s.vals[i]!, s.format));
            if (horizontal) labels.push(`<text x="${f(Math.max(p0, p1) + 3)}" y="${f(off + barW / 2 + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#404040">${t}</text>`);
            else labels.push(`<text x="${f(off + barW / 2)}" y="${f(Math.min(p0, p1) - 3)}" font-family="${FONT}" font-size="${f(fs)}" fill="#404040" text-anchor="middle">${t}</text>`);
          }
        }
      });
      if (prisms.length) {
        // painter's order: bars further back (left for columns, bottom for bars, lower in stacks) first
        const dir = d3 && d3.dx < 0 ? -1 : 1;
        prisms.sort((a, b) => (horizontal ? b.y - a.y || a.x - b.x : (a.x - b.x) * dir || b.y - a.y));
        for (const pr of prisms) emit(pr.svg);
        prisms.length = 0;
      }
      seriesIndex += n;
      continue;
    }
    if (g.type === 'stock') {
      // high-low lines, open-close (up/down) bars, then the series markers
      const barW = g.upDown ? slot / (1 + g.upDown.gap / 100) : 0;
      for (let i = 0; i < nCats; i++) {
        const vals = g.series.map((s) => s.vals[i]).filter((v): v is number => v !== null && v !== undefined);
        if (!vals.length) continue;
        const x = catPos(i);
        if (g.hiLow && vals.length > 1) emit(`<line x1="${f(x)}" y1="${f(vPos(Math.max(...vals), sc))}" x2="${f(x)}" y2="${f(vPos(Math.min(...vals), sc))}" stroke="#404040" stroke-width="0.75"/>`);
        const open = g.series[0]?.vals[i];
        const close = g.series[n - 1]?.vals[i];
        if (g.upDown && n >= 4 && open != null && close != null) {
          const up = close >= open;
          const y0 = vPos(Math.max(open, close), sc);
          const y1 = vPos(Math.min(open, close), sc);
          emit(`<rect x="${f(x - barW / 2)}" y="${f(y0)}" width="${f(barW)}" height="${f(Math.max(0.75, y1 - y0))}" fill="#${up ? (g.upDown.up ?? 'FFFFFF') : (g.upDown.down ?? '000000')}" stroke="#404040" stroke-width="0.75"/>`);
        }
      }
      g.series.forEach((s, si) => {
        const col = (s.lineColor ?? s.color) ?? auto(seriesIndex + si);
        const pts = s.vals.map((v, i) => (v === null ? null : ([catPos(i), vPos(v, sc)] as [number, number]))).filter((q): q is [number, number] => !!q);
        if (!s.noLine && pts.length > 1) emit(`<path d="M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}" fill="none" stroke="#${col}" stroke-width="${f(s.lineWidth ?? 2.25)}"/>`);
        if (s.marker) for (const [x, y] of pts) emit(markerSvg(s.symbol, x, y, Math.max(2, fs * 0.3), s.color ?? col, col));
      });
      seriesIndex += n;
      continue;
    }
    if (g.type === 'bubble' && xs) {
      // bubble area (or width) proportional to its size; the largest is a quarter of the plot
      const sizes = g.series.flatMap((s) => (s.sizes ?? []).filter((v): v is number => v !== null && v > 0));
      const maxSize = Math.max(...sizes, 1e-9);
      const maxR = (Math.min(plotW, plotH) * 0.125 * Math.max(0, Math.min(300, g.bubbleScale))) / 100;
      const items: { r: number; svg: string }[] = [];
      g.series.forEach((s, si) => {
        const col = s.color ?? auto(seriesIndex + si);
        s.vals.forEach((v, i) => {
          const size = s.sizes?.[i];
          if (v === null || size === null || size === undefined || size <= 0) return;
          const r = g.sizeArea ? maxR * Math.sqrt(size / maxSize) : (maxR * size) / maxSize;
          const x = xPos(s.xs?.[i] ?? i + 1);
          const y = vPos(v, sc);
          items.push({ r, svg: `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r)}" fill="#${s.pointColors.get(i) ?? col}" fill-opacity="0.8" stroke="#FFFFFF" stroke-width="0.75"/>` });
          if (g.showVal) labels.push(`<text x="${f(x)}" y="${f(y + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#404040" text-anchor="middle">${escapeXml(formatNumber(v, s.format))}</text>`);
        });
      });
      for (const it of items.sort((a, b) => b.r - a.r)) emit(it.svg);
      seriesIndex += n;
      continue;
    }
    if (g.threeD && d3 && (g.type === 'line' || g.type === 'area')) {
      // series as slabs (areas) or ribbons (lines) in depth rows, the back row first
      const stackedRows = g.grouping === 'stacked' || g.grouping === 'percentStacked';
      const rowsN = stackedRows ? 1 : n;
      const totals = Array.from({ length: nCats }, (_, i) => g.series.reduce((sum, s) => sum + Math.abs(s.vals[i] ?? 0), 0) || 1);
      const cum = new Array(nCats).fill(0);
      const shapes: { row: number; order: number; svg: string }[] = [];
      g.series.forEach((s, si) => {
        const row = stackedRows ? 0 : si;
        const o = { x: (d3.dx * row) / rowsN, y: (d3.dy * row) / rowsN };
        const t = { x: (d3.dx / rowsN) * 0.8, y: (d3.dy / rowsN) * 0.8 };
        const col = (g.type === 'area' ? s.color : (s.lineColor ?? s.color)) ?? auto(seriesIndex + si);
        const pts: [number, number][] = [];
        const base: [number, number][] = [];
        for (let i = 0; i < nCats; i++) {
          let v = s.vals[i];
          if (v === null || v === undefined) continue;
          if (g.grouping === 'percentStacked') v /= totals[i];
          const b = stackedRows ? cum[i] : Math.min(Math.max(0, sc.min), sc.max);
          if (stackedRows) cum[i] += v;
          const top = stackedRows ? cum[i] : v;
          pts.push([catPos(i) + o.x, vPos(top, sc) + o.y]);
          base.push([catPos(i) + o.x, vPos(b, sc) + o.y]);
        }
        if (!pts.length) return;
        let svg = '';
        for (let i = 0; i + 1 < pts.length; i++) {
          const a = pts[i];
          const b = pts[i + 1];
          svg += `<path d="${poly([a, b, [b[0] + t.x, b[1] + t.y], [a[0] + t.x, a[1] + t.y]])}" fill="#${g.type === 'area' ? tone(col, 1.2) : col}" stroke="#${tone(col, 0.7)}" stroke-width="0.4"/>`;
        }
        if (g.type === 'line' && pts.length > 1) svg += `<path d="M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}" fill="none" stroke="#${tone(col, 0.7)}" stroke-width="1.25"/>`;
        if (g.type === 'area') {
          const last = pts.length - 1;
          svg += `<path d="${poly([pts[last], base[last], [base[last][0] + t.x, base[last][1] + t.y], [pts[last][0] + t.x, pts[last][1] + t.y]])}" fill="#${tone(col, 0.72)}" stroke="#${tone(col, 0.6)}" stroke-width="0.4"/>`;
          svg += `<path d="M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}L${[...base].reverse().map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}Z" fill="#${col}" stroke="#${tone(col, 0.6)}" stroke-width="0.4"/>`;
        }
        shapes.push({ row, order: si, svg });
      });
      shapes.sort((a, b) => b.row - a.row || a.order - b.order);
      for (const sh of shapes) emit(sh.svg);
      seriesIndex += n;
      continue;
    }
    // line / area / scatter
    const cum = new Array(nCats).fill(0);
    const totals = Array.from({ length: nCats }, (_, i) => g.series.reduce((sum, s) => sum + Math.abs(s.vals[i] ?? 0), 0) || 1);
    const areaPolys: string[] = [];
    g.series.forEach((s, si) => {
      const k = seriesIndex + si;
      const col = (g.type === 'area' ? s.color : s.lineColor ?? s.color) ?? auto(k);
      const pts: [number, number][] = [];
      const base: [number, number][] = [];
      for (let i = 0; i < s.vals.length; i++) {
        let v = s.vals[i];
        if (v === null) continue;
        if (g.grouping === 'percentStacked') v /= totals[i];
        const b = g.grouping === 'stacked' || g.grouping === 'percentStacked' ? cum[i] : 0;
        if (g.grouping === 'stacked' || g.grouping === 'percentStacked') cum[i] += v;
        const vy = b + (g.grouping === 'standard' || g.grouping === 'clustered' ? v : v);
        const x = g.type === 'scatter' && xs ? xPos(s.xs?.[i] ?? i + 1) : catPos(i);
        pts.push([x, vPos(g.grouping === 'stacked' || g.grouping === 'percentStacked' ? cum[i] : vy, sc)]);
        base.push([x, vPos(g.grouping === 'stacked' || g.grouping === 'percentStacked' ? b : Math.min(Math.max(0, sc.min), sc.max), sc)]);
        if (g.showVal) labels.push(`<text x="${f(x)}" y="${f(pts[pts.length - 1][1] - 5)}" font-family="${FONT}" font-size="${f(fs)}" fill="#404040" text-anchor="middle">${escapeXml(formatNumber(s.vals[i]!, s.format))}</text>`);
      }
      if (!pts.length) return;
      if (g.type === 'area') {
        const d = `M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}L${[...base].reverse().map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}Z`;
        areaPolys.push(`<path d="${d}" fill="#${col}" fill-opacity="0.85"/>`);
        return;
      }
      if (!s.noLine && pts.length > 1) {
        let d = `M${f(pts[0][0])} ${f(pts[0][1])}`;
        if (s.smooth) {
          for (let i = 0; i + 1 < pts.length; i++) {
            const p0 = pts[i - 1] ?? pts[i];
            const p1 = pts[i];
            const p2 = pts[i + 1];
            const p3 = pts[i + 2] ?? p2;
            d += `C${f(p1[0] + (p2[0] - p0[0]) / 6)} ${f(p1[1] + (p2[1] - p0[1]) / 6)} ${f(p2[0] - (p3[0] - p1[0]) / 6)} ${f(p2[1] - (p3[1] - p1[1]) / 6)} ${f(p2[0])} ${f(p2[1])}`;
          }
        } else d += pts.slice(1).map(([x, y]) => `L${f(x)} ${f(y)}`).join('');
        emit(`<path d="${d}" fill="none" stroke="#${col}" stroke-width="${f(s.lineWidth ?? 2.25)}" stroke-linejoin="round" stroke-linecap="round"/>`);
      }
      if (s.marker) for (const [x, y] of pts) emit(markerSvg(s.symbol, x, y, Math.max(2, fs * 0.3), s.color ?? col, col));
    });
    // Stacked areas are drawn back to front so lower series stay visible
    for (const p of areaPolys.reverse()) emit(p);
    seriesIndex += n;
  }
  for (const l of labels) emit(l);
}

// ---------------------------------------------------------------------------- shared drawing helpers

/** A colour scaled towards black (k < 1) or white (k > 1). */
function tone(hex: string, k: number): string {
  return [0, 2, 4]
    .map((o) => {
      const v = parseInt(hex.slice(o, o + 2), 16) || 0;
      const t = k <= 1 ? v * k : v + (255 - v) * (k - 1);
      return Math.round(Math.max(0, Math.min(255, t)))
        .toString(16)
        .padStart(2, '0');
    })
    .join('')
    .toUpperCase();
}

const poly = (pts: [number, number][]): string => `M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}Z`;

/** A data marker (Office symbols) centred on x, y. */
function markerSvg(sym: string | undefined, x: number, y: number, r: number, fill: string, stroke: string): string {
  const st = ` fill="#${fill}" stroke="#${stroke}" stroke-width="0.75"`;
  switch (sym) {
    case 'square':
      return `<rect x="${f(x - r)}" y="${f(y - r)}" width="${f(2 * r)}" height="${f(2 * r)}"${st}/>`;
    case 'diamond':
      return `<path d="${poly([[x, y - r * 1.3], [x + r * 1.3, y], [x, y + r * 1.3], [x - r * 1.3, y]])}"${st}/>`;
    case 'triangle':
      return `<path d="${poly([[x, y - r * 1.3], [x + r * 1.2, y + r * 0.9], [x - r * 1.2, y + r * 0.9]])}"${st}/>`;
    case 'x':
      return `<path d="M${f(x - r)} ${f(y - r)}L${f(x + r)} ${f(y + r)}M${f(x + r)} ${f(y - r)}L${f(x - r)} ${f(y + r)}" fill="none" stroke="#${stroke}" stroke-width="1.25"/>`;
    case 'plus':
      return `<path d="M${f(x - r)} ${f(y)}L${f(x + r)} ${f(y)}M${f(x)} ${f(y - r)}L${f(x)} ${f(y + r)}" fill="none" stroke="#${stroke}" stroke-width="1.25"/>`;
    case 'star':
      return `<path d="M${f(x - r)} ${f(y)}L${f(x + r)} ${f(y)}M${f(x)} ${f(y - r)}L${f(x)} ${f(y + r)}M${f(x - r * 0.7)} ${f(y - r * 0.7)}L${f(x + r * 0.7)} ${f(y + r * 0.7)}M${f(x + r * 0.7)} ${f(y - r * 0.7)}L${f(x - r * 0.7)} ${f(y + r * 0.7)}" fill="none" stroke="#${stroke}" stroke-width="1"/>`;
    case 'dash':
      return `<path d="M${f(x - r * 1.2)} ${f(y)}L${f(x + r * 1.2)} ${f(y)}" fill="none" stroke="#${stroke}" stroke-width="2"/>`;
    case 'dot':
      return `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r * 0.5)}" fill="#${stroke}"/>`;
    default:
      return `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r)}"${st}/>`;
  }
}

/** Depth offset of 3-D charts: an oblique projection (Office's right-angle axes) from view3D. */
function depthOf(v: View3D | undefined, w: number, h: number): { dx: number; dy: number } {
  const d = Math.min(w, h) * 0.13 * Math.max(0.2, Math.min(2, (v?.depthPercent ?? 100) / 100));
  const rotY = (((v?.rotY ?? 20) % 360) + 360) % 360;
  const rotX = v?.rotX ?? 15;
  const sx = Math.cos((rotY * Math.PI) / 180) >= 0 ? 1 : -1;
  const sy = rotX >= 0 ? -1 : 1;
  return { dx: d * 0.85 * sx, dy: d * (0.3 + 0.4 * Math.min(1, Math.abs(rotX) / 30)) * sy };
}

// ---------------------------------------------------------------------------- radar

function drawRadar(c: ReturnType<typeof parseChart>, left: number, top: number, right: number, bottom: number, fs: number, auto: (i: number) => string, emit: (s: string) => void): void {
  const series = c.groups.flatMap((g) => g.series);
  const style = c.groups[0].radarStyle;
  const cats = series.reduce<string[]>((best, s) => (s.cats.length > best.length ? s.cats : best), []);
  const n = Math.max(cats.length, ...series.map((s) => s.vals.length));
  if (n < 1) return;
  const all = series.flatMap((s) => s.vals.filter((v): v is number => v !== null));
  const sc = niceScale(Math.min(0, ...all), Math.max(...all, 0), c.valAxis.min, c.valAxis.max, 5);
  const labelW = Math.max(0, ...cats.map((t) => textWidth(t, fs)));
  const R = Math.max(10, Math.min((right - left) / 2 - labelW - 8, (bottom - top) / 2 - fs * 1.8));
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2;
  const ang = (i: number): number => -Math.PI / 2 + (i * 2 * Math.PI) / n;
  const pt = (i: number, v: number): [number, number] => {
    const r = (R * (v - sc.min)) / (sc.max - sc.min || 1);
    return [cx + r * Math.cos(ang(i)), cy + r * Math.sin(ang(i))];
  };
  const ticks: number[] = [];
  for (let v = sc.min, k = 0; v <= sc.max + sc.step * 1e-6 && k < 50; v += sc.step, k++) ticks.push(Math.round(v / sc.step) * sc.step);
  // gridlines: polygons at the ticks, spokes to every category
  for (const t of ticks.slice(1)) emit(`<path d="${poly(Array.from({ length: n }, (_, i) => pt(i, t)))}" fill="none" stroke="#${GRID}" stroke-width="0.75"/>`);
  for (let i = 0; i < n; i++) {
    const [x, y] = pt(i, sc.max);
    emit(`<line x1="${f(cx)}" y1="${f(cy)}" x2="${f(x)}" y2="${f(y)}" stroke="#${AXIS}" stroke-width="0.75"/>`);
    const lx = cx + (R + 6) * Math.cos(ang(i));
    const ly = cy + (R + 6) * Math.sin(ang(i));
    const cos = Math.cos(ang(i));
    const anchor = cos > 0.25 ? 'start' : cos < -0.25 ? 'end' : 'middle';
    const dy = Math.sin(ang(i)) > 0.25 ? fs * 0.9 : Math.sin(ang(i)) < -0.25 ? -fs * 0.1 : fs * 0.35;
    if (!c.catAxis.deleted && cats[i] !== undefined) emit(`<text x="${f(lx)}" y="${f(ly + dy)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}"${anchor !== 'start' ? ` text-anchor="${anchor}"` : ''}>${escapeXml(cats[i])}</text>`);
  }
  if (!c.valAxis.deleted)
    for (const t of ticks) {
      const [x, y] = pt(0, t);
      emit(`<text x="${f(x - 4)}" y="${f(y + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs * 0.9)}" fill="#${LABEL}" text-anchor="end">${escapeXml(formatNumber(t, c.valAxis.format ?? series[0]?.format))}</text>`);
    }
  series.forEach((s, k) => {
    const col = (style === 'filled' ? s.color : (s.lineColor ?? s.color)) ?? auto(k);
    const pts = Array.from({ length: n }, (_, i) => pt(i, s.vals[i] ?? sc.min));
    if (style === 'filled') emit(`<path d="${poly(pts)}" fill="#${col}" fill-opacity="0.85" stroke="#${col}" stroke-width="0.75"/>`);
    else {
      emit(`<path d="${poly(pts)}" fill="none" stroke="#${col}" stroke-width="${f(s.lineWidth ?? 2.25)}" stroke-linejoin="round"/>`);
      if (s.marker) pts.forEach(([x, y], i) => s.vals[i] !== null && emit(markerSvg(s.symbol, x, y, Math.max(2, fs * 0.3), s.color ?? col, col)));
    }
  });
}

// ---------------------------------------------------------------------------- surface

type VPt = [number, number, number]; // x, y (data grid), value

/** Part of a convex polygon where value ≥ t (keepAbove) or ≤ t, with interpolated edges. */
function clipByValue(polyIn: VPt[], t: number, keepAbove: boolean): VPt[] {
  const inside = (p: VPt): boolean => (keepAbove ? p[2] >= t : p[2] <= t);
  const out: VPt[] = [];
  for (let i = 0; i < polyIn.length; i++) {
    const a = polyIn[i];
    const b = polyIn[(i + 1) % polyIn.length];
    if (inside(a)) out.push(a);
    if (inside(a) !== inside(b) && a[2] !== b[2]) {
      const k = (t - a[2]) / (b[2] - a[2]);
      out.push([a[0] + (b[0] - a[0]) * k, a[1] + (b[1] - a[1]) * k, t]);
    }
  }
  return out;
}

/** Value bands of a surface chart: [low, high) pairs from the value axis scale. */
function surfaceBands(c: ReturnType<typeof parseChart>): number[] {
  const all = c.groups.flatMap((g) => g.series.flatMap((s) => s.vals.filter((v): v is number => v !== null)));
  const sc = niceScale(Math.min(...all), Math.max(...all), c.valAxis.min, c.valAxis.max, 6);
  const t: number[] = [];
  for (let v = sc.min, k = 0; v <= sc.max + sc.step * 1e-6 && k < 40; v += sc.step, k++) t.push(Math.round(v / sc.step) * sc.step);
  return t;
}

function drawSurface(c: ReturnType<typeof parseChart>, left: number, top: number, right: number, bottom: number, fs: number, auto: (i: number) => string, emit: (s: string) => void): void {
  const g = c.groups[0];
  const rows = g.series;
  const cats = rows.reduce<string[]>((best, s) => (s.cats.length > best.length ? s.cats : best), []);
  const nx = Math.max(cats.length, ...rows.map((s) => s.vals.length));
  const ny = rows.length;
  if (nx < 2 || ny < 1) return;
  const bands = surfaceBands(c);
  const vmin = bands[0];
  const vmax = bands[bands.length - 1];
  const contour = !g.threeD || (c.view3D?.rotX ?? 15) >= 89;
  const labelW = Math.max(0, ...rows.map((s) => textWidth(s.name, fs))) + 6;
  const vLabelW = Math.max(...bands.map((t) => textWidth(formatNumber(t, c.valAxis.format), fs))) + 6;
  const plotL = left + (contour ? labelW : vLabelW);
  let plotR = right - 6;
  let plotT = top + fs;
  const plotB = bottom - fs * 1.8;
  const depth = contour ? { dx: 0, dy: 0 } : depthOf(c.view3D, plotR - plotL, plotB - plotT);
  if (!contour) {
    // the depth axis takes the series rows: give it a third of the room
    depth.dx *= 2.2;
    depth.dy *= 2.2;
    plotR -= Math.abs(depth.dx);
    plotT += Math.abs(depth.dy);
  }
  const W = plotR - plotL;
  const H = plotB - plotT;
  // grid → screen
  const scr = (x: number, y: number, v: number): [number, number] => {
    if (contour) return [plotL + (x * W) / (nx - 1), plotB - (ny > 1 ? (y * H) / (ny - 1) : 0)];
    const ky = ny > 1 ? y / (ny - 1) : 0;
    return [plotL + (x * W) / (nx - 1) + ky * depth.dx, plotB - ((v - vmin) / (vmax - vmin || 1)) * H + ky * depth.dy];
  };
  if (!contour) {
    // walls with value gridlines
    for (const t of bands) {
      const [x0, y0] = scr(0, 0, t);
      const [x1, y1] = scr(0, ny - 1, t);
      const [x2, y2] = scr(nx - 1, ny - 1, t);
      emit(`<path d="M${f(x0)} ${f(y0)}L${f(x1)} ${f(y1)}L${f(x2)} ${f(y2)}" fill="none" stroke="#${GRID}" stroke-width="0.75"/>`);
      if (!c.valAxis.deleted) emit(`<text x="${f(x0 - 4)}" y="${f(y0 + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="end">${escapeXml(formatNumber(t, c.valAxis.format ?? rows[0]?.format))}</text>`);
    }
  }
  const val = (x: number, y: number): number | null => rows[y]?.vals[x] ?? null;
  // painter's order: back rows first
  for (let y = (contour ? 0 : ny - 2); contour ? y < ny - 1 : y >= 0; y += contour ? 1 : -1) {
    for (let x = 0; x < nx - 1; x++) {
      const corners: [number, number][] = [
        [x, y],
        [x + 1, y],
        [x + 1, y + 1],
        [x, y + 1],
      ];
      if (corners.some(([cx, cy]) => val(cx, cy) === null)) continue;
      const tri = (a: number, b: number, d: number): VPt[] => [a, b, d].map((k) => [corners[k][0], corners[k][1], val(corners[k][0], corners[k][1])!] as VPt);
      for (const t of [tri(0, 1, 2), tri(0, 2, 3)]) {
        for (let k = 0; k + 1 < bands.length; k++) {
          let part = clipByValue(t, bands[k], true);
          if (k + 2 < bands.length) part = clipByValue(part, bands[k + 1], false);
          if (part.length < 3) continue;
          const pts = part.map(([px, py, pv]) => scr(px, py, pv));
          const col = auto(k);
          emit(g.wireframe ? `<path d="${poly(pts)}" fill="none" stroke="#${col}" stroke-width="0.75"/>` : `<path d="${poly(pts)}" fill="#${col}" stroke="#${col}" stroke-width="0.4"/>`);
        }
      }
    }
  }
  if (!contour)
    for (let y = ny - 1; y >= 0; y--)
      for (let x = 0; x < nx; x++) {
        const v = val(x, y);
        if (v === null) continue;
        const [x0, y0] = scr(x, y, v);
        const right = val(x + 1, y);
        const front = val(x, y - 1);
        if (right !== null && x + 1 < nx) {
          const [x1, y1] = scr(x + 1, y, right);
          emit(`<line x1="${f(x0)}" y1="${f(y0)}" x2="${f(x1)}" y2="${f(y1)}" stroke="#404040" stroke-width="0.35"/>`);
        }
        if (front !== null && y > 0) {
          const [x1, y1] = scr(x, y - 1, front);
          emit(`<line x1="${f(x0)}" y1="${f(y0)}" x2="${f(x1)}" y2="${f(y1)}" stroke="#404040" stroke-width="0.35"/>`);
        }
      }
  // axes labels: categories along the front, series along the side (contour: left)
  if (!c.catAxis.deleted)
    cats.forEach((t, i) => {
      const [x, y] = scr(i, 0, vmin);
      emit(`<text x="${f(x)}" y="${f((contour ? plotB : y) + fs * 1.3)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="middle">${escapeXml(t)}</text>`);
    });
  rows.forEach((s, j) => {
    if (!s.name) return;
    const [x, y] = scr(contour ? 0 : nx - 1, j, vmin);
    emit(contour ? `<text x="${f(x - 4)}" y="${f(y + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}" text-anchor="end">${escapeXml(s.name)}</text>` : `<text x="${f(x + 4)}" y="${f(y + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#${LABEL}">${escapeXml(s.name)}</text>`);
  });
}

// ---------------------------------------------------------------------------- 3-D pie and pie of pie

function pieSlices(vals: number[], start: number): { a1: number; a2: number; v: number; i: number }[] {
  const total = vals.reduce((a, b) => a + b, 0) || 1;
  let a = start;
  return vals.map((v, i) => {
    const sweep = (v / total) * Math.PI * 2;
    const r = { a1: a, a2: a + sweep, v, i };
    a += sweep;
    return r;
  });
}

function drawPie3D(g: Group, left: number, top: number, right: number, bottom: number, fs: number, auto: (i: number) => string, emit: (s: string) => void, view: View3D | undefined): void {
  const s = g.series[0];
  const vals = s.vals.map((v) => Math.max(0, v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (!total) return;
  const tilt = Math.max(0.2, Math.sin((Math.max(10, Math.min(80, view?.rotX ?? 30)) * Math.PI) / 180));
  const showLabels = g.showPercent || g.showVal || g.showCat;
  const pad = showLabels ? fs * 0.5 : 4;
  let r = Math.min((right - left) / 2 - pad, ((bottom - top) / 2 - pad) / tilt);
  const thick = r * 0.2 * (1 - tilt * 0.5);
  r = Math.max(10, Math.min(r, ((bottom - top) / 2 - pad - thick / 2) / tilt));
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2 - thick / 2;
  const ry = r * tilt;
  const p = (ang: number, dy = 0): string => `${f(cx + r * Math.cos(ang))} ${f(cy + dy + ry * Math.sin(ang))}`;
  const colorOfPoint = (i: number): string => s.pointColors.get(i) ?? (g.vary ? auto(i) : (s.color ?? auto(0)));
  const slices = pieSlices(vals, ((g.firstAngle - 90) * Math.PI) / 180).filter((x) => x.v > 0);
  // outer walls of the front half (sin > 0), from the back towards the front
  const walls: { mid: number; d: string; col: string }[] = [];
  for (const sl of slices) {
    for (let k = -1; k <= 2; k++) {
      const lo = Math.max(sl.a1, k * 2 * Math.PI);
      const hi = Math.min(sl.a2, k * 2 * Math.PI + Math.PI);
      if (hi - lo < 1e-4) continue;
      const large = hi - lo > Math.PI ? 1 : 0;
      walls.push({ mid: (lo + hi) / 2, col: tone(colorOfPoint(sl.i), 0.72), d: `M${p(lo)}A${f(r)} ${f(ry)} 0 ${large} 1 ${p(hi)}L${p(hi, thick)}A${f(r)} ${f(ry)} 0 ${large} 0 ${p(lo, thick)}Z` });
    }
  }
  walls.sort((a, b) => Math.abs(Math.sin(a.mid)) - Math.abs(Math.sin(b.mid)) || Math.cos(a.mid) - Math.cos(b.mid));
  for (const w of walls) emit(`<path d="${w.d}" fill="#${w.col}" stroke="#${tone(w.col, 0.9)}" stroke-width="0.5"/>`);
  for (const sl of slices) {
    const large = sl.a2 - sl.a1 > Math.PI ? 1 : 0;
    const d = sl.a2 - sl.a1 >= Math.PI * 2 - 1e-6 ? `M${p(0)}A${f(r)} ${f(ry)} 0 1 1 ${p(Math.PI)}A${f(r)} ${f(ry)} 0 1 1 ${p(0)}Z` : `M${f(cx)} ${f(cy)}L${p(sl.a1)}A${f(r)} ${f(ry)} 0 ${large} 1 ${p(sl.a2)}Z`;
    emit(`<path d="${d}" fill="#${colorOfPoint(sl.i)}" stroke="#FFFFFF" stroke-width="0.75"/>`);
  }
  if (showLabels)
    for (const sl of slices) {
      const mid = (sl.a1 + sl.a2) / 2;
      const parts: string[] = [];
      if (g.showCat && s.cats[sl.i]) parts.push(s.cats[sl.i]);
      if (g.showVal) parts.push(formatNumber(sl.v, s.format));
      if (g.showPercent) parts.push(`${Math.round((sl.v / total) * 100)}%`);
      emit(`<text x="${f(cx + r * 0.65 * Math.cos(mid))}" y="${f(cy + ry * 0.65 * Math.sin(mid) + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" font-weight="bold" fill="#FFFFFF" text-anchor="middle">${escapeXml(parts.join(' '))}</text>`);
    }
}

function drawOfPie(g: Group, left: number, top: number, right: number, bottom: number, fs: number, auto: (i: number) => string, emit: (s: string) => void): void {
  const s = g.series[0];
  const op = g.ofPie!;
  const vals = s.vals.map((v) => Math.max(0, v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (!total) return;
  const k = Math.min(op.splitPos, vals.length - 1);
  const main = vals.slice(0, vals.length - k);
  const second = vals.slice(vals.length - k);
  const other = second.reduce((a, b) => a + b, 0);
  const colorOfPoint = (i: number): string => s.pointColors.get(i) ?? (g.vary ? auto(i) : (s.color ?? auto(0)));
  const showLabels = g.showPercent || g.showVal || g.showCat;
  const pct = (v: number): string => {
    const parts: string[] = [];
    if (g.showVal) parts.push(formatNumber(v, s.format));
    if (g.showPercent || !parts.length) parts.push(`${Math.round((v / total) * 100)}%`);
    return parts.join(' ');
  };
  const W = right - left;
  const H = bottom - top;
  const r = Math.max(10, Math.min(W * 0.27, H / 2 - 6));
  const cx = left + W * 0.3;
  const cy = top + H / 2;
  const r2 = (r * Math.max(5, Math.min(200, op.secondSize))) / 100;
  const cx2 = left + W * 0.78;
  // main pie: the "other" slice faces the second plot
  const otherSweep = (other / total) * Math.PI * 2;
  const slices = pieSlices([...main, other], otherSweep / 2);
  const p = (x: number, y: number, rad: number, ang: number): [number, number] => [x + rad * Math.cos(ang), y + rad * Math.sin(ang)];
  slices.forEach((sl, n) => {
    if (sl.v <= 0) return;
    const isOther = n === slices.length - 1;
    const col = isOther ? auto(vals.length) : colorOfPoint(sl.i);
    const [x1, y1] = p(cx, cy, r, sl.a1);
    const [x2, y2] = p(cx, cy, r, sl.a2);
    const large = sl.a2 - sl.a1 > Math.PI ? 1 : 0;
    emit(`<path d="M${f(cx)} ${f(cy)}L${f(x1)} ${f(y1)}A${f(r)} ${f(r)} 0 ${large} 1 ${f(x2)} ${f(y2)}Z" fill="#${col}" stroke="#FFFFFF" stroke-width="1"/>`);
    if (showLabels) {
      const [tx, ty] = p(cx, cy, r * 0.65, (sl.a1 + sl.a2) / 2);
      emit(`<text x="${f(tx)}" y="${f(ty + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" font-weight="bold" fill="#FFFFFF" text-anchor="middle">${escapeXml(pct(sl.v))}</text>`);
    }
  });
  const [ox1, oy1] = p(cx, cy, r, -otherSweep / 2);
  const [ox2, oy2] = p(cx, cy, r, otherSweep / 2);
  if (op.bar) {
    // the split points stacked in a column, the first at the top
    const bh = Math.min(H - 8, r2 * 2);
    const bw = Math.max(14, r * 0.45);
    const bx = cx2 - bw / 2;
    let by = cy - bh / 2;
    emit(`<line x1="${f(ox1)}" y1="${f(oy1)}" x2="${f(bx)}" y2="${f(by)}" stroke="#${AXIS}" stroke-width="0.75"/><line x1="${f(ox2)}" y1="${f(oy2)}" x2="${f(bx)}" y2="${f(by + bh)}" stroke="#${AXIS}" stroke-width="0.75"/>`);
    second.forEach((v, j) => {
      const h = other ? (v / other) * bh : 0;
      emit(`<rect x="${f(bx)}" y="${f(by)}" width="${f(bw)}" height="${f(h)}" fill="#${colorOfPoint(vals.length - k + j)}" stroke="#FFFFFF" stroke-width="1"/>`);
      if (showLabels && h > fs) emit(`<text x="${f(cx2)}" y="${f(by + h / 2 + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" font-weight="bold" fill="#FFFFFF" text-anchor="middle">${escapeXml(pct(v))}</text>`);
      by += h;
    });
  } else {
    const sub = pieSlices(second, Math.PI - (second[0] / (other || 1)) * Math.PI);
    emit(`<line x1="${f(ox1)}" y1="${f(oy1)}" x2="${f(cx2)}" y2="${f(cy - r2)}" stroke="#${AXIS}" stroke-width="0.75"/><line x1="${f(ox2)}" y1="${f(oy2)}" x2="${f(cx2)}" y2="${f(cy + r2)}" stroke="#${AXIS}" stroke-width="0.75"/>`);
    sub.forEach((sl, j) => {
      if (sl.v <= 0) return;
      const [x1, y1] = p(cx2, cy, r2, sl.a1);
      const [x2, y2] = p(cx2, cy, r2, sl.a2);
      const large = sl.a2 - sl.a1 > Math.PI ? 1 : 0;
      const d = sl.a2 - sl.a1 >= Math.PI * 2 - 1e-6 ? `M${f(cx2 + r2)} ${f(cy)}A${f(r2)} ${f(r2)} 0 1 1 ${f(cx2 - r2)} ${f(cy)}A${f(r2)} ${f(r2)} 0 1 1 ${f(cx2 + r2)} ${f(cy)}Z` : `M${f(cx2)} ${f(cy)}L${f(x1)} ${f(y1)}A${f(r2)} ${f(r2)} 0 ${large} 1 ${f(x2)} ${f(y2)}Z`;
      emit(`<path d="${d}" fill="#${colorOfPoint(vals.length - k + j)}" stroke="#FFFFFF" stroke-width="1"/>`);
      if (showLabels) {
        const [tx, ty] = p(cx2, cy, r2 * 0.62, (sl.a1 + sl.a2) / 2);
        emit(`<text x="${f(tx)}" y="${f(ty + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" font-weight="bold" fill="#FFFFFF" text-anchor="middle">${escapeXml(pct(sl.v))}</text>`);
      }
    });
  }
}
