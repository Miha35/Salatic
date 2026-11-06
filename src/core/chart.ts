/**
 * Office charts (DrawingML chartSpace, as in word/charts/*.xml and ppt/charts/*.xml)
 * → SVG drawn from the cached values stored in the file. Covers column/bar
 * (clustered, stacked, 100 %), line, area, pie, doughnut and scatter charts,
 * combinations with a secondary value axis, titles, legends, gridlines, axis
 * number formats, per-point colours and value/percentage data labels.
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
  type: 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
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

function parseChart(xml: XNode, theme: Theme): { title?: string; titleSize: number; groups: Group[]; catAxis: Axis; valAxis: Axis; val2Axis?: Axis; xAxis?: Axis; legend?: string; bg?: string; plotBg?: string } {
  const chart = deep(xml, 'chart');
  const plot = kid(chart, 'plotArea');
  const map = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
  const color = (sp: XNode | undefined, kind: 'fill' | 'line'): string | undefined => {
    if (!sp) return undefined;
    const holder = kind === 'fill' ? sp : kid(sp, 'ln');
    if (!holder || kid(holder, 'noFill')) return undefined;
    const solid = kid(holder, 'solidFill') ?? kid(holder, 'gradFill');
    if (!solid) return undefined;
    const target = local(solid.name) === 'gradFill' ? deep(solid, 'gs') : solid;
    return colorOf(target, theme, map)?.hex;
  };
  const groups: Group[] = [];
  const axIds = new Map<string, XNode>();
  for (const ax of elements(plot)) if (/Ax$/.test(local(ax.name))) axIds.set(val(kid(ax, 'axId')) ?? '', ax);
  for (const g of elements(plot)) {
    const name = local(g.name);
    const type: Group['type'] | undefined = /^bar/.test(name) ? 'bar' : /^line/.test(name) || name === 'radarChart' || name === 'stockChart' ? 'line' : /^area/.test(name) ? 'area' : /^pie/.test(name) || name === 'ofPieChart' ? 'pie' : name === 'doughnutChart' ? 'doughnut' : name === 'scatterChart' || name === 'bubbleChart' ? 'scatter' : undefined;
    if (!type) continue;
    const dl = kid(g, 'dLbls');
    const grp: Group = {
      type,
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
          return type === 'line' ? val(kid(g, 'marker')) !== '0' : type === 'scatter';
        })(),
        smooth: on(kid(s, 'smooth')),
      };
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
  const scatterX = groups.some((g) => g.type === 'scatter') ? elements(plot).filter((a) => local(a.name) === 'valAx').find((a) => ['b', 't'].includes(val(kid(a, 'axPos')) ?? '')) : undefined;
  let title: string | undefined;
  const t = kid(chart, 'title');
  if (t) title = richText(kid(t, 'tx')) || (groups.length === 1 && groups[0].series.length === 1 ? groups[0].series[0].name : undefined);
  else if (!on(kid(chart, 'autoTitleDeleted')) && groups.length === 1 && groups[0].series.length === 1 && groups[0].type !== 'scatter') title = groups[0].series[0].name || undefined;
  const sz = deep(t, 'defRPr')?.attrs.sz ?? deep(t, 'rPr')?.attrs.sz;
  const legend = kid(chart, 'legend');
  return {
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
  const pieLike = c.groups.every((g) => g.type === 'pie' || g.type === 'doughnut');
  // Legend entries
  const entries: { name: string; color: string; line: boolean }[] = [];
  if (c.legend) {
    if (pieLike) {
      const s = c.groups[0].series[0];
      s.vals.forEach((_, i) => entries.push({ name: s.cats[i] ?? String(i + 1), color: s.pointColors.get(i) ?? auto(i), line: false }));
    } else {
      let k = 0;
      for (const g of c.groups)
        for (const s of g.series) {
          const lineLike = g.type === 'line' || g.type === 'scatter';
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
  if (pieLike) {
    drawPies(c.groups, left, top, right, bottom, fs, auto, (s) => (out += s));
  } else {
    drawAxes(c, left, top, right, bottom, fs, auto, (s) => (out += s));
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${f(W)} ${f(H)}" width="${f(W)}pt" height="${f(H)}pt">${out}</svg>`;
}

function drawPies(groups: Group[], left: number, top: number, right: number, bottom: number, fs: number, auto: (i: number) => string, emit: (s: string) => void): void {
  const g = groups[0];
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
  const groups = c.groups.filter((g) => g.type !== 'pie' && g.type !== 'doughnut');
  const scatter = groups.some((g) => g.type === 'scatter');
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
    xs = niceScale(autoMin(Math.min(...all), Math.max(...all)), Math.max(...all), c.xAxis?.min, c.xAxis?.max, intervals(xLen, true, Math.max(...all.map(Math.abs))));
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
  const plotW = plotR - plotL;
  const plotH = plotB - plotT;
  if (plotW < 10 || plotH < 10) return;
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
    if (c.valAxis.grid) emit(horizontal ? `<line x1="${f(p)}" y1="${f(plotT)}" x2="${f(p)}" y2="${f(plotB)}" stroke="#${GRID}" stroke-width="0.75"/>` : `<line x1="${f(plotL)}" y1="${f(p)}" x2="${f(plotR)}" y2="${f(p)}" stroke="#${GRID}" stroke-width="0.75"/>`);
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
  emit(horizontal ? `<line x1="${f(zero)}" y1="${f(plotT)}" x2="${f(zero)}" y2="${f(plotB)}" stroke="#${AXIS}" stroke-width="0.75"/>` : `<line x1="${f(plotL)}" y1="${f(zero)}" x2="${f(plotR)}" y2="${f(zero)}" stroke="#${AXIS}" stroke-width="0.75"/>`);
  // Series
  let seriesIndex = 0;
  const labels: string[] = [];
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
          if (horizontal) emit(`<rect x="${f(Math.min(p0, p1))}" y="${f(off)}" width="${f(Math.abs(p1 - p0))}" height="${f(barW)}" fill="#${fill}"/>`);
          else emit(`<rect x="${f(off)}" y="${f(Math.min(p0, p1))}" width="${f(barW)}" height="${f(Math.abs(p1 - p0))}" fill="#${fill}"/>`);
          if (g.showVal) {
            const t = escapeXml(g.grouping === 'percentStacked' ? `${Math.round(v * 100)}%` : formatNumber(s.vals[i]!, s.format));
            if (horizontal) labels.push(`<text x="${f(Math.max(p0, p1) + 3)}" y="${f(off + barW / 2 + fs * 0.35)}" font-family="${FONT}" font-size="${f(fs)}" fill="#404040">${t}</text>`);
            else labels.push(`<text x="${f(off + barW / 2)}" y="${f(Math.min(p0, p1) - 3)}" font-family="${FONT}" font-size="${f(fs)}" fill="#404040" text-anchor="middle">${t}</text>`);
          }
        }
      });
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
      if (s.marker) for (const [x, y] of pts) emit(`<circle cx="${f(x)}" cy="${f(y)}" r="${f(Math.max(2, fs * 0.3))}" fill="#${s.color ?? col}" stroke="#${col}" stroke-width="0.75"/>`);
    });
    // Stacked areas are drawn back to front so lower series stay visible
    for (const p of areaPolys.reverse()) emit(p);
    seriesIndex += n;
  }
  for (const l of labels) emit(l);
}
