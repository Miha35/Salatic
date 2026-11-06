/**
 * SVG → vector display list (the same one EMF/WMF produce), so SVG pictures and
 * generated charts are drawn as real vectors with real text in PDFs. Supports the
 * static subset documents use: shapes, paths (incl. arcs), transforms, groups,
 * <use>, presentation attributes, inline styles and simple class/type CSS rules,
 * text with tspans, embedded PNG/JPEG images, linear and radial gradient fills (drawn as PDF
 * shadings; writers without gradients use the average colour). Filters, masks and clip paths are ignored.
 */
import { familyOf, gradientAverage, type FontFamily, type Gradient, type GradientStop } from './model';
import type { ClipPath, Metafile, MetaCmd } from './metafile';
import { parseDataUri } from './util';
import { elements, parseXml, textContent, type XNode } from './xml';

type M = [number, number, number, number, number, number];
const IDENT: M = [1, 0, 0, 1, 0, 0];
const mul = (a: M, b: M): M => [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]];
const apply = (m: M, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const scaleOf = (m: M): number => Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;

const local = (n: string): string => n.replace(/^[a-z]+:/i, '');

const NAMED: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', orange: 'FFA500', gray: '808080', grey: '808080',
  silver: 'C0C0C0', navy: '000080', teal: '008080', maroon: '800000', purple: '800080', olive: '808000', lime: '00FF00', aqua: '00FFFF', cyan: '00FFFF',
  fuchsia: 'FF00FF', magenta: 'FF00FF', darkgray: 'A9A9A9', darkgrey: 'A9A9A9', lightgray: 'D3D3D3', lightgrey: 'D3D3D3', darkblue: '00008B', darkred: '8B0000',
  darkgreen: '006400', steelblue: '4682B4', skyblue: '87CEEB', gold: 'FFD700', pink: 'FFC0CB', brown: 'A52A2A', crimson: 'DC143C', indigo: '4B0082',
};

/** CSS colour → hex (null for "none"/"transparent", undefined when unknown). */
export function cssColor(v: string | undefined): string | null | undefined {
  if (v === undefined) return undefined;
  const s = v.trim().toLowerCase();
  if (!s || s === 'none' || s === 'transparent') return null;
  if (s.startsWith('#')) {
    const h = s.slice(1);
    if (/^[0-9a-f]{3,4}$/.test(h)) return (h[0] + h[0] + h[1] + h[1] + h[2] + h[2]).toUpperCase();
    if (/^[0-9a-f]{6}([0-9a-f]{2})?$/.test(h)) return h.slice(0, 6).toUpperCase();
    return undefined;
  }
  const rgb = /^rgba?\(\s*([\d.]+%?)[\s,]+([\d.]+%?)[\s,]+([\d.]+%?)/.exec(s);
  if (rgb) return rgb.slice(1, 4).map((c) => Math.round(Math.min(255, c.endsWith('%') ? (parseFloat(c) * 255) / 100 : parseFloat(c))).toString(16).padStart(2, '0')).join('').toUpperCase();
  return NAMED[s];
}

function parseTransform(t: string | undefined): M {
  let m: M = [...IDENT];
  if (!t) return m;
  for (const [, fn, args] of t.matchAll(/(\w+)\s*\(([^)]*)\)/g)) {
    const a = args.split(/[\s,]+/).filter(Boolean).map(Number);
    let n: M = [...IDENT];
    switch (fn) {
      case 'matrix':
        if (a.length === 6) n = a as M;
        break;
      case 'translate':
        n = [1, 0, 0, 1, a[0] ?? 0, a[1] ?? 0];
        break;
      case 'scale':
        n = [a[0] ?? 1, 0, 0, a[1] ?? a[0] ?? 1, 0, 0];
        break;
      case 'rotate': {
        const r = ((a[0] ?? 0) * Math.PI) / 180;
        const c = Math.cos(r);
        const s = Math.sin(r);
        n = [c, s, -s, c, 0, 0];
        if (a.length === 3) n = mul(mul([1, 0, 0, 1, a[1], a[2]], n), [1, 0, 0, 1, -a[1], -a[2]]);
        break;
      }
      case 'skewX':
        n = [1, 0, Math.tan(((a[0] ?? 0) * Math.PI) / 180), 1, 0, 0];
        break;
      case 'skewY':
        n = [1, Math.tan(((a[0] ?? 0) * Math.PI) / 180), 0, 1, 0, 0];
        break;
    }
    m = mul(m, n);
  }
  return m;
}

/** Path data → absolute segments (M, L, C, Z), arcs and quadratics converted to cubics. */
function parsePath(d: string): (['M' | 'L', number, number] | ['C', number, number, number, number, number, number] | ['Z'])[] {
  const out: (['M' | 'L', number, number] | ['C', number, number, number, number, number, number] | ['Z'])[] = [];
  const tokens = d.match(/[a-zA-Z]|[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g) ?? [];
  let i = 0;
  let cmd = '';
  let x = 0;
  let y = 0;
  let sx = 0;
  let sy = 0;
  let lastC: [number, number] | null = null;
  let lastQ: [number, number] | null = null;
  const num = (): number => Number(tokens[i++]);
  const has = (): boolean => i < tokens.length && !/^[a-zA-Z]$/.test(tokens[i]);
  let guard = 0;
  while (i < tokens.length && guard++ < 1_000_000) {
    if (/^[a-zA-Z]$/.test(tokens[i])) cmd = tokens[i++];
    else if (!cmd) {
      i++;
      continue;
    }
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    const ox = rel ? x : 0;
    const oy = rel ? y : 0;
    switch (C) {
      case 'M':
        if (!has()) break;
        x = ox + num();
        y = oy + num();
        sx = x;
        sy = y;
        out.push(['M', x, y]);
        cmd = rel ? 'l' : 'L';
        lastC = lastQ = null;
        continue;
      case 'L':
        if (!has()) break;
        x = ox + num();
        y = oy + num();
        out.push(['L', x, y]);
        lastC = lastQ = null;
        continue;
      case 'H':
        if (!has()) break;
        x = (rel ? x : 0) + num();
        out.push(['L', x, y]);
        lastC = lastQ = null;
        continue;
      case 'V':
        if (!has()) break;
        y = (rel ? y : 0) + num();
        out.push(['L', x, y]);
        lastC = lastQ = null;
        continue;
      case 'C': {
        if (!has()) break;
        const x1 = ox + num();
        const y1 = oy + num();
        const x2 = ox + num();
        const y2 = oy + num();
        x = ox + num();
        y = oy + num();
        out.push(['C', x1, y1, x2, y2, x, y]);
        lastC = [x2, y2];
        lastQ = null;
        continue;
      }
      case 'S': {
        if (!has()) break;
        const x1 = lastC ? 2 * x - lastC[0] : x;
        const y1 = lastC ? 2 * y - lastC[1] : y;
        const x2 = ox + num();
        const y2 = oy + num();
        x = ox + num();
        y = oy + num();
        out.push(['C', x1, y1, x2, y2, x, y]);
        lastC = [x2, y2];
        lastQ = null;
        continue;
      }
      case 'Q':
      case 'T': {
        if (!has()) break;
        let qx: number;
        let qy: number;
        if (C === 'Q') {
          qx = ox + num();
          qy = oy + num();
        } else {
          qx = lastQ ? 2 * x - lastQ[0] : x;
          qy = lastQ ? 2 * y - lastQ[1] : y;
        }
        const ex = ox + num();
        const ey = oy + num();
        out.push(['C', x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), ex + (2 / 3) * (qx - ex), ey + (2 / 3) * (qy - ey), ex, ey]);
        x = ex;
        y = ey;
        lastQ = [qx, qy];
        lastC = null;
        continue;
      }
      case 'A': {
        if (!has()) break;
        const rx = Math.abs(num());
        const ry = Math.abs(num());
        const rot = num();
        const large = num() !== 0;
        const sweep = num() !== 0;
        const ex = ox + num();
        const ey = oy + num();
        for (const c of arcToCubics(x, y, rx, ry, rot, large, sweep, ex, ey)) out.push(c);
        x = ex;
        y = ey;
        lastC = lastQ = null;
        continue;
      }
      case 'Z':
        out.push(['Z']);
        x = sx;
        y = sy;
        lastC = lastQ = null;
        cmd = '';
        continue;
    }
    // Command without (enough) parameters: skip it
    cmd = '';
  }
  return out;
}

/** SVG elliptical arc → cubic Béziers (SVG 1.1 implementation notes, F.6). */
function arcToCubics(x1: number, y1: number, rx: number, ry: number, phiDeg: number, large: boolean, sweep: boolean, x2: number, y2: number): ['C', number, number, number, number, number, number][] {
  if (!rx || !ry || (x1 === x2 && y1 === y2)) return [['C', x1, y1, x2, y2, x2, y2]];
  const phi = (phiDeg * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy;
  const y1p = -sin * dx + cos * dy;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let co = Math.sqrt(Math.max(0, num / den));
  if (large === sweep) co = -co;
  const cxp = (co * rx * y1p) / ry;
  const cyp = (-co * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const ang = (ux: number, uy: number, vx: number, vy: number): number => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sweep && dt > 0) dt -= 2 * Math.PI;
  else if (sweep && dt < 0) dt += 2 * Math.PI;
  const segs = Math.ceil(Math.abs(dt) / (Math.PI / 2));
  const out: ['C', number, number, number, number, number, number][] = [];
  const delta = dt / segs;
  const k = (4 / 3) * Math.tan(delta / 4);
  let t = t1;
  const pt = (a: number): [number, number] => [cx + rx * Math.cos(a) * cos - ry * Math.sin(a) * sin, cy + rx * Math.cos(a) * sin + ry * Math.sin(a) * cos];
  const der = (a: number): [number, number] => [-rx * Math.sin(a) * cos - ry * Math.cos(a) * sin, -rx * Math.sin(a) * sin + ry * Math.cos(a) * cos];
  for (let s = 0; s < segs; s++) {
    const a0 = t;
    const a1 = t + delta;
    const p0 = pt(a0);
    const p1 = pt(a1);
    const d0 = der(a0);
    const d1 = der(a1);
    out.push(['C', p0[0] + k * d0[0], p0[1] + k * d0[1], p1[0] - k * d1[0], p1[1] - k * d1[1], p1[0], p1[1]]);
    t = a1;
  }
  return out;
}

interface Style {
  fill: string | null;
  fillOpacity: number;
  stroke: string | null;
  strokeOpacity: number;
  strokeWidth: number;
  dash?: number[];
  opacity: number;
  evenOdd: boolean;
  fontSize: number;
  fontFamily: string;
  fontWeight: string;
  fontStyle: string;
  anchor: 'start' | 'middle' | 'end';
  display: boolean;
  /** id of the gradient painting the fill (url(#id)) */
  fillGrad?: string;
}

const len = (v: string | undefined, ref = 0): number | undefined => {
  if (v === undefined) return undefined;
  const m = /^\s*([-+]?(?:\d*\.\d+|\d+\.?)(?:e[-+]?\d+)?)\s*(px|pt|pc|mm|cm|in|em|%)?/i.exec(v);
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  switch ((m[2] ?? '').toLowerCase()) {
    case 'pt':
      return (n * 4) / 3;
    case 'pc':
      return n * 16;
    case 'mm':
      return (n * 96) / 25.4;
    case 'cm':
      return (n * 96) / 2.54;
    case 'in':
      return n * 96;
    case 'em':
      return n * 16;
    case '%':
      return (n / 100) * ref;
    default:
      return n;
  }
};

class SvgRenderer {
  cmds: MetaCmd[] = [];
  private ids = new Map<string, XNode>();
  private rules: { sel: string; decl: Record<string, string> }[] = [];
  private depth = 0;
  /** viewBox size: the reference for percentages in userSpaceOnUse gradients */
  rootW = 300;
  rootH = 150;

  constructor(private root: XNode) {
    const index = (n: XNode): void => {
      if (n.attrs.id) this.ids.set(n.attrs.id, n);
      for (const c of elements(n)) index(c);
    };
    index(root);
    const collect = (n: XNode): void => {
      if (local(n.name) === 'style') this.parseCss(textContent(n));
      for (const c of elements(n)) collect(c);
    };
    collect(root);
  }

  private parseCss(css: string): void {
    for (const [, sel, body] of css.replace(/\/\*[\s\S]*?\*\//g, '').matchAll(/([^{}]+)\{([^}]*)\}/g)) {
      const decl = declarations(body);
      for (const s of sel.split(',')) this.rules.push({ sel: s.trim(), decl });
    }
  }

  /** Effective properties: inherited style ← CSS rules ← presentation attributes ← style="". */
  private styleOf(n: XNode, parent: Style, viewW: number): Style {
    const props: Record<string, string> = {};
    for (const [k, v] of Object.entries(n.attrs)) props[k] = v;
    const classes = (n.attrs.class ?? '').split(/\s+/).filter(Boolean);
    const tag = local(n.name);
    for (const r of this.rules) {
      const sel = r.sel;
      if (sel === tag || (sel.startsWith('.') && classes.includes(sel.slice(1))) || (sel.startsWith('#') && n.attrs.id === sel.slice(1)) || sel === '*') Object.assign(props, r.decl);
    }
    Object.assign(props, declarations(n.attrs.style ?? ''));
    // font shorthand: [style] [weight] size[/line-height] family
    if (props.font) {
      const m = /^\s*(?:(italic|oblique|normal)\s+)?(?:(bold|bolder|lighter|normal|[1-9]00)\s+)?([\d.]+(?:px|pt|em|%)?)(?:\/[\d.]+\w*)?\s+(.+)$/i.exec(props.font);
      if (m) {
        if (m[1]) props['font-style'] ??= m[1];
        if (m[2]) props['font-weight'] ??= m[2];
        props['font-size'] ??= m[3];
        props['font-family'] ??= m[4];
      }
    }
    const s: Style = { ...parent, opacity: 1 };
    const paint = (v: string | undefined, current: string | null): string | null | undefined => {
      if (v === undefined) return undefined;
      if (v === 'currentColor') return cssColor(props.color) ?? current;
      const url = /url\(\s*#([^)\s]+)\s*\)/.exec(v);
      if (url) return this.gradientColor(url[1]) ?? null;
      return cssColor(v);
    };
    const f = paint(props.fill, parent.fill);
    if (f !== undefined) {
      s.fill = f;
      s.fillGrad = /url\(\s*#([^)\s]+)\s*\)/.exec(props.fill ?? '')?.[1];
    }
    const st = paint(props.stroke, parent.stroke);
    if (st !== undefined) s.stroke = st;
    if (props['fill-opacity'] !== undefined) s.fillOpacity = Number(props['fill-opacity']);
    if (props['stroke-opacity'] !== undefined) s.strokeOpacity = Number(props['stroke-opacity']);
    if (props.opacity !== undefined) s.opacity = Number(props.opacity);
    const sw = len(props['stroke-width'], viewW);
    if (sw !== undefined) s.strokeWidth = sw;
    if (props['stroke-dasharray'] !== undefined) s.dash = props['stroke-dasharray'] === 'none' ? undefined : props['stroke-dasharray'].split(/[\s,]+/).map(Number).filter((x) => x >= 0);
    if (props['fill-rule'] !== undefined) s.evenOdd = props['fill-rule'] === 'evenodd';
    const fs = len(props['font-size'], parent.fontSize);
    if (fs !== undefined) s.fontSize = fs;
    if (props['font-family']) s.fontFamily = props['font-family'];
    if (props['font-weight']) s.fontWeight = props['font-weight'];
    if (props['font-style']) s.fontStyle = props['font-style'];
    if (props['text-anchor']) s.anchor = props['text-anchor'] === 'middle' ? 'middle' : props['text-anchor'] === 'end' ? 'end' : 'start';
    if (props.display === 'none' || props.visibility === 'hidden') s.display = false;
    return s;
  }

  private gradientColor(id: string): string | undefined {
    let g = this.ids.get(id);
    let guard = 0;
    while (g && !elements(g).some((c) => local(c.name) === 'stop') && guard++ < 5) {
      const href = (g.attrs.href ?? g.attrs['xlink:href'] ?? '').replace(/^#/, '');
      g = href ? this.ids.get(href) : undefined;
    }
    const stops = g ? elements(g).filter((c) => local(c.name) === 'stop') : [];
    const cols = stops.map((s) => cssColor(declarations(s.attrs.style ?? '')['stop-color'] ?? s.attrs['stop-color'] ?? '#000')).filter((c): c is string => !!c);
    if (!cols.length) return undefined;
    const avg = [0, 2, 4].map((o) => Math.round(cols.reduce((sum, c) => sum + parseInt(c.slice(o, o + 2), 16), 0) / cols.length));
    return avg.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  private paint(d: string, s: Style, m: M): void {
    if (this.capture) {
      if (d) this.capture.push(d);
      return;
    }
    if (!d || !s.display) return;
    let fill = s.fill ?? undefined;
    const stroke = s.stroke && s.strokeWidth > 0 ? s.stroke : undefined;
    if (!fill && !stroke) return;
    const k = scaleOf(m);
    const grad = fill && s.fillGrad ? this.gradient(s.fillGrad, this.lastBox, m) : undefined;
    if (grad) fill = gradientAverage(grad);
    this.cmds.push({ k: 'path', d, fill, evenOdd: s.evenOdd, stroke, width: s.strokeWidth * k, dash: s.dash?.map((x) => x * k), opacity: s.opacity * s.fillOpacity, strokeOpacity: s.opacity * s.strokeOpacity, ...(grad ? { grad } : {}) });
  }

  /** Bounding box (user space) of the last path passed through pathD — the objectBoundingBox of gradients. */
  private lastBox: [number, number, number, number] = [0, 0, 0, 0];

  /** linearGradient / radialGradient → gradient in output space (follows href templates). */
  private gradient(id: string, box: [number, number, number, number], m: M): Gradient | undefined {
    const chain: XNode[] = [];
    for (let g = this.ids.get(id); g && chain.length < 8 && !chain.includes(g); ) {
      chain.push(g);
      const href = (g.attrs.href ?? g.attrs['xlink:href'] ?? '').replace(/^#/, '');
      g = href ? this.ids.get(href) : undefined;
    }
    const tag = chain[0] ? local(chain[0].name) : '';
    if (tag !== 'linearGradient' && tag !== 'radialGradient') return undefined;
    const attr = (k: string): string | undefined => chain.find((n) => n.attrs[k] !== undefined)?.attrs[k];
    const holder = chain.find((n) => elements(n).some((c) => local(c.name) === 'stop'));
    const stops: GradientStop[] = [];
    for (const st of holder ? elements(holder).filter((c) => local(c.name) === 'stop') : []) {
      const css = { ...st.attrs, ...declarations(st.attrs.style ?? '') };
      const off = css.offset ?? '0';
      const pos = Math.min(1, Math.max(stops.length ? stops[stops.length - 1].pos : 0, off.trim().endsWith('%') ? parseFloat(off) / 100 : parseFloat(off) || 0));
      const color = cssColor(css['stop-color'] ?? '#000') ?? '000000';
      const a = css['stop-opacity'] !== undefined ? Math.min(1, Math.max(0, Number(css['stop-opacity']))) : 1;
      stops.push({ pos, color, ...(a < 1 ? { alpha: a } : {}) });
    }
    if (!stops.length) return undefined;
    const bbox = attr('gradientUnits') !== 'userSpaceOnUse';
    const num = (k: string, def: number, ref: number): number => {
      const v = attr(k);
      if (v === undefined) return def;
      if (bbox) return v.trim().endsWith('%') ? parseFloat(v) / 100 : parseFloat(v) || 0;
      return len(v, ref) ?? def;
    };
    let g: M = parseTransform(attr('gradientTransform'));
    if (bbox) {
      const [x0, y0, x1, y1] = box;
      if (!(x1 - x0 > 0) || !(y1 - y0 > 0)) return undefined;
      g = mul([x1 - x0, 0, 0, y1 - y0, x0, y0], g);
    }
    const W = bbox ? 1 : this.rootW;
    const H = bbox ? 1 : this.rootH;
    const D = bbox ? 1 : Math.hypot(this.rootW, this.rootH) / Math.SQRT2;
    let coords: number[];
    if (tag === 'linearGradient') coords = [num('x1', 0, W), num('y1', 0, H), num('x2', W, W), num('y2', 0, H)];
    else {
      const cx = num('cx', W / 2, W);
      const cy = num('cy', H / 2, H);
      coords = [num('fx', cx, W), num('fy', cy, H), num('fr', 0, D), cx, cy, num('r', D / 2, D)];
    }
    return { kind: tag === 'linearGradient' ? 'linear' : 'radial', stops, coords, matrix: mul(m, g) };
  }

  private pathD(segs: ReturnType<typeof parsePath>, m: M): string {
    const f = (n: number): string => (Math.round(n * 100) / 100).toString();
    let d = '';
    const box: [number, number, number, number] = [Infinity, Infinity, -Infinity, -Infinity];
    for (const s of segs)
      for (let i = 1; i + 1 < s.length; i += 2) {
        const x = s[i] as number;
        const y = s[i + 1] as number;
        box[0] = Math.min(box[0], x);
        box[1] = Math.min(box[1], y);
        box[2] = Math.max(box[2], x);
        box[3] = Math.max(box[3], y);
      }
    this.lastBox = box;
    for (const s of segs) {
      if (s[0] === 'Z') d += 'Z ';
      else if (s[0] === 'C') {
        const a = apply(m, s[1], s[2]);
        const b = apply(m, s[3], s[4]);
        const c = apply(m, s[5], s[6]);
        d += `C${f(a[0])} ${f(a[1])} ${f(b[0])} ${f(b[1])} ${f(c[0])} ${f(c[1])} `;
      } else {
        const p = apply(m, s[1], s[2]);
        d += `${s[0]}${f(p[0])} ${f(p[1])} `;
      }
    }
    return d.trim();
  }

  /** Current clip (intersection of paths in output space) and, while reading a clipPath, its captured geometry. */
  private clip: ClipPath[] = [];
  private capture: string[] | null = null;

  walk(n: XNode, parent: Style, m: M, viewW: number, viewH: number): void {
    const ref = this.capture ? undefined : /url\(\s*#([^)\s]+)\s*\)/.exec(n.attrs['clip-path'] ?? declarations(n.attrs.style ?? '')['clip-path'] ?? '')?.[1];
    const cp = ref ? this.ids.get(ref) : undefined;
    if (!cp || local(cp.name) !== 'clipPath' || cp.attrs.clipPathUnits === 'objectBoundingBox') return this.walkInner(n, parent, m, viewW, viewH);
    // the clipPath's shapes, in the user space of the element (its transform included)
    const captured: string[] = [];
    this.capture = captured;
    const mm = mul(mul(m, parseTransform(n.attrs.transform)), parseTransform(cp.attrs.transform));
    try {
      for (const c of elements(cp)) this.walkInner(c, { ...parent, fill: '000000', display: true }, mm, viewW, viewH);
    } finally {
      this.capture = null;
    }
    if (!captured.length) return; // an empty clip hides the element
    const prev = this.clip;
    const rule = declarations(cp.attrs.style ?? '')['clip-rule'] ?? elements(cp)[0]?.attrs['clip-rule'];
    this.clip = [...prev, { d: captured.join(' '), evenOdd: rule === 'evenodd' }];
    this.cmds.push({ k: 'clip', paths: this.clip });
    try {
      this.walkInner(n, parent, m, viewW, viewH);
    } finally {
      this.clip = prev;
      this.cmds.push({ k: 'clip', paths: prev });
    }
  }

  private walkInner(n: XNode, parent: Style, m: M, viewW: number, viewH: number): void {
    if (++this.depth > 200) return;
    try {
      const tag = local(n.name);
      if (['defs', 'style', 'title', 'desc', 'metadata', 'clipPath', 'mask', 'linearGradient', 'radialGradient', 'pattern', 'marker', 'symbol', 'filter', 'script'].includes(tag)) return;
      const s = this.styleOf(n, parent, viewW);
      if (!s.display) return;
      const mm = mul(m, parseTransform(n.attrs.transform));
      const num = (k: string, ref = viewW): number => len(n.attrs[k], ref) ?? 0;
      switch (tag) {
        case 'svg':
        case 'g':
        case 'a':
        case 'switch':
          for (const c of elements(n)) this.walk(c, s, mm, viewW, viewH);
          return;
        case 'use': {
          const id = (n.attrs.href ?? n.attrs['xlink:href'] ?? '').replace(/^#/, '');
          const target = this.ids.get(id);
          if (!target) return;
          const tm = mul(mm, [1, 0, 0, 1, num('x'), num('y', viewH)]);
          if (local(target.name) === 'symbol') for (const c of elements(target)) this.walk(c, s, tm, viewW, viewH);
          else this.walk(target, s, tm, viewW, viewH);
          return;
        }
        case 'path':
          this.paint(this.pathD(parsePath(n.attrs.d ?? ''), mm), s, mm);
          return;
        case 'rect': {
          const x = num('x');
          const y = num('y', viewH);
          const w = num('width');
          const h = num('height', viewH);
          if (w <= 0 || h <= 0) return;
          let rx = len(n.attrs.rx, viewW) ?? len(n.attrs.ry, viewH) ?? 0;
          let ry = len(n.attrs.ry, viewH) ?? rx;
          rx = Math.min(rx, w / 2);
          ry = Math.min(ry, h / 2);
          const d = rx || ry ? `M${x + rx} ${y}H${x + w - rx}A${rx} ${ry} 0 0 1 ${x + w} ${y + ry}V${y + h - ry}A${rx} ${ry} 0 0 1 ${x + w - rx} ${y + h}H${x + rx}A${rx} ${ry} 0 0 1 ${x} ${y + h - ry}V${y + ry}A${rx} ${ry} 0 0 1 ${x + rx} ${y}Z` : `M${x} ${y}H${x + w}V${y + h}H${x}Z`;
          this.paint(this.pathD(parsePath(d), mm), s, mm);
          return;
        }
        case 'circle':
        case 'ellipse': {
          const cx = num('cx');
          const cy = num('cy', viewH);
          const rx = tag === 'circle' ? num('r') : num('rx');
          const ry = tag === 'circle' ? rx : num('ry', viewH);
          if (rx <= 0 || ry <= 0) return;
          this.paint(this.pathD(parsePath(`M${cx + rx} ${cy}A${rx} ${ry} 0 1 1 ${cx - rx} ${cy}A${rx} ${ry} 0 1 1 ${cx + rx} ${cy}Z`), mm), s, mm);
          return;
        }
        case 'line':
          this.paint(this.pathD(parsePath(`M${num('x1')} ${num('y1', viewH)}L${num('x2')} ${num('y2', viewH)}`), mm), { ...s, fill: null }, mm);
          return;
        case 'polyline':
        case 'polygon': {
          const pts = (n.attrs.points ?? '').trim().split(/[\s,]+/).map(Number);
          if (pts.length < 4) return;
          let d = `M${pts[0]} ${pts[1]}`;
          for (let i = 2; i + 1 < pts.length; i += 2) d += `L${pts[i]} ${pts[i + 1]}`;
          if (tag === 'polygon') d += 'Z';
          this.paint(this.pathD(parsePath(d), mm), tag === 'polyline' ? { ...s, fill: s.fill } : s, mm);
          return;
        }
        case 'text':
          this.text(n, s, mm, viewW, viewH);
          return;
        case 'image': {
          const href = n.attrs.href ?? n.attrs['xlink:href'] ?? '';
          const img = parseDataUri(href);
          if (!img || !/png|jpe?g/.test(img.mime)) return;
          const x = num('x');
          const y = num('y', viewH);
          const w = num('width');
          const h = num('height', viewH);
          const [x0, y0] = apply(mm, x, y);
          const [x1, y1] = apply(mm, x + w, y + h);
          // PNG/JPEG only: the PDF writer embeds them directly
          this.cmds.push({ k: 'image', x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0), png: img.data });
          return;
        }
        default:
          return;
      }
    } finally {
      this.depth--;
    }
  }

  private text(n: XNode, s: Style, m: M, viewW: number, viewH: number): void {
    let x = len(n.attrs.x?.split(/[\s,]+/)[0], viewW) ?? 0;
    let y = len(n.attrs.y?.split(/[\s,]+/)[0], viewH) ?? 0;
    // Runs without their own position continue after the previous one (measured when drawn)
    let flow = false;
    let first = true;
    const emit = (raw: string, st: Style, xx: number, yy: number): void => {
      let t = raw.replace(/\s+/g, ' ');
      if (first) t = t.replace(/^ /, '');
      if (!t.trim()) {
        if (t) flow = true;
        return;
      }
      first = false;
      if (!st.fill || !st.display) return;
      const [px, py] = apply(m, xx, yy);
      const k = scaleOf(m);
      const angle = (-Math.atan2(m[1], m[0]) * 180) / Math.PI;
      const fam: FontFamily = /mono|courier|consol/i.test(st.fontFamily) ? 'mono' : (familyOf(st.fontFamily.split(',')[0].replace(/['"]/g, '').trim()) ?? 'sans');
      const bold = /bold|[6-9]00/.test(st.fontWeight);
      this.cmds.push({
        k: 'text',
        x: px,
        y: py,
        text: t,
        size: st.fontSize * k,
        family: /serif/i.test(st.fontFamily) && !/sans/i.test(st.fontFamily) && fam === 'sans' ? 'serif' : fam,
        bold,
        italic: /italic|oblique/.test(st.fontStyle),
        underline: false,
        color: st.fill,
        angle,
        anchor: flow ? 'start' : st.anchor,
        baseline: 'alphabetic',
        flow,
      });
      flow = true;
    };
    for (const c of n.children) {
      if (typeof c === 'string') emit(c, s, x, y);
      else if (local(c.name) === 'tspan') {
        const ts = this.styleOf(c, s, viewW);
        if (c.attrs.x !== undefined || c.attrs.y !== undefined) flow = false;
        if (c.attrs.x !== undefined) x = len(c.attrs.x.split(/[\s,]+/)[0], viewW) ?? x;
        if (c.attrs.y !== undefined) y = len(c.attrs.y.split(/[\s,]+/)[0], viewH) ?? y;
        if (c.attrs.dx !== undefined) x += len(c.attrs.dx, viewW) ?? 0;
        if (c.attrs.dy !== undefined) {
          y += len(c.attrs.dy, ts.fontSize) ?? 0;
          flow = false;
        }
        emit(textContent(c), ts, x, y);
      }
    }
  }
}

function declarations(body: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of body.split(';')) {
    const i = part.indexOf(':');
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).replace(/!important/, '').trim();
  }
  return out;
}

/** Parse an SVG document into a display list in its viewBox units. */
export function parseSvg(data: Uint8Array | string): Metafile | null {
  const src = typeof data === 'string' ? data : new TextDecoder().decode(data);
  let root: XNode | undefined;
  try {
    const doc = parseXml(src);
    const find = (n: XNode): XNode | undefined => (local(n.name) === 'svg' ? n : elements(n).map(find).find(Boolean));
    root = find(doc);
  } catch {
    return null;
  }
  if (!root) return null;
  const vb = (root.attrs.viewBox ?? '').trim().split(/[\s,]+/).map(Number);
  const wAttr = len(root.attrs.width);
  const hAttr = len(root.attrs.height);
  let [vx, vy, vw, vh] = vb.length === 4 && vb.every((v) => Number.isFinite(v)) ? vb : [0, 0, wAttr ?? 300, hAttr ?? 150];
  if (!(vw > 0 && vh > 0)) [vw, vh] = [300, 150];
  const outW = wAttr && !/%/.test(root.attrs.width ?? '') ? wAttr : vw;
  const outH = hAttr && !/%/.test(root.attrs.height ?? '') ? hAttr : vh;
  const r = new SvgRenderer(root);
  r.rootW = vw;
  r.rootH = vh;
  const base: Style = { fill: '000000', fillOpacity: 1, stroke: null, strokeOpacity: 1, strokeWidth: 1, opacity: 1, evenOdd: false, fontSize: 16, fontFamily: 'sans-serif', fontWeight: '400', fontStyle: 'normal', anchor: 'start', display: true };
  // viewBox → output units (uniform scale, centred: preserveAspectRatio xMidYMid meet)
  const k = Math.min(outW / vw, outH / vh);
  const m: M = [k, 0, 0, k, -vx * k + (outW - vw * k) / 2, -vy * k + (outH - vh * k) / 2];
  for (const c of elements(root)) r.walk(c, base, m, vw, vh);
  return { width: outW, height: outH, widthPt: (outW * 72) / 96, heightPt: (outH * 72) / 96, cmds: r.cmds };
}
