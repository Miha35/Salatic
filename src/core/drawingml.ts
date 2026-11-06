/** DrawingML colours: scheme/preset/system colours with lumMod, lumOff, tint, shade and alpha. */
import { child, childrenNamed, elements, type XNode } from './xml';
import type { Gradient } from './model';


export function hexToRgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(0, 2), 16) / 255, parseInt(hex.slice(2, 4), 16) / 255, parseInt(hex.slice(4, 6), 16) / 255];
}
export function rgbToHex(c: [number, number, number]): string {
  return c.map((v) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}
function rgbToHsl([r, g, b]: [number, number, number]): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  h /= 6;
  return [h, s, l];
}
function hslToRgb([h, s, l]: [number, number, number]): [number, number, number] {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number): number => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)];
}

const PRESET_COLORS: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500' };

export interface Theme {
  colors: Record<string, string>;
  major: string;
  minor: string;
  fills: XNode[];
  /** Background fill styles (bgRef idx 1001…). */
  bgFills?: XNode[];
}

/**
 * a:gradFill → gradient in the local space of a w × h box (y down).
 * Linear: `a:lin ang` (60000ths of a degree, clockwise; `scaled` stretches the angle with the box),
 * spanning the box corner to corner along the gradient direction.
 * Path (circle/rect/shape): stop 0 at the focus (`a:fillToRect`), the last stop at the farthest corner;
 * all paths are drawn as circles.
 */
export function gradientOf(grad: XNode, w: number, h: number, theme: Theme, map: Record<string, string>, phClr?: string): Gradient | undefined {
  const stops = childrenNamed(child(grad, 'a:gsLst'), 'a:gs')
    .map((gs) => ({ gs, c: colorOf(gs, theme, map, phClr) }))
    .filter((x): x is { gs: XNode; c: { hex: string; alpha: number } } => !!x.c)
    .map(({ gs, c }) => ({ pos: Math.min(1, Math.max(0, +(gs.attrs.pos ?? 0) / 100000)), color: c.hex, ...(c.alpha < 1 ? { alpha: c.alpha } : {}) }))
    .sort((a, b) => a.pos - b.pos);
  if (!stops.length || !(w > 0) || !(h > 0)) return undefined;
  const path = child(grad, 'a:path');
  if (path) {
    const r = child(path, 'a:fillToRect')?.attrs ?? {};
    const pct = (k: string): number => +(r[k] ?? 0) / 100000;
    const fx0 = pct('l') * w;
    const fy0 = pct('t') * h;
    const fx1 = Math.max(fx0, w - pct('r') * w);
    const fy1 = Math.max(fy0, h - pct('b') * h);
    // rect (and shape, drawn like LibreOffice does): concentric rectangles from the focus to the edges
    if (path.attrs.path === 'rect' || path.attrs.path === 'shape') return { kind: 'rect', stops, coords: [0, 0, w, h, fx0, fy0, fx1, fy1] };
    const fx = (fx0 + fx1) / 2;
    const fy = (fy0 + fy1) / 2;
    // a true circle around the focus, reaching the farthest corner (like PowerPoint and LibreOffice)
    const reach = Math.max(...[0, w].flatMap((cx) => [0, h].map((cy) => Math.hypot(cx - fx, cy - fy))));
    return { kind: 'radial', stops, coords: [fx, fy, 0, fx, fy, reach] };
  }
  const lin = child(grad, 'a:lin');
  const ang = ((+(lin?.attrs.ang ?? 0) / 60000) * Math.PI) / 180;
  const scaled = lin?.attrs.scaled === '1' || lin?.attrs.scaled === 'true';
  let dx = Math.cos(ang);
  let dy = Math.sin(ang);
  if (scaled) [dx, dy] = [dx * h, dy * w];
  const len = Math.hypot(dx, dy) || 1;
  dx /= len;
  dy /= len;
  const proj = [0, w].flatMap((x) => [0, h].map((y) => (x - w / 2) * dx + (y - h / 2) * dy));
  const lo = Math.min(...proj);
  const hi = Math.max(...proj);
  return { kind: 'linear', stops, coords: [w / 2 + dx * lo, h / 2 + dy * lo, w / 2 + dx * hi, h / 2 + dy * hi] };
}

const toLinear = (c: number): number => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c: number): number => Math.min(1, Math.max(0, c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055));

/** Colour of a fill/colour container (a:solidFill, a:fontRef, …), honouring modifiers. */
export function colorOf(node: XNode | undefined, theme: Theme, map: Record<string, string>, phClr?: string): { hex: string; alpha: number } | undefined {
  if (!node) return undefined;
  const c = elements(node).find((e) => /^a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(e.name));
  if (!c) return undefined;
  let hex: string | undefined;
  switch (c.name) {
    case 'a:srgbClr':
      hex = c.attrs.val;
      break;
    case 'a:sysClr':
      hex = c.attrs.lastClr ?? (c.attrs.val === 'window' ? 'FFFFFF' : '000000');
      break;
    case 'a:prstClr':
      hex = PRESET_COLORS[c.attrs.val] ?? '000000';
      break;
    case 'a:scrgbClr':
      hex = rgbToHex([+c.attrs.r / 100000, +c.attrs.g / 100000, +c.attrs.b / 100000]);
      break;
    case 'a:schemeClr': {
      const v = c.attrs.val;
      if (v === 'phClr') hex = phClr;
      else hex = theme.colors[map[v] ?? v] ?? theme.colors[v];
      break;
    }
    case 'a:hslClr':
      hex = rgbToHex(hslToRgb([+c.attrs.hue / 21600000, +c.attrs.sat / 100000, +c.attrs.lum / 100000]));
      break;
  }
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return undefined;
  let rgb = hexToRgb(hex);
  let alpha = 1;
  for (const m of elements(c)) {
    const v = +(m.attrs.val ?? 0) / 100000;
    switch (m.name) {
      case 'a:lumMod': {
        const hsl = rgbToHsl(rgb);
        hsl[2] *= v;
        rgb = hslToRgb(hsl);
        break;
      }
      case 'a:lumOff': {
        const hsl = rgbToHsl(rgb);
        hsl[2] = Math.min(1, hsl[2] + v);
        rgb = hslToRgb(hsl);
        break;
      }
      // tint/shade work in linear RGB (as Office and LibreOffice do), the HSL modifiers in sRGB
      case 'a:tint':
        rgb = rgb.map((x) => toSrgb(1 - (1 - toLinear(x)) * v)) as [number, number, number];
        break;
      case 'a:shade':
        rgb = rgb.map((x) => toSrgb(toLinear(x) * v)) as [number, number, number];
        break;
      case 'a:satMod':
      case 'a:satOff':
      case 'a:sat':
      case 'a:hueMod':
      case 'a:hueOff':
      case 'a:hue':
      case 'a:lum': {
        const hsl = rgbToHsl(rgb);
        const deg = +(m.attrs.val ?? 0) / 60000 / 360;
        if (m.name === 'a:satMod') hsl[1] *= v;
        else if (m.name === 'a:satOff') hsl[1] += v;
        else if (m.name === 'a:sat') hsl[1] = v;
        else if (m.name === 'a:hueMod') hsl[0] *= v;
        else if (m.name === 'a:hueOff') hsl[0] += deg;
        else if (m.name === 'a:hue') hsl[0] = deg;
        else hsl[2] = v;
        hsl[0] = ((hsl[0] % 1) + 1) % 1;
        hsl[1] = Math.min(1, Math.max(0, hsl[1]));
        hsl[2] = Math.min(1, Math.max(0, hsl[2]));
        rgb = hslToRgb(hsl);
        break;
      }
      case 'a:comp': {
        const hsl = rgbToHsl(rgb);
        hsl[0] = (hsl[0] + 0.5) % 1;
        rgb = hslToRgb(hsl);
        break;
      }
      case 'a:inv':
        rgb = rgb.map((x) => 1 - x) as [number, number, number];
        break;
      case 'a:gray': {
        const y = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2];
        rgb = [y, y, y];
        break;
      }
      case 'a:alpha':
        alpha = v;
        break;
    }
  }
  return { hex: rgbToHex(rgb), alpha };
}

