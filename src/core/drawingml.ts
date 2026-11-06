/** DrawingML colours: scheme/preset/system colours with lumMod, lumOff, tint, shade and alpha. */
import { elements, type XNode } from './xml';


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
}

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
      case 'a:tint':
        rgb = rgb.map((x) => 1 - (1 - x) * v) as [number, number, number];
        break;
      case 'a:shade':
        rgb = rgb.map((x) => x * v) as [number, number, number];
        break;
      case 'a:alpha':
        alpha = v;
        break;
    }
  }
  return { hex: rgbToHex(rgb), alpha };
}

