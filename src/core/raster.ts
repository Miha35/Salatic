/**
 * A small anti-aliased scanline rasteriser for display-list paths (absolute M/L/C/Z) and gradient
 * evaluation. Used where a format has no vector transparency or gradients (EMF): the fill is
 * rendered into a 32-bit premultiplied bitmap that GDI alpha-blends.
 */
import type { Gradient } from './model';

type Edge = { x0: number; y0: number; x1: number; y1: number; dir: number };

/** Path data → line segments (curves flattened), in the path's own units. */
function edgesOf(d: string): Edge[] {
  const edges: Edge[] = [];
  const re = /([MLCZmlcz])|(-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)/g;
  let op = '';
  let nums: number[] = [];
  let cx = 0;
  let cy = 0;
  let sx = 0;
  let sy = 0;
  const line = (x: number, y: number): void => {
    if (y !== cy) edges.push(cy < y ? { x0: cx, y0: cy, x1: x, y1: y, dir: 1 } : { x0: x, y0: y, x1: cx, y1: cy, dir: -1 });
    cx = x;
    cy = y;
  };
  const flush = (): void => {
    if (op === 'M') {
      for (let i = 0; i + 1 < nums.length; i += 2) {
        if (i === 0) {
          // an open subpath still closes for filling
          if (cx !== sx || cy !== sy) line(sx, sy);
          cx = sx = nums[0];
          cy = sy = nums[1];
        } else line(nums[i], nums[i + 1]);
      }
    } else if (op === 'L') for (let i = 0; i + 1 < nums.length; i += 2) line(nums[i], nums[i + 1]);
    else if (op === 'C')
      for (let i = 0; i + 5 < nums.length; i += 6) {
        const [x1, y1, x2, y2, x3, y3] = nums.slice(i, i + 6);
        const x0 = cx;
        const y0 = cy;
        const n = Math.max(4, Math.min(64, Math.ceil(Math.hypot(x3 - x0, y3 - y0) / 2)));
        for (let k = 1; k <= n; k++) {
          const t = k / n;
          const u = 1 - t;
          line(u * u * u * x0 + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3, u * u * u * y0 + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3);
        }
      }
    else if (op === 'Z') {
      line(sx, sy);
    }
    nums = [];
  };
  for (let m; (m = re.exec(d)); ) {
    if (m[1]) {
      flush();
      op = m[1].toUpperCase();
      if (op === 'Z') flush();
    } else nums.push(+m[2]);
  }
  flush();
  line(sx, sy);
  return edges;
}

/**
 * Coverage (0…1) of a path over a w × h pixel grid whose pixel (0,0) covers [ox, ox + 1/scale) ×
 * [oy, oy + 1/scale) in path units: exact along x, 4 samples per pixel along y.
 */
export function coverage(d: string, evenOdd: boolean, ox: number, oy: number, w: number, h: number, scale: number): Float32Array {
  const out = new Float32Array(w * h);
  const edges = edgesOf(d).map((e) => ({ x0: (e.x0 - ox) * scale, y0: (e.y0 - oy) * scale, x1: (e.x1 - ox) * scale, y1: (e.y1 - oy) * scale, dir: e.dir }));
  const SUB = 4;
  const xs: { x: number; dir: number }[] = [];
  for (let py = 0; py < h; py++)
    for (let s = 0; s < SUB; s++) {
      const y = py + (s + 0.5) / SUB;
      xs.length = 0;
      for (const e of edges) if (y >= e.y0 && y < e.y1) xs.push({ x: e.x0 + ((y - e.y0) * (e.x1 - e.x0)) / (e.y1 - e.y0), dir: e.dir });
      if (xs.length < 2) continue;
      xs.sort((a, b) => a.x - b.x);
      let wind = 0;
      for (let k = 0; k + 1 < xs.length; k++) {
        wind += evenOdd ? 1 : xs[k].dir;
        const inside = evenOdd ? wind % 2 === 1 : wind !== 0;
        if (!inside) continue;
        // span [a, b): add its length to each pixel it crosses
        const a = Math.max(0, xs[k].x);
        const b = Math.min(w, xs[k + 1].x);
        if (b <= a) continue;
        const row = py * w;
        for (let px = Math.floor(a); px < Math.ceil(b) && px < w; px++) out[row + px] += (Math.min(b, px + 1) - Math.max(a, px)) / SUB;
      }
    }
  for (let i = 0; i < out.length; i++) out[i] = Math.min(1, out[i]);
  return out;
}

const rgbOf = (hex: string): number[] => [0, 2, 4].map((o) => parseInt(hex.slice(o, o + 2), 16) || 0);

/** Colour and alpha of a gradient at a point of the drawing space. */
export function gradientAt(g: Gradient, x: number, y: number): [number, number, number, number] {
  let px = x;
  let py = y;
  if (g.matrix) {
    const [a, b, c, d, e, f] = g.matrix;
    const det = a * d - b * c || 1;
    px = (d * (x - e) - c * (y - f)) / det;
    py = (a * (y - f) - b * (x - e)) / det;
  }
  let t: number;
  const k = g.coords;
  if (g.kind === 'linear') {
    const dx = k[2] - k[0];
    const dy = k[3] - k[1];
    t = ((px - k[0]) * dx + (py - k[1]) * dy) / (dx * dx + dy * dy || 1);
  } else if (g.kind === 'radial') {
    // largest t with |p − c(t)| = r(t), c(t) = f + t(c − f), r(t) = fr + t(r − fr)
    const [fx, fy, fr, cx, cy, r] = k;
    const cdx = cx - fx;
    const cdy = cy - fy;
    const dr = r - fr;
    const pdx = px - fx;
    const pdy = py - fy;
    const A = cdx * cdx + cdy * cdy - dr * dr;
    const B = pdx * cdx + pdy * cdy + fr * dr;
    const C = pdx * pdx + pdy * pdy - fr * fr;
    if (Math.abs(A) < 1e-9) t = C / (2 * B || 1);
    else {
      const disc = B * B - A * C;
      t = disc < 0 ? 1 : (B + Math.sqrt(disc)) / A;
    }
  } else {
    const [x0, y0, x1, y1, fx0, fy0, fx1, fy1] = k;
    t = Math.max(px < fx0 ? (fx0 - px) / (fx0 - x0 || 1) : 0, px > fx1 ? (px - fx1) / (x1 - fx1 || 1) : 0, py < fy0 ? (fy0 - py) / (fy0 - y0 || 1) : 0, py > fy1 ? (py - fy1) / (y1 - fy1 || 1) : 0);
  }
  t = Math.max(0, Math.min(1, t));
  const s = g.stops;
  if (!s.length) return [0, 0, 0, 1];
  if (t <= s[0].pos) return [...rgbOf(s[0].color), s[0].alpha ?? 1] as [number, number, number, number];
  for (let i = 1; i < s.length; i++)
    if (t <= s[i].pos) {
      const u = (t - s[i - 1].pos) / (s[i].pos - s[i - 1].pos || 1);
      const a = rgbOf(s[i - 1].color);
      const b = rgbOf(s[i].color);
      return [a[0] + (b[0] - a[0]) * u, a[1] + (b[1] - a[1]) * u, a[2] + (b[2] - a[2]) * u, (s[i - 1].alpha ?? 1) + ((s[i].alpha ?? 1) - (s[i - 1].alpha ?? 1)) * u];
    }
  const last = s[s.length - 1];
  return [...rgbOf(last.color), last.alpha ?? 1] as [number, number, number, number];
}

/** Bounding box of path data (control points included). */
export function pathBox(d: string): [number, number, number, number] {
  const nums = (d.match(/-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) ?? []).map(Number);
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (let i = 0; i + 1 < nums.length; i += 2) {
    x0 = Math.min(x0, nums[i]);
    x1 = Math.max(x1, nums[i]);
    y0 = Math.min(y0, nums[i + 1]);
    y1 = Math.max(y1, nums[i + 1]);
  }
  return [x0, y0, x1, y1];
}
