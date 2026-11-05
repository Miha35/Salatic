/**
 * PowerPoint (.pptx, PresentationML) → Deck. Resolves the theme (colour scheme
 * with lumMod/lumOff/tint/shade, major/minor fonts), the slide → layout → master
 * chain for backgrounds, placeholder geometry and text styles (txStyles,
 * lstStyle levels, bodyPr), group transforms, pictures with cropping, tables and
 * speaker notes. `deckToDoc` flattens the deck for text formats.
 */
import { unzipSync } from 'fflate';
import { familyOf, normalizeRuns, type Align, type Block, type Doc, type ListItem, type TextRun } from '../model';
import type { Deck, Slide, SlideBullet, SlideElement, SlideLine, SlidePara, SlideRole, SlideTableCell, SlideTableRow, SlideText } from '../slides';
import { child, childrenNamed, elements, find, parseXml, textContent, type XNode } from '../xml';
import { ConvertError, sniffImageMime, unzipOffice } from '../util';

const decoder = new TextDecoder('utf-8');
const EMU = 12700; // EMU per point
const pt = (v: string | undefined, def = 0): number => (v === undefined || v === '' || !Number.isFinite(+v) ? def : +v / EMU);

// ---------------------------------------------------------------------------
// Colours

function hexToRgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(0, 2), 16) / 255, parseInt(hex.slice(2, 4), 16) / 255, parseInt(hex.slice(4, 6), 16) / 255];
}
function rgbToHex(c: [number, number, number]): string {
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

interface Theme {
  colors: Record<string, string>;
  major: string;
  minor: string;
  fills: XNode[];
}

/** Colour of a fill/colour container (a:solidFill, a:fontRef, …), honouring modifiers. */
function colorOf(node: XNode | undefined, theme: Theme, map: Record<string, string>, phClr?: string): { hex: string; alpha: number } | undefined {
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

// ---------------------------------------------------------------------------
// Text styles

interface LevelStyle {
  align?: Align;
  marL?: number;
  indent?: number;
  lineHeight?: number;
  linePts?: number;
  spaceBefore?: number;
  spaceBeforePct?: number;
  spaceAfter?: number;
  spaceAfterPct?: number;
  bullet?: { kind: 'none' } | { kind: 'char'; char: string } | { kind: 'auto'; scheme: string; start: number };
  buColor?: string;
  buFont?: string;
  buScale?: number;
  run: RunStyle;
}

interface RunStyle {
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  font?: string;
  baseline?: number;
  highlight?: string;
}

const ALIGN: Record<string, Align> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };

interface Part {
  path: string;
  xml: XNode;
  rels: Map<string, { target: string; external: boolean; type: string }>;
}

interface Placeholder {
  type: string;
  idx?: string;
  node: XNode;
}

interface MasterLike {
  part: Part;
  placeholders: Placeholder[];
  map: Record<string, string>;
}

const DEFAULT_MAP: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };

class PptxReader {
  warnings = new Set<string>();
  private theme: Theme = { colors: {}, major: 'Calibri Light', minor: 'Calibri', fills: [] };
  private defaultTextStyle?: XNode;
  private parts = new Map<string, Part>();

  constructor(private files: Record<string, Uint8Array>) {}

  private part(path: string): Part | undefined {
    const cached = this.parts.get(path);
    if (cached) return cached;
    const data = this.files[path];
    if (!data) return undefined;
    const xml = parseXml(decoder.decode(data));
    const rels = new Map<string, { target: string; external: boolean; type: string }>();
    const dir = path.split('/').slice(0, -1).join('/');
    const relsPath = `${dir}/_rels/${path.split('/').pop()}.rels`;
    const relData = this.files[relsPath];
    if (relData) {
      for (const r of childrenNamed(find(parseXml(decoder.decode(relData)), 'Relationships'), 'Relationship')) {
        const external = r.attrs.TargetMode === 'External';
        rels.set(r.attrs.Id, { target: external ? r.attrs.Target : resolvePath(dir, r.attrs.Target), external, type: (r.attrs.Type ?? '').split('/').pop() ?? '' });
      }
    }
    const p = { path, xml, rels };
    this.parts.set(path, p);
    return p;
  }

  private relOfType(p: Part, type: string): string | undefined {
    for (const r of p.rels.values()) if (r.type === type && !r.external) return r.target;
    return undefined;
  }

  read(): Deck {
    const pres = this.part('ppt/presentation.xml');
    if (!pres) throw new ConvertError('В архиве нет презентации PowerPoint (ppt/presentation.xml).');
    const sz = find(pres.xml, 'p:sldSz');
    const width = pt(sz?.attrs.cx, 720);
    const height = pt(sz?.attrs.cy, 540);
    this.defaultTextStyle = find(pres.xml, 'p:defaultTextStyle');
    const slides: Slide[] = [];
    const ids = childrenNamed(find(pres.xml, 'p:sldIdLst'), 'p:sldId');
    let themeLoaded = false;
    for (const id of ids) {
      const target = pres.rels.get(id.attrs['r:id'])?.target;
      const slidePart = target ? this.part(target) : undefined;
      if (!slidePart) continue;
      const layoutPart = this.part(this.relOfType(slidePart, 'slideLayout') ?? '');
      const masterPart = layoutPart ? this.part(this.relOfType(layoutPart, 'slideMaster') ?? '') : undefined;
      if (!themeLoaded && masterPart) {
        this.loadTheme(this.part(this.relOfType(masterPart, 'theme') ?? ''));
        themeLoaded = true;
      }
      slides.push(this.slide(slidePart, layoutPart, masterPart, slides.length + 1));
    }
    const core = this.part('docProps/core.xml');
    const title = textContent(find(core?.xml, 'dc:title')).trim() || undefined;
    if (!slides.length) throw new ConvertError('В презентации нет слайдов.');
    return { title, width, height, slides };
  }

  private loadTheme(p: Part | undefined): void {
    if (!p) return;
    const scheme = find(p.xml, 'a:clrScheme');
    for (const c of elements(scheme)) {
      const col = colorOf(c, { colors: {}, major: '', minor: '', fills: [] }, {});
      if (col) this.theme.colors[c.name.replace(/^a:/, '')] = col.hex;
    }
    const major = find(find(p.xml, 'a:majorFont'), 'a:latin')?.attrs.typeface;
    const minor = find(find(p.xml, 'a:minorFont'), 'a:latin')?.attrs.typeface;
    if (major) this.theme.major = major;
    if (minor) this.theme.minor = minor;
    this.theme.fills = elements(find(p.xml, 'a:fillStyleLst'));
  }

  private masterLike(p: Part | undefined, parentMap?: Record<string, string>): MasterLike | undefined {
    if (!p) return undefined;
    const placeholders: Placeholder[] = [];
    for (const sp of elements(find(p.xml, 'p:spTree'))) {
      const ph = find(child(sp, 'p:nvSpPr') ?? child(sp, 'p:nvPicPr') ?? sp, 'p:ph');
      if (ph) placeholders.push({ type: ph.attrs.type ?? 'body', idx: ph.attrs.idx, node: sp });
    }
    const cm = find(p.xml, 'p:clrMap');
    const ovr = find(find(p.xml, 'p:clrMapOvr'), 'a:overrideClrMapping');
    const map: Record<string, string> = { ...DEFAULT_MAP, ...(parentMap ?? {}) };
    for (const src of [cm, ovr]) if (src) Object.assign(map, src.attrs);
    return { part: p, placeholders, map };
  }

  private slide(sp: Part, lp: Part | undefined, mp: Part | undefined, slideNo: number): Slide {
    const master = this.masterLike(mp);
    const layout = this.masterLike(lp, master?.map);
    const slide = this.masterLike(sp, layout?.map ?? master?.map)!;
    const map = slide.map;
    const ctx: ShapeCtx = { part: sp, map, layout, master, kind: 'slide', groups: [], slideNo };
    const out: Slide = { elements: [] };
    if (sp.xml && find(sp.xml, 'p:sld')?.attrs.show === '0') out.hidden = true;
    out.background = this.background(sp, map) ?? (lp ? this.background(lp, map) : undefined) ?? (mp ? this.background(mp, map) : undefined);
    // Decorative (non-placeholder) shapes of the master and the layout are drawn behind the slide
    const showMaster = (p: Part | undefined): boolean => find(p?.xml, 'p:sld')?.attrs.showMasterSp !== '0' && find(p?.xml, 'p:sldLayout')?.attrs.showMasterSp !== '0';
    if (mp && master && showMaster(sp) && showMaster(lp)) this.tree(find(mp.xml, 'p:spTree'), { part: mp, map, layout: undefined, master: undefined, kind: 'master', groups: [], slideNo }, out.elements);
    if (lp && layout && find(sp.xml, 'p:sld')?.attrs.showMasterSp !== '0') this.tree(find(lp.xml, 'p:spTree'), { part: lp, map, layout: undefined, master, kind: 'layout', groups: [], slideNo }, out.elements);
    this.tree(find(sp.xml, 'p:spTree'), ctx, out.elements);
    // Speaker notes
    const notesPart = this.part(this.relOfType(sp, 'notesSlide') ?? '');
    if (notesPart) {
      const blocks: Block[] = [];
      for (const s of elements(find(notesPart.xml, 'p:spTree'))) {
        const ph = find(s, 'p:ph');
        if (ph?.attrs.type !== 'body') continue;
        const body = child(s, 'p:txBody');
        for (const p of childrenNamed(body, 'a:p')) {
          const runs = normalizeRuns(childrenNamed(p, 'a:r').map((r) => ({ text: textContent(child(r, 'a:t')) })));
          if (runs.some((r) => r.text.trim())) blocks.push({ type: 'paragraph', runs });
        }
      }
      if (blocks.length) out.notes = blocks;
    }
    return out;
  }

  private background(p: Part, map: Record<string, string>): Slide['background'] | undefined {
    const bg = find(find(p.xml, 'p:cSld'), 'p:bg');
    if (!bg) return undefined;
    const pr = child(bg, 'p:bgPr');
    if (pr) {
      const solid = child(pr, 'a:solidFill');
      if (solid) return { color: colorOf(solid, this.theme, map)?.hex };
      const grad = child(pr, 'a:gradFill');
      if (grad) {
        const stops = childrenNamed(find(grad, 'a:gsLst'), 'a:gs');
        const mid = stops[Math.floor(stops.length / 2)] ?? stops[0];
        return { color: colorOf(mid, this.theme, map)?.hex };
      }
      const blip = child(pr, 'a:blipFill');
      if (blip) {
        const img = this.blip(p, find(blip, 'a:blip'));
        if (img) return { image: img };
      }
      return undefined;
    }
    const ref = child(bg, 'p:bgRef');
    if (ref) {
      const color = colorOf(ref, this.theme, map)?.hex;
      const idx = +(ref.attrs.idx ?? 0);
      // 1001+ refer to the background fill style list; a solid style takes the reference colour
      const style = this.theme.fills[idx - 1001] ?? this.theme.fills[idx - 1];
      if (style?.name === 'a:solidFill' || !style) return { color };
      if (style.name === 'a:gradFill') {
        const stops = childrenNamed(find(style, 'a:gsLst'), 'a:gs');
        return { color: colorOf(stops[0], this.theme, map, color)?.hex ?? color };
      }
      return { color };
    }
    return undefined;
  }

  private blip(p: Part, blip: XNode | undefined): { data: Uint8Array; mime: string } | undefined {
    const rid = blip?.attrs['r:embed'];
    const rel = rid ? p.rels.get(rid) : undefined;
    const data = rel && !rel.external ? this.files[rel.target] : undefined;
    if (!data) return undefined;
    const mime = sniffImageMime(data) ?? (/\.svg$/i.test(rel!.target) ? 'image/svg+xml' : /\.tiff?$/i.test(rel!.target) ? 'image/tiff' : /\.emf$/i.test(rel!.target) ? 'image/emf' : /\.wmf$/i.test(rel!.target) ? 'image/wmf' : 'application/octet-stream');
    return { data, mime };
  }

  // -- shapes ---------------------------------------------------------------------

  private tree(root: XNode | undefined, ctx: ShapeCtx, out: SlideElement[]): void {
    for (const n of elements(root)) {
      switch (n.name) {
        case 'p:sp':
          this.shape(n, ctx, out);
          break;
        case 'p:pic':
          this.picture(n, ctx, out);
          break;
        case 'p:cxnSp':
          this.connector(n, ctx, out);
          break;
        case 'p:grpSp': {
          const xfrm = find(child(n, 'p:grpSpPr'), 'a:xfrm');
          const off = child(xfrm, 'a:off');
          const ext = child(xfrm, 'a:ext');
          const chOff = child(xfrm, 'a:chOff');
          const chExt = child(xfrm, 'a:chExt');
          const g = {
            x: pt(off?.attrs.x),
            y: pt(off?.attrs.y),
            sx: chExt && +chExt.attrs.cx ? +(ext?.attrs.cx ?? 0) / +chExt.attrs.cx : 1,
            sy: chExt && +chExt.attrs.cy ? +(ext?.attrs.cy ?? 0) / +chExt.attrs.cy : 1,
            cx: pt(chOff?.attrs.x),
            cy: pt(chOff?.attrs.y),
          };
          this.tree(n, { ...ctx, groups: [...ctx.groups, g] }, out);
          break;
        }
        case 'p:graphicFrame':
          this.graphicFrame(n, ctx, out);
          break;
        case 'mc:AlternateContent': {
          // The fallback is what older consumers render (often a picture of SmartArt/equations)
          const choice = child(n, 'mc:Choice');
          const fallback = child(n, 'mc:Fallback');
          const req = choice?.attrs.Requires ?? '';
          const pick = /^(p14|a14|p15)$/.test(req) || !fallback ? choice : fallback;
          if (pick) this.tree(pick, ctx, out);
          break;
        }
      }
    }
  }

  /** Geometry with group transforms applied. */
  private place(x: number, y: number, w: number, h: number, ctx: ShapeCtx): { x: number; y: number; w: number; h: number } {
    for (let i = ctx.groups.length - 1; i >= 0; i--) {
      const g = ctx.groups[i];
      x = g.x + (x - g.cx) * g.sx;
      y = g.y + (y - g.cy) * g.sy;
      w *= g.sx;
      h *= g.sy;
    }
    return { x, y, w, h };
  }

  private phOf(n: XNode): Placeholder | undefined {
    const ph = find(child(n, 'p:nvSpPr') ?? child(n, 'p:nvPicPr') ?? child(n, 'p:nvGraphicFramePr'), 'p:ph');
    return ph ? { type: ph.attrs.type ?? 'body', idx: ph.attrs.idx, node: n } : undefined;
  }

  private match(list: Placeholder[] | undefined, ph: Placeholder): Placeholder | undefined {
    if (!list) return undefined;
    const t = ph.type === 'ctrTitle' ? 'title' : ph.type;
    if (ph.idx !== undefined) {
      const byIdx = list.find((p) => p.idx === ph.idx);
      if (byIdx) return byIdx;
    }
    return list.find((p) => p.type === ph.type) ?? list.find((p) => (p.type === 'ctrTitle' ? 'title' : p.type) === t) ?? (t === 'subTitle' || t === 'obj' ? list.find((p) => p.type === 'body') : undefined);
  }

  /** Inherited placeholders: [master, layout] for a slide shape. */
  private inherited(n: XNode, ctx: ShapeCtx): { ph?: Placeholder; chain: XNode[] } {
    const ph = this.phOf(n);
    if (!ph) return { chain: [] };
    const chain: XNode[] = [];
    if (ctx.kind === 'slide') {
      const l = this.match(ctx.layout?.placeholders, ph);
      const m = this.match(ctx.master?.placeholders, l ? { ...ph, type: l.type, idx: undefined } : ph);
      if (m) chain.push(m.node);
      if (l) chain.push(l.node);
    } else if (ctx.kind === 'layout') {
      const m = this.match(ctx.master?.placeholders, { ...ph, idx: undefined });
      if (m) chain.push(m.node);
    }
    return { ph, chain };
  }

  private xfrmOf(nodes: XNode[]): { x: number; y: number; w: number; h: number; rot: number } | undefined {
    for (let i = nodes.length - 1; i >= 0; i--) {
      const sp = child(nodes[i], 'p:spPr') ?? child(nodes[i], 'p:xfrm') ?? undefined;
      const xfrm = nodes[i].name === 'p:graphicFrame' ? child(nodes[i], 'p:xfrm') : find(sp, 'a:xfrm');
      const off = child(xfrm, 'a:off');
      const ext = child(xfrm, 'a:ext');
      if (off && ext) return { x: pt(off.attrs.x), y: pt(off.attrs.y), w: pt(ext.attrs.cx), h: pt(ext.attrs.cy), rot: +(xfrm?.attrs.rot ?? 0) / 60000 };
    }
    return undefined;
  }

  private shape(n: XNode, ctx: ShapeCtx, out: SlideElement[]): void {
    const { ph, chain } = this.inherited(n, ctx);
    // Placeholders of masters/layouts are prompts, not content
    if (ph && ctx.kind !== 'slide') return;
    const all = [...chain, n];
    const xf = this.xfrmOf(all);
    if (!xf) return;
    const geo = this.place(xf.x, xf.y, xf.w, xf.h, ctx);
    const spPr = child(n, 'p:spPr');
    const style = child(n, 'p:style');
    const geom = find(spPr, 'a:prstGeom')?.attrs.prst ?? (find(spPr, 'a:custGeom') ? 'rect' : 'rect');
    const fill = this.fill(spPr, style, ctx);
    const line = this.line(spPr, style, ctx);
    const type = ph?.type;
    const role: SlideRole = type === 'title' || type === 'ctrTitle' ? 'title' : type === 'subTitle' ? 'subtitle' : type === 'dt' || type === 'ftr' || type === 'sldNum' ? 'footer' : ph ? 'body' : 'other';
    const fontColor = colorOf(find(style, 'a:fontRef'), this.theme, ctx.map)?.hex;
    const text = this.text(child(n, 'p:txBody'), [...chain.map((c) => child(c, 'p:txBody'))], type, ctx, fontColor);
    // Blip-filled shapes are pictures
    const blipFill = child(spPr, 'a:blipFill');
    if (blipFill) {
      const img = this.blip(ctx.part, find(blipFill, 'a:blip'));
      if (img) out.push({ kind: 'image', ...geo, ...img, rot: xf.rot || undefined });
    }
    if (geom === 'line' || geom === 'straightConnector1') {
      if (line) out.push({ kind: 'line', x1: geo.x, y1: geo.y, x2: geo.x + geo.w, y2: geo.y + geo.h, line });
      return;
    }
    if (!fill && !line && !text) return;
    out.push({ kind: 'shape', ...geo, geom, rot: xf.rot || undefined, fill: fill?.hex, fillOpacity: fill && fill.alpha < 1 ? fill.alpha : undefined, line, text, role });
  }

  private fill(spPr: XNode | undefined, style: XNode | undefined, ctx: ShapeCtx): { hex: string; alpha: number } | undefined {
    if (child(spPr, 'a:noFill')) return undefined;
    const solid = child(spPr, 'a:solidFill');
    if (solid) return colorOf(solid, this.theme, ctx.map);
    const grad = child(spPr, 'a:gradFill');
    if (grad) {
      const stops = childrenNamed(find(grad, 'a:gsLst'), 'a:gs');
      return colorOf(stops[Math.floor(stops.length / 2)] ?? stops[0], this.theme, ctx.map);
    }
    if (child(spPr, 'a:blipFill') || child(spPr, 'a:pattFill') || child(spPr, 'a:grpFill')) return undefined;
    const ref = find(style, 'a:fillRef');
    if (ref && +(ref.attrs.idx ?? 0) > 0) return colorOf(ref, this.theme, ctx.map);
    return undefined;
  }

  private line(spPr: XNode | undefined, style: XNode | undefined, ctx: ShapeCtx): SlideLine | undefined {
    const ln = child(spPr, 'a:ln');
    if (child(ln, 'a:noFill')) return undefined;
    // w="0" is a hairline in PowerPoint
    const width = ln?.attrs.w !== undefined ? Math.max(0.25, pt(ln.attrs.w)) : 0.75;
    const dash = !!child(ln, 'a:prstDash') && child(ln, 'a:prstDash')!.attrs.val !== 'solid';
    const solid = child(ln, 'a:solidFill');
    if (solid) {
      const c = colorOf(solid, this.theme, ctx.map);
      return c ? { color: c.hex, width, dash, opacity: c.alpha < 1 ? c.alpha : undefined } : undefined;
    }
    const ref = find(style, 'a:lnRef');
    if (ref && +(ref.attrs.idx ?? 0) > 0) {
      const c = colorOf(ref, this.theme, ctx.map);
      return c ? { color: c.hex, width, dash } : undefined;
    }
    return undefined;
  }

  private picture(n: XNode, ctx: ShapeCtx, out: SlideElement[]): void {
    const { ph, chain } = this.inherited(n, ctx);
    if (ph && ctx.kind !== 'slide') return;
    const xf = this.xfrmOf([...chain, n]);
    if (!xf) return;
    const geo = this.place(xf.x, xf.y, xf.w, xf.h, ctx);
    const blipFill = child(n, 'p:blipFill');
    const img = this.blip(ctx.part, find(blipFill, 'a:blip'));
    if (!img) return;
    const src = child(blipFill, 'a:srcRect');
    const crop = src ? { l: +(src.attrs.l ?? 0) / 100000, t: +(src.attrs.t ?? 0) / 100000, r: +(src.attrs.r ?? 0) / 100000, b: +(src.attrs.b ?? 0) / 100000 } : undefined;
    const alt = find(n, 'p:cNvPr')?.attrs.descr || undefined;
    out.push({ kind: 'image', ...geo, ...img, crop: crop && (crop.l || crop.t || crop.r || crop.b) ? crop : undefined, rot: xf.rot || undefined, alt });
  }

  private connector(n: XNode, ctx: ShapeCtx, out: SlideElement[]): void {
    const xf = this.xfrmOf([n]);
    if (!xf) return;
    const g = this.place(xf.x, xf.y, xf.w, xf.h, ctx);
    const xfrm = find(child(n, 'p:spPr'), 'a:xfrm');
    const line = this.line(child(n, 'p:spPr'), child(n, 'p:style'), ctx);
    if (!line) return;
    const fh = xfrm?.attrs.flipH === '1';
    const fv = xfrm?.attrs.flipV === '1';
    out.push({ kind: 'line', x1: fh ? g.x + g.w : g.x, y1: fv ? g.y + g.h : g.y, x2: fh ? g.x : g.x + g.w, y2: fv ? g.y : g.y + g.h, line });
  }

  private graphicFrame(n: XNode, ctx: ShapeCtx, out: SlideElement[]): void {
    const data = find(n, 'a:graphicData');
    const uri = data?.attrs.uri ?? '';
    const xf = this.xfrmOf([...this.inherited(n, ctx).chain, n]);
    if (!xf) return;
    const geo = this.place(xf.x, xf.y, xf.w, xf.h, ctx);
    if (uri.endsWith('/table')) {
      const tbl = child(data, 'a:tbl');
      if (tbl) out.push(this.table(tbl, geo, ctx));
      return;
    }
    if (uri.endsWith('/chart')) this.warnings.add('Диаграммы (графики PowerPoint) не отображаются: сохраните их как изображения.');
    else if (uri.endsWith('/diagram')) this.warnings.add('Некоторые SmartArt-схемы не отображаются.');
    else if (uri.includes('ole')) {
      // Embedded objects carry a preview picture
      const pic = find(n, 'p:pic');
      if (pic) this.picture(pic, ctx, out);
    }
  }

  private table(tbl: XNode, geo: { x: number; y: number; w: number; h: number }, ctx: ShapeCtx): SlideElement {
    const cols = childrenNamed(child(tbl, 'a:tblGrid'), 'a:gridCol').map((c) => pt(c.attrs.w));
    const pr = child(tbl, 'a:tblPr');
    const firstRow = pr?.attrs.firstRow === '1';
    const bandRow = pr?.attrs.bandRow === '1';
    const styled = !!find(pr, 'a:tableStyleId') || firstRow || bandRow;
    const accent = this.theme.colors.accent1 ?? '4472C4';
    const light = (k: number): string => rgbToHex(hexToRgb(accent).map((v) => 1 - (1 - v) * k) as [number, number, number]);
    const rows: SlideTableRow[] = [];
    childrenNamed(tbl, 'a:tr').forEach((tr, ri) => {
      const cells: SlideTableCell[] = [];
      const header = firstRow && ri === 0;
      for (const tc of childrenNamed(tr, 'a:tc')) {
        const tcPr = child(tc, 'a:tcPr');
        let fill = this.fill(tcPr, undefined, ctx)?.hex;
        // Default "Medium Style 2 – Accent 1" look when the cell has no fill of its own
        if (!fill && styled && !child(tcPr, 'a:noFill')) fill = header ? accent : bandRow && (ri - (firstRow ? 1 : 0)) % 2 === 0 ? light(0.2) : light(0.1);
        const text = this.text(child(tc, 'a:txBody'), [], 'tbl', ctx, header ? 'FFFFFF' : undefined) ?? emptyText();
        if (header) for (const p of text.paragraphs) for (const r of p.runs) r.bold ??= true;
        text.insets = [pt(tcPr?.attrs.marL, 7.2), pt(tcPr?.attrs.marT, 3.6), pt(tcPr?.attrs.marR, 7.2), pt(tcPr?.attrs.marB, 3.6)];
        text.anchor = tcPr?.attrs.anchor === 'ctr' ? 'ctr' : tcPr?.attrs.anchor === 'b' ? 'b' : 't';
        const border = (name: string): SlideLine | undefined => {
          const ln = child(tcPr, name);
          if (!ln || child(ln, 'a:noFill')) return styled ? { color: 'FFFFFF', width: 1 } : undefined;
          const c = colorOf(child(ln, 'a:solidFill'), this.theme, ctx.map);
          return c ? { color: c.hex, width: ln.attrs.w ? pt(ln.attrs.w) : 0.75 } : undefined;
        };
        cells.push({
          text,
          fill,
          span: +(tc.attrs.gridSpan ?? 1) || 1,
          rowSpan: +(tc.attrs.rowSpan ?? 1) || 1,
          merged: tc.attrs.hMerge === '1' || tc.attrs.vMerge === '1',
          borders: { l: border('a:lnL'), r: border('a:lnR'), t: border('a:lnT'), b: border('a:lnB') },
        });
      }
      rows.push({ h: pt(tr.attrs.h), cells });
    });
    return { kind: 'table', ...geo, cols, rows };
  }

  // -- text -------------------------------------------------------------------------

  /** Level style chain for a placeholder type (master txStyles → placeholder lstStyles → shape). */
  private levelStyles(type: string | undefined, ctx: ShapeCtx, inheritedBodies: (XNode | undefined)[], own: XNode | undefined): XNode[][] {
    const master = ctx.kind === 'master' ? ctx.part : ctx.master?.part;
    const txStyles = find(master?.xml, 'p:txStyles');
    let base: XNode | undefined;
    if (type === 'tbl') base = this.defaultTextStyle;
    else if (!type) base = this.defaultTextStyle ?? child(txStyles, 'p:otherStyle');
    else if (type === 'title' || type === 'ctrTitle') base = child(txStyles, 'p:titleStyle');
    else if (type === 'body' || type === 'subTitle' || type === 'obj' || type === 'pic' || type === 'tbl') base = child(txStyles, 'p:bodyStyle');
    else base = child(txStyles, 'p:otherStyle');
    const lists = [base, ...inheritedBodies.map((b) => child(b, 'a:lstStyle')), child(own, 'a:lstStyle')];
    const levels: XNode[][] = [];
    for (let l = 1; l <= 9; l++) levels.push(lists.map((ls) => child(ls, `a:lvl${l}pPr`)).filter((x): x is XNode => !!x));
    return levels;
  }

  private runStyleOf(rPr: XNode | undefined, ctx: ShapeCtx, into: RunStyle): void {
    if (!rPr) return;
    const a = rPr.attrs;
    if (a.sz) into.size = +a.sz / 100;
    if (a.b !== undefined) into.bold = a.b === '1' || a.b === 'true';
    if (a.i !== undefined) into.italic = a.i === '1' || a.i === 'true';
    if (a.u !== undefined) into.underline = a.u !== 'none';
    if (a.strike !== undefined) into.strike = a.strike !== 'noStrike';
    if (a.baseline !== undefined) into.baseline = +a.baseline;
    const fill = child(rPr, 'a:solidFill');
    if (fill) into.color = colorOf(fill, this.theme, ctx.map)?.hex ?? into.color;
    const hl = child(rPr, 'a:highlight');
    if (hl) into.highlight = colorOf(hl, this.theme, ctx.map)?.hex;
    const latin = child(rPr, 'a:latin')?.attrs.typeface ?? child(rPr, 'a:ea')?.attrs.typeface;
    if (latin) into.font = latin;
  }

  private levelOf(nodes: XNode[], ctx: ShapeCtx): LevelStyle {
    const s: LevelStyle = { run: {} };
    for (const n of nodes) this.applyPPr(n, ctx, s);
    return s;
  }

  private applyPPr(n: XNode, ctx: ShapeCtx, s: LevelStyle): void {
    const a = n.attrs;
    if (a.algn && ALIGN[a.algn]) s.align = ALIGN[a.algn];
    if (a.marL !== undefined) s.marL = pt(a.marL);
    if (a.indent !== undefined) s.indent = pt(a.indent);
    const spc = (name: string): { pct?: number; pts?: number } | undefined => {
      const e = child(n, name);
      if (!e) return undefined;
      const pct = child(e, 'a:spcPct')?.attrs.val;
      const pts = child(e, 'a:spcPts')?.attrs.val;
      return { pct: pct !== undefined ? +pct / 100000 : undefined, pts: pts !== undefined ? +pts / 100 : undefined };
    };
    const ln = spc('a:lnSpc');
    if (ln) {
      s.lineHeight = ln.pct;
      s.linePts = ln.pts;
    }
    const sb = spc('a:spcBef');
    if (sb) {
      s.spaceBefore = sb.pts;
      s.spaceBeforePct = sb.pct;
    }
    const sa = spc('a:spcAft');
    if (sa) {
      s.spaceAfter = sa.pts;
      s.spaceAfterPct = sa.pct;
    }
    if (child(n, 'a:buNone')) s.bullet = { kind: 'none' };
    const bc = child(n, 'a:buChar');
    if (bc) s.bullet = { kind: 'char', char: bc.attrs.char ?? '•' };
    const ba = child(n, 'a:buAutoNum');
    if (ba) s.bullet = { kind: 'auto', scheme: ba.attrs.type ?? 'arabicPeriod', start: +(ba.attrs.startAt ?? 1) || 1 };
    const bclr = child(n, 'a:buClr');
    if (bclr) s.buColor = colorOf(bclr, this.theme, ctx.map)?.hex;
    if (child(n, 'a:buClrTx')) s.buColor = undefined;
    const bf = child(n, 'a:buFont');
    if (bf) s.buFont = bf.attrs.typeface;
    const bsz = child(n, 'a:buSzPct');
    if (bsz) s.buScale = +(bsz.attrs.val ?? 100000) / 100000;
    this.runStyleOf(child(n, 'a:defRPr'), ctx, s.run);
  }

  private fontFamily(name: string | undefined): TextRun['font'] {
    if (!name) return undefined;
    if (name.startsWith('+mj')) name = this.theme.major;
    else if (name.startsWith('+mn')) name = this.theme.minor;
    return familyOf(name);
  }

  private text(body: XNode | undefined, inherited: (XNode | undefined)[], type: string | undefined, ctx: ShapeCtx, fontColor?: string): SlideText | undefined {
    if (!body) return undefined;
    const paras = childrenNamed(body, 'a:p');
    if (!paras.some((p) => childrenNamed(p, 'a:r').some((r) => textContent(child(r, 'a:t')).trim()) || child(p, 'a:fld'))) return undefined;
    // bodyPr inherits from the placeholders
    const bodyPr: Record<string, string> = {};
    let fontScale = 1;
    let lineReduction = 0;
    for (const b of [...inherited, body]) {
      const pr = child(b, 'a:bodyPr');
      if (!pr) continue;
      Object.assign(bodyPr, pr.attrs);
      const fit = child(pr, 'a:normAutofit');
      if (fit) {
        fontScale = +(fit.attrs.fontScale ?? 100000) / 100000;
        lineReduction = +(fit.attrs.lnSpcReduction ?? 0) / 100000;
      } else if (child(pr, 'a:noAutofit') || child(pr, 'a:spAutoFit')) {
        fontScale = 1;
        lineReduction = 0;
      }
    }
    const levels = this.levelStyles(type, ctx, inherited, body).map((nodes) => this.levelOf(nodes, ctx));
    const counters: number[] = [];
    const out: SlidePara[] = [];
    for (const p of paras) {
      const pPr = child(p, 'a:pPr');
      const lvl = Math.min(8, +(pPr?.attrs.lvl ?? 0) || 0);
      const st: LevelStyle = { ...levels[lvl], run: { ...levels[lvl].run } };
      if (pPr) this.applyPPr(pPr, ctx, st);
      const size = (st.run.size ?? 18) * fontScale;
      const runs: TextRun[] = [];
      const pushRun = (text: string, rPr: XNode | undefined, link?: string): void => {
        const rs: RunStyle = { ...st.run };
        this.runStyleOf(rPr, ctx, rs);
        const r: TextRun = { text };
        if (rs.bold) r.bold = true;
        if (rs.italic) r.italic = true;
        if (rs.underline) r.underline = true;
        if (rs.strike) r.strike = true;
        if (rs.baseline && rs.baseline > 0) r.sup = true;
        else if (rs.baseline && rs.baseline < 0) r.sub = true;
        const color = rs.color ?? fontColor ?? this.theme.colors[ctx.map.tx1 ?? 'dk1'] ?? '000000';
        r.color = color;
        if (rs.highlight) r.highlight = rs.highlight;
        r.font = this.fontFamily(rs.font ?? (type === 'title' || type === 'ctrTitle' ? '+mj-lt' : '+mn-lt'));
        r.size = Math.round((rs.size ?? 18) * fontScale * 10) / 10;
        if (link) r.link = link;
        runs.push(r);
      };
      for (const r of elements(p)) {
        if (r.name === 'a:r') {
          const rPr = child(r, 'a:rPr');
          const hlink = child(rPr, 'a:hlinkClick')?.attrs['r:id'];
          const rel = hlink ? ctx.part.rels.get(hlink) : undefined;
          pushRun(textContent(child(r, 'a:t')), rPr, rel?.external && /^(https?:|mailto:)/i.test(rel.target) ? rel.target : undefined);
        } else if (r.name === 'a:br') pushRun('\n', child(r, 'a:rPr'));
        else if (r.name === 'a:fld') {
          const t = r.attrs.type === 'slidenum' ? String(ctx.slideNo ?? textContent(child(r, 'a:t'))) : textContent(child(r, 'a:t'));
          pushRun(t, child(r, 'a:rPr'));
        }
      }
      const hasText = runs.some((r) => r.text.trim());
      let bullet: SlideBullet | undefined;
      if (hasText && st.bullet && st.bullet.kind !== 'none') {
        while (counters.length > lvl + 1) counters.pop();
        const first = runs[0];
        const symbolFont = !!st.buFont && SYMBOL_FONT.test(st.buFont);
        if (st.bullet.kind === 'char') bullet = { text: st.bullet.char, color: st.buColor ?? first.color, font: symbolFont ? first.font : (this.fontFamily(st.buFont) ?? first.font), scale: st.buScale ?? 1 };
        else {
          counters[lvl] = (counters[lvl] ?? st.bullet.start - 1) + 1;
          bullet = { text: autoNumber(st.bullet.scheme, counters[lvl]), color: st.buColor ?? first.color, font: first.font, scale: st.buScale ?? 1 };
        }
        // Symbol-font bullets (Wingdings "§", "Ø", private-use code points, …) are mapped to Unicode
        const cp = bullet.text.codePointAt(0) ?? 0;
        if (symbolFont || (cp >= 0xe000 && cp <= 0xf8ff)) {
          const ch = cp >= 0xf000 && cp <= 0xf0ff ? String.fromCharCode(cp - 0xf000) : bullet.text;
          bullet.text = WINGDINGS[ch] ?? '•';
        }
      } else if (!st.bullet || st.bullet.kind === 'none') counters.length = lvl;
      const lineFactor = Math.max(0.5, (st.lineHeight ?? 1) * (1 - lineReduction));
      out.push({
        runs,
        align: st.align,
        level: lvl,
        marL: st.marL ?? 0,
        indent: st.indent ?? 0,
        lineHeight: st.linePts === undefined ? lineFactor : undefined,
        linePts: st.linePts,
        spaceBefore: st.spaceBefore ?? (st.spaceBeforePct ?? 0) * size * 1.2,
        spaceAfter: st.spaceAfter ?? (st.spaceAfterPct ?? 0) * size * 1.2,
        bullet,
        size,
      });
    }
    const ins = (k: string, def: number): number => (bodyPr[k] !== undefined ? pt(bodyPr[k]) : def);
    return {
      paragraphs: out,
      anchor: bodyPr.anchor === 'ctr' ? 'ctr' : bodyPr.anchor === 'b' ? 'b' : 't',
      insets: [ins('lIns', 7.2), ins('tIns', 3.6), ins('rIns', 7.2), ins('bIns', 3.6)],
      wrap: bodyPr.wrap !== 'none',
      fontScale,
      lineReduction,
    };
  }
}

interface ShapeCtx {
  part: Part;
  map: Record<string, string>;
  layout?: MasterLike;
  master?: MasterLike;
  kind: 'slide' | 'layout' | 'master';
  groups: { x: number; y: number; sx: number; sy: number; cx: number; cy: number }[];
  slideNo?: number;
}

const SYMBOL_FONT = /wingdings|webdings|symbol|starbats|opensymbol|marlett/i;
const WINGDINGS: Record<string, string> = { '§': '▪', Ø: '➢', 'ü': '✓', 'q': '❑', 'v': '❖', 'n': '■', 'l': '●', 'Ÿ': '•', 'o': '□', 'è': '➔' };

function emptyText(): SlideText {
  return { paragraphs: [], anchor: 't', insets: [7.2, 3.6, 7.2, 3.6], wrap: true, fontScale: 1, lineReduction: 0 };
}

function roman(n: number): string {
  const map: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of map) while (n >= v) {
    s += r;
    n -= v;
  }
  return s;
}

function autoNumber(scheme: string, n: number): string {
  const alpha = (k: number): string => {
    let s = '';
    while (k > 0) {
      s = String.fromCharCode(97 + ((k - 1) % 26)) + s;
      k = Math.floor((k - 1) / 26);
    }
    return s;
  };
  let core = String(n);
  if (/^alphaLc/.test(scheme)) core = alpha(n);
  else if (/^alphaUc/.test(scheme)) core = alpha(n).toUpperCase();
  else if (/^romanLc/.test(scheme)) core = roman(n);
  else if (/^romanUc/.test(scheme)) core = roman(n).toUpperCase();
  if (/ParenBoth$/.test(scheme)) return `(${core})`;
  if (/ParenR$/.test(scheme)) return `${core})`;
  if (/Plain$/.test(scheme)) return core;
  return `${core}.`;
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base ? base.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

export function pptxToDeck(bytes: Uint8Array): { deck: Deck; warnings: string[] } {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipOffice(bytes, unzipSync);
  } catch (e) {
    if (e instanceof ConvertError) throw e;
    throw new ConvertError('Файл PPTX повреждён.');
  }
  const r = new PptxReader(files);
  const deck = r.read();
  return { deck, warnings: [...r.warnings] };
}

// ---------------------------------------------------------------------------
// Deck → Doc (text formats)

function paragraphsToBlocks(text: SlideText, role: SlideRole, out: Block[]): void {
  // Sizes and theme colours belong to the slide look, not to a flowing document
  const plain = (runs: TextRun[]): TextRun[] =>
    normalizeRuns(
      runs.map((r) => {
        const { size: _s, font: _f, color, ...rest } = r;
        return color && !/^(000000|FFFFFF|[0-3][0-9A-F][0-3][0-9A-F][0-3][0-9A-F])$/i.test(color) && role === 'other' ? { ...rest, color } : rest;
      }),
    );
  let list: { block: Extract<Block, { type: 'list' }>; level: number }[] = [];
  for (const p of text.paragraphs) {
    const runs = plain(p.runs);
    if (!runs.some((r) => r.text.trim())) {
      list = [];
      continue;
    }
    if (role === 'title') {
      out.push({ type: 'heading', level: 2, runs: runs.map(({ bold: _b, ...r }) => r) });
      continue;
    }
    if (p.bullet) {
      const ordered = /^[(\dIVXLCivxlcA-Za-z]+[.)]?$/.test(p.bullet.text) && !/^[•◦▪■●➢✓❑❖□➔–\-*]$/.test(p.bullet.text);
      while (list.length && list[list.length - 1].level > p.level) list.pop();
      if (!list.length || list[list.length - 1].level < p.level || list[list.length - 1].block.ordered !== ordered) {
        const fresh: Extract<Block, { type: 'list' }> = { type: 'list', ordered, items: [] };
        const parent = list[list.length - 1];
        if (parent && parent.level < p.level) {
          const last: ListItem | undefined = parent.block.items[parent.block.items.length - 1];
          if (last) last.blocks.push(fresh);
          else out.push(fresh);
        } else {
          if (parent) list.pop();
          out.push(fresh);
        }
        list.push({ block: fresh, level: p.level });
      }
      list[list.length - 1].block.items.push({ blocks: [{ type: 'paragraph', runs }] });
      continue;
    }
    list = [];
    out.push({ type: 'paragraph', runs, ...(p.align && p.align !== 'left' ? { align: p.align } : {}) });
  }
}

export function deckToDoc(deck: Deck): Doc {
  const blocks: Block[] = [];
  deck.slides.forEach((s, i) => {
    if (s.hidden) return;
    if (blocks.length) blocks.push({ type: 'pagebreak' });
    const start = blocks.length;
    // Reading order: top-to-bottom, then left-to-right; the title first
    const els = s.elements.filter((e) => !(e.kind === 'shape' && e.role === 'footer'));
    const order = (e: SlideElement): number => (e.kind === 'shape' && e.role === 'title' ? -1e9 : (e.kind === 'line' ? e.y1 : e.y) * 10000 + (e.kind === 'line' ? e.x1 : e.x));
    for (const e of [...els].sort((a, b) => order(a) - order(b))) {
      if (e.kind === 'shape' && e.text) paragraphsToBlocks(e.text, e.role, blocks);
      else if (e.kind === 'image' && e.w > 20 && e.h > 20 && /^image\/(png|jpeg|gif|webp|bmp|svg\+xml|tiff|emf|wmf)$/.test(e.mime)) {
        blocks.push({ type: 'image', data: e.data, mime: e.mime, width: Math.round((e.w * 96) / 72), height: Math.round((e.h * 96) / 72), alt: e.alt });
      } else if (e.kind === 'table') {
        const rows = e.rows.map((r, ri) => ({
          cells: r.cells
            .filter((c) => !c.merged)
            .map((c) => {
              const inner: Block[] = [];
              paragraphsToBlocks(c.text, 'body', inner);
              return { blocks: inner, ...(ri === 0 && c.text.paragraphs.some((p) => p.runs.some((x) => x.bold)) ? { header: true } : {}), ...(c.span > 1 ? { colspan: c.span } : {}) };
            }),
        }));
        blocks.push({ type: 'table', rows });
      }
    }
    if (!blocks.slice(start).some((b) => b.type === 'heading')) blocks.splice(start, 0, { type: 'heading', level: 2, runs: [{ text: `Слайд ${i + 1}` }] });
    if (s.notes?.length) blocks.push({ type: 'quote', blocks: [{ type: 'paragraph', runs: [{ text: 'Заметки докладчика', bold: true }] }, ...s.notes] });
  });
  return { title: deck.title, blocks, font: 'sans', size: 12 };
}
