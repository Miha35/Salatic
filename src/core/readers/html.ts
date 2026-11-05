import type { Align, Block, Doc, ListItem, ParaFormat, TableRow, TextRun } from '../model';
import { familyOf, trimRuns } from '../model';
import { findTag, htmlText, parseHtml, parseStyle, type HNode } from '../html';
import { imageSize, parseDataUri, sniffImageMime } from '../util';

type InlineStyle = Omit<TextRun, 'text'>;

const INLINE = new Set([
  '#text', 'a', 'abbr', 'b', 'bdi', 'bdo', 'big', 'br', 'cite', 'code', 'data', 'del', 'dfn', 'em', 'font', 'i', 'img', 'ins', 'kbd',
  'label', 'mark', 'q', 's', 'samp', 'small', 'span', 'strike', 'strong', 'sub', 'sup', 'time', 'tt', 'u', 'var', 'wbr', 'nobr',
]);
const IGNORED = new Set(['head', 'script', 'style', 'meta', 'link', 'title', 'noscript', 'template', 'button', 'input', 'select', 'textarea', 'svg', 'canvas', 'iframe', 'object', 'video', 'audio', 'colgroup', 'col']);

const NAMED_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', orange: 'FFA500',
  purple: '800080', gray: '808080', grey: '808080', brown: 'A52A2A', navy: '000080', teal: '008080', maroon: '800000',
  olive: '808000', lime: '00FF00', aqua: '00FFFF', fuchsia: 'FF00FF', silver: 'C0C0C0', darkred: '8B0000', darkblue: '00008B',
  darkgreen: '006400',
};

export function parseColor(v: string | undefined): string | undefined {
  if (!v) return undefined;
  v = v.trim().toLowerCase();
  let hex: string | undefined;
  const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
  if (short) hex = short[1] + short[1] + short[2] + short[2] + short[3] + short[3];
  else if (/^#[0-9a-f]{6}$/.test(v)) hex = v.slice(1);
  else {
    const m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(v);
    if (m) hex = [m[1], m[2], m[3]].map((x) => Math.min(255, +x).toString(16).padStart(2, '0')).join('');
    else hex = NAMED_COLORS[v]?.toLowerCase();
  }
  if (!hex || hex === '000000') return undefined;
  return hex.toUpperCase();
}

function styleFor(node: HNode, parent: InlineStyle): InlineStyle {
  const s: InlineStyle = { ...parent };
  switch (node.tag) {
    case 'b':
    case 'strong':
      s.bold = true;
      break;
    case 'i':
    case 'em':
    case 'cite':
    case 'dfn':
    case 'var':
      s.italic = true;
      break;
    case 'u':
    case 'ins':
      s.underline = true;
      break;
    case 's':
    case 'strike':
    case 'del':
      s.strike = true;
      break;
    case 'code':
    case 'kbd':
    case 'samp':
    case 'tt':
      s.code = true;
      break;
    case 'sup':
      s.sup = true;
      break;
    case 'sub':
      s.sub = true;
      break;
    case 'a': {
      const href = node.attrs.href?.trim();
      if (href && /^(https?:|mailto:|tel:)/i.test(href)) s.link = href;
      break;
    }
    case 'font': {
      const c = parseColor(node.attrs.color);
      if (c) s.color = c;
      break;
    }
  }
  const st = parseStyle(node.attrs.style);
  const fw = st['font-weight'];
  if (fw) {
    if (fw === 'bold' || fw === 'bolder' || +fw >= 600) s.bold = true;
    else if (fw === 'normal' || fw === 'lighter' || +fw <= 500) s.bold = false;
  }
  const fs = st['font-style'];
  if (fs) s.italic = fs === 'italic' || fs === 'oblique';
  const td = st['text-decoration'] ?? st['text-decoration-line'];
  if (td) {
    if (td.includes('underline')) s.underline = true;
    if (td.includes('line-through')) s.strike = true;
    if (td === 'none') {
      s.underline = false;
      s.strike = false;
    }
  }
  const va = st['vertical-align'];
  if (va === 'super') s.sup = true;
  if (va === 'sub') s.sub = true;
  if (st.color) {
    const c = parseColor(st.color);
    s.color = c;
  }
  const face = st['font-family'] ?? (node.tag === 'font' ? node.attrs.face?.toLowerCase() : undefined);
  if (face) {
    s.font = familyOf(face);
    if (s.font === 'mono' && /monospace|courier|consolas/.test(face)) s.code = true;
  }
  const size = st['font-size'] ? lengthPt(st['font-size']) : node.tag === 'font' && node.attrs.size ? FONT_TAG_SIZES[+node.attrs.size] : undefined;
  if (size && size > 2 && size < 200) s.size = size;
  const bg = st['background-color'] ?? (st.background && !/url\(/.test(st.background) ? st.background.split(/\s+/)[0] : undefined);
  if (bg && !/transparent|inherit|initial|none|^#fff(fff)?$|white/.test(bg)) {
    const c = parseColor(bg) ?? (/^#?0{3,6}$|black/.test(bg) ? '000000' : undefined);
    if (c) s.highlight = c;
  }
  if (node.tag === 'mark') s.highlight ??= 'FFFF00';
  return s;
}

const FONT_TAG_SIZES: Record<number, number> = { 1: 7.5, 2: 10, 3: 12, 4: 13.5, 5: 18, 6: 24, 7: 36 };

/** CSS length → points (relative units against a 12pt base). */
export function lengthPt(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const named: Record<string, number> = { 'xx-small': 7, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, 'xxx-large': 36 };
  if (named[v]) return named[v];
  const m = /^(-?[\d.]+)\s*(pt|px|em|rem|cm|mm|in|pc|%)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  if (!Number.isFinite(n)) return undefined;
  switch (m[2]) {
    case 'px':
      return Math.round(n * 0.75 * 100) / 100;
    case 'em':
    case 'rem':
      return n * 12;
    case '%':
      return (n / 100) * 12;
    case 'cm':
      return n * 28.3465;
    case 'mm':
      return n * 2.83465;
    case 'in':
      return n * 72;
    case 'pc':
      return n * 12;
    case 'pt':
      return n;
    default:
      return n === 0 ? 0 : n * 0.75; // unitless lengths in HTML attributes are pixels
  }
}

/** Paragraph formatting from an element's inline style. */
function paraFormatOf(node: HNode): ParaFormat {
  const st = parseStyle(node.attrs.style);
  const f: ParaFormat = {};
  const a = alignOf(node);
  if (a) f.align = a;
  const ml = lengthPt(st['margin-left'] ?? st['padding-left']);
  if (ml) f.indent = ml;
  const ti = lengthPt(st['text-indent']);
  if (ti) f.firstLine = ti;
  const mt = lengthPt(st['margin-top']);
  if (mt !== undefined && st['margin-top']) f.spaceBefore = mt;
  const mb = lengthPt(st['margin-bottom']);
  if (mb !== undefined && st['margin-bottom']) f.spaceAfter = mb;
  const lh = st['line-height'];
  if (lh && lh !== 'normal') {
    const m = /^([\d.]+)(%)?$/.exec(lh);
    if (m) f.lineHeight = Math.round(((m[2] ? +m[1] / 100 : +m[1]) / 1.15) * 100) / 100;
  }
  return f;
}

function alignOf(node: HNode): Align | undefined {
  const a = (parseStyle(node.attrs.style)['text-align'] ?? node.attrs.align ?? '').toLowerCase();
  if (a === 'center' || a === 'right' || a === 'justify' || a === 'left') return a;
  if (a === 'start') return 'left';
  if (a === 'end') return 'right';
  return undefined;
}

const isPageBreak = (node: HNode, where: 'before' | 'after'): boolean => {
  const st = parseStyle(node.attrs.style);
  return st[`page-break-${where}`] === 'always' || st[`break-${where}`] === 'page';
};

function imageFromNode(node: HNode): Block | null {
  const src = node.attrs.src ?? '';
  const parsed = parseDataUri(src);
  if (!parsed) return null;
  const mime = sniffImageMime(parsed.data) ?? parsed.mime;
  if (!mime.startsWith('image/')) return null;
  const st = parseStyle(node.attrs.style);
  const natural = imageSize(parsed.data);
  let w = parseFloat(node.attrs.width ?? st.width ?? '') || 0;
  let h = parseFloat(node.attrs.height ?? st.height ?? '') || 0;
  if (natural) {
    if (w && !h) h = (w * natural.height) / natural.width;
    else if (h && !w) w = (h * natural.width) / natural.height;
    else if (!w && !h) {
      w = natural.width;
      h = natural.height;
    }
  }
  if (!w || !h) {
    w = w || 300;
    h = h || 200;
  }
  return { type: 'image', data: parsed.data, mime, width: Math.round(w), height: Math.round(h), alt: node.attrs.alt || undefined };
}

class Ctx {
  title?: string;
  header?: Block[];
  footer?: Block[];
  /** Footnote bodies by element id (pandoc/markdown-it/our own HTML). */
  notes = new Map<string, HNode>();
}

let currentCtx: Ctx | null = null;

function noteFor(href: string | undefined): HNode | undefined {
  if (!href?.startsWith('#') || !currentCtx) return undefined;
  return currentCtx.notes.get(href.slice(1));
}

/** Fix whitespace across runs: collapse, drop spaces around line breaks. */
function tidyRuns(runs: TextRun[]): TextRun[] {
  const out: TextRun[] = [];
  let prevEndsSpace = true;
  for (const r of runs) {
    if (r.note || r.field) {
      out.push(r);
      prevEndsSpace = false;
      continue;
    }
    let t = r.text;
    if (!r.code || !t.includes('\n')) {
      if (prevEndsSpace) t = t.replace(/^ +/, '');
      t = t.replace(/ *\n */g, '\n');
    }
    if (!t) continue;
    prevEndsSpace = / $/.test(t) || t.endsWith('\n');
    out.push({ ...r, text: t });
  }
  return trimRuns(out);
}

function inlineRuns(node: HNode, style: InlineStyle, runs: TextRun[], pre: boolean, sink: (b: Block) => void): void {
  if (node.tag === '#text') {
    let t = node.text ?? '';
    if (!pre) t = t.replace(/[ \t\r\n\f]+/g, ' ');
    if (t) runs.push({ text: t, ...clean(style) });
    return;
  }
  if (node.tag === 'br') {
    runs.push({ text: '\n', ...clean(style) });
    return;
  }
  if (node.attrs['data-field'] === 'page' || node.attrs['data-field'] === 'pages') {
    runs.push({ text: '', ...clean(style), field: node.attrs['data-field'] });
    return;
  }
  if (node.tag === 'a') {
    const li = noteFor(node.attrs.href);
    if (li && currentCtx) {
      const { sup: _s, link: _l, ...rest } = clean(style);
      const ctx = currentCtx;
      ctx.notes.delete(node.attrs.href!.slice(1)); // a note body is used once
      const body = collectBlocks(li.children.filter((c) => !(c.tag === 'a' && /^#fnref|footnote-back/.test((c.attrs.href ?? '') + ' ' + (c.attrs.class ?? '')))), {}, ctx);
      runs.push({ text: '', ...rest, note: stripBackrefs(body) });
      return;
    }
    if (/^#fnref/.test(node.attrs.href ?? '') || /footnote-back/.test(node.attrs.class ?? '')) return;
  }
  if (node.tag === 'img') {
    const img = imageFromNode(node);
    if (img) sink(img);
    else if (node.attrs.alt) runs.push({ text: node.attrs.alt, ...clean(style) });
    return;
  }
  if (IGNORED.has(node.tag)) return;
  const st = styleFor(node, style);
  const blockish = !INLINE.has(node.tag);
  if (blockish && runs.length) runs.push({ text: '\n', ...clean(style) });
  if (node.tag === 'q') runs.push({ text: '«', ...clean(st) });
  for (const c of node.children) inlineRuns(c, st, runs, pre || node.tag === 'pre', sink);
  if (node.tag === 'q') runs.push({ text: '»', ...clean(st) });
}

function clean(s: InlineStyle): InlineStyle {
  const o: InlineStyle = {};
  for (const [k, v] of Object.entries(s) as [keyof InlineStyle, unknown][]) if (v) (o as Record<string, unknown>)[k] = v;
  return o;
}

function stripBackrefs(blocks: Block[]): Block[] {
  return blocks
    .map((b) => (b.type === 'paragraph' ? { ...b, runs: trimRuns(b.runs.filter((r) => r.text.trim() !== '↩' && r.text.trim() !== '↩︎')) } : b))
    .filter((b) => b.type !== 'paragraph' || b.runs.length);
}

function collectBlocks(children: HNode[], style: InlineStyle, ctx: Ctx, fmt?: ParaFormat | Align): Block[] {
  const out: Block[] = [];
  let runs: TextRun[] = [];
  const f: ParaFormat = typeof fmt === 'string' ? { align: fmt } : (fmt ?? {});
  const flush = (): void => {
    const r = tidyRuns(runs);
    if (r.length) out.push({ type: 'paragraph', runs: r, ...f });
    runs = [];
  };
  const sink = (b: Block): void => {
    flush();
    out.push(b);
  };
  for (const c of children) {
    if (INLINE.has(c.tag)) inlineRuns(c, style, runs, false, sink);
    else {
      flush();
      blockFrom(c, style, ctx, out, f.align);
    }
  }
  flush();
  return out;
}

function headingRuns(node: HNode, style: InlineStyle): { runs: TextRun[]; images: Block[] } {
  const runs: TextRun[] = [];
  const images: Block[] = [];
  for (const c of node.children) inlineRuns(c, style, runs, false, (b) => images.push(b));
  return { runs: tidyRuns(runs), images };
}

function blockFrom(node: HNode, style: InlineStyle, ctx: Ctx, out: Block[], inheritedAlign?: Align): void {
  const tag = node.tag;
  if (tag === 'title') {
    ctx.title ??= htmlText(node).trim() || undefined;
    return;
  }
  if (IGNORED.has(tag)) return;
  if (isPageBreak(node, 'before')) out.push({ type: 'pagebreak' });
  const align = alignOf(node) ?? inheritedAlign;
  const st = styleFor(node, style);
  switch (tag) {
    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6': {
      const { runs, images } = headingRuns(node, st);
      if (runs.length) {
        const h: Block = { type: 'heading', level: +tag[1] as 1, runs, ...paraFormatOf(node) };
        if (align) h.align = align;
        out.push(h);
      }
      out.push(...images);
      break;
    }
    case 'p':
    case 'summary':
    case 'figcaption':
    case 'dt':
    case 'dd':
    case 'address': {
      const extra = tag === 'dt' ? { bold: true } : tag === 'figcaption' ? { italic: true } : {};
      out.push(...collectBlocks(node.children, { ...st, ...extra }, ctx, { ...paraFormatOf(node), ...(align ? { align } : {}) }));
      break;
    }
    case 'ul':
    case 'ol':
    case 'menu': {
      const items: ListItem[] = [];
      for (const c of node.children) {
        if (c.tag === 'li') items.push({ blocks: collectBlocks(c.children, st, ctx) });
        else if (c.tag === '#text') {
          if (c.text?.trim()) items.push({ blocks: [{ type: 'paragraph', runs: [{ text: c.text.trim() }] }] });
        } else if (c.tag === 'ul' || c.tag === 'ol') {
          // malformed nesting: <ul><li>a</li><ul>…</ul></ul> → attach to previous item
          const nested: Block[] = [];
          blockFrom(c, st, ctx, nested);
          if (items.length) items[items.length - 1].blocks.push(...nested);
          else items.push({ blocks: nested });
        } else items.push({ blocks: collectBlocks([c], st, ctx) });
      }
      for (const it of items) if (!it.blocks.length) it.blocks.push({ type: 'paragraph', runs: [] });
      if (items.length) out.push({ type: 'list', ordered: tag === 'ol', items });
      break;
    }
    case 'li':
      out.push({ type: 'list', ordered: false, items: [{ blocks: collectBlocks(node.children, st, ctx) }] });
      break;
    case 'table': {
      const rows: TableRow[] = [];
      const caption = node.children.find((c) => c.tag === 'caption');
      if (caption) out.push(...collectBlocks(caption.children, { ...st, bold: true }, ctx, 'center'));
      const visitRows = (n: HNode, inHead: boolean): void => {
        for (const c of n.children) {
          if (c.tag === 'tr') {
            const cells = c.children
              .filter((x) => x.tag === 'td' || x.tag === 'th')
              .map((cell) => {
                const span = parseInt(cell.attrs.colspan ?? '1', 10);
                const header = cell.tag === 'th' || inHead;
                return {
                  blocks: collectBlocks(cell.children, header ? { ...st, bold: true } : st, ctx, alignOf(cell)),
                  ...(header ? { header: true } : {}),
                  ...(span > 1 && span < 1000 ? { colspan: span } : {}),
                };
              });
            if (cells.length) rows.push({ cells });
          } else if (c.tag === 'thead' || c.tag === 'tbody' || c.tag === 'tfoot') visitRows(c, c.tag === 'thead');
        }
      };
      visitRows(node, false);
      if (rows.length) out.push({ type: 'table', rows });
      break;
    }
    case 'pre': {
      let text = htmlText(node);
      if (text.startsWith('\n')) text = text.slice(1);
      text = text.replace(/\s+$/, '');
      if (text) out.push({ type: 'code', text });
      break;
    }
    case 'blockquote': {
      const blocks = collectBlocks(node.children, st, ctx);
      if (blocks.length) out.push({ type: 'quote', blocks });
      break;
    }
    case 'hr':
      out.push(/\bpage-break\b/.test(node.attrs.class ?? '') || 'data-page-break' in node.attrs ? { type: 'pagebreak' } : { type: 'hr' });
      break;
    case 'img': {
      const img = imageFromNode(node);
      if (img) out.push(img);
      break;
    }
    case 'header':
    case 'footer':
    case 'section':
    case 'div':
    case 'aside': {
      const cls = node.attrs.class ?? '';
      if (/\bdoc-header\b/.test(cls)) {
        ctx.header = collectBlocks(node.children, st, ctx);
        break;
      }
      if (/\bdoc-footer\b/.test(cls)) {
        ctx.footer = collectBlocks(node.children, st, ctx);
        break;
      }
      if (/\bfootnotes\b/.test(cls) || node.attrs.role === 'doc-endnotes') break; // bodies were attached to their references
      out.push(...collectBlocks(node.children, st, ctx, align));
      break;
    }
    default:
      out.push(...collectBlocks(node.children, st, ctx, align));
  }
  if (isPageBreak(node, 'after')) out.push({ type: 'pagebreak' });
}

/** Parse an HTML string (full document or fragment) into a Doc. */
export function htmlToDoc(html: string): Doc {
  const tree = parseHtml(html);
  const ctx = new Ctx();
  const titleNode = findTag(tree, 'title');
  if (titleNode) ctx.title = htmlText(titleNode).trim() || undefined;
  // Footnote bodies: <li id="fn1"> inside a .footnotes container
  const visit = (n: HNode, inNotes: boolean): void => {
    const here = inNotes || /\bfootnotes\b/.test(n.attrs.class ?? '') || n.attrs.role === 'doc-endnotes';
    if (here && n.tag === 'li' && n.attrs.id) ctx.notes.set(n.attrs.id, n);
    for (const c of n.children) visit(c, here);
  };
  visit(tree, false);
  const body = findTag(tree, 'body') ?? tree;
  // Base font from a body{...} rule or the body's own style
  let base: Record<string, string> = {};
  const styleTag = findAllTags(tree, 'style').map((n) => htmlText(n)).join('\n');
  for (const m of styleTag.matchAll(/(?:^|[}\s])body\s*\{([^}]*)\}/g)) base = { ...base, ...parseStyle(m[1]) };
  base = { ...base, ...parseStyle(body.attrs.style) };
  const saved = currentCtx;
  currentCtx = ctx;
  let blocks: Block[];
  try {
    blocks = collectBlocks(body.children, {}, ctx);
  } finally {
    currentCtx = saved;
  }
  const doc: Doc = { title: ctx.title, blocks };
  const family = base['font-family'] ? familyOf(base['font-family']) : undefined;
  if (family) doc.font = family;
  const size = lengthPt(base['font-size']);
  if (size && size > 4 && size < 72) doc.size = size;
  if (ctx.header?.length) doc.header = ctx.header;
  if (ctx.footer?.length) doc.footer = ctx.footer;
  return doc;
}

function findAllTags(node: HNode, tag: string, out: HNode[] = []): HNode[] {
  for (const c of node.children) {
    if (c.tag === tag) out.push(c);
    findAllTags(c, tag, out);
  }
  return out;
}
