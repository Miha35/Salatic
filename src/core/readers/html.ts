import type { Align, Block, Doc, ListItem, TableRow, TextRun } from '../model';
import { trimRuns } from '../model';
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
  if (st['font-family']?.includes('monospace')) s.code = true;
  return s;
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
}

/** Fix whitespace across runs: collapse, drop spaces around line breaks. */
function tidyRuns(runs: TextRun[]): TextRun[] {
  const out: TextRun[] = [];
  let prevEndsSpace = true;
  for (const r of runs) {
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

function collectBlocks(children: HNode[], style: InlineStyle, ctx: Ctx, align?: Align): Block[] {
  const out: Block[] = [];
  let runs: TextRun[] = [];
  const flush = (): void => {
    const r = tidyRuns(runs);
    if (r.length) out.push(align ? { type: 'paragraph', runs: r, align } : { type: 'paragraph', runs: r });
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
      blockFrom(c, style, ctx, out, align);
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
        const h: Block = { type: 'heading', level: +tag[1] as 1, runs };
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
      out.push(...collectBlocks(node.children, { ...st, ...extra }, ctx, align));
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
  const body = findTag(tree, 'body') ?? tree;
  const blocks = collectBlocks(body.children, {}, ctx);
  return { title: ctx.title, blocks };
}
