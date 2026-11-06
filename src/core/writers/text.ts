/** Doc → plain text, Markdown, HTML. */
import { metafileToSvgBytes } from '../metafile';
import type { Align, Block, Doc, FontFamily, ParaFormat, TextRun } from '../model';
import { blockText } from '../model';
import { bytesToBase64, escapeHtml, maxOf } from '../util';

/** Footnotes collected while writing; references are numbered in reading order. */
/** Image as a data URI; EMF/WMF pictures (which browsers cannot show) become SVG. */
function dataUri(b: Extract<Block, { type: 'image' }>): string {
  if (b.mime === 'image/emf' || b.mime === 'image/wmf') {
    const svg = metafileToSvgBytes(b.data);
    if (svg) return `data:image/svg+xml;base64,${bytesToBase64(svg)}`;
  }
  return `data:${b.mime};base64,${bytesToBase64(b.data)}`;
}

class Notes {
  list: Block[][] = [];
  add(blocks: Block[]): number {
    this.list.push(blocks);
    return this.list.length;
  }
}

// ---------------------------------------------------------------------------
// Plain text

function plainRuns(runs: TextRun[], notes: Notes): string {
  let s = '';
  for (const r of runs) {
    if (r.note) s += `[${notes.add(r.note)}]`;
    else if (r.field) s += '#';
    else s += r.text;
  }
  return s;
}

export function docToText(doc: Doc): string {
  const notes = new Notes();
  const emitAll = (blocks: Block[]): string => {
    const out: string[] = [];
    const emit = (list: Block[], indent: string): void => {
      for (const b of list) {
        switch (b.type) {
          case 'paragraph':
          case 'heading': {
            const t = plainRuns(b.runs, notes).trim();
            if (t) out.push(indent + t.replace(/\n/g, '\n' + indent));
            break;
          }
          case 'list':
            b.items.forEach((it, i) => {
              const marker = b.ordered ? `${(b.start ?? 1) + i}. ` : '• ';
              const saved = out.length;
              emit(it.blocks, '');
              const lines = out.splice(saved);
              const pad = ' '.repeat(marker.length);
              const body = lines.join('\n').split('\n');
              out.push(indent + marker + (body[0] ?? '') + body.slice(1).map((l) => '\n' + indent + (l ? pad + l : '')).join(''));
            });
            break;
          case 'table':
            out.push(
              b.rows
                .map((r) => indent + r.cells.map((c) => c.blocks.map((x) => (x.type === 'paragraph' || x.type === 'heading' ? plainRuns(x.runs, notes) : blockText(x))).join(' ').replace(/\s+/g, ' ').trim()).join('\t'))
                .join('\n'),
            );
            break;
          case 'code':
            out.push(b.text.split('\n').map((l) => indent + l).join('\n'));
            break;
          case 'quote': {
            const saved = out.length;
            emit(b.blocks, '');
            const inner = out.splice(saved).join('\n\n');
            out.push(inner.split('\n').map((l) => indent + '> ' + l).join('\n'));
            break;
          }
          case 'hr':
            out.push(indent + '———');
            break;
          case 'pagebreak':
            out.push('\f');
            break;
          case 'image':
            if (b.alt) out.push(`${indent}[${b.alt}]`);
            break;
        }
      }
    };
    emit(blocks, '');
    return out.join('\n\n').replace(/\n\n\f\n\n/g, '\n\n');
  };
  let text = emitAll(doc.blocks);
  if (notes.list.length) {
    const parts: string[] = [];
    for (let i = 0; i < notes.list.length; i++) parts.push(`[${i + 1}] ${emitAll(notes.list[i]).replace(/\n+/g, ' ')}`);
    text += '\n\n———\n' + parts.join('\n');
  }
  return text + '\n';
}

// ---------------------------------------------------------------------------
// Markdown

function mdEscape(s: string): string {
  return s.replace(/([\\`*_[\]<>#|~])/g, '\\$1').replace(/^(\s*)([-+]|\d+[.)])(\s)/gm, '$1\\$2$3');
}

function mdRuns(runs: TextRun[], notes: Notes, inTable = false): string {
  let out = '';
  for (const r of runs) {
    if (r.note) {
      out += `[^${notes.add(r.note)}]`;
      continue;
    }
    if (r.field) {
      out += '#';
      continue;
    }
    if (!r.text) continue;
    if (r.code) {
      const t = r.text.replace(/\n/g, ' ');
      const fence = t.includes('`') ? '``' : '`';
      out += `${fence}${t.includes('`') ? ' ' + t + ' ' : t}${fence}`;
      continue;
    }
    // Keep surrounding whitespace outside of emphasis markers
    const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(r.text)!;
    let core = mdEscape(m[2]);
    if (inTable) core = core.replace(/\n/g, '<br>');
    else core = core.replace(/\n/g, '  \n');
    if (core) {
      if (r.strike) core = `~~${core}~~`;
      if (r.italic) core = `*${core}*`;
      if (r.bold) core = `**${core}**`;
      if (r.highlight) core = `<mark>${core}</mark>`;
      if (r.sup) core = `<sup>${core}</sup>`;
      if (r.sub) core = `<sub>${core}</sub>`;
      if (r.link) core = `[${core}](${r.link.replace(/[()\s]/g, encodeURIComponent)})`;
    }
    out += m[1] + core + m[3];
  }
  return out;
}

export function docToMarkdown(doc: Doc): string {
  const notes = new Notes();
  const render = (blocks: Block[]): string => {
    const out: string[] = [];
    const emit = (list: Block[], prefix: string): void => {
      for (const b of list) {
        switch (b.type) {
          case 'heading': {
            const t = mdRuns(b.runs, notes).replace(/\s*\n\s*/g, ' ').trim();
            if (t) out.push(prefix + '#'.repeat(b.level) + ' ' + t);
            break;
          }
          case 'paragraph': {
            const t = mdRuns(b.runs, notes).trim();
            if (t) out.push(t.split('\n').map((l) => prefix + l).join('\n'));
            break;
          }
          case 'list':
            out.push(
              b.items
                .map((it, i) => {
                  const marker = b.ordered ? `${(b.start ?? 1) + i}. ` : '- ';
                  const saved = out.length;
                  emit(it.blocks, '');
                  const body = out.splice(saved).join('\n');
                  const pad = ' '.repeat(marker.length);
                  return body
                    .split('\n')
                    .map((l, k) => prefix + (k === 0 ? marker + l : l ? pad + l : ''))
                    .join('\n');
                })
                .join('\n'),
            );
            break;
          case 'table': {
            const width = maxOf(b.rows.map((r) => r.cells.reduce((s, c) => s + (c.colspan ?? 1), 0)), 0);
            if (!width) break;
            const cellText = (cb: Block[]): string =>
              cb
                .map((x) => (x.type === 'paragraph' || x.type === 'heading' ? mdRuns(x.runs, notes, true) : mdEscape(blockText(x))))
                .join('<br>')
                .replace(/\|/g, '\\|')
                .replace(/\n/g, '<br>')
                .trim();
            const rows = b.rows.map((r) => {
              const cells: string[] = [];
              for (const c of r.cells) {
                cells.push(cellText(c.blocks));
                for (let k = 1; k < (c.colspan ?? 1); k++) cells.push('');
              }
              while (cells.length < width) cells.push('');
              return '| ' + cells.join(' | ') + ' |';
            });
            rows.splice(1, 0, '|' + ' --- |'.repeat(width));
            out.push(rows.map((r) => prefix + r).join('\n'));
            break;
          }
          case 'code': {
            const fence = b.text.includes('```') ? '~~~' : '```';
            out.push([fence, ...b.text.split('\n'), fence].map((l) => prefix + l).join('\n'));
            break;
          }
          case 'quote':
            emit(b.blocks, prefix + '> ');
            break;
          case 'hr':
          case 'pagebreak':
            out.push(prefix + '---');
            break;
          case 'image':
            out.push(`${prefix}![${mdEscape(b.alt ?? '')}](${dataUri(b)})`);
            break;
        }
      }
    };
    emit(blocks, '');
    return out.join('\n\n');
  };
  let md = render(doc.blocks);
  const defs: string[] = [];
  for (let i = 0; i < notes.list.length; i++) {
    const body = render(notes.list[i]).split('\n');
    defs.push(`[^${i + 1}]: ${body[0] ?? ''}` + body.slice(1).map((l) => '\n' + (l ? '    ' + l : '')).join(''));
  }
  if (defs.length) md += '\n\n' + defs.join('\n');
  return md + '\n';
}

// ---------------------------------------------------------------------------
// HTML

export const CSS_FONT: Record<FontFamily, string> = {
  sans: 'Arial, "Liberation Sans", Helvetica, sans-serif',
  serif: '"Times New Roman", "Liberation Serif", Times, serif',
  mono: '"Courier New", "Liberation Mono", monospace',
  calibri: 'Calibri, Carlito, "Segoe UI", sans-serif',
};

const num = (n: number): string => String(Math.round(n * 100) / 100);

/** Markup flavour: HTML for pages, XHTML (self-closed void tags, no named entities) for EPUB. */
export interface HtmlCtx {
  xhtml: boolean;
  /** Image source for a block; defaults to an inline data URI. */
  imageSrc?: (b: Extract<Block, { type: 'image' }>) => string;
  /** Keep source data (Office chart parts) in data attributes, so an editing round trip loses nothing. */
  sources?: boolean;
}
const HTML: HtmlCtx = { xhtml: false };

function htmlRuns(runs: TextRun[], notes: Notes, ctx: HtmlCtx = HTML): string {
  const br = ctx.xhtml ? '<br/>' : '<br>';
  let s = '';
  for (const r of runs) {
    if (r.note) {
      const n = notes.add(r.note);
      s += `<sup class="footnote-ref"><a href="#fn${n}" id="fnref${n}"${ctx.xhtml ? ' epub:type="noteref"' : ''}>${n}</a></sup>`;
      continue;
    }
    if (r.field) {
      s += `<span data-field="${r.field}">#</span>`;
      continue;
    }
    let t = escapeHtml(r.text).replace(/\n/g, br).replace(/\t/g, ctx.xhtml ? '&#8195;' : '&emsp;');
    if (!t) continue;
    if (r.code) t = `<code>${t}</code>`;
    if (r.bold) t = `<strong>${t}</strong>`;
    if (r.italic) t = `<em>${t}</em>`;
    if (r.underline) t = `<u>${t}</u>`;
    if (r.strike) t = `<s>${t}</s>`;
    if (r.sup) t = `<sup>${t}</sup>`;
    if (r.sub) t = `<sub>${t}</sub>`;
    const css: string[] = [];
    if (r.color) css.push(`color:#${r.color}`);
    if (r.highlight) css.push(`background-color:#${r.highlight}`);
    if (r.font && !r.code) css.push(`font-family:${CSS_FONT[r.font].replace(/"/g, "'")}`);
    if (r.size) css.push(`font-size:${num(r.size)}pt`);
    if (css.length) t = `<span style="${css.join(';')}">${t}</span>`;
    if (r.link) t = `<a href="${escapeHtml(r.link)}">${t}</a>`;
    s += t;
  }
  return s;
}

function paraStyle(f: ParaFormat): string {
  const css: string[] = [];
  if (f.align && f.align !== 'left') css.push(`text-align:${f.align}`);
  if (f.indent) css.push(`margin-left:${num(f.indent)}pt`);
  if (f.firstLine) css.push(`text-indent:${num(f.firstLine)}pt`);
  if (f.spaceBefore !== undefined) css.push(`margin-top:${num(f.spaceBefore)}pt`);
  if (f.spaceAfter !== undefined) css.push(`margin-bottom:${num(f.spaceAfter)}pt`);
  // Single spacing is ~1.15 of the font size for common fonts
  if (f.lineHeight !== undefined) css.push(`line-height:${num(f.lineHeight * 1.15)}`);
  return css.length ? ` style="${css.join(';')}"` : '';
}

const alignAttr = (a?: Align): string => (a && a !== 'left' ? ` style="text-align:${a}"` : '');

function blocksHtml(blocks: Block[], notes: Notes, ctx: HtmlCtx = HTML): string {
  const br = ctx.xhtml ? '<br/>' : '<br>';
  const hr = ctx.xhtml ? '<hr/>' : '<hr>';
  let s = '';
  let sectionOpen = false;
  for (const b of blocks) {
    switch (b.type) {
      case 'section': {
        // e-book readers paginate with CSS columns themselves: sections flow as one column there
        if (ctx.xhtml) break;
        // a section with columns (or starting a page) wraps the blocks up to the next section
        if (sectionOpen) s += '</div>\n';
        sectionOpen = false;
        const n = Math.max(1, Math.round(b.columns));
        if (n > 1 || b.start === 'page') {
          const css = n > 1 ? ` style="column-count:${n};column-gap:${num(b.gap)}pt${b.separator ? ';column-rule:0.5pt solid currentColor' : ''}"` : '';
          // unequal widths and an own page travel in data attributes (CSS columns are always equal)
          const widths = n > 1 && b.widths?.length === n ? ` data-widths="${b.widths.map(num).join(',')}" data-gaps="${(b.gaps ?? []).map(num).join(',')}"` : '';
          const pg = b.page ? ` data-page="${[b.page.width, b.page.height, b.page.margins.top, b.page.margins.right, b.page.margins.bottom, b.page.margins.left].map(num).join(',')}"` : '';
          s += `<div class="doc-section" data-columns="${n}" data-gap="${num(b.gap)}" data-start="${b.start}"${b.separator ? ' data-separator="1"' : ''}${widths}${pg}${css}>\n`;
          sectionOpen = true;
        }
        break;
      }
      case 'paragraph':
        s += `<p${paraStyle(b)}>${htmlRuns(b.runs, notes, ctx) || br}</p>\n`;
        break;
      case 'heading':
        s += `<h${b.level}${paraStyle(b)}>${htmlRuns(b.runs, notes, ctx)}</h${b.level}>\n`;
        break;
      case 'list': {
        const tag = b.ordered ? 'ol' : 'ul';
        const start = b.ordered && b.start && b.start !== 1 ? ` start="${b.start}"` : '';
        s += `<${tag}${start}>\n`;
        for (const it of b.items) {
          // A single paragraph item is rendered inline to keep the markup compact
          const [first, ...rest] = it.blocks;
          if (first?.type === 'paragraph') s += `<li>${htmlRuns(first.runs, notes, ctx)}${rest.length ? '\n' + blocksHtml(rest, notes, ctx) : ''}</li>\n`;
          else s += `<li>${blocksHtml(it.blocks, notes, ctx)}</li>\n`;
        }
        s += `</${tag}>\n`;
        break;
      }
      case 'table':
        s += '<table>\n';
        for (const r of b.rows) {
          s += '<tr>';
          for (const c of r.cells) {
            const tag = c.header ? 'th' : 'td';
            const span = c.colspan && c.colspan > 1 ? ` colspan="${c.colspan}"` : '';
            const only = c.blocks.length === 1 && c.blocks[0].type === 'paragraph' ? c.blocks[0] : null;
            const align = only ? alignAttr(only.align) : '';
            s += `<${tag}${span}${align}>${only ? htmlRuns(only.runs, notes, ctx) : blocksHtml(c.blocks, notes, ctx)}</${tag}>`;
          }
          s += '</tr>\n';
        }
        s += '</table>\n';
        break;
      case 'image': {
        // a floating picture: CSS float to its side (text around it), its placement kept in data-float
        const f = b.float;
        let css = '';
        let pAttr = '';
        if (f) {
          const [t, r, bt, l] = (f.dist ?? [0, 9, 0, 9]).map(num);
          const side = f.h.align === 'right' || f.h.align === 'outside' ? 'right' : f.h.align === 'center' ? 'center' : f.h.align ? 'left' : (f.h.offset ?? 0) > 200 ? 'right' : 'left';
          if (f.wrap === 'square' && side !== 'center') css = ` style="float:${side};margin:${t}pt ${r}pt ${bt}pt ${l}pt"`;
          else if (side !== 'left') pAttr = ` style="text-align:${side}"`;
          pAttr += ' class="doc-float"';
        }
        const data = f ? ` data-float="${escapeHtml(JSON.stringify(f))}"` : '';
        s += `<p${pAttr}><img src="${ctx.imageSrc ? escapeHtml(ctx.imageSrc(b)) : dataUri(b)}" width="${b.width}" height="${b.height}" alt="${escapeHtml(b.alt ?? '')}"${css}${data}${ctx.sources && b.chart ? ` data-chart="${escapeHtml(JSON.stringify(b.chart))}"` : ''}${ctx.xhtml ? '/' : ''}></p>\n`;
        break;
      }
      case 'code':
        s += `<pre><code>${escapeHtml(b.text)}</code></pre>\n`;
        break;
      case 'quote':
        s += `<blockquote>\n${blocksHtml(b.blocks, notes, ctx)}</blockquote>\n`;
        break;
      case 'hr':
        s += `${hr}\n`;
        break;
      case 'pagebreak':
        s += `<hr class="page-break"${ctx.xhtml ? '/' : ''}>\n`;
        break;
    }
  }
  if (sectionOpen) s += '</div>\n';
  return s;
}

function notesHtml(notes: Notes, ctx: HtmlCtx = HTML): string {
  if (!notes.list.length) return '';
  let s = `<section class="footnotes"${ctx.xhtml ? ' epub:type="footnotes"' : ''}>\n${ctx.xhtml ? '<hr/>' : '<hr>'}\n<ol>\n`;
  // Notes may reference further notes: iterate while the list grows
  for (let i = 0; i < notes.list.length; i++) {
    const body = blocksHtml(notes.list[i], notes, ctx).trim();
    s += `<li id="fn${i + 1}"${ctx.xhtml ? ' epub:type="footnote"' : ''}>${body}</li>\n`;
  }
  return s + '</ol>\n</section>\n';
}

/** XHTML body of a chapter (blocks + its footnotes) for EPUB. */
export function blocksToXhtml(blocks: Block[], imageSrc: HtmlCtx['imageSrc']): string {
  const notes = new Notes();
  const ctx: HtmlCtx = { xhtml: true, imageSrc };
  return blocksHtml(blocks, notes, ctx) + notesHtml(notes, ctx);
}

/** Blocks only (no header, footer or footnotes) – used where a bare fragment is needed. */
export function blocksToHtml(blocks: Block[], opts: { sources?: boolean } = {}): string {
  const notes = new Notes();
  return blocksHtml(blocks, notes, { ...HTML, ...opts }) + notesHtml(notes);
}

/** Editable/printable body: header, content, footnotes, footer. */
export function docToHtmlFragment(doc: Doc, opts: { sources?: boolean } = {}): string {
  const notes = new Notes();
  const ctx: HtmlCtx = { ...HTML, ...opts };
  let s = '';
  if (doc.header?.length) s += `<header class="doc-header">\n${blocksHtml(doc.header, notes, ctx)}</header>\n`;
  s += blocksHtml(doc.blocks, notes, ctx);
  s += notesHtml(notes);
  if (doc.footer?.length) s += `<footer class="doc-footer">\n${blocksHtml(doc.footer, new Notes())}</footer>\n`;
  return s;
}

const HTML_CSS = `
:root{color-scheme:light}
body{font-family:Arial,"Liberation Sans","Helvetica Neue",sans-serif;font-size:16px;line-height:1.55;color:#1d1d1f;max-width:820px;margin:40px auto;padding:0 20px}
h1,h2,h3,h4,h5,h6{line-height:1.25;margin:1.4em 0 .5em}
p{margin:.6em 0}
img{max-width:100%;height:auto}
table{border-collapse:collapse;margin:1em 0;width:100%}
td,th{border:1px solid #c8c8cc;padding:6px 10px;vertical-align:top;text-align:left}
th{background:#f2f2f5}
pre{background:#f5f5f7;padding:12px 14px;border-radius:6px;overflow:auto}
code{font-family:Consolas,"Liberation Mono",monospace;font-size:.92em}
blockquote{margin:1em 0;padding:.2em 1em;border-left:4px solid #d0d0d6;color:#444}
.doc-header,.doc-footer{color:#666;font-size:.85em}
.doc-header{border-bottom:1px solid #e3e3e8;margin-bottom:1.5em}
.doc-footer{border-top:1px solid #e3e3e8;margin-top:2em}
.footnotes{font-size:.9em;color:#333}
.footnotes hr{border:0;border-top:1px solid #ccc;width:30%;margin:2em 0 .5em}
hr.page-break{border:0;page-break-after:always;break-after:page;margin:2em 0;border-top:1px dashed #bbb}
@media print{hr.page-break{border:0;margin:0}}
`.trim();

export function docToHtml(doc: Doc): string {
  const title = escapeHtml(doc.title ?? 'Документ');
  const base: string[] = [];
  if (doc.font) base.push(`font-family:${CSS_FONT[doc.font]}`);
  if (doc.size) base.push(`font-size:${num(doc.size)}pt`);
  const bodyCss = base.length ? `\nbody{${base.join(';')}}` : '';
  return `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<style>
${HTML_CSS}${bodyCss}
</style>
</head>
<body>
${docToHtmlFragment(doc)}</body>
</html>
`;
}
