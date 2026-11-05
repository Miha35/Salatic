/** Doc → plain text, Markdown, HTML. */
import type { Align, Block, Doc, TextRun } from '../model';
import { blockText, runsText } from '../model';
import { bytesToBase64, escapeHtml, maxOf } from '../util';

// ---------------------------------------------------------------------------
// Plain text

export function docToText(doc: Doc): string {
  const out: string[] = [];
  const emit = (blocks: Block[], indent: string): void => {
    for (const b of blocks) {
      switch (b.type) {
        case 'paragraph':
        case 'heading': {
          const t = runsText(b.runs).trim();
          if (t) out.push(indent + t.replace(/\n/g, '\n' + indent));
          break;
        }
        case 'list':
          b.items.forEach((it, i) => {
            const marker = b.ordered ? `${(b.start ?? 1) + i}. ` : '• ';
            const lines: string[] = [];
            const saved = out.length;
            emit(it.blocks, '');
            lines.push(...out.splice(saved));
            const pad = ' '.repeat(marker.length);
            const body = lines.join('\n').split('\n');
            out.push(indent + marker + (body[0] ?? '') + body.slice(1).map((l) => '\n' + indent + (l ? pad + l : '')).join(''));
          });
          break;
        case 'table':
          out.push(b.rows.map((r) => indent + r.cells.map((c) => c.blocks.map(blockText).join(' ').replace(/\s+/g, ' ').trim()).join('\t')).join('\n'));
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
  emit(doc.blocks, '');
  return out.join('\n\n').replace(/\n\n\f\n\n/g, '\n\n') + '\n';
}

// ---------------------------------------------------------------------------
// Markdown

function mdEscape(s: string): string {
  return s.replace(/([\\`*_[\]<>#|~])/g, '\\$1').replace(/^(\s*)([-+]|\d+[.)])(\s)/gm, '$1\\$2$3');
}

function mdRuns(runs: TextRun[], inTable = false): string {
  let out = '';
  for (const r of runs) {
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
      if (r.sup) core = `<sup>${core}</sup>`;
      if (r.sub) core = `<sub>${core}</sub>`;
      if (r.link) core = `[${core}](${r.link.replace(/[()\s]/g, encodeURIComponent)})`;
    }
    out += m[1] + core + m[3];
  }
  return out;
}

export function docToMarkdown(doc: Doc): string {
  const out: string[] = [];
  const emit = (blocks: Block[], prefix: string): void => {
    for (const b of blocks) {
      switch (b.type) {
        case 'heading': {
          const t = mdRuns(b.runs).replace(/\s*\n\s*/g, ' ').trim();
          if (t) out.push(prefix + '#'.repeat(b.level) + ' ' + t);
          break;
        }
        case 'paragraph': {
          const t = mdRuns(b.runs).trim();
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
          const cellText = (blocks: Block[]): string =>
            blocks
              .map((x) => (x.type === 'paragraph' || x.type === 'heading' ? mdRuns(x.runs, true) : mdEscape(blockText(x))))
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
          out.push(`${prefix}![${mdEscape(b.alt ?? '')}](data:${b.mime};base64,${bytesToBase64(b.data)})`);
          break;
      }
    }
  };
  emit(doc.blocks, '');
  return out.join('\n\n') + '\n';
}

// ---------------------------------------------------------------------------
// HTML

function htmlRuns(runs: TextRun[]): string {
  let s = '';
  for (const r of runs) {
    let t = escapeHtml(r.text).replace(/\n/g, '<br>').replace(/\t/g, '&emsp;');
    if (!t) continue;
    if (r.code) t = `<code>${t}</code>`;
    if (r.bold) t = `<strong>${t}</strong>`;
    if (r.italic) t = `<em>${t}</em>`;
    if (r.underline) t = `<u>${t}</u>`;
    if (r.strike) t = `<s>${t}</s>`;
    if (r.sup) t = `<sup>${t}</sup>`;
    if (r.sub) t = `<sub>${t}</sub>`;
    if (r.color) t = `<span style="color:#${r.color}">${t}</span>`;
    if (r.link) t = `<a href="${escapeHtml(r.link)}">${t}</a>`;
    s += t;
  }
  return s;
}

const alignAttr = (a?: Align): string => (a && a !== 'left' ? ` style="text-align:${a}"` : '');

export function blocksToHtml(blocks: Block[]): string {
  let s = '';
  for (const b of blocks) {
    switch (b.type) {
      case 'paragraph':
        s += `<p${alignAttr(b.align)}>${htmlRuns(b.runs) || '<br>'}</p>\n`;
        break;
      case 'heading':
        s += `<h${b.level}${alignAttr(b.align)}>${htmlRuns(b.runs)}</h${b.level}>\n`;
        break;
      case 'list': {
        const tag = b.ordered ? 'ol' : 'ul';
        const start = b.ordered && b.start && b.start !== 1 ? ` start="${b.start}"` : '';
        s += `<${tag}${start}>\n`;
        for (const it of b.items) {
          // A single paragraph item is rendered inline to keep the markup compact
          const [first, ...rest] = it.blocks;
          if (first?.type === 'paragraph') s += `<li>${htmlRuns(first.runs)}${rest.length ? '\n' + blocksToHtml(rest) : ''}</li>\n`;
          else s += `<li>${blocksToHtml(it.blocks)}</li>\n`;
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
            s += `<${tag}${span}${align}>${only ? htmlRuns(only.runs) : blocksToHtml(c.blocks)}</${tag}>`;
          }
          s += '</tr>\n';
        }
        s += '</table>\n';
        break;
      case 'image':
        s += `<p><img src="data:${b.mime};base64,${bytesToBase64(b.data)}" width="${b.width}" height="${b.height}" alt="${escapeHtml(b.alt ?? '')}"></p>\n`;
        break;
      case 'code':
        s += `<pre><code>${escapeHtml(b.text)}</code></pre>\n`;
        break;
      case 'quote':
        s += `<blockquote>\n${blocksToHtml(b.blocks)}</blockquote>\n`;
        break;
      case 'hr':
        s += '<hr>\n';
        break;
      case 'pagebreak':
        s += '<hr class="page-break">\n';
        break;
    }
  }
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
code{font-family:Consolas,"DejaVu Sans Mono",monospace;font-size:.92em}
blockquote{margin:1em 0;padding:.2em 1em;border-left:4px solid #d0d0d6;color:#444}
hr.page-break{border:0;page-break-after:always;break-after:page;margin:2em 0;border-top:1px dashed #bbb}
@media print{hr.page-break{border:0;margin:0}}
`.trim();

export function docToHtml(doc: Doc): string {
  const title = escapeHtml(doc.title ?? 'Документ');
  return `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<style>
${HTML_CSS}
</style>
</head>
<body>
${blocksToHtml(doc.blocks)}</body>
</html>
`;
}
