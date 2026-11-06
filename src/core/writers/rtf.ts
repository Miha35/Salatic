/**
 * Doc → Rich Text Format (RTF 1.9, as read by Word, LibreOffice and WordPad).
 * Non-ASCII text is written as \uN escapes with '?' fallbacks, lists use the
 * Word list table (with \listtext for simple readers), images are embedded as
 * PNG/JPEG blips.
 */
import type { Block, Doc, FontFamily, ParaFormat, TableCell, TextRun } from '../model';
import { maxOf } from '../util';
import { parseSvg } from '../svg-render';
import { metafileToEmf } from './emf';

const FONTS: FontFamily[] = ['sans', 'serif', 'mono', 'calibri'];
/** Sections only exist at the top level of the body. */
const depthZero = (indent: number): boolean => indent === 0;
const FONT_DEF: Record<FontFamily, string> = {
  sans: '{\\f0\\fswiss\\fcharset0\\fprq2 Arial;}',
  serif: '{\\f1\\froman\\fcharset0\\fprq2 Times New Roman;}',
  mono: '{\\f2\\fmodern\\fcharset0\\fprq1 Courier New;}',
  calibri: '{\\f3\\fswiss\\fcharset0\\fprq2 Calibri;}',
};
const BULLETS = ['•', '◦', '▪'];
const tw = (pt: number): number => Math.round(pt * 20);

/** RTF-escape text: \ { } and everything outside printable ASCII. */
export function rtfEscape(s: string): string {
  let out = '';
  for (const ch of s) {
    const cp = ch.codePointAt(0)!;
    if (ch === '\\' || ch === '{' || ch === '}') out += '\\' + ch;
    else if (ch === '\n') out += '\\line ';
    else if (ch === '\t') out += '\\tab ';
    else if (cp === 0xa0) out += '\\~';
    else if (cp >= 0x20 && cp < 0x80) out += ch;
    else if (cp < 0x20) continue;
    else if (cp > 0xffff) {
      const v = cp - 0x10000;
      const hi = 0xd800 + (v >> 10);
      const lo = 0xdc00 + (v & 0x3ff);
      out += `\\u${hi - 65536}?\\u${lo - 65536}?`;
    } else out += `\\u${cp > 32767 ? cp - 65536 : cp}?`;
  }
  return out;
}

class RtfWriter {
  warnings = new Set<string>();
  private colors: string[] = [];
  private lists: string[] = [];
  private overrides: string[] = [];
  private listCount = 0;
  readonly pageW: number;
  readonly pageH: number;
  readonly margins: { top: number; right: number; bottom: number; left: number };
  private font: FontFamily;
  private size: number;

  constructor(
    private doc: Doc,
    opts: { pageSize?: 'auto' | 'a4' | 'letter' },
  ) {
    const auto = (opts.pageSize ?? 'auto') === 'auto';
    if (auto && doc.page) {
      this.pageW = tw(doc.page.width);
      this.pageH = tw(doc.page.height);
    } else {
      const letter = opts.pageSize === 'letter';
      this.pageW = letter ? 12240 : 11906;
      this.pageH = letter ? 15840 : 16838;
    }
    const m = doc.page?.margins;
    this.margins = m ? { top: tw(m.top), right: tw(m.right), bottom: tw(m.bottom), left: tw(m.left) } : { top: 1134, right: 1134, bottom: 1134, left: 1134 };
    this.font = doc.font ?? 'sans';
    this.size = doc.size ?? 11;
  }

  get contentW(): number {
    return this.pageW - this.margins.left - this.margins.right;
  }

  private color(hex: string): number {
    const h = hex.toUpperCase();
    let i = this.colors.indexOf(h);
    if (i < 0) {
      this.colors.push(h);
      i = this.colors.length - 1;
    }
    return i + 1; // index 0 is "auto"
  }

  private charFmt(r: Omit<TextRun, 'text'>, size = this.size): string {
    let s = '';
    const family: FontFamily = r.code && !r.font ? 'mono' : (r.font ?? this.font);
    s += `\\f${FONTS.indexOf(family)}\\fs${Math.round((r.size ?? size) * 2)}`;
    if (r.bold) s += '\\b';
    if (r.italic) s += '\\i';
    if (r.underline) s += '\\ul';
    if (r.strike) s += '\\strike';
    if (r.sup) s += '\\super';
    else if (r.sub) s += '\\sub';
    if (r.color) s += `\\cf${this.color(r.color)}`;
    if (r.highlight) s += `\\chcbpat${this.color(r.highlight)}\\highlight${this.color(r.highlight)}`;
    return s;
  }

  private run(r: TextRun, size?: number): string {
    const fmt = this.charFmt(r, size);
    if (r.field) return `{\\field{\\*\\fldinst{${fmt} ${r.field === 'page' ? 'PAGE' : 'NUMPAGES'} }}{\\fldrslt{${fmt} 1}}}`;
    if (r.note) {
      const body = this.blocks(r.note.length ? r.note : [{ type: 'paragraph', runs: [] }], 0, this.contentW, { note: true });
      return `{${fmt}\\super\\chftn}{\\footnote\\pard\\plain{\\fs${Math.round(this.size * 1.64)}\\super\\chftn} ${body.replace(/\\par\s*$/, '')}}`;
    }
    if (!r.text) return '';
    const text = rtfEscape(r.text);
    if (r.link) return `{\\field{\\*\\fldinst{HYPERLINK "${r.link.replace(/["\\{}]/g, (c) => encodeURIComponent(c))}"}}{\\fldrslt{${fmt}\\ul\\cf${this.color('0563C1')} ${text}}}}`;
    return `{${fmt} ${text}}`;
  }

  private runs(runs: TextRun[], size?: number): string {
    return runs.map((r) => this.run(r, size)).join('');
  }

  private paraFmt(f: ParaFormat | undefined, indent = 0): string {
    let s = '';
    const align = f?.align;
    if (align === 'center') s += '\\qc';
    else if (align === 'right') s += '\\qr';
    else if (align === 'justify') s += '\\qj';
    const left = indent + (f?.indent ? tw(f.indent) : 0);
    if (left) s += `\\li${left}`;
    if (f?.firstLine) s += `\\fi${tw(f.firstLine)}`;
    const d = this.doc.defaults ?? {};
    s += `\\sb${tw(f?.spaceBefore ?? d.spaceBefore ?? 0)}\\sa${tw(f?.spaceAfter ?? d.spaceAfter ?? 6)}`;
    s += `\\sl${Math.round((f?.lineHeight ?? d.lineHeight ?? 1.15) * 240)}\\slmult1`;
    return s;
  }

  private listDef(ordered: boolean, start: number): number {
    const id = ++this.listCount;
    let levels = '';
    for (let lv = 0; lv < 9; lv++) {
      const li = 720 * (lv + 1);
      if (ordered) levels += `{\\listlevel\\levelnfc0\\levelnfcn0\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat${lv === 0 ? start : 1}{\\leveltext\\'02\\'0${lv}.;}{\\levelnumbers\\'01;}\\fi-360\\li${li}\\lin${li}}`;
      else levels += `{\\listlevel\\levelnfc23\\levelnfcn23\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1{\\leveltext\\'01${rtfEscape(BULLETS[lv % 3])};}{\\levelnumbers;}\\f${FONTS.indexOf('sans')}\\fi-360\\li${li}\\lin${li}}`;
    }
    this.lists.push(`{\\list\\listtemplateid${1000 + id}\\listhybrid${levels}{\\listname ;}\\listid${id}}`);
    this.overrides.push(`{\\listoverride\\listid${id}\\listoverridecount0\\ls${id}}`);
    return id;
  }

  private image(b: Extract<Block, { type: 'image' }>, maxW: number): string {
    if (b.mime === 'image/svg+xml') {
      // RTF has no SVG picture type: draw it as an EMF
      const mf = parseSvg(b.data);
      if (mf) b = { ...b, data: metafileToEmf(mf), mime: 'image/emf' };
    }
    const kind = b.mime === 'image/png' ? 'pngblip' : b.mime === 'image/jpeg' ? 'jpegblip' : b.mime === 'image/emf' ? 'emfblip' : b.mime === 'image/wmf' ? 'wmetafile8' : null;
    if (!kind) {
      this.warnings.add('Некоторые изображения пропущены: RTF поддерживает PNG, JPEG, SVG, EMF и WMF.');
      return '';
    }
    const maxH = this.pageH - this.margins.top - this.margins.bottom - 600;
    let w = Math.max(1, b.width) * 15;
    let h = Math.max(1, b.height) * 15;
    const k = Math.min(1, maxW / w, maxH / h);
    w = Math.round(w * k);
    h = Math.round(h * k);
    // RTF carries WMF without the 22-byte "placeable" header
    const data = kind === 'wmetafile8' && b.data[0] === 0xd7 && b.data[1] === 0xcd ? b.data.subarray(22) : b.data;
    let hex = '';
    const parts: string[] = [];
    for (let i = 0; i < data.length; i++) {
      hex += data[i].toString(16).padStart(2, '0');
      if (hex.length >= 128) {
        parts.push(hex);
        hex = '';
      }
    }
    if (hex) parts.push(hex);
    const alt = b.alt ? `{\\*\\picprop{\\sp{\\sn wzDescription}{\\sv ${rtfEscape(b.alt)}}}}` : '';
    return `{\\*\\shppict{\\pict${alt}\\${kind}\\picw${b.width}\\pich${b.height}\\picwgoal${w}\\pichgoal${h}\n${parts.join('\n')}\n}}`;
  }

  private table(b: Extract<Block, { type: 'table' }>, width: number, indent: number): string {
    const cols = maxOf(b.rows.map((r) => r.cells.reduce((s, c) => s + (c.colspan ?? 1), 0)), 1);
    const weights = new Array(cols).fill(1);
    for (const r of b.rows) {
      let c = 0;
      for (const cell of r.cells) {
        const span = cell.colspan ?? 1;
        if (span === 1) {
          const len = cell.blocks.reduce((s, x) => s + (x.type === 'paragraph' || x.type === 'heading' ? x.runs.reduce((a, rr) => a + rr.text.length, 0) : 10), 0);
          weights[c] = Math.max(weights[c], Math.min(40, Math.max(4, len)));
        }
        c += span;
      }
    }
    const total = weights.reduce((a, w) => a + w, 0);
    const widths = weights.map((w) => Math.floor((w / total) * width));
    let s = '';
    for (const r of b.rows) {
      const header = r.cells.length > 0 && r.cells.every((c) => c.header);
      const cells: TableCell[] = [...r.cells];
      let used = cells.reduce((a, x) => a + (x.colspan ?? 1), 0);
      while (used < cols) {
        cells.push({ blocks: [] });
        used++;
      }
      let def = `\\trowd\\trgaph108\\trleft${indent - 108}${header ? '\\trhdr' : ''}`;
      let x = indent;
      let c = 0;
      const spans: number[] = [];
      for (const cell of cells) {
        const span = Math.min(cell.colspan ?? 1, cols - c);
        if (span <= 0) break;
        x += widths.slice(c, c + span).reduce((a, v) => a + v, 0);
        c += span;
        spans.push(span);
        const brd = '\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10';
        def += `${brd}${cell.header ? `\\clcbpat${this.color('F2F2F2')}` : ''}\\cellx${x}`;
      }
      let body = '';
      cells.slice(0, spans.length).forEach((cell, i) => {
        const inner = cell.blocks.length ? cell.blocks : [{ type: 'paragraph', runs: [] } as Block];
        const w = widths.slice(0, spans[i]).reduce((a, v) => a + v, 0);
        let content = this.blocks(inner, 0, Math.max(500, w - 216), { cell: true });
        content = content.replace(/\\par\s*$/, '');
        body += `${content}\\cell `;
      });
      s += `${def}\n${body}\n${def}\\row\n`;
    }
    return s + '\\pard\\plain';
  }

  blocks(blocks: Block[], indent: number, width: number, ctx: { note?: boolean; cell?: boolean; list?: { id: number; level: number; ordered: boolean; n: number } } = {}): string {
    let s = '';
    const intbl = ctx.cell ? '\\intbl' : '';
    const noteSize = this.size * 0.82;
    let list = ctx.list;
    for (const b of blocks) {
      switch (b.type) {
        case 'paragraph':
        case 'heading': {
          const heading = b.type === 'heading' ? b.level : 0;
          const hs = [0, 2, 1.64, 1.36, 1.18, 1.09, 1][heading] * this.size;
          let pre = `\\pard\\plain${intbl}`;
          if (heading) pre += `\\s${heading}\\outlinelevel${heading - 1}\\keepn`;
          let listText = '';
          if (list) {
            const li = 720 * (list.level + 1);
            pre += `\\ls${list.id}\\ilvl${list.level}\\fi-360\\li${li + indent}`;
            const label = list.ordered ? `${list.n}.` : BULLETS[list.level % 3];
            listText = `{\\listtext\\pard\\plain\\f0\\fs${Math.round(this.size * 2)} ${rtfEscape(label)}\\tab}`;
            pre += this.paraFmt({ ...b, indent: undefined, firstLine: undefined, spaceAfter: b.spaceAfter ?? 2 });
            list = undefined;
          } else if (ctx.cell) pre += this.paraFmt({ ...b, spaceBefore: 0, spaceAfter: 0, lineHeight: 1 }, indent);
          else if (ctx.note) pre += this.paraFmt({ ...b, spaceBefore: 0, spaceAfter: 0, lineHeight: 1 }, indent);
          else if (heading) pre += this.paraFmt({ ...b, spaceBefore: b.spaceBefore ?? [0, 18, 15, 12, 10, 10, 8][heading], spaceAfter: b.spaceAfter ?? 6 }, indent);
          else pre += this.paraFmt(b, indent);
          const runs = heading ? b.runs.map((r) => ({ ...r, bold: true })) : b.runs;
          s += `${pre}${listText}${this.runs(runs, heading ? hs : ctx.note ? noteSize : undefined)}\\par\n`;
          break;
        }
        case 'list': {
          const level = ctx.list ? ctx.list.level + 1 : Math.min(8, Math.round(indent / 720));
          const id = this.listDef(b.ordered, b.start ?? 1);
          b.items.forEach((it, i) => {
            const inner = it.blocks.length ? it.blocks : [{ type: 'paragraph', runs: [] } as Block];
            const first = inner[0];
            const info = { id, level, ordered: b.ordered, n: (b.start ?? 1) + i };
            if (first.type !== 'paragraph' && first.type !== 'heading') s += this.blocks([{ type: 'paragraph', runs: [] }], indent, width, { ...ctx, list: info });
            for (const [k, x] of inner.entries()) {
              if (x.type === 'list') s += this.blocks([x], indent, width, { ...ctx, list: { ...info } });
              else s += this.blocks([x], k === 0 ? indent : indent + 720 * (level + 1), width, { ...ctx, list: k === 0 && (x.type === 'paragraph' || x.type === 'heading') ? info : undefined });
            }
          });
          break;
        }
        case 'table':
          if (ctx.cell) {
            // RTF nested tables are rarely supported by readers: flatten the inner table to text
            for (const r of b.rows) s += `\\pard\\plain${intbl}${this.paraFmt({ spaceBefore: 0, spaceAfter: 0 })}{\\f${FONTS.indexOf(this.font)}\\fs${Math.round(this.size * 2)} ${rtfEscape(r.cells.map((c) => c.blocks.map(blockPlain).join(' ')).join(' | '))}}\\par\n`;
          } else s += this.table(b, width - indent, indent) + '\n';
          break;
        case 'image': {
          const pict = this.image(b, width - indent);
          if (pict) s += `\\pard\\plain${intbl}${this.paraFmt({ align: b.width * 15 < width * 0.9 ? undefined : 'center', spaceAfter: 6 }, indent)}${pict}\\par\n`;
          break;
        }
        case 'code':
          s += `\\pard\\plain${intbl}${this.paraFmt({ spaceBefore: 0, spaceAfter: 6, lineHeight: 1 }, indent)}\\cbpat${this.color('F3F4F6')}{\\f2\\fs19 ${rtfEscape(b.text)}}\\par\n`;
          break;
        case 'quote':
          for (const inner of b.blocks) {
            if (inner.type === 'paragraph') s += `\\pard\\plain${intbl}\\brdrl\\brdrs\\brdrw30\\brsp160\\brdrcf${this.color('C8C8CC')}${this.paraFmt(inner, indent + 567)}${this.runs(inner.runs.map((r) => ({ ...r, italic: true, color: r.color ?? '404040' })))}\\par\n`;
            else s += this.blocks([inner], indent + 567, width, ctx);
          }
          break;
        case 'hr':
          s += `\\pard\\plain${intbl}\\brdrb\\brdrs\\brdrw10\\brsp20\\brdrcf${this.color('A0A0A0')}${this.paraFmt({ spaceBefore: 0, spaceAfter: 6 }, indent)}{\\fs4 }\\par\n`;
          break;
        case 'pagebreak':
          if (!ctx.cell && !ctx.note) s += '\\pard\\plain\\page\n';
          break;
        case 'section':
          // a new section with its own columns (only in the body)
          if (!ctx.cell && !ctx.note && depthZero(indent)) s += `\\sect\\sectd${this.sectionProps(b)}\n`;
          break;
      }
    }
    return s;
  }

  private sectionProps(b: Extract<Block, { type: 'section' }>): string {
    const n = Math.max(1, Math.round(b.columns));
    return `${b.start === 'page' ? '\\sbkpage' : '\\sbknone'}\\pgwsxn${this.pageW}\\pghsxn${this.pageH}${n > 1 ? `\\cols${n}\\colsx${Math.round(b.gap * 20)}${b.separator ? '\\linebetcol' : ''}` : ''}`;
  }

  build(): Uint8Array {
    const doc = this.doc;
    const blocks = [...doc.blocks];
    while (blocks.length && (blocks[blocks.length - 1].type === 'pagebreak' || blocks[blocks.length - 1].type === 'section')) blocks.pop();
    while (blocks.length && blocks[0].type === 'pagebreak') blocks.shift();
    // a leading section sets the first section's columns
    const lead = blocks[0]?.type === 'section' ? (blocks.shift() as Extract<Block, { type: 'section' }>) : undefined;
    const body = this.blocks(blocks, 0, this.contentW);
    const header = doc.header?.length ? `{\\header ${this.blocks(doc.header, 0, this.contentW)}}` : '';
    const footer = doc.footer?.length ? `{\\footer ${this.blocks(doc.footer, 0, this.contentW)}}` : '';
    const fs = Math.round(this.size * 2);
    const f = FONTS.indexOf(this.font);
    const stylesheet =
      `{\\stylesheet{\\ql\\f${f}\\fs${fs} Normal;}` +
      [1, 2, 3, 4, 5, 6].map((lv) => `{\\s${lv}\\ql\\keepn\\outlinelevel${lv - 1}\\sbasedon0\\snext0\\b\\f${f}\\fs${Math.round(fs * [0, 2, 1.64, 1.36, 1.18, 1.09, 1][lv])} heading ${lv};}`).join('') +
      '}';
    const colors = `{\\colortbl;${this.colors.map((c) => `\\red${parseInt(c.slice(0, 2), 16)}\\green${parseInt(c.slice(2, 4), 16)}\\blue${parseInt(c.slice(4, 6), 16)};`).join('')}}`;
    const lists = this.lists.length ? `{\\*\\listtable${this.lists.join('')}}{\\*\\listoverridetable${this.overrides.join('')}}` : '';
    const info = `{\\info${doc.title ? `{\\title ${rtfEscape(doc.title)}}` : ''}{\\*\\company Salatic}}`;
    const m = this.margins;
    const rtf =
      `{\\rtf1\\ansi\\ansicpg1252\\deff${f}\\uc1\\deflang1049\n` +
      `{\\fonttbl${FONTS.map((x) => FONT_DEF[x]).join('')}}\n${colors}\n${stylesheet}\n${lists}\n${info}\n` +
      `\\paperw${this.pageW}\\paperh${this.pageH}\\margl${m.left}\\margr${m.right}\\margt${m.top}\\margb${m.bottom}${this.pageW > this.pageH ? '\\landscape' : ''}\\viewkind4\n` +
      `\\sectd\\pgwsxn${this.pageW}\\pghsxn${this.pageH}${lead ? this.sectionProps(lead).replace(/\\sbk\w+/, '') : ''}\\headery709\\footery709${header}${footer}\n${body}}`;
    // RTF is 7-bit: every non-ASCII character was escaped above
    const out = new Uint8Array(rtf.length);
    for (let i = 0; i < rtf.length; i++) out[i] = rtf.charCodeAt(i) & 0x7f;
    return out;
  }
}

function blockPlain(b: Block): string {
  if (b.type === 'paragraph' || b.type === 'heading') return b.runs.map((r) => r.text).join('');
  if (b.type === 'code') return b.text;
  return '';
}

export function docToRtf(doc: Doc, opts: { pageSize?: 'auto' | 'a4' | 'letter' } = {}): { data: Uint8Array; warnings: string[] } {
  const w = new RtfWriter(doc, opts);
  const data = w.build();
  return { data, warnings: [...w.warnings] };
}
