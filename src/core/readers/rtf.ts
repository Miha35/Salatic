/**
 * Rich Text Format (RTF 1.9.1) → Doc. A single-pass tokenizer drives a group
 * stack: character formatting is group-scoped, paragraph formatting lives in the
 * current output "sink" (body, footnote, header, footer). Handles code pages and
 * \u escapes, fonts, colours, styles, Word lists (\listtable) and legacy \pn
 * lists, tables, fields (hyperlinks, page numbers), footnotes, headers/footers,
 * PNG/JPEG pictures and page breaks.
 */
import { familyOf, normalizeRuns, trimRuns, type Align, type Block, type Doc, type FontFamily, type ParaFormat, type TableCell, type TableRow, type TextRun } from '../model';
import { ConvertError, imageSize } from '../util';
import { ListNester } from './lists';

interface CharState {
  b: boolean;
  i: boolean;
  ul: boolean;
  strike: boolean;
  sup: boolean;
  sub: boolean;
  hidden: boolean;
  fs: number; // half-points
  f: number;
  cf: number;
  hl: number;
  uc: number;
  link?: string;
  field?: 'page' | 'pages';
}

type Dest =
  | 'body'
  | 'skip'
  | 'fonttbl'
  | 'colortbl'
  | 'stylesheet'
  | 'listtable'
  | 'listoverridetable'
  | 'info'
  | 'title'
  | 'pict'
  | 'fldinst'
  | 'shpinst'
  | 'pn'
  | 'sn'
  | 'sv'
  | 'star';

interface Group {
  ch: CharState;
  dest: Dest;
  sink: Sink;
  /** Set on the group that opened a new sink (footnote, header, footer). */
  opens?: 'footnote' | 'header' | 'footer';
  fieldInst?: string;
  isField?: boolean;
  text?: string[];
  pict?: Pict;
  font?: { id: number; family?: FontFamily; charset?: number; name: string };
  style?: { id: number; name: string; outline?: number; char?: boolean };
  list?: { id: number; levels: { nfc: number; start: number }[] };
  level?: { nfc: number; start: number };
  override?: { listId: number; ls: number };
  spName?: string;
  alt?: string;
}

interface Pict {
  kind?: 'png' | 'jpeg' | 'other';
  w?: number;
  h?: number;
  wGoal?: number;
  hGoal?: number;
  scaleX: number;
  scaleY: number;
  hex: string[];
  bin?: Uint8Array;
  alt?: string;
}

interface ParaState {
  align?: Align;
  li: number;
  fi: number;
  sb?: number;
  sa?: number;
  sl?: number;
  slmult: boolean;
  style: number;
  intbl: boolean;
  ls: number;
  ilvl: number;
  outline?: number;
  pagebb: boolean;
  pnBullet?: boolean;
  pnOrdered?: boolean;
  shaded?: boolean;
  brdrl?: boolean;
  brdrb?: boolean;
}

const freshPara = (): ParaState => ({ li: 0, fi: 0, slmult: false, style: 0, intbl: false, ls: 0, ilvl: 0, pagebb: false });

interface TableState {
  rows: TableRow[];
  cells: TableCell[];
  cellBlocks: Block[];
  header: boolean;
  merges: ('first' | 'cont' | null)[];
  nextMerge: 'first' | 'cont' | null;
}

class Sink {
  blocks: Block[] = [];
  nester: ListNester;
  runs: TextRun[] = [];
  para: ParaState = freshPara();
  table: TableState | null = null;
  constructor(readonly kind: 'body' | 'footnote' | 'header' | 'footer') {
    this.nester = new ListNester(this.blocks);
  }
}

const CHARSET_CP: Record<number, string> = {
  0: 'windows-1252',
  77: 'macintosh',
  128: 'shift_jis',
  129: 'euc-kr',
  134: 'gbk',
  136: 'big5',
  161: 'windows-1253',
  162: 'windows-1254',
  163: 'windows-1258',
  177: 'windows-1255',
  178: 'windows-1256',
  186: 'windows-1257',
  204: 'windows-1251',
  222: 'windows-874',
  238: 'windows-1250',
};
const ANSI_CP: Record<number, string> = { 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 10000: 'macintosh', 65001: 'utf-8' };

const decoders = new Map<string, TextDecoder>();
function decoderFor(label: string): TextDecoder {
  let d = decoders.get(label);
  if (!d) {
    try {
      d = new TextDecoder(label);
    } catch {
      d = new TextDecoder('windows-1252');
    }
    decoders.set(label, d);
  }
  return d;
}

/** Destinations whose content is never visible text. */
const SKIP = new Set([
  'filetbl', 'revtbl', 'rsidtbl', 'generator', 'xmlnstbl', 'listtext', 'pntext', 'pntxta', 'pntxtb', 'themedata', 'colorschememapping',
  'datastore', 'latentstyles', 'pgdsctbl', 'nonshppict', 'objdata', 'objclass', 'objname', 'bkmkstart', 'bkmkend', 'fldtype', 'datafield',
  'formfield', 'atnid', 'atnauthor', 'annotation', 'atnref', 'atndate', 'atrfstart', 'atrfend', 'xe', 'tc', 'txe', 'rxe', 'mmathPr', 'listpicture',
  'shprslt', 'docvar', 'userprops', 'passwordhash', 'protusertbl', 'wgrffmtfilter', 'ftnsep', 'ftnsepc', 'ftncn', 'aftnsep', 'aftnsepc',
  'aftncn', 'template', 'nonesttables', 'nesttableprops', 'fchars', 'lchars', 'pnseclvl', 'background', 'do', 'stylesheet_', 'oldcprops',
  'oldpprops', 'oldtprops', 'oldsprops', 'blipuid', 'mhtmltag', 'htmltag', 'factoidname', 'bookmark', 'defchp', 'defpap',
]);

export class RtfReader {
  warnings = new Set<string>();
  private data: Uint8Array;
  private pos = 0;
  private stack: Group[] = [];
  private fonts = new Map<number, { family: FontFamily; charset?: number; name: string }>();
  private colors: (string | null)[] = [];
  private styles = new Map<number, { name: string; outline?: number }>();
  private lists = new Map<number, { nfc: number; start: number }[]>();
  private overrides = new Map<number, number>();
  private ansiCp = 'windows-1252';
  private defFont = 0;
  private body = new Sink('body');
  private header?: Block[];
  private footer?: Block[];
  private title?: string;
  private page = { w: 12240, h: 15840, l: 1800, r: 1800, t: 1440, b: 1440, set: false };
  private pendingBytes: number[] = [];
  private pendingCp = 'windows-1252';
  private skipChars = 0;
  private highSurrogate = 0;
  private pendingSect = false;
  private imageCount = 0;

  constructor(bytes: Uint8Array) {
    this.data = bytes;
  }

  private get g(): Group {
    return this.stack[this.stack.length - 1];
  }

  read(): Doc {
    const d = this.data;
    let start = 0;
    while (start < d.length && (d[start] === 0x20 || d[start] === 0x0d || d[start] === 0x0a || d[start] === 0xef || d[start] === 0xbb || d[start] === 0xbf)) start++;
    if (!(d[start] === 0x7b && d[start + 1] === 0x5c && d[start + 2] === 0x72 && d[start + 3] === 0x74 && d[start + 4] === 0x66)) throw new ConvertError('Файл не является документом RTF.');
    this.pos = start;
    const root: Group = { ch: { b: false, i: false, ul: false, strike: false, sup: false, sub: false, hidden: false, fs: 24, f: -1, cf: 0, hl: 0, uc: 1 }, dest: 'body', sink: this.body };
    this.stack = [root];
    let depth = 0;
    while (this.pos < d.length) {
      const c = d[this.pos];
      if (c === 0x7b) {
        // {
        this.flushBytes();
        this.pos++;
        const parent = this.g;
        this.stack.push({ ch: { ...parent.ch }, dest: parent.dest, sink: parent.sink });
        this.skipChars = 0;
        if (++depth > 2000) throw new ConvertError('Документ RTF повреждён: слишком глубокая вложенность.');
      } else if (c === 0x7d) {
        // }
        this.flushBytes();
        this.pos++;
        this.skipChars = 0;
        if (this.stack.length > 1) this.closeGroup(this.stack.pop()!);
        depth--;
        if (this.stack.length === 1 && depth <= 0) break; // end of the document group
      } else if (c === 0x5c) this.control();
      else if (c === 0x0d || c === 0x0a) this.pos++;
      else this.text();
    }
    this.flushBytes();
    this.endParagraph(this.body, false);
    this.flushTable(this.body);
    return this.buildDoc();
  }

  // -- tokenizer ------------------------------------------------------------

  private text(): void {
    const d = this.data;
    const g = this.g;
    let end = this.pos;
    while (end < d.length && d[end] !== 0x5c && d[end] !== 0x7b && d[end] !== 0x7d) end++;
    const dest = g.dest;
    if (dest === 'pict') {
      g.pict?.hex.push(String.fromCharCode(...d.subarray(this.pos, Math.min(end, this.pos + 60000))));
      // very long runs are split to stay within the argument limit of fromCharCode
      for (let p = this.pos + 60000; p < end; p += 60000) g.pict?.hex.push(String.fromCharCode(...d.subarray(p, Math.min(end, p + 60000))));
      this.pos = end;
      return;
    }
    for (let p = this.pos; p < end; p++) {
      const b = d[p];
      if (b === 0x0d || b === 0x0a) continue;
      if (this.skipChars > 0) {
        this.skipChars--;
        continue;
      }
      if (b === 0x09) {
        this.flushBytes();
        this.emit('\t');
      } else this.byte(b);
    }
    this.pos = end;
  }

  private control(): void {
    const d = this.data;
    const p = this.pos + 1;
    const c = d[p];
    if (c === undefined) {
      this.pos = p;
      return;
    }
    const isLetter = (x: number): boolean => (x >= 0x41 && x <= 0x5a) || (x >= 0x61 && x <= 0x7a);
    if (!isLetter(c)) {
      this.pos = p + 1;
      switch (c) {
        case 0x27: {
          // \'hh
          const h = String.fromCharCode(d[p + 1] ?? 0, d[p + 2] ?? 0);
          this.pos = p + 3;
          const v = parseInt(h, 16);
          if (this.skipChars > 0) {
            this.skipChars--;
            return;
          }
          if (Number.isFinite(v)) this.byte(v);
          return;
        }
        case 0x2a: // \*
          if (this.g.dest === 'body' || this.g.dest === 'shpinst' || this.g.dest === 'pict') this.g.dest = 'star';
          return;
        case 0x5c:
        case 0x7b:
        case 0x7d:
          if (this.skipChars > 0) {
            this.skipChars--;
            return;
          }
          this.byte(c);
          return;
        case 0x7e:
          this.flushBytes();
          this.emit(' ');
          return;
        case 0x5f:
          this.flushBytes();
          this.emit('‑');
          return;
        case 0x2d: // optional hyphen
          return;
        case 0x0a:
        case 0x0d:
          this.word('par', null);
          return;
        case 0x7c: // formula character
        case 0x3a: // index subentry
          return;
        default:
          return;
      }
    }
    let e = p;
    while (e < d.length && isLetter(d[e]) && e - p < 32) e++;
    const name = String.fromCharCode(...d.subarray(p, e));
    let param: number | null = null;
    let q = e;
    if (d[q] === 0x2d || (d[q] >= 0x30 && d[q] <= 0x39)) {
      const s = q;
      if (d[q] === 0x2d) q++;
      while (q < d.length && d[q] >= 0x30 && d[q] <= 0x39 && q - s < 11) q++;
      param = parseInt(String.fromCharCode(...d.subarray(s, q)), 10);
      if (!Number.isFinite(param)) param = null;
    }
    if (d[q] === 0x20) q++;
    this.pos = q;
    if (name === 'bin' && param !== null && param > 0) {
      const bytes = d.subarray(this.pos, Math.min(d.length, this.pos + param));
      this.pos += param;
      if (this.g.dest === 'pict' && this.g.pict) this.g.pict.bin = bytes;
      return;
    }
    if (this.skipChars > 0 && name !== 'u') {
      // A control word in the fallback text of \u counts as one character
      this.skipChars--;
      return;
    }
    this.word(name, param);
  }

  // -- text output -------------------------------------------------------------

  private byte(b: number): void {
    const g = this.g;
    const cp = this.codepage(g.ch.f);
    if (this.pendingBytes.length && cp !== this.pendingCp) this.flushBytes();
    this.pendingCp = cp;
    this.pendingBytes.push(b);
  }

  private codepage(font: number): string {
    const f = this.fonts.get(font < 0 ? this.defFont : font);
    if (f?.charset !== undefined && f.charset !== 1 && f.charset !== 2) return CHARSET_CP[f.charset] ?? this.ansiCp;
    return this.ansiCp;
  }

  private flushBytes(): void {
    if (!this.pendingBytes.length) return;
    const s = decoderFor(this.pendingCp).decode(new Uint8Array(this.pendingBytes));
    this.pendingBytes = [];
    this.emit(s);
  }

  private emit(s: string): void {
    const g = this.g;
    switch (g.dest) {
      case 'body':
        break;
      case 'fonttbl':
      case 'colortbl':
      case 'stylesheet':
      case 'title':
      case 'fldinst':
      case 'listtable':
      case 'listoverridetable':
      case 'sn':
      case 'sv':
        (g.text ??= []).push(s);
        if (g.dest === 'colortbl') this.colorText(s);
        else if ((g.dest === 'fonttbl' || g.dest === 'stylesheet') && s.includes(';')) this.registerEntry(g);
        return;
      default:
        return;
    }
    if (g.ch.hidden) return;
    if (this.pendingSect) {
      this.pendingSect = false;
      this.endParagraph(g.sink, false);
      this.pushBlock(g.sink, { type: 'pagebreak' });
    }
    const style = this.runStyle(g.ch);
    const runs = g.sink.runs;
    const last = runs[runs.length - 1];
    if (last && !last.note && sameStyle(last, style)) last.text += s;
    else runs.push({ text: s, ...style });
  }

  private runStyle(ch: CharState): Omit<TextRun, 'text'> {
    const s: Omit<TextRun, 'text'> = {};
    if (ch.b) s.bold = true;
    if (ch.i) s.italic = true;
    if (ch.ul) s.underline = true;
    if (ch.strike) s.strike = true;
    if (ch.sup) s.sup = true;
    else if (ch.sub) s.sub = true;
    const color = this.colors[ch.cf];
    if (ch.cf > 0 && color && color !== '000000') s.color = color;
    const hl = this.colors[ch.hl];
    if (ch.hl > 0 && hl && hl !== 'FFFFFF') s.highlight = hl;
    const font = this.fonts.get(ch.f < 0 ? this.defFont : ch.f);
    if (font) s.font = font.family;
    s.size = ch.fs / 2;
    if (ch.link) s.link = ch.link;
    if (ch.field) s.field = ch.field;
    return s;
  }

  private colorText(s: string): void {
    // Each ';' closes a colour entry; an entry without components is "auto"
    const g = this.g;
    for (const ch of s) {
      if (ch !== ';') continue;
      const c = (g as Group & { rgb?: number[] }).rgb;
      this.colors.push(c ? c.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase() : null);
      (g as Group & { rgb?: number[] }).rgb = undefined;
    }
  }

  // -- control words -----------------------------------------------------------

  private word(name: string, param: number | null): void {
    // Text buffered before a control word must be emitted before its effect
    if (name !== 'u') this.flushBytes();
    const g = this.g;
    const ch = g.ch;
    const on = param === null || param !== 0;
    // Destination after \* that we do not know: skip the whole group
    if (g.dest === 'star') {
      if (name === 'shppict' || name === 'ud') g.dest = 'body';
      else if (name === 'fldinst') g.dest = 'fldinst';
      else if (name === 'pn') g.dest = 'pn';
      else if (name === 'picprop') g.dest = 'shpinst';
      else if (name === 'shpinst') g.dest = 'shpinst';
      else if (name === 'footnote') {
        g.dest = 'body';
        this.word(name, param);
        return;
      } else if (name === 'listtable' || name === 'listoverridetable') g.dest = name;
      else g.dest = 'skip';
      if (g.dest !== 'body' || name === 'shppict' || name === 'ud') return;
    }
    if (g.dest === 'skip') return;
    if (SKIP.has(name)) {
      g.dest = 'skip';
      return;
    }
    switch (g.dest) {
      case 'fonttbl':
        return this.fontWord(name, param);
      case 'colortbl': {
        const rgb = ((g as Group & { rgb?: number[] }).rgb ??= [0, 0, 0]);
        if (name === 'red') rgb[0] = param ?? 0;
        else if (name === 'green') rgb[1] = param ?? 0;
        else if (name === 'blue') rgb[2] = param ?? 0;
        return;
      }
      case 'stylesheet':
        return this.styleWord(name, param);
      case 'listtable':
        return this.listWord(name, param);
      case 'listoverridetable':
        if (name === 'listoverride') g.override = { listId: 0, ls: 0 };
        else if (name === 'listid' && g.override) g.override.listId = param ?? 0;
        else if (name === 'ls' && g.override) g.override.ls = param ?? 0;
        return;
      case 'info':
        if (name === 'title') {
          g.dest = 'title';
          g.text = [];
        } else if (name !== 'info') g.dest = 'skip';
        return;
      case 'pict':
        return this.pictWord(name, param);
      case 'fldinst':
      case 'title':
        if (name === 'u' && param !== null) this.unicode(param);
        else if (name === 'uc') ch.uc = param ?? 1;
        return;
      case 'pn':
        // legacy (Word 6/95, WordPad) paragraph numbering
        if (name === 'pnlvlblt') g.sink.para.pnBullet = true;
        else if (name === 'pnlvlbody' || name === 'pndec' || name === 'pnucrm' || name === 'pnlcrm' || name === 'pnucltr' || name === 'pnlcltr') g.sink.para.pnOrdered = true;
        return;
      case 'shpinst':
        if (name === 'sn' || name === 'sv') {
          g.dest = name;
          g.text = [];
          return;
        }
        if (name !== 'pict' && name !== 'shptxt') return;
        break;
      case 'sn':
      case 'sv':
        if (name === 'u' && param !== null) this.unicode(param);
        return;
      default:
        break;
    }
    const sink = g.sink;
    const para = sink.para;
    switch (name) {
      // document
      case 'ansicpg':
        if (param) this.ansiCp = ANSI_CP[param] ?? `windows-${param}`;
        return;
      case 'mac':
        this.ansiCp = 'macintosh';
        return;
      case 'deff':
        this.defFont = param ?? 0;
        return;
      case 'fonttbl':
      case 'colortbl':
      case 'stylesheet':
      case 'info':
        g.dest = name;
        g.text = [];
        return;
      case 'listtable':
      case 'listoverridetable':
        g.dest = name;
        return;
      case 'pict':
        g.dest = 'pict';
        g.pict = { scaleX: 100, scaleY: 100, hex: [] };
        return;
      case 'paperw':
        this.page.w = param ?? this.page.w;
        this.page.set = true;
        return;
      case 'paperh':
        this.page.h = param ?? this.page.h;
        this.page.set = true;
        return;
      case 'margl':
        this.page.l = param ?? this.page.l;
        return;
      case 'margr':
        this.page.r = param ?? this.page.r;
        return;
      case 'margt':
        this.page.t = param ?? this.page.t;
        return;
      case 'margb':
        this.page.b = param ?? this.page.b;
        return;
      // sinks
      case 'footnote':
        g.opens = 'footnote';
        g.sink = new Sink('footnote');
        g.ch = { ...g.ch, sup: false, sub: false, link: undefined, field: undefined };
        return;
      case 'header':
      case 'headerr':
      case 'headerl':
      case 'headerf':
        g.opens = 'header';
        g.sink = new Sink('header');
        (g as Group & { which?: string }).which = name;
        return;
      case 'footer':
      case 'footerr':
      case 'footerl':
      case 'footerf':
        g.opens = 'footer';
        g.sink = new Sink('footer');
        (g as Group & { which?: string }).which = name;
        return;
      case 'chftn':
      case 'chpgn':
        if (name === 'chpgn') this.emitField('page');
        return;
      // fields
      case 'field':
        g.isField = true;
        return;
      case 'fldinst':
        g.dest = 'fldinst';
        g.text = [];
        return;
      case 'fldrslt': {
        const field = this.stack.slice().reverse().find((x) => x.isField);
        const inst = (field?.fieldInst ?? '').trim();
        const m = /^HYPERLINK\s+(?:\\[a-z]\s+)*"([^"]+)"/i.exec(inst) ?? /^HYPERLINK\s+(\S+)/i.exec(inst);
        if (m && /^(https?:|mailto:)/i.test(m[1])) ch.link = m[1];
        else if (/^PAGE\b/i.test(inst)) ch.field = 'page';
        else if (/^(NUMPAGES|SECTIONPAGES)\b/i.test(inst)) ch.field = 'pages';
        return;
      }
      case 'shpinst':
        g.dest = 'shpinst';
        return;
      case 'shptxt':
        g.dest = 'body';
        return;
      // characters
      case 'par':
        this.flushBytes();
        this.endParagraph(sink, true);
        return;
      case 'line':
        this.emit('\n');
        return;
      case 'tab':
        this.emit('\t');
        return;
      case 'page':
        this.flushBytes();
        this.endParagraph(sink, false);
        this.pushBlock(sink, { type: 'pagebreak' });
        return;
      case 'sect':
        this.flushBytes();
        this.endParagraph(sink, false);
        if (sink.kind === 'body') this.pendingSect = true;
        return;
      case 'sbknone':
        this.pendingSect = false;
        return;
      case 'emdash':
        return this.emit('—');
      case 'endash':
        return this.emit('–');
      case 'bullet':
        return this.emit('•');
      case 'lquote':
        return this.emit('‘');
      case 'rquote':
        return this.emit('’');
      case 'ldblquote':
        return this.emit('“');
      case 'rdblquote':
        return this.emit('”');
      case 'emspace':
      case 'enspace':
      case 'qmspace':
        return this.emit(' ');
      case 'zwj':
        return this.emit('‍');
      case 'zwnj':
        return this.emit('‌');
      case 'u':
        if (param !== null) this.unicode(param);
        return;
      case 'uc':
        ch.uc = Math.max(0, param ?? 1);
        return;
      // character formatting
      case 'plain':
        Object.assign(ch, { b: false, i: false, ul: false, strike: false, sup: false, sub: false, hidden: false, fs: 24, f: -1, cf: 0, hl: 0 });
        return;
      case 'b':
        ch.b = on;
        return;
      case 'i':
        ch.i = on;
        return;
      case 'ul':
      case 'uld':
      case 'uldash':
      case 'uldb':
      case 'ulth':
      case 'ulw':
      case 'ulwave':
      case 'uldashd':
      case 'uldashdd':
        ch.ul = on;
        return;
      case 'ulnone':
        ch.ul = false;
        return;
      case 'strike':
      case 'striked':
        ch.strike = on;
        return;
      case 'super':
        ch.sup = on;
        ch.sub = false;
        return;
      case 'sub':
        ch.sub = on;
        ch.sup = false;
        return;
      case 'nosupersub':
        ch.sup = ch.sub = false;
        return;
      case 'up':
        ch.sup = (param ?? 6) > 0;
        return;
      case 'dn':
        ch.sub = (param ?? 6) > 0;
        return;
      case 'v':
        ch.hidden = on;
        return;
      case 'fs':
        if (param && param > 0 && param < 3277) ch.fs = param;
        return;
      case 'f':
        ch.f = param ?? 0;
        return;
      case 'cf':
        ch.cf = param ?? 0;
        return;
      case 'highlight':
      case 'cb':
      case 'chcbpat':
        ch.hl = param ?? 0;
        return;
      // paragraph formatting
      case 'pard':
        sink.para = { ...freshPara(), intbl: false };
        return;
      case 'ql':
        para.align = 'left';
        return;
      case 'qc':
        para.align = 'center';
        return;
      case 'qr':
        para.align = 'right';
        return;
      case 'qj':
      case 'qd':
        para.align = 'justify';
        return;
      case 'li':
      case 'lin':
        para.li = param ?? 0;
        return;
      case 'fi':
        para.fi = param ?? 0;
        return;
      case 'sb':
        para.sb = param ?? 0;
        return;
      case 'sa':
        para.sa = param ?? 0;
        return;
      case 'sl':
        para.sl = param ?? undefined;
        return;
      case 'slmult':
        para.slmult = on;
        return;
      case 's':
        para.style = param ?? 0;
        return;
      case 'outlinelevel':
        para.outline = param ?? undefined;
        return;
      case 'intbl':
        para.intbl = true;
        return;
      case 'ls':
        para.ls = param ?? 0;
        return;
      case 'ilvl':
        para.ilvl = param ?? 0;
        return;
      case 'pagebb':
        para.pagebb = true;
        return;
      case 'cbpat':
        para.shaded = (param ?? 0) > 0;
        return;
      case 'brdrl':
        para.brdrl = true;
        return;
      case 'brdrb':
        para.brdrb = true;
        return;
      case 'pnlvlblt':
        para.pnBullet = true;
        return;
      case 'pnlvlbody':
      case 'pndec':
      case 'pnlvlcont':
        if (name !== 'pnlvlcont') para.pnOrdered = true;
        return;
      // tables
      case 'trowd':
        sink.table ??= { rows: [], cells: [], cellBlocks: [], header: false, merges: [], nextMerge: null };
        sink.table.header = false;
        sink.table.merges = [];
        return;
      case 'trhdr':
        if (sink.table) sink.table.header = true;
        return;
      case 'clmgf':
        if (sink.table) sink.table.nextMerge = 'first';
        return;
      case 'clmrg':
        if (sink.table) sink.table.nextMerge = 'cont';
        return;
      case 'cellx':
        if (sink.table) {
          sink.table.merges.push(sink.table.nextMerge);
          sink.table.nextMerge = null;
        }
        return;
      case 'cell':
        this.flushBytes();
        this.endCell(sink);
        return;
      case 'nestcell':
        this.emit('\t');
        return;
      case 'nestrow':
        this.emit('\n');
        return;
      case 'row':
        this.flushBytes();
        this.endRow(sink);
        return;
      default:
        return;
    }
  }

  private unicode(param: number): void {
    let code = param < 0 ? param + 65536 : param;
    this.skipChars = this.g.ch.uc;
    this.flushBytes();
    if (code >= 0xd800 && code <= 0xdbff) {
      this.highSurrogate = code;
      return;
    }
    if (code >= 0xdc00 && code <= 0xdfff && this.highSurrogate) {
      code = 0x10000 + ((this.highSurrogate - 0xd800) << 10) + (code - 0xdc00);
      this.highSurrogate = 0;
    }
    if (code === 0) return;
    this.emit(String.fromCodePoint(code));
  }

  private emitField(field: 'page' | 'pages'): void {
    const g = this.g;
    const saved = g.ch.field;
    g.ch.field = field;
    this.emit('1');
    g.ch.field = saved;
  }

  // -- tables ------------------------------------------------------------------

  private fontWord(name: string, param: number | null): void {
    const g = this.g;
    if (name === 'f') {
      g.font = { id: param ?? 0, name: '' };
      g.text = [];
    } else if (g.font) {
      if (name === 'froman') g.font.family = 'serif';
      else if (name === 'fswiss') g.font.family = 'sans';
      else if (name === 'fmodern') g.font.family = 'mono';
      else if (name === 'fcharset') g.font.charset = param ?? 0;
    }
  }

  private styleWord(name: string, param: number | null): void {
    const g = this.g;
    if (name === 's') {
      g.style = { id: param ?? 0, name: '' };
      g.text = [];
    } else if (name === 'cs' || name === 'ts' || name === 'ds') g.style = { id: -1, name: '', char: true };
    else if (name === 'outlinelevel' && g.style) g.style.outline = param ?? undefined;
  }

  private listWord(name: string, param: number | null): void {
    const g = this.g;
    if (name === 'list') g.list = { id: 0, levels: [] };
    else if (name === 'listlevel') g.level = { nfc: 0, start: 1 };
    else if (name === 'levelnfc' || name === 'levelnfcn') {
      if (g.level) g.level.nfc = param ?? 0;
    } else if (name === 'levelstartat') {
      if (g.level) g.level.start = param ?? 1;
    } else if (name === 'listid') {
      const list = this.stack.slice().reverse().find((x) => x.list)?.list;
      if (list) list.id = param ?? 0;
    } else if (name === 'leveltext' || name === 'levelnumbers' || name === 'listname') g.dest = 'skip';
  }

  private pictWord(name: string, param: number | null): void {
    const p = this.g.pict;
    if (!p) return;
    switch (name) {
      case 'pngblip':
        p.kind = 'png';
        break;
      case 'jpegblip':
        p.kind = 'jpeg';
        break;
      case 'emfblip':
      case 'wmetafile':
      case 'macpict':
      case 'pmmetafile':
      case 'dibitmap':
      case 'wbitmap':
        p.kind ??= 'other';
        break;
      case 'picw':
        p.w = param ?? undefined;
        break;
      case 'pich':
        p.h = param ?? undefined;
        break;
      case 'picwgoal':
        p.wGoal = param ?? undefined;
        break;
      case 'pichgoal':
        p.hGoal = param ?? undefined;
        break;
      case 'picscalex':
        p.scaleX = param ?? 100;
        break;
      case 'picscaley':
        p.scaleY = param ?? 100;
        break;
    }
  }

  private closeGroup(g: Group): void {
    const parent = this.g;
    switch (g.dest) {
      case 'fonttbl':
      case 'stylesheet':
        this.registerEntry(g);
        break;
      case 'listtable':
        if (g.level) {
          const list = this.stack.slice().reverse().find((x) => x.list)?.list;
          list?.levels.push(g.level);
        }
        if (g.list) this.lists.set(g.list.id, g.list.levels);
        break;
      case 'listoverridetable':
        if (g.override) this.overrides.set(g.override.ls, g.override.listId);
        break;
      case 'title':
        this.title = (g.text ?? []).join('').trim() || this.title;
        break;
      case 'fldinst': {
        // Nested groups inside fldinst append to the enclosing instruction
        if (parent.dest === 'fldinst') (parent.text ??= []).push(...(g.text ?? []));
        else {
          const field = this.stack.slice().reverse().find((x) => x.isField);
          if (field) field.fieldInst = (field.fieldInst ?? '') + (g.text ?? []).join('');
        }
        break;
      }
      case 'pict':
        if (g.pict && parent.dest !== 'pict') {
          g.pict.alt ??= this.stack.slice().reverse().find((x) => x.alt)?.alt;
          this.picture(g.pict, parent.sink);
        }
        break;
      case 'sn':
        parent.spName = (g.text ?? []).join('').trim();
        break;
      case 'sv': {
        const value = (g.text ?? []).join('').trim();
        if (parent.spName === 'wzDescription' && value) {
          // the property belongs to the enclosing picture or shape
          const owner = this.stack.slice().reverse().find((x) => x.pict) ?? this.stack.find((x) => x.dest === 'shpinst') ?? parent;
          if (owner.pict) owner.pict.alt = value;
          else owner.alt = value;
        }
        break;
      }
    }
    if (g.opens) {
      const sink = g.sink;
      this.endParagraph(sink, false);
      this.flushTable(sink);
      sink.nester.close();
      if (g.opens === 'footnote') {
        const note = sink.blocks.length ? sink.blocks : [{ type: 'paragraph', runs: [{ text: ' ' }] } as Block];
        parent.sink.runs.push({ text: '', note });
      } else {
        const which = (g as Group & { which?: string }).which ?? '';
        const target = g.opens;
        const primary = which === 'header' || which === 'headerr' || which === 'footer' || which === 'footerr';
        if (sink.blocks.length && (primary || !this[target])) this[target] = sink.blocks;
      }
    }
  }

  /** Font and style table entries end with ';' (with or without their own group). */
  private registerEntry(g: Group): void {
    const name = (g.text ?? []).join('').replace(/;.*$/s, '').trim();
    if (g.dest === 'fonttbl' && g.font) {
      const fam = familyOf(name);
      // A named font decides; the generic class only helps for unknown names
      const family: FontFamily = /arial|helvet|verdana|tahoma|segoe|calibri|carlito|times|georgia|cambria|garamond|courier|consolas|mono|sans|serif/i.test(name) && fam ? fam : (g.font.family ?? fam ?? 'sans');
      this.fonts.set(g.font.id, { family, charset: g.font.charset, name });
      g.font = undefined;
    } else if (g.dest === 'stylesheet' && g.style) {
      if (!g.style.char) this.styles.set(g.style.id, { name, outline: g.style.outline });
      g.style = undefined;
    }
    g.text = [];
  }

  private picture(p: Pict, sink: Sink): void {
    if (p.kind !== 'png' && p.kind !== 'jpeg') {
      if (p.kind) this.warnings.add('Некоторые изображения (WMF/EMF) пропущены: формат не поддерживается.');
      return;
    }
    let data = p.bin;
    if (!data) {
      const hex = p.hex.join('').replace(/[^0-9a-fA-F]/g, '');
      data = new Uint8Array(hex.length >> 1);
      for (let i = 0; i < data.length; i++) data[i] = parseInt(hex.substr(i * 2, 2), 16);
    }
    if (data.length < 8) return;
    const natural = imageSize(data) ?? { width: p.w ?? 300, height: p.h ?? 200 };
    // goal sizes are twips; 15 twips per CSS pixel
    let width = p.wGoal ? p.wGoal / 15 : natural.width;
    let height = p.hGoal ? p.hGoal / 15 : natural.height;
    width = (width * p.scaleX) / 100;
    height = (height * p.scaleY) / 100;
    this.imageCount++;
    const block: Block = { type: 'image', data, mime: p.kind === 'png' ? 'image/png' : 'image/jpeg', width: Math.round(width), height: Math.round(height) };
    if (p.alt) block.alt = p.alt;
    // An image is a block of its own: split the paragraph around it
    const runs = sink.runs;
    sink.runs = [];
    if (runs.some((r) => r.text.trim() || r.note)) {
      sink.runs = runs;
      this.endParagraph(sink, false, true);
    }
    this.pushBlock(sink, block);
  }

  // -- paragraphs ----------------------------------------------------------------

  private pushBlock(sink: Sink, b: Block): void {
    if (sink.para.intbl && sink.table) {
      sink.table.cellBlocks.push(b);
      return;
    }
    this.flushTable(sink);
    if (b.type === 'image' && sink.para.ls && sink.nester.append([b])) return;
    sink.nester.block(b);
  }

  private headingLevel(p: ParaState): number | undefined {
    const st = this.styles.get(p.style);
    const name = st?.name ?? '';
    const m = /(?:heading|заголовок)\s*(\d)/i.exec(name);
    if (m) return Math.min(6, Math.max(1, +m[1]));
    if (/^title$|^название$/i.test(name)) return 1;
    const outline = p.outline ?? st?.outline;
    if (outline !== undefined && outline >= 0 && outline < 9) return Math.min(6, outline + 1);
    return undefined;
  }

  /** Close the current paragraph. `explicit` is true for \par (empty paragraphs count). */
  private endParagraph(sink: Sink, explicit: boolean, keepPara = false): void {
    const p = sink.para;
    const runs = trimRuns(sink.runs);
    sink.runs = [];
    const hasText = runs.some((r) => r.text.trim() || r.note || r.field);
    if (!hasText && !explicit) return;
    if (p.pagebb && hasText) {
      this.pushBlock(sink, { type: 'pagebreak' });
      if (!keepPara) p.pagebb = false;
    }
    const fmt: ParaFormat = {};
    if (p.align && p.align !== 'left') fmt.align = p.align;
    const listed = !!p.ls || p.pnBullet || p.pnOrdered;
    if (!listed) {
      if (p.li) fmt.indent = p.li / 20;
      if (p.fi) fmt.firstLine = p.fi / 20;
    }
    if (p.sb !== undefined) fmt.spaceBefore = p.sb / 20;
    if (p.sa !== undefined) fmt.spaceAfter = p.sa / 20;
    if (p.sl && p.slmult && p.sl > 0) fmt.lineHeight = Math.round((p.sl / 240) * 100) / 100;
    const level = this.headingLevel(p);
    if (!hasText && p.brdrb && !p.intbl && !listed) {
      this.flushTable(sink);
      sink.nester.block({ type: 'hr' });
      return;
    }
    let block: Block;
    const mono = this.fonts.get(this.defFont)?.family !== 'mono' && runs.every((r) => !r.text.trim() || (r.font === 'mono' && !r.bold && !r.italic && !r.link && !r.note));
    if (!level && hasText && mono && p.shaded && !p.ls) {
      // A shaded monospace paragraph is a code block (this is how we and Word-like editors write them)
      const code: Block = { type: 'code', text: runs.map((r) => r.text).join('') };
      const prev = sink.blocks[sink.blocks.length - 1];
      if (p.intbl && sink.table) sink.table.cellBlocks.push(code);
      else if (prev?.type === 'code' && !sink.table) prev.text += '\n' + code.text;
      else {
        this.flushTable(sink);
        sink.nester.block(code);
      }
      return;
    }
    if (level && hasText) {
      const { indent: _i, firstLine: _f, ...hf } = fmt;
      block = { type: 'heading', level: level as 1, runs: runs.map((r) => ({ ...r, bold: undefined })).map(stripUndefined), ...hf };
    } else block = { type: 'paragraph', runs: hasText ? runs : [], ...fmt };
    if (p.intbl && sink.table) {
      sink.table.cellBlocks.push(block);
      return;
    }
    this.flushTable(sink);
    if (listed && hasText) {
      let ordered = false;
      let start: number | undefined;
      let key = `pn`;
      if (p.ls) {
        const listId = this.overrides.get(p.ls);
        const levels = listId !== undefined ? this.lists.get(listId) : undefined;
        const lv = levels?.[Math.min(p.ilvl, (levels?.length ?? 1) - 1)];
        ordered = lv ? lv.nfc !== 23 && lv.nfc !== 255 : false;
        start = lv?.start;
        key = `ls${p.ls}`;
      } else ordered = !!p.pnOrdered && !p.pnBullet;
      sink.nester.item([block], { key, level: p.ls ? p.ilvl : Math.max(0, Math.round(p.li / 720) - 1), ordered, start });
      return;
    }
    if (!hasText && sink.kind !== 'body') return;
    if (p.brdrl && p.li > 0 && hasText && block.type === 'paragraph' && !p.intbl) {
      // An indented paragraph with a left rule is a quotation
      const allItalic = block.runs.every((r) => r.italic || !r.text.trim());
      block.runs = normalizeRuns(block.runs.map((r) => stripUndefined({ ...r, italic: allItalic ? undefined : r.italic, color: r.color === '404040' ? undefined : r.color })));
      delete block.indent;
      delete block.firstLine;
      const prev = sink.blocks[sink.blocks.length - 1];
      if (prev?.type === 'quote') prev.blocks.push(block);
      else sink.nester.block({ type: 'quote', blocks: [block] });
      return;
    }
    sink.nester.block(block);
  }

  private endCell(sink: Sink): void {
    const t = sink.table;
    if (!t) {
      // \cell outside of a table definition: treat as a paragraph end
      this.endParagraph(sink, false);
      return;
    }
    const runs = sink.runs;
    if (runs.some((r) => r.text.trim() || r.note)) {
      const saved = sink.para.intbl;
      sink.para.intbl = true;
      this.endParagraph(sink, false);
      sink.para.intbl = saved;
    } else sink.runs = [];
    t.cells.push({ blocks: t.cellBlocks });
    t.cellBlocks = [];
  }

  private endRow(sink: Sink): void {
    const t = sink.table;
    if (!t) return;
    if (sink.runs.length) this.endCell(sink);
    const cells: TableCell[] = [];
    t.cells.forEach((c, i) => {
      if (t.merges[i] === 'cont' && cells.length) {
        const prev = cells[cells.length - 1];
        prev.colspan = (prev.colspan ?? 1) + 1;
        if (c.blocks.some((b) => b.type !== 'paragraph' || b.runs.length)) prev.blocks.push(...c.blocks);
        return;
      }
      if (t.header) c.header = true;
      cells.push(c);
    });
    if (cells.length) t.rows.push({ cells });
    t.cells = [];
    t.cellBlocks = [];
  }

  private flushTable(sink: Sink): void {
    const t = sink.table;
    if (!t) return;
    if (t.cells.length) this.endRow(sink);
    sink.table = null;
    if (!t.rows.length) return;
    // Cells keep only meaningful paragraphs
    for (const r of t.rows) for (const c of r.cells) c.blocks = c.blocks.filter((b) => b.type !== 'paragraph' || b.runs.length);
    sink.nester.block({ type: 'table', rows: t.rows });
  }

  // -- result ----------------------------------------------------------------------

  private buildDoc(): Doc {
    const doc: Doc = { blocks: this.body.blocks };
    if (this.title) doc.title = this.title;
    if (this.header) doc.header = this.header;
    if (this.footer) doc.footer = this.footer;
    while (doc.blocks.length && doc.blocks[doc.blocks.length - 1].type === 'pagebreak') doc.blocks.pop();
    // Dominant font and size become document defaults; runs keep only deviations
    const fontCount = new Map<FontFamily, number>();
    const sizeCount = new Map<number, number>();
    const visit = (blocks: Block[] | undefined): void => {
      for (const b of blocks ?? []) {
        if (b.type === 'paragraph') for (const r of b.runs) {
          if (r.font) fontCount.set(r.font, (fontCount.get(r.font) ?? 0) + r.text.length);
          if (r.size) sizeCount.set(r.size, (sizeCount.get(r.size) ?? 0) + r.text.length);
        }
        else if (b.type === 'list') b.items.forEach((i) => visit(i.blocks));
        else if (b.type === 'table') b.rows.forEach((r) => r.cells.forEach((c) => visit(c.blocks)));
      }
    };
    visit(doc.blocks);
    const top = <T>(m: Map<T, number>): T | undefined => [...m.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
    doc.font = top(fontCount) ?? this.fonts.get(this.defFont)?.family ?? 'serif';
    doc.size = top(sizeCount) ?? 12;
    const clean = (blocks: Block[] | undefined, note = false): void => {
      for (const b of blocks ?? []) {
        if (b.type === 'paragraph' || b.type === 'heading') {
          if (note) for (const k of ['indent', 'firstLine', 'spaceBefore', 'spaceAfter', 'lineHeight'] as const) delete (b as ParaFormat)[k];
          b.runs = normalizeRuns(
            b.runs.map((r) => {
              const x = { ...r };
              if (x.font === doc.font) delete x.font;
              // headings and notes take their size from their styles
              if (x.size === doc.size || b.type === 'heading' || note) delete x.size;
              if (x.note) clean(x.note, true);
              return x;
            }),
          );
        } else if (b.type === 'list') b.items.forEach((i) => clean(i.blocks, note));
        else if (b.type === 'table') b.rows.forEach((r) => r.cells.forEach((c) => clean(c.blocks, note)));
        else if (b.type === 'quote') clean(b.blocks, note);
      }
    };
    clean(doc.blocks);
    clean(doc.header);
    clean(doc.footer);
    // The most common paragraph spacing becomes the document default
    const spacing = new Map<string, number>();
    const collect = (blocks: Block[] | undefined): void => {
      for (const b of blocks ?? []) {
        if (b.type === 'paragraph') {
          const k = `${b.spaceBefore}|${b.spaceAfter}|${b.lineHeight}`;
          spacing.set(k, (spacing.get(k) ?? 0) + 1);
        } else if (b.type === 'list') b.items.forEach((i) => collect(i.blocks));
        else if (b.type === 'quote') collect(b.blocks);
      }
    };
    collect(doc.blocks);
    const best = [...spacing.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
    if (best) {
      const [sb, sa, lh] = best.split('|').map((v) => (v === 'undefined' ? undefined : +v));
      const defaults: ParaFormat = {};
      if (sb !== undefined) defaults.spaceBefore = sb;
      if (sa !== undefined) defaults.spaceAfter = sa;
      if (lh !== undefined) defaults.lineHeight = lh;
      if (Object.keys(defaults).length) doc.defaults = defaults;
      const strip = (blocks: Block[] | undefined): void => {
        for (const b of blocks ?? []) {
          if (b.type === 'paragraph' || b.type === 'heading') {
            if (b.spaceBefore === sb) delete b.spaceBefore;
            if (b.spaceAfter === sa) delete b.spaceAfter;
            if (b.lineHeight === lh) delete b.lineHeight;
          } else if (b.type === 'list') b.items.forEach((i) => strip(i.blocks));
          else if (b.type === 'table') b.rows.forEach((r) => r.cells.forEach((c) => strip(c.blocks)));
          else if (b.type === 'quote') strip(b.blocks);
        }
      };
      strip(doc.blocks);
      strip(doc.header);
      strip(doc.footer);
    }
    const pg = this.page;
    if (pg.set) {
      doc.page = { width: pg.w / 20, height: pg.h / 20, margins: { top: pg.t / 20, right: pg.r / 20, bottom: pg.b / 20, left: pg.l / 20 } };
    }
    return doc;
  }
}

function stripUndefined<T extends object>(o: T): T {
  for (const k of Object.keys(o) as (keyof T)[]) if (o[k] === undefined) delete o[k];
  return o;
}

function sameStyle(a: TextRun, b: Omit<TextRun, 'text'>): boolean {
  const keys: (keyof Omit<TextRun, 'text' | 'note'>)[] = ['bold', 'italic', 'underline', 'strike', 'sup', 'sub', 'color', 'highlight', 'font', 'size', 'link', 'field', 'code'];
  return keys.every((k) => a[k] === b[k]);
}

export function rtfToDoc(bytes: Uint8Array): { doc: Doc; warnings: string[] } {
  const r = new RtfReader(bytes);
  const doc = r.read();
  return { doc, warnings: [...r.warnings] };
}
