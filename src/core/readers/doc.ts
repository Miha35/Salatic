/**
 * Word 97–2003 binary (.doc) → Doc, following MS-DOC.
 *
 * Text comes from the piece table (CLX); paragraph properties (style, alignment,
 * indents, lists, table marks) from PAPX FKPs; character properties (bold, italic,
 * underline, strike, size, colour, super/subscript) from CHPX FKPs; heading levels from
 * the style sheet (built-in style ids 1–9); pictures from the Data stream.
 */
import { Cfb } from '../cfb';
import type { Align, Block, Doc, FontFamily, ListItem, TableRow, TextRun } from '../model';
import { familyOf, forEachRun, normalizeRuns, trimRuns } from '../model';
import { ConvertError, imageSize, sniffImageMime } from '../util';

interface Piece {
  cpStart: number;
  cpEnd: number;
  fc: number;
  compressed: boolean;
}

interface Prop {
  fcStart: number;
  fcEnd: number;
  sprms: Uint8Array;
  istd?: number;
}

interface ParaProps {
  istd: number;
  align?: Align;
  inTable: boolean;
  ttp: boolean;
  ilfo: number;
  ilvl: number;
  outline?: number;
  indent?: number;
  firstLine?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  pageBreak?: boolean;
}

type CharProps = Omit<TextRun, 'text'> & { hidden?: boolean; picture?: number; ftc?: number };

const ICO: string[] = ['', '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];

/** Operand length of a sprm (MS-DOC 2.6.1). */
function operandSize(sprm: number, data: Uint8Array, pos: number): number {
  switch (sprm >> 13) {
    case 0:
    case 1:
      return 1;
    case 2:
    case 4:
    case 5:
      return 2;
    case 3:
      return 4;
    case 7:
      return 3;
    default:
      if (sprm === 0xd608 || sprm === 0xd606) return (data[pos] | (data[pos + 1] << 8)) + 1;
      if (sprm === 0xc615) {
        const n = data[pos];
        return n === 255 ? 1 + data[pos + 1] * 4 + 1 + data[pos + 2 + data[pos + 1] * 4] * 3 + 1 : n + 1;
      }
      return data[pos] + 1;
  }
}

function* sprms(data: Uint8Array): Generator<{ sprm: number; at: number; len: number }> {
  let i = 0;
  while (i + 2 <= data.length) {
    const sprm = data[i] | (data[i + 1] << 8);
    i += 2;
    if (i >= data.length) break;
    const len = operandSize(sprm, data, i);
    yield { sprm, at: i, len };
    i += len;
  }
}

class DocReader {
  private wd: Uint8Array;
  private table: Uint8Array;
  private data: Uint8Array | null;
  private wdv: DataView;
  private tdv: DataView;
  private pieces: Piece[] = [];
  private papx: Prop[] = [];
  private chpx: Prop[] = [];
  private headingOf = new Map<number, number>(); // istd → level
  private styleBold = new Set<number>();
  private lists = new Map<number, boolean>(); // ilfo → ordered
  private fibRgFcLcb = 0;
  private ccpText = 0;
  private fonts: string[] = [];
  private defaultFont = 'Times New Roman';
  private defaultSize = 10;
  warnings = new Set<string>();

  constructor(bytes: Uint8Array) {
    const cfb = new Cfb(bytes);
    const wd = cfb.stream('WordDocument');
    if (!wd) throw new ConvertError('Это не документ Word: нет потока WordDocument.');
    this.wd = wd;
    this.wdv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    const ident = this.wdv.getUint16(0, true);
    if (ident !== 0xa5ec) throw new ConvertError('Документы Word 6.0/95 не поддерживаются — пересохраните файл в формате Word 97 или новее.');
    const flags = this.wdv.getUint16(0x0a, true);
    if (flags & 0x0100) throw new ConvertError('Документ Word защищён паролем. Снимите защиту и попробуйте снова.', 'password');
    const table = cfb.stream(flags & 0x0200 ? '1Table' : '0Table');
    if (!table) throw new ConvertError('Повреждённый документ Word: нет таблицы свойств.');
    this.table = table;
    this.tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
    this.data = cfb.stream('Data');
    // FIB: base (32) + csw + fibRgW + cslw + fibRgLw + cbRgFcLcb + fibRgFcLcb
    const csw = this.wdv.getUint16(0x20, true);
    const lwStart = 0x22 + csw * 2 + 2;
    const cslw = this.wdv.getUint16(0x22 + csw * 2, true);
    this.ccpText = this.wdv.getInt32(lwStart + 3 * 4, true);
    this.fibRgFcLcb = lwStart + cslw * 4 + 2;
    this.readPieces();
    this.readFonts();
    this.papx = this.readFkps(13, true);
    this.chpx = this.readFkps(12, false);
    this.readStyles();
    this.readLists();
  }

  private fcLcb(index: number): { fc: number; lcb: number } {
    const off = this.fibRgFcLcb + index * 8;
    if (off + 8 > this.wd.length) return { fc: 0, lcb: 0 };
    return { fc: this.wdv.getUint32(off, true), lcb: this.wdv.getUint32(off + 4, true) };
  }

  private readPieces(): void {
    const { fc, lcb } = this.fcLcb(33);
    if (!lcb || fc + lcb > this.table.length) throw new ConvertError('Повреждённый документ Word: нет таблицы фрагментов текста.');
    let i = fc;
    const end = fc + lcb;
    while (i < end && this.table[i] === 0x01) i += 3 + this.tdv.getInt16(i + 1, true); // skip Prc
    if (this.table[i] !== 0x02) throw new ConvertError('Повреждённый документ Word: неверный формат таблицы фрагментов.');
    const len = this.tdv.getUint32(i + 1, true);
    const plc = i + 5;
    const n = Math.floor((len - 4) / 12);
    for (let k = 0; k < n; k++) {
      const cpStart = this.tdv.getUint32(plc + k * 4, true);
      const cpEnd = this.tdv.getUint32(plc + (k + 1) * 4, true);
      const pcd = plc + (n + 1) * 4 + k * 8;
      const fcRaw = this.tdv.getUint32(pcd + 2, true);
      const compressed = (fcRaw & 0x40000000) !== 0;
      this.pieces.push({ cpStart, cpEnd, fc: compressed ? (fcRaw & 0x3fffffff) / 2 : fcRaw & 0x3fffffff, compressed });
    }
  }

  private text(cpStart: number, cpEnd: number): string {
    let s = '';
    const dec1252 = new TextDecoder('windows-1252');
    for (const p of this.pieces) {
      const a = Math.max(cpStart, p.cpStart);
      const b = Math.min(cpEnd, p.cpEnd);
      if (a >= b) continue;
      if (p.compressed) {
        const off = p.fc + (a - p.cpStart);
        s += dec1252.decode(this.wd.subarray(off, off + (b - a)));
      } else {
        const off = p.fc + (a - p.cpStart) * 2;
        for (let k = 0; k < b - a && off + k * 2 + 1 < this.wd.length; k++) s += String.fromCharCode(this.wdv.getUint16(off + k * 2, true));
      }
    }
    return s;
  }

  /** File position of a character. */
  private fcOf(cp: number): number {
    for (const p of this.pieces) if (cp >= p.cpStart && cp < p.cpEnd) return p.compressed ? p.fc + (cp - p.cpStart) : p.fc + (cp - p.cpStart) * 2;
    return -1;
  }

  /** Read PAPX (index 13) or CHPX (index 12) formatted disk pages. */
  private readFkps(index: number, para: boolean): Prop[] {
    const { fc, lcb } = this.fcLcb(index);
    const out: Prop[] = [];
    if (!lcb || fc + lcb > this.table.length) return out;
    const n = (lcb - 4) / 8;
    for (let k = 0; k < n; k++) {
      const pn = this.tdv.getUint32(fc + (n + 1) * 4 + k * 4, true) & 0x3fffff;
      const page = pn * 512;
      if (page + 512 > this.wd.length) continue;
      const crun = this.wd[page + 511];
      for (let r = 0; r < crun; r++) {
        const fcStart = this.wdv.getUint32(page + r * 4, true);
        const fcEnd = this.wdv.getUint32(page + (r + 1) * 4, true);
        if (para) {
          const bOff = this.wd[page + (crun + 1) * 4 + r * 13] * 2;
          if (!bOff) {
            out.push({ fcStart, fcEnd, sprms: new Uint8Array(0), istd: 0 });
            continue;
          }
          let cb = this.wd[page + bOff];
          let start = page + bOff + 1;
          if (cb === 0) {
            cb = this.wd[start];
            start++;
            cb *= 2;
          } else cb = cb * 2 - 1;
          const istd = this.wdv.getUint16(start, true);
          out.push({ fcStart, fcEnd, istd, sprms: this.wd.subarray(start + 2, Math.min(start + cb, page + 511)) });
        } else {
          const bOff = this.wd[page + (crun + 1) * 4 + r] * 2;
          if (!bOff) {
            out.push({ fcStart, fcEnd, sprms: new Uint8Array(0) });
            continue;
          }
          const cb = this.wd[page + bOff];
          out.push({ fcStart, fcEnd, sprms: this.wd.subarray(page + bOff + 1, page + bOff + 1 + cb) });
        }
      }
    }
    return out.sort((a, b) => a.fcStart - b.fcStart);
  }

  private propAt(list: Prop[], fc: number): Prop | undefined {
    let lo = 0;
    let hi = list.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const p = list[mid];
      if (fc < p.fcStart) hi = mid - 1;
      else if (fc >= p.fcEnd) lo = mid + 1;
      else return p;
    }
    return undefined;
  }

  /** Font table (SttbfFfn): names by index. */
  private readFonts(): void {
    const { fc, lcb } = this.fcLcb(15);
    if (!lcb || fc + lcb > this.table.length) return;
    let pos = fc;
    let cData = this.tdv.getUint16(pos, true);
    pos += 2;
    if (cData === 0xffff) {
      cData = this.tdv.getUint16(pos, true);
      pos += 2;
    }
    pos += 2; // cbExtra
    for (let i = 0; i < cData && pos < fc + lcb; i++) {
      const len = this.table[pos];
      let name = '';
      for (let k = pos + 1 + 39; k + 1 < pos + 1 + len && k + 1 < this.table.length; k += 2) {
        const c = this.tdv.getUint16(k, true);
        if (!c) break;
        name += String.fromCharCode(c);
      }
      this.fonts.push(name);
      pos += len + 1;
    }
  }

  /** Style sheet: built-in heading styles (sti 1–9), default font and size of "Normal". */
  private readStyles(): void {
    const { fc, lcb } = this.fcLcb(1);
    if (!lcb || fc + lcb > this.table.length) return;
    const cbStshi = this.tdv.getUint16(fc, true);
    const cstd = this.tdv.getUint16(fc + 2, true);
    const cbBase = this.tdv.getUint16(fc + 4, true);
    const ftcAscii = this.tdv.getUint16(fc + 2 + 12, true);
    if (this.fonts[ftcAscii]) this.defaultFont = this.fonts[ftcAscii];
    let pos = fc + 2 + cbStshi;
    for (let istd = 0; istd < cstd && pos + 2 <= fc + lcb; istd++) {
      const cbStd = this.tdv.getUint16(pos, true);
      pos += 2;
      if (cbStd) {
        const sti = this.tdv.getUint16(pos, true) & 0x0fff;
        const stk = this.tdv.getUint16(pos + 2, true) & 0x000f;
        if (sti >= 1 && sti <= 9) {
          this.headingOf.set(istd, Math.min(6, sti));
          this.styleBold.add(istd);
        }
        if (sti === 62) this.headingOf.set(istd, 1); // Title
        if (istd === 0) this.readNormalStyle(pos, cbStd, cbBase, stk);
      }
      pos += cbStd;
    }
  }

  /** Character formatting of the "Normal" style → document defaults. */
  private readNormalStyle(pos: number, cbStd: number, cbBase: number, stk: number): void {
    const end = pos + cbStd;
    let p = pos + cbBase;
    const cch = this.tdv.getUint16(p, true);
    p += 2 + cch * 2 + 2;
    const upx = (): Uint8Array | null => {
      if (p + 2 > end) return null;
      const cb = this.tdv.getUint16(p, true);
      const data = this.table.subarray(p + 2, p + 2 + cb);
      p += 2 + cb + (cb % 2);
      return data;
    };
    if (stk === 1) upx(); // paragraph properties
    const chpx = upx();
    if (!chpx) return;
    const dv = new DataView(chpx.buffer, chpx.byteOffset, chpx.byteLength);
    for (const { sprm, at, len } of sprms(chpx)) {
      if (at + len > chpx.length) break;
      if (sprm === 0x4a43) this.defaultSize = dv.getUint16(at, true) / 2;
      else if (sprm === 0x4a4f && this.fonts[dv.getUint16(at, true)]) this.defaultFont = this.fonts[dv.getUint16(at, true)];
    }
  }

  /** Lists: ilfo → ordered/bulleted, from PlfLst (LSTF + LVL) and PlfLfo. */
  private readLists(): void {
    const lst = this.fcLcb(73);
    const lfo = this.fcLcb(74);
    if (!lst.lcb || !lfo.lcb || lst.fc + lst.lcb > this.table.length || lfo.fc + lfo.lcb > this.table.length) return;
    const cLst = this.tdv.getInt16(lst.fc, true);
    const ordered = new Map<number, boolean>(); // lsid → first level is numbered
    // LSTF array, then for each list its LVLs (1 or 9)
    const lstfs: { lsid: number; simple: boolean }[] = [];
    for (let i = 0; i < cLst; i++) {
      const o = lst.fc + 2 + i * 28;
      lstfs.push({ lsid: this.tdv.getInt32(o, true), simple: (this.table[o + 26] & 0x01) !== 0 });
    }
    let pos = lst.fc + 2 + cLst * 28;
    for (const l of lstfs) {
      const levels = l.simple ? 1 : 9;
      for (let k = 0; k < levels && pos + 28 <= this.table.length; k++) {
        const nfc = this.table[pos + 4];
        const cbGrpprlChpx = this.table[pos + 24];
        const cbGrpprlPapx = this.table[pos + 25];
        if (k === 0) ordered.set(l.lsid, nfc !== 0x17 && nfc !== 0xff);
        pos += 28 + cbGrpprlPapx + cbGrpprlChpx;
        const cch = this.tdv.getUint16(pos, true);
        pos += 2 + cch * 2;
      }
    }
    const cLfo = this.tdv.getInt32(lfo.fc, true);
    for (let i = 0; i < cLfo; i++) {
      const lsid = this.tdv.getInt32(lfo.fc + 4 + i * 16, true);
      this.lists.set(i + 1, ordered.get(lsid) ?? false);
    }
  }

  private paraProps(cpMark: number): ParaProps {
    const fc = this.fcOf(cpMark);
    const p = this.propAt(this.papx, fc);
    const props: ParaProps = { istd: p?.istd ?? 0, inTable: false, ttp: false, ilfo: 0, ilvl: 0 };
    if (!p) return props;
    const d = p.sprms;
    const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
    for (const { sprm, at, len } of sprms(d)) {
      if (at + len > d.length) break;
      switch (sprm) {
        case 0x2403:
        case 0x2461: {
          const jc = d[at];
          props.align = jc === 1 ? 'center' : jc === 2 ? 'right' : jc === 3 || jc === 4 ? 'justify' : undefined;
          break;
        }
        case 0x2416:
          props.inTable = d[at] !== 0;
          break;
        case 0x2417:
          props.ttp = d[at] !== 0;
          break;
        case 0x260a:
          props.ilvl = d[at];
          break;
        case 0x460b:
          props.ilfo = dv.getInt16(at, true);
          break;
        case 0x2640:
          props.outline = d[at];
          break;
        case 0x2407:
          props.pageBreak = d[at] !== 0;
          break;
        case 0x840f:
        case 0x845e:
          props.indent = dv.getInt16(at, true) / 20;
          break;
        case 0x8411:
        case 0x8460:
          props.firstLine = dv.getInt16(at, true) / 20;
          break;
        case 0xa413:
          props.spaceBefore = dv.getUint16(at, true) / 20;
          break;
        case 0xa414:
          props.spaceAfter = dv.getUint16(at, true) / 20;
          break;
      }
    }
    return props;
  }

  private charProps(cp: number, base: CharProps): CharProps {
    const fc = this.fcOf(cp);
    const p = this.propAt(this.chpx, fc);
    const s: CharProps = { ...base };
    if (!p) return s;
    const d = p.sprms;
    const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
    const toggle = (cur: boolean | undefined, v: number): boolean => (v === 0 ? false : v === 1 ? true : v === 0x81 ? !cur : !!cur);
    for (const { sprm, at, len } of sprms(d)) {
      if (at + len > d.length) break;
      switch (sprm) {
        case 0x0835:
          s.bold = toggle(base.bold, d[at]);
          break;
        case 0x0836:
          s.italic = toggle(base.italic, d[at]);
          break;
        case 0x0837:
        case 0x2a53:
          s.strike = d[at] !== 0;
          break;
        case 0x083c:
          s.hidden = d[at] !== 0;
          break;
        case 0x2a3e:
          s.underline = d[at] !== 0;
          break;
        case 0x4a43:
          s.size = dv.getUint16(at, true) / 2;
          break;
        case 0x2a42: {
          const c = ICO[d[at]];
          if (c && c !== '000000') s.color = c;
          break;
        }
        case 0x6870: {
          const auto = d[at + 3] === 0xff;
          const hex = [d[at], d[at + 1], d[at + 2]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
          if (!auto && hex !== '000000') s.color = hex;
          break;
        }
        case 0x2a48:
          if (d[at] === 1) s.sup = true;
          else if (d[at] === 2) s.sub = true;
          break;
        case 0x6a03:
          s.picture = dv.getInt32(at, true);
          break;
        case 0x4a4f:
          s.ftc = dv.getUint16(at, true);
          break;
      }
    }
    return s;
  }

  /** Extract a picture from the Data stream (PICF + OfficeArt BLIP). */
  private picture(offset: number): Block | null {
    const data = this.data;
    if (!data || offset < 0 || offset + 4 > data.length) return null;
    const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const lcb = dv.getInt32(offset, true);
    const end = Math.min(data.length, offset + Math.max(0, lcb));
    for (let i = offset + 4; i < end - 8; i++) {
      // PNG
      if (data[i] === 0x89 && data[i + 1] === 0x50 && data[i + 2] === 0x4e && data[i + 3] === 0x47) {
        let k = i + 8;
        while (k + 12 <= end) {
          const len = dv.getUint32(k);
          const type = String.fromCharCode(data[k + 4], data[k + 5], data[k + 6], data[k + 7]);
          k += 12 + len;
          if (type === 'IEND') break;
        }
        return this.imageBlock(data.slice(i, Math.min(k, end)));
      }
      // JPEG
      if (data[i] === 0xff && data[i + 1] === 0xd8 && data[i + 2] === 0xff) {
        let last = -1;
        for (let k = end - 2; k > i; k--) {
          if (data[k] === 0xff && data[k + 1] === 0xd9) {
            last = k + 2;
            break;
          }
        }
        if (last > 0) return this.imageBlock(data.slice(i, last));
      }
    }
    this.warnings.add('Некоторые изображения (векторные WMF/EMF) из документа Word 97–2003 пропущены.');
    return null;
  }

  private imageBlock(bytes: Uint8Array): Block | null {
    const mime = sniffImageMime(bytes);
    if (!mime) return null;
    const size = imageSize(bytes) ?? { width: 300, height: 200 };
    return { type: 'image', data: bytes, mime, width: size.width, height: size.height };
  }

  read(): Doc {
    const full = this.text(0, this.ccpText);
    type Para = { runs: TextRun[]; props: ParaProps; images: Block[]; cellEnd: boolean };
    const paras: Para[] = [];
    let runs: TextRun[] = [];
    let images: Block[] = [];
    let field: { phase: 'instr' | 'result'; depth: number; instr: string; link?: string }[] = [];
    let cur: { style: CharProps; text: string } | null = null;
    const flushRun = (): void => {
      if (cur && cur.text) {
        const { hidden: _h, picture: _p, ...style } = cur.style;
        runs.push({ text: cur.text, ...style });
      }
      cur = null;
    };
    let paraStart = 0;
    for (let i = 0; i < full.length; i++) {
      const ch = full[i];
      const code = ch.charCodeAt(0);
      if (code === 0x0d || code === 0x07) {
        flushRun();
        const props = this.paraProps(i);
        paras.push({ runs: normalizeRuns(runs), props, images, cellEnd: code === 0x07 });
        runs = [];
        images = [];
        paraStart = i + 1;
        continue;
      }
      if (code === 0x13) {
        flushRun();
        field.push({ phase: 'instr', depth: field.length, instr: '' });
        continue;
      }
      if (code === 0x14) {
        flushRun();
        const f = field[field.length - 1];
        if (f) {
          f.phase = 'result';
          const m = /HYPERLINK\s+"([^"]+)"/.exec(f.instr);
          if (m && /^(https?:|mailto:)/i.test(m[1])) f.link = m[1];
        }
        continue;
      }
      if (code === 0x15) {
        flushRun();
        field.pop();
        continue;
      }
      const open = field[field.length - 1];
      if (open?.phase === 'instr') {
        open.instr += ch;
        continue;
      }
      if (field.some((f) => f.phase === 'instr')) continue;
      const style = this.charProps(i, {});
      if (style.hidden) continue;
      if (code === 0x01 || code === 0x08) {
        flushRun();
        if (style.picture !== undefined) {
          const img = this.picture(style.picture);
          if (img) images.push(img);
        }
        continue;
      }
      let t = ch;
      if (code === 0x0b) t = '\n';
      else if (code === 0x0c) {
        flushRun();
        images.push({ type: 'pagebreak' });
        continue;
      } else if (code === 0x1e) t = '-';
      else if (code === 0x1f || code === 0x02 || code === 0x05 || code < 0x09) continue;
      else if (code === 0x09) t = '\t';
      else if (code === 0xa0) t = ' ';
      const { picture: _p, ftc, ...clean } = style;
      if (ftc !== undefined && this.fonts[ftc]) {
        const fam: FontFamily | undefined = familyOf(this.fonts[ftc]);
        if (fam) clean.font = fam;
      }
      const link = field.find((f) => f.link)?.link;
      if (link) clean.link = link;
      if (cur && JSON.stringify(cur.style) === JSON.stringify(clean)) cur.text += t;
      else {
        flushRun();
        cur = { style: clean, text: t };
      }
    }
    flushRun();
    if (runs.length || images.length) paras.push({ runs: normalizeRuns(runs), props: this.paraProps(Math.max(0, full.length - 1)), images, cellEnd: false });
    void paraStart;
    field = [];

    // Assemble blocks: tables (cell marks + TTP rows), lists, headings, paragraphs
    const blocks: Block[] = [];
    let table: TableRow[] | null = null;
    let cell: Block[] = [];
    let row: TableRow = { cells: [] };
    let listStack: { list: Extract<Block, { type: 'list' }>; ilfo: number }[] = [];
    const paraBlock = (p: Para): Block[] => {
      const out: Block[] = [];
      if (p.props.pageBreak) out.push({ type: 'pagebreak' });
      for (const img of p.images) if (img.type === 'pagebreak') out.push(img);
      const runs2 = trimRuns(p.runs);
      const level = this.headingOf.get(p.props.istd) ?? (p.props.outline !== undefined && p.props.outline < 9 ? Math.min(6, p.props.outline + 1) : 0);
      const fmt: Record<string, unknown> = {};
      if (p.props.align) fmt.align = p.props.align;
      if (p.props.indent) fmt.indent = p.props.indent;
      if (p.props.firstLine) fmt.firstLine = p.props.firstLine;
      if (p.props.spaceBefore) fmt.spaceBefore = p.props.spaceBefore;
      if (p.props.spaceAfter !== undefined) fmt.spaceAfter = p.props.spaceAfter;
      if (runs2.length && level) out.push({ type: 'heading', level: level as 1, runs: runs2.map(({ bold: _b, ...r }) => r), ...(p.props.align ? { align: p.props.align } : {}) });
      else if (runs2.length || !p.images.length) out.push({ type: 'paragraph', runs: runs2, ...fmt });
      for (const img of p.images) if (img.type !== 'pagebreak') out.push(img);
      return out;
    };
    for (const p of paras) {
      if (p.props.inTable || p.cellEnd) {
        listStack = [];
        table ??= [];
        if (p.props.ttp) {
          // row end mark
          if (row.cells.length) table.push(row);
          row = { cells: [] };
          cell = [];
          continue;
        }
        cell.push(...paraBlock(p));
        if (p.cellEnd) {
          row.cells.push({ blocks: cell.length ? cell : [{ type: 'paragraph', runs: [] }] });
          cell = [];
        }
        continue;
      }
      if (table) {
        if (row.cells.length) table.push(row);
        if (table.length) blocks.push({ type: 'table', rows: table });
        table = null;
        row = { cells: [] };
        cell = [];
      }
      if (p.props.ilfo > 0 && p.runs.length) {
        const ordered = this.lists.get(p.props.ilfo) ?? false;
        const lvl = Math.min(p.props.ilvl, 8);
        while (listStack.length > lvl + 1) listStack.pop();
        if (listStack.length === lvl + 1 && (listStack[lvl].list.ordered !== ordered || (lvl === 0 && listStack[0].ilfo !== p.props.ilfo))) listStack.pop();
        while (listStack.length < lvl + 1) {
          const fresh: Extract<Block, { type: 'list' }> = { type: 'list', ordered, items: [] };
          if (!listStack.length) blocks.push(fresh);
          else {
            const parent = listStack[listStack.length - 1].list;
            if (!parent.items.length) parent.items.push({ blocks: [] });
            parent.items[parent.items.length - 1].blocks.push(fresh);
          }
          listStack.push({ list: fresh, ilfo: p.props.ilfo });
        }
        const item: ListItem = { blocks: paraBlock({ ...p, props: { ...p.props, indent: undefined, firstLine: undefined } }) };
        listStack[lvl].list.items.push(item);
        continue;
      }
      listStack = [];
      blocks.push(...paraBlock(p));
    }
    if (table) {
      if (row.cells.length) table.push(row);
      if (table.length) blocks.push({ type: 'table', rows: table });
    }
    // Drop trailing empty paragraphs
    while (blocks.length && blocks[blocks.length - 1].type === 'paragraph' && !(blocks[blocks.length - 1] as { runs: TextRun[] }).runs.length) blocks.pop();
    // Body size: the size carrying most characters; runs keep only differing sizes
    const def = this.defaultSize;
    const sizes = new Map<number, number>();
    forEachRun(blocks, (r) => sizes.set(r.size ?? def, (sizes.get(r.size ?? def) ?? 0) + r.text.length));
    const size = [...sizes].sort((a, b) => b[1] - a[1])[0]?.[0] ?? def;
    const font = familyOf(this.defaultFont) ?? 'serif';
    forEachRun(blocks, (r) => {
      if ((r.size ?? def) === size) delete r.size;
      else r.size ??= def;
      if (r.font === font) delete r.font;
    });
    return { blocks, font, size };
  }
}

export function docBinaryToDoc(bytes: Uint8Array): { doc: Doc; warnings: string[] } {
  const r = new DocReader(bytes);
  const doc = r.read();
  return { doc, warnings: [...r.warnings] };
}
