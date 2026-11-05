/** Excel 97–2003 (.xls, BIFF8) → Workbook, following MS-XLS. */
import { Cfb } from '../cfb';
import type { CellValue, Sheet, Workbook } from '../model';
import { ConvertError } from '../util';
import { BUILTIN_DATE_FORMATS, excelDate, isDateFormat } from './data';

interface Rec {
  type: number;
  data: Uint8Array;
}

function records(stream: Uint8Array, start = 0): Rec[] {
  const out: Rec[] = [];
  const dv = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  let pos = start;
  while (pos + 4 <= stream.length) {
    const type = dv.getUint16(pos, true);
    const len = dv.getUint16(pos + 2, true);
    out.push({ type, data: stream.subarray(pos + 4, Math.min(stream.length, pos + 4 + len)) });
    pos += 4 + len;
    if (type === 0x000a && start !== 0) break; // EOF of a sheet substream
  }
  return out;
}

/** Reads across a record and its CONTINUE records (strings may be split between them). */
class Segments {
  private i = 0;
  private pos = 0;
  constructor(private segs: Uint8Array[]) {}
  get done(): boolean {
    return this.i >= this.segs.length;
  }
  private seg(): Uint8Array {
    return this.segs[this.i];
  }
  private advance(): void {
    while (this.i < this.segs.length && this.pos >= this.seg().length) {
      this.i++;
      this.pos = 0;
    }
  }
  u8(): number {
    this.advance();
    if (this.done) return 0;
    return this.seg()[this.pos++];
  }
  u16(): number {
    return this.u8() | (this.u8() << 8);
  }
  i32(): number {
    return this.u16() | (this.u16() << 16);
  }
  skip(n: number): void {
    while (n > 0 && !this.done) {
      this.advance();
      if (this.done) return;
      const take = Math.min(n, this.seg().length - this.pos);
      this.pos += take;
      n -= take;
    }
  }
  /** Characters of a string; a CONTINUE boundary re-states the high-byte flag. */
  chars(count: number, highByte: boolean): string {
    let s = '';
    let wide = highByte;
    while (count > 0 && !this.done) {
      if (this.pos >= this.seg().length) {
        this.i++;
        this.pos = 0;
        if (this.done) break;
        wide = (this.seg()[this.pos++] & 0x01) !== 0;
        continue;
      }
      if (wide) {
        const lo = this.seg()[this.pos++];
        const hi = this.pos < this.seg().length ? this.seg()[this.pos++] : 0;
        s += String.fromCharCode(lo | (hi << 8));
      } else s += String.fromCharCode(this.seg()[this.pos++]);
      count--;
    }
    return s;
  }
  /** XLUnicodeRichExtendedString (SST entries). */
  richString(): string {
    const cch = this.u16();
    const flags = this.u8();
    const runs = flags & 0x08 ? this.u16() : 0;
    const ext = flags & 0x04 ? this.i32() : 0;
    const s = this.chars(cch, (flags & 0x01) !== 0);
    this.skip(runs * 4 + ext);
    return s;
  }
}

function xlString(d: Uint8Array, pos: number, cchBytes: 1 | 2): { s: string; end: number } {
  const cch = cchBytes === 1 ? d[pos] : d[pos] | (d[pos + 1] << 8);
  pos += cchBytes;
  const wide = (d[pos++] & 0x01) !== 0;
  let s = '';
  for (let k = 0; k < cch && pos < d.length; k++) {
    if (wide) {
      s += String.fromCharCode(d[pos] | (d[pos + 1] << 8));
      pos += 2;
    } else s += String.fromCharCode(d[pos++]);
  }
  return { s, end: pos };
}

function rk(v: number): number {
  let n: number;
  if (v & 0x02) n = v >> 2;
  else {
    const buf = new DataView(new ArrayBuffer(8));
    buf.setUint32(4, v & 0xfffffffc, true);
    buf.setUint32(0, 0, true);
    n = buf.getFloat64(0, true);
  }
  return v & 0x01 ? n / 100 : n;
}

export function xlsToBook(bytes: Uint8Array): Workbook {
  const cfb = new Cfb(bytes);
  const stream = cfb.stream('Workbook') ?? cfb.stream('Book');
  if (!stream) throw new ConvertError('Это не книга Excel: нет потока Workbook.');
  if (cfb.has('EncryptionInfo') || cfb.has('Encryption')) throw new ConvertError('Книга Excel защищена паролем. Снимите защиту и попробуйте снова.', 'password');
  const recs = records(stream);
  if (!recs.length || recs[0].type !== 0x0809) throw new ConvertError('Файл .xls повреждён.');
  const bofVersion = recs[0].data[0] | (recs[0].data[1] << 8);
  if (bofVersion !== 0x0600) throw new ConvertError('Поддерживаются книги Excel 97–2003 (BIFF8). Пересохраните файл в формате .xlsx.');
  let date1904 = false;
  const formats = new Map<number, string>();
  const xfFormat: number[] = [];
  const sheetsMeta: { name: string; pos: number; type: number }[] = [];
  let sst: string[] = [];
  for (let i = 0; i < recs.length; i++) {
    const { type, data } = recs[i];
    const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
    if (type === 0x000a) break; // end of globals
    switch (type) {
      case 0x002f:
        throw new ConvertError('Книга Excel защищена паролем. Снимите защиту и попробуйте снова.', 'password');
      case 0x0022:
        date1904 = dv.getUint16(0, true) === 1;
        break;
      case 0x041e:
        formats.set(dv.getUint16(0, true), xlString(data, 2, 2).s);
        break;
      case 0x00e0:
        xfFormat.push(dv.getUint16(2, true));
        break;
      case 0x0085:
        sheetsMeta.push({ pos: dv.getUint32(0, true), type: data[5], name: xlString(data, 6, 1).s });
        break;
      case 0x00fc: {
        const segs = [data];
        while (recs[i + 1]?.type === 0x003c) segs.push(recs[++i].data);
        const r = new Segments(segs);
        r.skip(4);
        const unique = r.i32();
        sst = [];
        for (let k = 0; k < unique && !r.done; k++) sst.push(r.richString());
        break;
      }
    }
  }
  const isDate = (ixfe: number): boolean => {
    const f = xfFormat[ixfe];
    return f !== undefined && (BUILTIN_DATE_FORMATS.has(f) || (formats.has(f) && isDateFormat(formats.get(f)!)));
  };
  const num = (ixfe: number, v: number): CellValue => (isDate(ixfe) ? excelDate(v, date1904) : v);
  const sheets: Sheet[] = [];
  for (const meta of sheetsMeta) {
    if (meta.type !== 0) continue; // charts, macros
    const rows: CellValue[][] = [];
    const set = (r: number, c: number, v: CellValue): void => {
      if (r > 1_048_575 || c > 16_383) return;
      while (rows.length <= r) rows.push([]);
      const row = rows[r];
      while (row.length < c) row.push(null);
      row[c] = v;
    };
    const sr = records(stream, meta.pos);
    for (let i = 1; i < sr.length; i++) {
      const { type, data } = sr[i];
      const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
      if (data.length < 6 && type !== 0x000a) continue;
      switch (type) {
        case 0x00fd: // LABELSST
          set(dv.getUint16(0, true), dv.getUint16(2, true), sst[dv.getUint32(6, true)] ?? '');
          break;
        case 0x0204: // LABEL
          set(dv.getUint16(0, true), dv.getUint16(2, true), xlString(data, 6, 2).s);
          break;
        case 0x0203: // NUMBER
          set(dv.getUint16(0, true), dv.getUint16(2, true), num(dv.getUint16(4, true), dv.getFloat64(6, true)));
          break;
        case 0x027e: // RK
          set(dv.getUint16(0, true), dv.getUint16(2, true), num(dv.getUint16(4, true), rk(dv.getUint32(6, true))));
          break;
        case 0x00bd: {
          // MULRK
          const r = dv.getUint16(0, true);
          let c = dv.getUint16(2, true);
          for (let p = 4; p + 6 <= data.length - 2; p += 6) set(r, c++, num(dv.getUint16(p, true), rk(dv.getUint32(p + 2, true))));
          break;
        }
        case 0x0205: // BOOLERR
          if (data[7] === 0) set(dv.getUint16(0, true), dv.getUint16(2, true), data[6] !== 0);
          break;
        case 0x0006: {
          // FORMULA: cached result
          const r = dv.getUint16(0, true);
          const c = dv.getUint16(2, true);
          if (dv.getUint16(12, true) === 0xffff) {
            const kind = data[6];
            if (kind === 0) {
              const next = sr[i + 1];
              if (next?.type === 0x0207) {
                set(r, c, xlString(next.data, 0, 2).s);
                i++;
              }
            } else if (kind === 1) set(r, c, data[8] !== 0);
          } else set(r, c, num(dv.getUint16(4, true), dv.getFloat64(6, true)));
          break;
        }
      }
    }
    while (rows.length && rows[rows.length - 1].every((v) => v === null || v === '')) rows.pop();
    sheets.push({ name: meta.name || `Лист${sheets.length + 1}`, rows });
  }
  if (!sheets.length) throw new ConvertError('В книге Excel не найдено листов.');
  return { sheets };
}
