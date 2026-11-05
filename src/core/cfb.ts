/**
 * Compound File Binary (OLE2) reader – the container of legacy Office files
 * (.doc, .xls, .ppt). Implements MS-CFB: FAT/DIFAT chains, directory tree and the
 * mini stream for small streams.
 */
import { ConvertError } from './util';

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export interface CfbEntry {
  name: string;
  type: number; // 1 storage, 2 stream, 5 root
  start: number;
  size: number;
  left: number;
  right: number;
  child: number;
}

export class Cfb {
  private dv: DataView;
  private sectorSize: number;
  private miniSectorSize: number;
  private miniCutoff: number;
  private fat: number[] = [];
  private miniFat: number[] = [];
  private miniStream: Uint8Array = new Uint8Array(0);
  readonly entries: CfbEntry[] = [];

  constructor(private data: Uint8Array) {
    if (data.length < 512 || SIGNATURE.some((b, i) => data[i] !== b)) throw new ConvertError('Файл не является документом Office старого формата (OLE2).');
    this.dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
    this.sectorSize = 1 << this.dv.getUint16(0x1e, true);
    this.miniSectorSize = 1 << this.dv.getUint16(0x20, true);
    this.miniCutoff = this.dv.getUint32(0x38, true);
    if (this.sectorSize !== 512 && this.sectorSize !== 4096) throw new ConvertError('Повреждённый документ Office: неверный размер сектора.');
    this.readFat();
    this.readDirectory();
    const root = this.entries[0];
    if (root) this.miniStream = this.chain(root.start, root.size, false);
    const miniFatStart = this.dv.getUint32(0x3c, true);
    const miniFatCount = this.dv.getUint32(0x40, true);
    if (miniFatCount && miniFatStart !== ENDOFCHAIN) {
      const bytes = this.chain(miniFatStart, miniFatCount * this.sectorSize, false);
      const mdv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      for (let i = 0; i + 4 <= bytes.length; i += 4) this.miniFat.push(mdv.getUint32(i, true));
    }
  }

  private sectorOffset(n: number): number {
    return (n + 1) * this.sectorSize;
  }

  private readFat(): void {
    const difat: number[] = [];
    for (let i = 0; i < 109; i++) {
      const s = this.dv.getUint32(0x4c + i * 4, true);
      if (s !== FREESECT && s !== ENDOFCHAIN) difat.push(s);
    }
    let next = this.dv.getUint32(0x44, true);
    let guard = 0;
    const perSector = this.sectorSize / 4 - 1;
    while (next !== ENDOFCHAIN && next !== FREESECT && guard++ < 50_000) {
      const off = this.sectorOffset(next);
      if (off + this.sectorSize > this.data.length) break;
      for (let i = 0; i < perSector; i++) {
        const s = this.dv.getUint32(off + i * 4, true);
        if (s !== FREESECT && s !== ENDOFCHAIN) difat.push(s);
      }
      next = this.dv.getUint32(off + perSector * 4, true);
    }
    for (const s of difat) {
      const off = this.sectorOffset(s);
      if (off + this.sectorSize > this.data.length) continue;
      for (let i = 0; i < this.sectorSize / 4; i++) this.fat.push(this.dv.getUint32(off + i * 4, true));
    }
  }

  /** Read a chain of sectors (regular or mini) up to `size` bytes. */
  private chain(start: number, size: number, mini: boolean): Uint8Array {
    const unit = mini ? this.miniSectorSize : this.sectorSize;
    const table = mini ? this.miniFat : this.fat;
    const out = new Uint8Array(Math.max(0, size));
    let pos = 0;
    let s = start;
    let guard = 0;
    const limit = Math.ceil(size / unit) + 2;
    while (s !== ENDOFCHAIN && s !== FREESECT && pos < size && guard++ < limit) {
      const len = Math.min(unit, size - pos);
      if (mini) {
        const off = s * unit;
        if (off + len > this.miniStream.length) break;
        out.set(this.miniStream.subarray(off, off + len), pos);
      } else {
        const off = this.sectorOffset(s);
        if (off >= this.data.length) break;
        out.set(this.data.subarray(off, Math.min(off + len, this.data.length)), pos);
      }
      pos += len;
      s = table[s] ?? ENDOFCHAIN;
    }
    return out;
  }

  private readDirectory(): void {
    const start = this.dv.getUint32(0x30, true);
    let s = start;
    let guard = 0;
    while (s !== ENDOFCHAIN && s !== FREESECT && guard++ < 100_000) {
      const off = this.sectorOffset(s);
      if (off + this.sectorSize > this.data.length) break;
      for (let i = 0; i < this.sectorSize / 128; i++) {
        const e = off + i * 128;
        const nameLen = this.dv.getUint16(e + 0x40, true);
        let name = '';
        for (let k = 0; k + 2 < nameLen && k < 64; k += 2) name += String.fromCharCode(this.dv.getUint16(e + k, true));
        this.entries.push({
          name,
          type: this.data[e + 0x42],
          left: this.dv.getUint32(e + 0x44, true),
          right: this.dv.getUint32(e + 0x48, true),
          child: this.dv.getUint32(e + 0x4c, true),
          start: this.dv.getUint32(e + 0x74, true),
          size: this.dv.getUint32(e + 0x78, true),
        });
      }
      s = this.fat[s] ?? ENDOFCHAIN;
    }
  }

  /** Stream by name (top-level storages are searched too; names are case-insensitive). */
  stream(name: string): Uint8Array | null {
    const e = this.entries.find((x) => x.type === 2 && x.name.toLowerCase() === name.toLowerCase());
    if (!e) return null;
    return this.chain(e.start, e.size, e.size < this.miniCutoff);
  }

  has(name: string): boolean {
    return this.entries.some((x) => x.name.toLowerCase() === name.toLowerCase());
  }
}

export const isCfb = (b: Uint8Array): boolean => b.length >= 8 && SIGNATURE.every((v, i) => b[i] === v);
