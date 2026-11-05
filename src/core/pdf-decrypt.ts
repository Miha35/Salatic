/**
 * PDF Standard Security Handler (ISO 32000-2 §7.6): removes encryption so that
 * pdf-lib based tools can work with protected files. Supports RC4 40/128-bit
 * (R2–R4), AES-128 (AESV2) and AES-256 (R5 and R6). Files that only restrict
 * permissions (empty user password) open without asking; otherwise the caller
 * must supply the user or owner password.
 *
 * Works on raw bytes: objects are scanned sequentially, encrypted strings are
 * rewritten as hex strings and stream data is replaced (with a direct /Length).
 * Cross-reference offsets become stale, which is fine for pdf-lib (it scans
 * objects itself) and for pdf.js (it reconstructs the table).
 */
import { aesCbcDecrypt, aesCbcEncrypt, md5, rc4, sha } from './crypto';
import { ConvertError } from './util';

type PVal =
  | { t: 'dict'; s: number; e: number; entries: { key: string; ks: number; v: PVal }[] }
  | { t: 'arr'; s: number; e: number; items: PVal[] }
  | { t: 'str'; s: number; e: number; bytes: Uint8Array }
  | { t: 'name'; s: number; e: number; v: string }
  | { t: 'num'; s: number; e: number; v: number }
  | { t: 'ref'; s: number; e: number; n: number; g: number }
  | { t: 'kw'; s: number; e: number; v: string };

interface PObj {
  n: number;
  g: number;
  value: PVal;
  stream?: { s: number; e: number };
}

const isWs = (c: number): boolean => c === 0 || c === 9 || c === 10 || c === 12 || c === 13 || c === 32;
const isDelim = (c: number): boolean => c === 40 || c === 41 || c === 60 || c === 62 || c === 91 || c === 93 || c === 123 || c === 125 || c === 47 || c === 37;
const isDigit = (c: number): boolean => c >= 48 && c <= 57;

class Lexer {
  pos = 0;
  constructor(readonly d: Uint8Array) {}

  skipWs(): void {
    const d = this.d;
    while (this.pos < d.length) {
      const c = d[this.pos];
      if (isWs(c)) this.pos++;
      else if (c === 37) {
        while (this.pos < d.length && d[this.pos] !== 10 && d[this.pos] !== 13) this.pos++;
      } else break;
    }
  }

  private word(): string {
    const s = this.pos;
    while (this.pos < this.d.length && !isWs(this.d[this.pos]) && !isDelim(this.d[this.pos])) this.pos++;
    return String.fromCharCode(...this.d.subarray(s, this.pos));
  }

  private tryInt(): number | null {
    const s = this.pos;
    while (this.pos < this.d.length && isDigit(this.d[this.pos])) this.pos++;
    if (this.pos === s) return null;
    return +String.fromCharCode(...this.d.subarray(s, this.pos));
  }

  value(depth = 0): PVal {
    if (depth > 200) throw new ConvertError('PDF повреждён: слишком глубокая вложенность объектов.');
    this.skipWs();
    const d = this.d;
    const s = this.pos;
    const c = d[s];
    if (c === 60 && d[s + 1] === 60) {
      this.pos += 2;
      const entries: { key: string; ks: number; v: PVal }[] = [];
      for (;;) {
        this.skipWs();
        if (this.pos >= d.length) break;
        if (d[this.pos] === 62 && d[this.pos + 1] === 62) {
          this.pos += 2;
          break;
        }
        if (d[this.pos] !== 47) {
          // malformed entry: skip a token
          this.value(depth + 1);
          continue;
        }
        const ks = this.pos;
        this.pos++;
        const key = this.word();
        const v = this.value(depth + 1);
        entries.push({ key, ks, v });
      }
      return { t: 'dict', s, e: this.pos, entries };
    }
    if (c === 60) {
      this.pos++;
      let hex = '';
      while (this.pos < d.length && d[this.pos] !== 62) {
        const ch = d[this.pos++];
        if (!isWs(ch)) hex += String.fromCharCode(ch);
      }
      this.pos++;
      if (hex.length % 2) hex += '0';
      const bytes = new Uint8Array(hex.length / 2);
      for (let i = 0; i < bytes.length; i++) bytes[i] = parseInt(hex.substr(i * 2, 2), 16) || 0;
      return { t: 'str', s, e: this.pos, bytes };
    }
    if (c === 40) {
      this.pos++;
      const out: number[] = [];
      let depthP = 1;
      while (this.pos < d.length) {
        const ch = d[this.pos++];
        if (ch === 92) {
          const n = d[this.pos++];
          const map: Record<number, number> = { 110: 10, 114: 13, 116: 9, 98: 8, 102: 12, 40: 40, 41: 41, 92: 92 };
          if (map[n] !== undefined) out.push(map[n]);
          else if (n >= 48 && n <= 55) {
            let v = n - 48;
            for (let k = 0; k < 2 && d[this.pos] >= 48 && d[this.pos] <= 55; k++) v = v * 8 + (d[this.pos++] - 48);
            out.push(v & 255);
          } else if (n === 13) {
            if (d[this.pos] === 10) this.pos++;
          } else if (n !== 10) out.push(n);
          continue;
        }
        if (ch === 40) depthP++;
        else if (ch === 41 && --depthP === 0) break;
        out.push(ch);
      }
      return { t: 'str', s, e: this.pos, bytes: Uint8Array.from(out) };
    }
    if (c === 91) {
      this.pos++;
      const items: PVal[] = [];
      for (;;) {
        this.skipWs();
        if (this.pos >= d.length) break;
        if (d[this.pos] === 93) {
          this.pos++;
          break;
        }
        items.push(this.value(depth + 1));
      }
      return { t: 'arr', s, e: this.pos, items };
    }
    if (c === 47) {
      this.pos++;
      return { t: 'name', s, e: this.pos, v: this.word() };
    }
    if (isDigit(c) || c === 43 || c === 45 || c === 46) {
      const w = this.word();
      const num = { t: 'num' as const, s, e: this.pos, v: parseFloat(w) || 0 };
      // "n g R" reference
      if (/^\d+$/.test(w)) {
        const save = this.pos;
        this.skipWs();
        const g = this.tryInt();
        if (g !== null) {
          this.skipWs();
          if (d[this.pos] === 82 && (this.pos + 1 >= d.length || isWs(d[this.pos + 1]) || isDelim(d[this.pos + 1]))) {
            this.pos++;
            return { t: 'ref', s, e: this.pos, n: +w, g };
          }
        }
        this.pos = save;
      }
      return num;
    }
    if (c === undefined) return { t: 'kw', s, e: s, v: '' };
    if (c === 41 || c === 62 || c === 93 || c === 125 || c === 123) {
      this.pos++;
      return { t: 'kw', s, e: this.pos, v: String.fromCharCode(c) };
    }
    return { t: 'kw', s, e: this.pos, v: this.word() || (this.pos++, '') };
  }
}

const dictGet = (v: PVal | undefined, key: string): PVal | undefined => (v?.t === 'dict' ? v.entries.find((e) => e.key === key)?.v : undefined);

interface Scan {
  objects: PObj[];
  /** Latest definition of each object number. */
  byNum: Map<number, PObj>;
  trailers: PVal[];
}

const OBJ_RE = /(\d+)[\0\t\n\f\r ]+(\d+)[\0\t\n\f\r ]+obj\b/y;

function scan(d: Uint8Array): Scan {
  const text = latin1(d);
  const lx = new Lexer(d);
  const objects: PObj[] = [];
  const byNum = new Map<number, PObj>();
  const trailers: PVal[] = [];
  let pos = 0;
  while (pos < d.length) {
    const c = d[pos];
    if (isDigit(c)) {
      OBJ_RE.lastIndex = pos;
      const m = OBJ_RE.exec(text);
      if (m) {
        lx.pos = pos + m[0].length;
        const value = lx.value();
        const obj: PObj = { n: +m[1], g: +m[2], value };
        lx.skipWs();
        if (text.startsWith('stream', lx.pos)) {
          let s = lx.pos + 6;
          if (d[s] === 13 && d[s + 1] === 10) s += 2;
          else if (d[s] === 10 || d[s] === 13) s += 1;
          const len = dictGet(value, 'Length');
          let e = -1;
          if (len?.t === 'num') {
            const cand = s + len.v;
            const after = text.slice(cand, cand + 12).replace(/^[\r\n ]+/, '');
            if (after.startsWith('endstream')) e = cand;
          }
          if (e < 0) {
            const k = text.indexOf('endstream', s);
            e = k < 0 ? d.length : k;
            // trailing EOL before endstream is not data
            if (d[e - 1] === 10) e--;
            if (d[e - 1] === 13) e--;
          }
          obj.stream = { s, e };
          const k = text.indexOf('endstream', e);
          lx.pos = k < 0 ? d.length : k + 9;
        }
        lx.skipWs();
        if (text.startsWith('endobj', lx.pos)) lx.pos += 6;
        objects.push(obj);
        byNum.set(obj.n, obj);
        pos = Math.max(lx.pos, pos + 1);
        continue;
      }
    }
    if (c === 116 && text.startsWith('trailer', pos)) {
      lx.pos = pos + 7;
      const t = lx.value();
      if (t.t === 'dict') trailers.push(t);
      pos = Math.max(lx.pos, pos + 1);
      continue;
    }
    pos++;
  }
  // Cross-reference streams carry the trailer entries
  for (const o of objects) if (o.stream && dictGet(o.value, 'Type')?.t === 'name' && (dictGet(o.value, 'Type') as { v: string }).v === 'XRef') trailers.push(o.value);
  return { objects, byNum, trailers };
}

function latin1(d: Uint8Array): string {
  let s = '';
  for (let i = 0; i < d.length; i += 0x8000) s += String.fromCharCode(...d.subarray(i, i + 0x8000));
  return s;
}

/** Quick byte-level check: does the file declare an /Encrypt dictionary (direct or by reference)? */
export function isEncryptedPdf(d: Uint8Array): boolean {
  const pat = [47, 69, 110, 99, 114, 121, 112, 116]; // "/Encrypt"
  outer: for (let i = d.indexOf(47); i >= 0 && i + 9 < d.length; i = d.indexOf(47, i + 1)) {
    for (let k = 1; k < 8; k++) if (d[i + k] !== pat[k]) continue outer;
    let j = i + 8;
    if (!isWs(d[j]) && d[j] !== 60 && !isDigit(d[j])) continue; // e.g. "/EncryptMetadata"
    while (j < d.length && isWs(d[j])) j++;
    if (isDigit(d[j]) || (d[j] === 60 && d[j + 1] === 60)) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// Keys

const PAD = Uint8Array.from([0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08, 0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a]);

const cat = (...parts: Uint8Array<ArrayBufferLike>[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
};
const eq = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((v, i) => v === b[i]);

function padPassword(pw: string): Uint8Array {
  const bytes = Uint8Array.from([...pw].map((ch) => (ch.charCodeAt(0) < 256 ? ch.charCodeAt(0) : 63))).subarray(0, 32);
  return cat(bytes, PAD.subarray(0, 32 - bytes.length));
}

interface Handler {
  R: number;
  V: number;
  keyLen: number;
  O: Uint8Array;
  U: Uint8Array;
  OE?: Uint8Array;
  UE?: Uint8Array;
  P: number;
  id0: Uint8Array;
  encryptMetadata: boolean;
  stm: 'rc4' | 'aes' | 'aes256' | 'none';
  str: 'rc4' | 'aes' | 'aes256' | 'none';
}

/** Algorithm 2: file key from a (user) password, R2–R4. */
function fileKeyR4(h: Handler, padded: Uint8Array): Uint8Array {
  const p = new Uint8Array(4);
  new DataView(p.buffer).setInt32(0, h.P, true);
  let hash = md5(cat(padded, h.O.subarray(0, 32), p, h.id0, h.R >= 4 && !h.encryptMetadata ? Uint8Array.from([255, 255, 255, 255]) : new Uint8Array(0)));
  const n = h.R === 2 ? 5 : h.keyLen;
  if (h.R >= 3) for (let i = 0; i < 50; i++) hash = md5(hash.subarray(0, n));
  return hash.subarray(0, n);
}

/** Algorithms 4/5: does `key` reproduce /U? */
function userOk(h: Handler, key: Uint8Array): boolean {
  if (h.R === 2) return eq(rc4(key, PAD), h.U.subarray(0, 32));
  let x = rc4(key, md5(cat(PAD, h.id0)));
  for (let i = 1; i <= 19; i++) x = rc4(key.map((b) => b ^ i), x);
  return eq(x, h.U.subarray(0, 16));
}

/** Algorithm 7: recover the user password from /O with the owner password. */
function userFromOwner(h: Handler, ownerPw: string): Uint8Array {
  let hash = md5(padPassword(ownerPw));
  if (h.R >= 3) for (let i = 0; i < 50; i++) hash = md5(hash);
  const key = hash.subarray(0, h.R === 2 ? 5 : h.keyLen);
  if (h.R === 2) return rc4(key, h.O.subarray(0, 32));
  let x = h.O.subarray(0, 32);
  for (let i = 19; i >= 0; i--) x = rc4(key.map((b) => b ^ i), x);
  return x;
}

/** Algorithm 2.B (R6) / plain SHA-256 (R5). */
async function hash2B(h: Handler, pw: Uint8Array, salt: Uint8Array, udata: Uint8Array): Promise<Uint8Array> {
  let k = await sha(256, cat(pw, salt, udata));
  if (h.R === 5) return k;
  for (let round = 0; ; round++) {
    const k1one = cat(pw, k, udata);
    const k1 = new Uint8Array(k1one.length * 64);
    for (let i = 0; i < 64; i++) k1.set(k1one, i * k1one.length);
    const e = aesCbcEncrypt(k.subarray(0, 16), k.subarray(16, 32), k1);
    let sum = 0;
    for (let i = 0; i < 16; i++) sum += e[i];
    const mod = sum % 3;
    k = await sha(mod === 0 ? 256 : mod === 1 ? 384 : 512, e);
    if (round >= 63 && e[e.length - 1] <= round - 31) break;
  }
  return k.subarray(0, 32);
}

async function fileKeyR6(h: Handler, password: string): Promise<Uint8Array | null> {
  const pw = new TextEncoder().encode(password.normalize('NFKC')).subarray(0, 127);
  const U = h.U;
  const O = h.O;
  if (eq(await hash2B(h, pw, U.subarray(32, 40), new Uint8Array(0)), U.subarray(0, 32))) {
    const ik = await hash2B(h, pw, U.subarray(40, 48), new Uint8Array(0));
    return aesCbcDecrypt(ik, new Uint8Array(16), h.UE!.subarray(0, 32), false);
  }
  if (eq(await hash2B(h, pw, O.subarray(32, 40), U.subarray(0, 48)), O.subarray(0, 32))) {
    const ik = await hash2B(h, pw, O.subarray(40, 48), U.subarray(0, 48));
    return aesCbcDecrypt(ik, new Uint8Array(16), h.OE!.subarray(0, 32), false);
  }
  return null;
}

async function fileKey(h: Handler, password: string): Promise<Uint8Array | null> {
  if (h.R >= 5) return fileKeyR6(h, password);
  const asUser = fileKeyR4(h, padPassword(password));
  if (userOk(h, asUser)) return asUser;
  const user = userFromOwner(h, password);
  const asOwner = fileKeyR4(h, cat(user, PAD).subarray(0, 32));
  return userOk(h, asOwner) ? asOwner : null;
}

function objectKey(h: Handler, key: Uint8Array, n: number, g: number, aes: boolean): Uint8Array {
  const ext = Uint8Array.from([n & 255, (n >> 8) & 255, (n >> 16) & 255, g & 255, (g >> 8) & 255]);
  const k = md5(cat(key, ext, aes ? Uint8Array.from([0x73, 0x41, 0x6c, 0x54]) : new Uint8Array(0)));
  return k.subarray(0, Math.min(16, key.length + 5));
}

/** Bulk AES-CBC through WebCrypto (fast); falls back to the JS cipher on padding errors. */
async function aesDecrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  if (data.length >= 512 && data.length % 16 === 0 && globalThis.crypto?.subtle) {
    try {
      const k = await crypto.subtle.importKey('raw', key as BufferSource, 'AES-CBC', false, ['decrypt']);
      return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-CBC', iv: iv as BufferSource }, k, data as BufferSource));
    } catch {
      /* invalid padding: JS path below keeps the data */
    }
  }
  return aesCbcDecrypt(key, iv, data);
}

async function decryptData(h: Handler, key: Uint8Array, method: Handler['stm'], n: number, g: number, data: Uint8Array): Promise<Uint8Array> {
  switch (method) {
    case 'none':
      return data;
    case 'rc4':
      return rc4(objectKey(h, key, n, g, false), data);
    case 'aes':
    case 'aes256': {
      if (data.length < 32) return data.length === 16 ? new Uint8Array(0) : data;
      const k = method === 'aes256' ? key : objectKey(h, key, n, g, true);
      return aesDecrypt(k, data.subarray(0, 16), data.subarray(16));
    }
  }
}

// ---------------------------------------------------------------------------

function handlerOf(enc: PVal, id0: Uint8Array, scanned: Scan): Handler {
  const resolve = (v: PVal | undefined): PVal | undefined => (v?.t === 'ref' ? scanned.byNum.get(v.n)?.value : v);
  const num = (k: string, def = 0): number => {
    const v = resolve(dictGet(enc, k));
    return v?.t === 'num' ? v.v : def;
  };
  const str = (k: string): Uint8Array | undefined => {
    const v = resolve(dictGet(enc, k));
    return v?.t === 'str' ? v.bytes : undefined;
  };
  const filter = resolve(dictGet(enc, 'Filter'));
  if (filter?.t !== 'name' || filter.v !== 'Standard') throw new ConvertError('PDF защищён сертификатом или нестандартным способом шифрования — откройте его в программе, которой он создан, и сохраните без защиты.', 'password_unsupported');
  const V = num('V');
  const R = num('R');
  const O = str('O');
  const U = str('U');
  if (!O || !U || ![1, 2, 4, 5].includes(V)) throw new ConvertError('Неподдерживаемая версия шифрования PDF.', 'password_unsupported');
  const emd = resolve(dictGet(enc, 'EncryptMetadata'));
  const h: Handler = { R, V, keyLen: Math.max(5, Math.min(16, num('Length', 40) / 8)), O, U, OE: str('OE'), UE: str('UE'), P: num('P') | 0, id0, encryptMetadata: !(emd?.t === 'kw' && emd.v === 'false'), stm: 'rc4', str: 'rc4' };
  if (V === 1) h.keyLen = 5;
  if (V >= 4) {
    const cf = resolve(dictGet(enc, 'CF'));
    const method = (name: string): Handler['stm'] => {
      const fname = resolve(dictGet(enc, name));
      if (fname?.t !== 'name' || fname.v === 'Identity') return 'none';
      const f = resolve(dictGet(cf, fname.v));
      const cfm = resolve(dictGet(f, 'CFM'));
      const len = resolve(dictGet(f, 'Length'));
      if (len?.t === 'num' && V === 4) h.keyLen = len.v > 32 ? len.v / 8 : len.v;
      if (cfm?.t !== 'name' || cfm.v === 'None') return 'none';
      return cfm.v === 'AESV3' ? 'aes256' : cfm.v === 'AESV2' ? 'aes' : 'rc4';
    };
    h.stm = method('StmF');
    h.str = method('StrF');
  }
  if (V === 5) h.keyLen = 32;
  return h;
}

/**
 * Remove the encryption of a PDF. Throws ConvertError with code
 * `password_required` (no password given and one is needed) or
 * `password_invalid` (the given password is wrong).
 */
export async function decryptPdf(bytes: Uint8Array, password = ''): Promise<Uint8Array> {
  const scanned = scan(bytes);
  let encVal: PVal | undefined;
  const encEntry: { s: number; e: number }[] = [];
  let id0: Uint8Array = new Uint8Array(0);
  for (const t of scanned.trailers) {
    if (t.t !== 'dict') continue;
    const e = t.entries.find((x) => x.key === 'Encrypt');
    if (e) {
      encVal = e.v.t === 'ref' ? scanned.byNum.get(e.v.n)?.value : e.v;
      encEntry.push({ s: e.ks, e: e.v.e });
    }
    const id = dictGet(t, 'ID');
    if (id?.t === 'arr' && id.items[0]?.t === 'str') id0 = id.items[0].bytes;
  }
  if (!encVal || encVal.t !== 'dict') return bytes;
  const encRefNum = (() => {
    for (const t of scanned.trailers) {
      const e = dictGet(t, 'Encrypt');
      if (e?.t === 'ref') return e.n;
    }
    return -1;
  })();
  const h = handlerOf(encVal, id0, scanned);
  const key = await fileKey(h, password);
  if (!key) throw new ConvertError(password ? 'Неверный пароль PDF.' : 'PDF защищён паролем. Введите пароль, чтобы открыть его.', password ? 'password_invalid' : 'password_required');

  const edits: { s: number; e: number; data: Uint8Array<ArrayBufferLike> }[] = [];
  const hex = (b: Uint8Array): Uint8Array => {
    let s = '<';
    for (const x of b) s += x.toString(16).padStart(2, '0');
    return new TextEncoder().encode(s + '>');
  };
  const visit = async (v: PVal, n: number, g: number, skipContents: boolean): Promise<void> => {
    if (v.t === 'str') edits.push({ s: v.s, e: v.e, data: hex(await decryptData(h, key, h.str, n, g, v.bytes)) });
    else if (v.t === 'arr') for (const x of v.items) await visit(x, n, g, false);
    else if (v.t === 'dict') {
      const isSig = v.entries.some((x) => x.key === 'ByteRange');
      for (const e of v.entries) if (!(skipContents && isSig && e.key === 'Contents')) await visit(e.v, n, g, true);
    }
  };
  const latest = new Set(scanned.byNum.values());
  for (const o of scanned.objects) {
    if (o.n === encRefNum && !latest.has(o) && o.value === encVal) continue;
    if (o.value === encVal) continue;
    const type = dictGet(o.value, 'Type');
    const typeName = type?.t === 'name' ? type.v : '';
    if (o.stream && typeName === 'XRef') continue;
    await visit(o.value, o.n, o.g, true);
    if (o.stream) {
      if (typeName === 'Metadata' && !h.encryptMetadata) continue;
      const data = await decryptData(h, key, h.stm, o.n, o.g, bytes.subarray(o.stream.s, o.stream.e));
      edits.push({ s: o.stream.s, e: o.stream.e, data });
      const len = o.value.t === 'dict' ? o.value.entries.find((x) => x.key === 'Length') : undefined;
      if (len) edits.push({ s: len.v.s, e: len.v.e, data: new TextEncoder().encode(String(data.length)) });
    }
  }
  // Drop /Encrypt from every trailer so readers treat the file as plain
  for (const e of encEntry) edits.push({ s: e.s, e: e.e, data: new Uint8Array(e.e - e.s).fill(32) });
  edits.sort((a, b) => a.s - b.s);
  const parts: Uint8Array<ArrayBufferLike>[] = [];
  let pos = 0;
  for (const ed of edits) {
    if (ed.s < pos) continue; // nested/overlapping edit (cannot happen for well-formed files)
    parts.push(bytes.subarray(pos, ed.s), ed.data);
    pos = ed.e;
  }
  parts.push(bytes.subarray(pos));
  return cat(...parts);
}
