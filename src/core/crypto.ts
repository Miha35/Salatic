/**
 * Small synchronous primitives needed by the PDF Standard Security Handler:
 * MD5 (RFC 1321), RC4 and AES (FIPS-197) with CBC. SHA-2 comes from WebCrypto.
 */

// ---------------------------------------------------------------------------
// MD5

const S = [7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21];
const K = Array.from({ length: 64 }, (_, i) => Math.floor(Math.abs(Math.sin(i + 1)) * 2 ** 32) >>> 0);

export function md5(data: Uint8Array): Uint8Array {
  const len = data.length;
  const total = ((len + 8) >> 6) * 64 + 64;
  const buf = new Uint8Array(total);
  buf.set(data);
  buf[len] = 0x80;
  const bits = len * 8;
  const dv = new DataView(buf.buffer);
  dv.setUint32(total - 8, bits >>> 0, true);
  dv.setUint32(total - 4, Math.floor(bits / 2 ** 32), true);
  let a0 = 0x67452301;
  let b0 = 0xefcdab89;
  let c0 = 0x98badcfe;
  let d0 = 0x10325476;
  const M = new Uint32Array(16);
  for (let off = 0; off < total; off += 64) {
    for (let i = 0; i < 16; i++) M[i] = dv.getUint32(off + i * 4, true);
    let A = a0;
    let B = b0;
    let C = c0;
    let D = d0;
    for (let i = 0; i < 64; i++) {
      let F: number;
      let g: number;
      if (i < 16) {
        F = (B & C) | (~B & D);
        g = i;
      } else if (i < 32) {
        F = (D & B) | (~D & C);
        g = (5 * i + 1) % 16;
      } else if (i < 48) {
        F = B ^ C ^ D;
        g = (3 * i + 5) % 16;
      } else {
        F = C ^ (B | ~D);
        g = (7 * i) % 16;
      }
      F = (F + A + K[i] + M[g]) >>> 0;
      A = D;
      D = C;
      C = B;
      B = (B + ((F << S[i]) | (F >>> (32 - S[i])))) >>> 0;
    }
    a0 = (a0 + A) >>> 0;
    b0 = (b0 + B) >>> 0;
    c0 = (c0 + C) >>> 0;
    d0 = (d0 + D) >>> 0;
  }
  const out = new Uint8Array(16);
  const odv = new DataView(out.buffer);
  [a0, b0, c0, d0].forEach((v, i) => odv.setUint32(i * 4, v, true));
  return out;
}

// ---------------------------------------------------------------------------
// RC4

export function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) s[i] = i;
  for (let i = 0, j = 0; i < 256; i++) {
    j = (j + s[i] + key[i % key.length]) & 255;
    [s[i], s[j]] = [s[j], s[i]];
  }
  const out = new Uint8Array(data.length);
  for (let n = 0, i = 0, j = 0; n < data.length; n++) {
    i = (i + 1) & 255;
    j = (j + s[i]) & 255;
    [s[i], s[j]] = [s[j], s[i]];
    out[n] = data[n] ^ s[(s[i] + s[j]) & 255];
  }
  return out;
}

// ---------------------------------------------------------------------------
// AES

const SBOX = new Uint8Array(256);
const INV = new Uint8Array(256);
(() => {
  // Build the S-box from the multiplicative inverse in GF(2^8) and the affine map
  let p = 1;
  let q = 1;
  do {
    p = p ^ ((p << 1) & 0xff) ^ (p & 0x80 ? 0x1b : 0);
    q ^= q << 1;
    q ^= q << 2;
    q ^= q << 4;
    q &= 0xff;
    if (q & 0x80) q ^= 0x09;
    const x = q ^ ((q << 1) | (q >> 7)) ^ ((q << 2) | (q >> 6)) ^ ((q << 3) | (q >> 5)) ^ ((q << 4) | (q >> 4));
    SBOX[p] = (x ^ 0x63) & 0xff;
  } while (p !== 1);
  SBOX[0] = 0x63;
  for (let i = 0; i < 256; i++) INV[SBOX[i]] = i;
})();

const xt = (b: number): number => ((b << 1) ^ (b & 0x80 ? 0x1b : 0)) & 0xff;
const mul = (a: number, b: number): number => {
  let r = 0;
  while (b) {
    if (b & 1) r ^= a;
    a = xt(a);
    b >>= 1;
  }
  return r;
};

function expandKey(key: Uint8Array): Uint8Array {
  const nk = key.length / 4;
  const rounds = nk + 6;
  const w = new Uint8Array(16 * (rounds + 1));
  w.set(key);
  let rcon = 1;
  for (let i = nk; i < 4 * (rounds + 1); i++) {
    const t = w.slice((i - 1) * 4, i * 4);
    if (i % nk === 0) {
      const t0 = t[0];
      t[0] = SBOX[t[1]] ^ rcon;
      t[1] = SBOX[t[2]];
      t[2] = SBOX[t[3]];
      t[3] = SBOX[t0];
      rcon = xt(rcon);
    } else if (nk > 6 && i % nk === 4) for (let k = 0; k < 4; k++) t[k] = SBOX[t[k]];
    for (let k = 0; k < 4; k++) w[i * 4 + k] = w[(i - nk) * 4 + k] ^ t[k];
  }
  return w;
}

function encryptBlock(w: Uint8Array, rounds: number, b: Uint8Array): Uint8Array {
  const s = Uint8Array.from(b);
  for (let i = 0; i < 16; i++) s[i] ^= w[i];
  for (let r = 1; r <= rounds; r++) {
    for (let i = 0; i < 16; i++) s[i] = SBOX[s[i]];
    // ShiftRows (column-major state)
    const t = s.slice();
    for (let c = 0; c < 4; c++) for (let rr = 0; rr < 4; rr++) s[c * 4 + rr] = t[((c + rr) % 4) * 4 + rr];
    if (r < rounds) {
      for (let c = 0; c < 4; c++) {
        const [a0, a1, a2, a3] = [s[c * 4], s[c * 4 + 1], s[c * 4 + 2], s[c * 4 + 3]];
        s[c * 4] = xt(a0) ^ xt(a1) ^ a1 ^ a2 ^ a3;
        s[c * 4 + 1] = a0 ^ xt(a1) ^ xt(a2) ^ a2 ^ a3;
        s[c * 4 + 2] = a0 ^ a1 ^ xt(a2) ^ xt(a3) ^ a3;
        s[c * 4 + 3] = xt(a0) ^ a0 ^ a1 ^ a2 ^ xt(a3);
      }
    }
    for (let i = 0; i < 16; i++) s[i] ^= w[r * 16 + i];
  }
  return s;
}

function decryptBlock(w: Uint8Array, rounds: number, b: Uint8Array): Uint8Array {
  // Uint8Array.from copies even when `b` is a Node Buffer (whose slice() is a view)
  const s = Uint8Array.from(b);
  for (let i = 0; i < 16; i++) s[i] ^= w[rounds * 16 + i];
  for (let r = rounds - 1; r >= 0; r--) {
    const t = s.slice();
    for (let c = 0; c < 4; c++) for (let rr = 0; rr < 4; rr++) s[((c + rr) % 4) * 4 + rr] = t[c * 4 + rr];
    for (let i = 0; i < 16; i++) s[i] = INV[s[i]];
    for (let i = 0; i < 16; i++) s[i] ^= w[r * 16 + i];
    if (r > 0) {
      for (let c = 0; c < 4; c++) {
        const [a0, a1, a2, a3] = [s[c * 4], s[c * 4 + 1], s[c * 4 + 2], s[c * 4 + 3]];
        s[c * 4] = mul(a0, 14) ^ mul(a1, 11) ^ mul(a2, 13) ^ mul(a3, 9);
        s[c * 4 + 1] = mul(a0, 9) ^ mul(a1, 14) ^ mul(a2, 11) ^ mul(a3, 13);
        s[c * 4 + 2] = mul(a0, 13) ^ mul(a1, 9) ^ mul(a2, 14) ^ mul(a3, 11);
        s[c * 4 + 3] = mul(a0, 11) ^ mul(a1, 13) ^ mul(a2, 9) ^ mul(a3, 14);
      }
    }
  }
  return s;
}

/** AES-CBC encryption without padding (input length must be a multiple of 16). */
export function aesCbcEncrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Uint8Array {
  const w = expandKey(key);
  const rounds = key.length / 4 + 6;
  const out = new Uint8Array(data.length);
  let prev = iv;
  for (let off = 0; off + 16 <= data.length; off += 16) {
    const block = Uint8Array.from(data.subarray(off, off + 16));
    for (let i = 0; i < 16; i++) block[i] ^= prev[i];
    prev = encryptBlock(w, rounds, block);
    out.set(prev, off);
  }
  return out;
}

/** AES-CBC decryption; `unpad` strips valid PKCS#7 padding. */
export function aesCbcDecrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array, unpad = true): Uint8Array {
  const w = expandKey(key);
  const rounds = key.length / 4 + 6;
  const n = data.length - (data.length % 16);
  const out = new Uint8Array(n);
  let prev = iv;
  for (let off = 0; off < n; off += 16) {
    const block = data.subarray(off, off + 16);
    const d = decryptBlock(w, rounds, block);
    for (let i = 0; i < 16; i++) out[off + i] = d[i] ^ prev[i];
    prev = block;
  }
  if (unpad && n) {
    const p = out[n - 1];
    if (p >= 1 && p <= 16 && out.subarray(n - p).every((x) => x === p)) return out.subarray(0, n - p);
  }
  return out;
}

export async function sha(bits: 256 | 384 | 512, data: Uint8Array): Promise<Uint8Array> {
  return new Uint8Array(await crypto.subtle.digest(`SHA-${bits}`, data as BufferSource));
}
