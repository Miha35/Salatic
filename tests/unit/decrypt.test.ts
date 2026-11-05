/** Encrypted PDFs: our Standard Security Handler against files encrypted by pypdf (and LibreOffice). */
import { describe, expect, it, beforeAll } from 'vitest';
import { createCipheriv, createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { PDFDocument } from 'pdf-lib';
import { aesCbcDecrypt, aesCbcEncrypt, md5, rc4 } from '../../src/core/crypto';
import { decryptPdf, isEncryptedPdf } from '../../src/core/pdf-decrypt';
import { docToPdf } from '../../src/core/writers/pdf';
import { convert } from '../../src/core/convert';
import { docToDocx } from '../../src/core/writers/docx';
import { ConvertError } from '../../src/core/util';
import { env, hasCmd, loadFonts, loadPdfjs, pdfToText, soffice, text, withTemp } from './helpers';
import { sample } from './sample';

describe('crypto primitives (vs Node crypto)', () => {
  it('MD5', () => {
    for (const s of ['', 'a', 'abc', 'x'.repeat(55), 'y'.repeat(56), 'z'.repeat(64), 'w'.repeat(1000)]) {
      const d = new TextEncoder().encode(s);
      expect(Buffer.from(md5(d)).toString('hex')).toBe(createHash('md5').update(d).digest('hex'));
    }
  });
  it('AES-128/256 CBC with and without padding', () => {
    for (const kl of [16, 32]) {
      const key = new Uint8Array(kl).map((_, i) => i * 7 + 1);
      const iv = new Uint8Array(16).map((_, i) => i * 3);
      const data = new Uint8Array(64).map((_, i) => i * 5);
      const c = createCipheriv(kl === 16 ? 'aes-128-cbc' : 'aes-256-cbc', key, iv);
      c.setAutoPadding(false);
      const ref = Buffer.concat([c.update(data), c.final()]);
      expect(Buffer.from(aesCbcEncrypt(key, iv, data)).toString('hex')).toBe(ref.toString('hex'));
      expect([...aesCbcDecrypt(key, iv, ref, false)]).toEqual([...data]);
      const c2 = createCipheriv(kl === 16 ? 'aes-128-cbc' : 'aes-256-cbc', key, iv);
      const padded = Buffer.concat([c2.update(data.subarray(0, 37)), c2.final()]);
      expect([...aesCbcDecrypt(key, iv, padded)]).toEqual([...data.subarray(0, 37)]);
    }
  });
  it('RC4 test vector', () => {
    expect(Buffer.from(rc4(new TextEncoder().encode('Key'), new TextEncoder().encode('Plaintext'))).toString('hex')).toBe('bbf316e8d940af0ad3');
  });
});

describe.skipIf(!hasCmd('python3') || !hasCmd('pdftotext'))('PDF decryption', () => {
  let plain: Uint8Array;
  beforeAll(async () => {
    plain = (await docToPdf(sample, { fonts: loadFonts })).data;
  });
  const encrypt = (algo: string, user: string, owner: string): Uint8Array =>
    withTemp((dir) => {
      const script = `
import sys
from pypdf import PdfReader, PdfWriter
r = PdfReader(sys.argv[1]); w = PdfWriter()
for p in r.pages: w.add_page(p)
w.add_metadata({'/Title': 'Секретный отчёт'})
w.encrypt(user_password=sys.argv[4], owner_password=sys.argv[5] or None, algorithm=sys.argv[3])
w.write(sys.argv[2])`;
      const src = join(dir, 'in.pdf');
      const dst = join(dir, 'out.pdf');
      writeFileSync(src, plain);
      execFileSync('python3', ['-I', '-c', script, src, dst, algo, user, owner]);
      return new Uint8Array(readFileSync(dst));
    });

  const check = async (enc: Uint8Array, password: string): Promise<void> => {
    expect(isEncryptedPdf(enc)).toBe(true);
    const dec = await decryptPdf(enc, password);
    expect(isEncryptedPdf(dec)).toBe(false);
    const doc = await PDFDocument.load(dec); // pdf-lib refuses encrypted files
    expect(doc.getPageCount()).toBe(2);
    expect(doc.getTitle()).toBe('Секретный отчёт'); // an encrypted string in the Info dictionary
    expect(pdfToText(await doc.save())).toContain('Съешь же ещё этих мягких');
    const pdfjs = await loadPdfjs();
    const p = await pdfjs.getDocument({ data: dec.slice(), verbosity: 0 }).promise;
    expect(p.numPages).toBe(2);
  };

  const cases: [string, string, string, string][] = [
    ['RC4-40', '', 'owner', ''],
    ['RC4-128', 'user', 'owner', 'user'],
    ['RC4-128', 'user', 'owner', 'owner'],
    ['AES-128', '', 'owner', ''],
    ['AES-128', 'secret', 'owner', 'secret'],
    ['AES-256-R5', 'u5', 'o5', 'u5'],
    ['AES-256', 'пароль', 'владелец', 'пароль'],
    ['AES-256', 'user', 'владелец', 'владелец'],
  ];
  for (const [algo, user, owner, pw] of cases) {
    it(`${algo}, ${user ? 'user password' : 'permissions only'}${pw === owner ? ' (owner password)' : ''}`, async () => {
      await check(encrypt(algo, user, owner), pw);
    });
  }

  it('asks for a password and rejects a wrong one', async () => {
    const enc = encrypt('AES-256', 'right', 'owner');
    await expect(decryptPdf(enc, '')).rejects.toMatchObject({ code: 'password_required' });
    await expect(decryptPdf(enc, 'wrong')).rejects.toMatchObject({ code: 'password_invalid' });
    await expect(decryptPdf(enc, 'wrong')).rejects.toBeInstanceOf(ConvertError);
    const rc4enc = encrypt('RC4-128', 'right', 'owner');
    await expect(decryptPdf(rc4enc, 'nope')).rejects.toMatchObject({ code: 'password_invalid' });
  });

  it('leaves plain PDFs untouched and does not mistake /EncryptMetadata for encryption', async () => {
    expect(isEncryptedPdf(plain)).toBe(false);
    expect(await decryptPdf(plain)).toBe(plain);
    expect(isEncryptedPdf(new TextEncoder().encode('%PDF-1.7\n<< /EncryptMetadata false >>'))).toBe(false);
  });

  it('converts a decrypted PDF to Word', async () => {
    const dec = await decryptPdf(encrypt('AES-128', '', 'owner'));
    const r = await convert(dec, 'pdf', 'md', env);
    expect(text(r.data)).toContain('Отчёт о работе');
  });

  it.skipIf(!hasCmd('soffice'))('decrypts a PDF encrypted by LibreOffice', async () => {
    const opts = JSON.stringify({ EncryptFile: { type: 'boolean', value: 'true' }, DocumentOpenPassword: { type: 'string', value: 'lo-secret' } });
    const enc = soffice(docToDocx(sample).data, 'docx', `pdf:writer_pdf_Export:${opts}`);
    expect(isEncryptedPdf(enc)).toBe(true);
    await expect(decryptPdf(enc)).rejects.toMatchObject({ code: 'password_required' });
    const dec = await decryptPdf(enc, 'lo-secret');
    const doc = await PDFDocument.load(dec);
    expect(pdfToText(await doc.save())).toContain('Съешь же ещё этих мягких');
  });
});
