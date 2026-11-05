/** EPUB: our writer (structure, well-formed XHTML, navigation) and reader (ours and hand-made books). */
import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { docToEpub } from '../../src/core/writers/epub';
import { epubToDoc } from '../../src/core/readers/epub';
import { docToMarkdown } from '../../src/core/writers/text';
import { detectFormat } from '../../src/core/formats';
import { convert } from '../../src/core/convert';
import { blockText, type Doc } from '../../src/core/model';
import { env, hasCmd, pdfToText, text } from './helpers';
import { sample } from './sample';

/** Strict XML well-formedness check with Python's expat-based parser. */
function wellFormed(xml: Uint8Array): void {
  execFileSync('python3', ['-I', '-c', 'import sys,xml.dom.minidom; xml.dom.minidom.parseString(sys.stdin.buffer.read())'], { input: xml });
}

const book: Doc = {
  ...sample,
  title: 'Тестовая книга',
  blocks: [
    ...sample.blocks,
    { type: 'paragraph', runs: [{ text: 'Сноска' }, { text: '', note: [{ type: 'paragraph', runs: [{ text: 'Текст сноски' }] }] }, { text: ' & <теги> "кавычки"\tтаб' }] },
  ],
};

describe('EPUB writer', () => {
  const { data, warnings } = docToEpub(book);
  const files = unzipSync(data);

  it('has a valid OCF container', () => {
    expect(text(data.subarray(30, 38))).toBe('mimetype');
    expect(data[8] | (data[9] << 8)).toBe(0); // stored, not deflated
    expect(text(files['mimetype'])).toBe('application/epub+zip');
    expect(text(files['META-INF/container.xml'])).toContain('full-path="OEBPS/content.opf"');
    expect(detectFormat('book.zip', data).format).toBe('epub');
    expect(warnings).toEqual([]);
  });

  it.skipIf(!hasCmd('python3'))('contains only well-formed XML documents', () => {
    const xmlFiles = Object.keys(files).filter((p) => /\.(xhtml|opf|ncx|xml)$/.test(p));
    expect(xmlFiles.length).toBeGreaterThanOrEqual(6);
    for (const p of xmlFiles) wellFormed(files[p]);
  });

  it('declares every resource in the manifest and splits chapters at top headings', () => {
    const opf = text(files['OEBPS/content.opf']);
    const hrefs = [...opf.matchAll(/href="([^"]+)"/g)].map((m) => `OEBPS/${m[1]}`);
    for (const h of hrefs) expect(files[h], h).toBeDefined();
    const content = Object.keys(files).filter((p) => p.startsWith('OEBPS/') && !p.endsWith('content.opf'));
    for (const p of content) expect(hrefs).toContain(p);
    expect(opf).toContain('properties="nav"');
    expect(opf).toMatch(/<dc:identifier id="bookid">urn:uuid:[0-9a-f-]{36}<\/dc:identifier>/);
    expect(opf).toMatch(/<meta property="dcterms:modified">\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ<\/meta>/);
    // the sample has two top-level sections (h1 + "Вторая страница" h2 after a page break)
    const nav = text(files['OEBPS/nav.xhtml']);
    expect(nav).toContain('Отчёт о работе');
    expect(Object.keys(files).some((p) => p.startsWith('OEBPS/images/'))).toBe(true);
    const chapters = Object.keys(files).filter((p) => p.startsWith('OEBPS/text/'));
    expect(chapters.length).toBe(3);
    expect(chapters.map((p) => text(files[p])).join('')).toContain('epub:type="noteref"');
  });

  it('round-trips through our reader', () => {
    const back = epubToDoc(data).doc;
    expect(back.title).toBe('Тестовая книга');
    const md = docToMarkdown(back);
    expect(md).toContain('# Отчёт о работе');
    expect(md).toContain('[ссылка](https://example.com/)');
    expect(md).toMatch(/- Второй пункт\n {2}1\. Вложенный один\n {2}2\. Вложенный два/);
    expect(md).toContain('| Москва | 13 100 000 |');
    expect(md).toContain('[^1]: Текст сноски');
    expect(back.blocks.map(blockText).join('\n')).toContain('& <теги> "кавычки"');
    expect(back.blocks.some((b) => b.type === 'image')).toBe(true);
    expect(md).not.toContain('Содержание');
  });

  it('converts EPUB to PDF and DOCX', async () => {
    const pdf = await convert(data, 'epub', 'pdf', env);
    if (hasCmd('pdftotext')) expect(pdfToText(pdf.data)).toContain('Съешь же ещё');
    const docx = await convert(data, 'epub', 'docx', env);
    expect(docx.data.length).toBeGreaterThan(1000);
  });
});

describe('EPUB reader', () => {
  const xhtml = (body: string): string =>
    `<?xml version="1.0" encoding="utf-8"?><html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"><head><title>x</title></head><body>${body}</body></html>`;
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 4, 0, 0, 0, 2, 8, 2, 0, 0, 0]);
  const make = (extra: Record<string, Uint8Array> = {}): Uint8Array =>
    zipSync({
      mimetype: strToU8('application/epub+zip'),
      'META-INF/container.xml': strToU8('<container><rootfiles><rootfile full-path="OPS/book.opf"/></rootfiles></container>'),
      'OPS/book.opf': strToU8(
        '<package xmlns="http://www.idpf.org/2007/opf" xmlns:dc="http://purl.org/dc/elements/1.1/"><metadata><dc:title>Книга</dc:title></metadata>' +
          '<manifest><item id="c" href="Text/cover.xhtml" media-type="application/xhtml+xml"/><item id="a" href="Text/ch%201.xhtml" media-type="application/xhtml+xml"/>' +
          '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/><item id="i" href="Images/p.png" media-type="image/png"/></manifest>' +
          '<spine><itemref idref="c"/><itemref idref="nav"/><itemref idref="a"/></spine></package>',
      ),
      'OPS/nav.xhtml': strToU8(xhtml('<nav epub:type="toc"><ol><li>TOC</li></ol></nav>')),
      'OPS/Text/cover.xhtml': strToU8(xhtml('<svg xmlns:xlink="http://www.w3.org/1999/xlink"><image width="4" height="2" xlink:href="../Images/p.png"/></svg>')),
      'OPS/Text/ch 1.xhtml': strToU8(xhtml('<h1>Глава</h1><p>Текст<a epub:type="noteref" href="#n1">1</a> и <img src="../Images/p.png" alt="рис"/></p><aside epub:type="footnote" id="n1"><p>Примечание</p></aside>')),
      'OPS/Images/p.png': png,
      ...extra,
    });

  it('follows the spine, resolves relative and escaped paths, SVG covers and EPUB 3 notes', () => {
    const { doc } = epubToDoc(make());
    expect(doc.title).toBe('Книга');
    const types = doc.blocks.map((b) => b.type);
    expect(types[0]).toBe('image'); // cover
    expect(types).toContain('pagebreak');
    expect(doc.blocks.map(blockText).join('|')).not.toContain('TOC');
    const md = docToMarkdown(doc);
    expect(md).toContain('# Глава');
    expect(md).toContain('[^1]: Примечание');
    expect(md).not.toMatch(/^Примечание$/m);
    expect(doc.blocks.filter((b) => b.type === 'image').length).toBe(2);
  });

  it('rejects DRM-protected books but accepts font obfuscation', () => {
    const enc = (alg: string): Uint8Array =>
      strToU8(`<encryption xmlns="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:enc="http://www.w3.org/2001/04/xmlenc#"><enc:EncryptedData><enc:EncryptionMethod Algorithm="${alg}"/></enc:EncryptedData></encryption>`);
    expect(() => epubToDoc(make({ 'META-INF/encryption.xml': enc('http://www.w3.org/2001/04/xmlenc#aes128-cbc') }))).toThrow(/DRM/);
    expect(epubToDoc(make({ 'META-INF/encryption.xml': enc('http://www.idpf.org/2008/embedding') })).doc.blocks.length).toBeGreaterThan(0);
  });

  it('reports broken packages clearly', () => {
    expect(() => epubToDoc(zipSync({ mimetype: strToU8('application/epub+zip') }))).toThrow(/OPF/);
    expect(() => epubToDoc(new Uint8Array([1, 2, 3]))).toThrow();
  });
});
