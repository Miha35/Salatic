import { describe, expect, it } from 'vitest';
import { zipSync } from 'fflate';
import { docxToDoc } from '../../src/core/readers/docx';
import { xlsxToBook } from '../../src/core/readers/data';
import { markdownToDoc, parseInline } from '../../src/core/readers/markdown';
import { htmlToDoc } from '../../src/core/readers/html';
import { docToDocx } from '../../src/core/writers/docx';
import { docToMarkdown } from '../../src/core/writers/text';
import { bookToXlsx } from '../../src/core/writers/data';
import { bookToDoc } from '../../src/core/model';
import { detectFormat } from '../../src/core/formats';
import { bytes } from './helpers';

describe('hostile and huge inputs', () => {
  it('rejects zip bombs before inflating them', () => {
    const bomb = zipSync({ 'word/document.xml': [new Uint8Array(60 * 1024 * 1024), { level: 9 }], '[Content_Types].xml': bytes('<Types/>') });
    expect(bomb.length).toBeLessThan(1024 * 1024);
    expect(detectFormat('x.docx', bomb).format).toBe('docx');
    expect(() => docxToDoc(bomb)).toThrow(/подозрительно/);
    const xbomb = zipSync({ 'xl/workbook.xml': [new Uint8Array(60 * 1024 * 1024), { level: 9 }] });
    expect(() => xlsxToBook(xbomb)).toThrow(/подозрительно/);
  });

  it('reports corrupted archives with a readable message', () => {
    const broken = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5, 6, 7, 8]);
    expect(() => docxToDoc(broken)).toThrow(/повреждён/);
    expect(() => xlsxToBook(broken)).toThrow(/повреждён/);
  });

  it('parses pathological Markdown in linear time', () => {
    const t0 = performance.now();
    parseInline('*'.repeat(200_000));
    parseInline('_a '.repeat(100_000));
    markdownToDoc('['.repeat(50_000) + 'x' + '('.repeat(50_000));
    expect(performance.now() - t0).toBeLessThan(3000);
  });

  it('parses deeply broken HTML without throwing', () => {
    const doc = htmlToDoc('<div>'.repeat(5000) + 'текст' + '<table><tr><td>'.repeat(200) + '</p></li></b>');
    expect(JSON.stringify(doc)).toContain('текст');
  });

  it('handles tables with 150 000 rows (no call-stack overflow)', () => {
    const rows = Array.from({ length: 150_000 }, (_, i) => [`r${i}`, i]);
    const book = { sheets: [{ name: 'Big', rows, header: false }] };
    const xlsx = bookToXlsx(book);
    expect(xlsxToBook(xlsx).sheets[0].rows.length).toBe(150_000);
    const doc = bookToDoc(book);
    expect(docToMarkdown(doc).length).toBeGreaterThan(1_000_000);
    expect(docToDocx(doc).data.length).toBeGreaterThan(100_000);
  }, 60_000);
});
