/** Rich Text Format: our writer, our reader, LibreOffice interoperability and hand-written edge cases. */
import { describe, expect, it } from 'vitest';
import { docToRtf, rtfEscape } from '../../src/core/writers/rtf';
import { rtfToDoc } from '../../src/core/readers/rtf';
import { docToDocx } from '../../src/core/writers/docx';
import { docxToDoc } from '../../src/core/readers/docx';
import { docToMarkdown } from '../../src/core/writers/text';
import { detectFormat } from '../../src/core/formats';
import { convert } from '../../src/core/convert';
import { ConvertError } from '../../src/core/util';
import { blockText, type Block, type Doc, type TextRun } from '../../src/core/model';
import { env, hasCmd, pdfToText, soffice, text } from './helpers';
import { sample } from './sample';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
const read = (s: string): Doc => rtfToDoc(enc(s)).doc;
const runsOf = (b: Block | undefined): TextRun[] => (b && (b.type === 'paragraph' || b.type === 'heading') ? b.runs : []);

const withExtras: Doc = {
  ...sample,
  header: [{ type: 'paragraph', runs: [{ text: 'Шапка документа' }] }],
  footer: [{ type: 'paragraph', align: 'center', runs: [{ text: 'Стр. ' }, { text: '1', field: 'page' }, { text: ' из ' }, { text: '1', field: 'pages' }] }],
  blocks: [
    ...sample.blocks,
    {
      type: 'paragraph',
      runs: [{ text: 'Сноска тут' }, { text: '', note: [{ type: 'paragraph', runs: [{ text: 'Текст сноски' }] }] }, { text: ' эмодзи 😀 и 中文', color: 'FF0000', highlight: 'FFFF00', size: 16, font: 'serif' }],
    },
  ],
};

describe('RTF writer', () => {
  const { data, warnings } = docToRtf(withExtras);

  it('produces 7-bit RTF that is detected as RTF', () => {
    expect(text(data.subarray(0, 6))).toBe('{\\rtf1');
    expect(data.every((b) => b < 0x80)).toBe(true);
    expect(detectFormat('file.bin', data).format).toBe('rtf');
    expect(warnings).toEqual([]);
  });

  it('escapes specials, Unicode (incl. astral) and control characters', () => {
    expect(rtfEscape('a{b}\\c')).toBe('a\\{b\\}\\\\c');
    expect(rtfEscape('Ж')).toBe('\\u1046?');
    expect(rtfEscape('я')).toBe('\\u1103?');
    expect(rtfEscape('😀')).toBe('\\u-10179?\\u-8704?');
    expect(rtfEscape('x\ny\tz\u0001')).toBe('x\\line y\\tab z');
  });

  it('round-trips through our reader without losing structure', () => {
    expect(docToMarkdown(rtfToDoc(docToRtf(sample).data).doc)).toBe(docToMarkdown(sample));
    const back = rtfToDoc(data).doc;
    expect(back.header?.map(blockText)).toEqual(['Шапка документа']);
    expect(runsOf(back.footer?.[0]).filter((r) => r.field).map((r) => r.field)).toEqual(['page', 'pages']);
    const last = runsOf(back.blocks[back.blocks.length - 1]);
    expect(last.find((r) => r.note)?.note?.map(blockText)).toEqual(['Текст сноски']);
    expect(last.find((r) => r.text.includes('эмодзи'))).toMatchObject({ text: ' эмодзи 😀 и 中文', color: 'FF0000', highlight: 'FFFF00', size: 16, font: 'serif' });
    expect(back.font).toBe('sans');
    expect(back.size).toBe(11);
  });

  it.skipIf(!hasCmd('soffice'))('opens in LibreOffice with lists, tables, notes and page numbers', () => {
    const t = pdfToText(soffice(data, 'rtf', 'pdf'));
    expect(t).toContain('Шапка документа');
    expect(t).toMatch(/Стр\. 1 из 2/);
    expect(t).toContain('Текст сноски');
    expect(t).toContain('Москва');
    const viaDocx = docxToDoc(soffice(data, 'rtf', 'docx:MS Word 2007 XML')).doc;
    const md = docToMarkdown(viaDocx);
    expect(md).toContain('# Отчёт о работе');
    expect(md).toMatch(/- Второй пункт\n {2}1\. Вложенный один\n {2}2\. Вложенный два/);
    expect(md).toContain('| Москва | 13 100 000 |');
    expect(md).toContain('[ссылка](https://example.com/)');
    expect(viaDocx.blocks.some((b) => b.type === 'image')).toBe(true);
  });
});

describe.skipIf(!hasCmd('soffice'))('RTF reader on LibreOffice files', () => {
  const rtf = soffice(docToDocx(withExtras).data, 'docx', 'rtf');
  const { doc } = rtfToDoc(rtf);

  it('reads structure, formatting, notes and the master page', () => {
    const md = docToMarkdown(doc);
    expect(md).toContain('# Отчёт о работе');
    expect(md).toContain('## Список задач');
    expect(md).toContain('**жирный**');
    expect(md).toContain('*курсив*');
    expect(md).toContain('[ссылка](https://example.com/)');
    expect(md).toMatch(/- Второй пункт\n {2}1\. Вложенный один\n {2}2\. Вложенный два/);
    expect(md).toContain('| Москва | 13 100 000 |');
    expect(md).toContain('[^1]: Текст сноски');
    // runs may be split by script (LibreOffice assigns East Asian fonts separately)
    expect(doc.blocks.map(blockText).join('\n')).toContain('эмодзи 😀 и 中文');
    expect(doc.blocks.some((b) => b.type === 'image')).toBe(true);
    expect(doc.blocks.some((b) => b.type === 'pagebreak')).toBe(true);
    expect(doc.header?.map(blockText)).toEqual(['Шапка документа']);
    expect(runsOf(doc.footer?.[0]).filter((r) => r.field).length).toBe(2);
    expect(doc.font).toBe('sans');
    expect(doc.size).toBe(11);
  });

  it('converts .rtf to every document target', async () => {
    for (const to of ['docx', 'pdf', 'odt', 'txt', 'md', 'html'] as const) {
      const r = await convert(rtf, 'rtf', to, env);
      expect(r.data.length).toBeGreaterThan(100);
    }
  });
});

describe('RTF reader edge cases', () => {
  it('decodes code pages: \\ansicpg1251 and per-font \\fcharset', () => {
    const d = read("{\\rtf1\\ansi\\ansicpg1251\\deff0{\\fonttbl{\\f0\\fswiss\\fcharset204 Arial;}{\\f1\\froman\\fcharset0 Times;}}\\pard \\'cf\\'f0\\'e8\\'e2\\'e5\\'f2 {\\f1 caf\\'e9}\\par}");
    expect(blockText(d.blocks[0])).toBe('Привет café');
  });

  it('decodes double-byte code pages (Shift-JIS)', () => {
    const d = read("{\\rtf1\\ansi\\ansicpg932{\\fonttbl{\\f0\\fnil\\fcharset128 MS Mincho;}}\\f0 \\'93\\'fa\\'96\\'7b\\par}");
    expect(blockText(d.blocks[0])).toBe('日本');
  });

  it('handles \\u with \\uc skipping, surrogate pairs and \\uc0', () => {
    expect(blockText(read('{\\rtf1 \\uc1\\u1046?\\u-10179?\\u-8704?\\par}').blocks[0])).toBe('Ж😀');
    expect(blockText(read("{\\rtf1 \\uc2\\u1046\\'3f\\'3fA\\par}").blocks[0])).toBe('ЖA');
    expect(blockText(read('{\\rtf1 {\\uc0\\u1046}x\\par}').blocks[0])).toBe('Жx');
  });

  it('reads legacy WordPad lists (\\pn) as real lists', () => {
    const d = read(
      '{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}' +
        "{\\pntext\\f0\\'B7\\tab}{\\*\\pn\\pnlvlblt\\pnf0\\pnindent0{\\pntxtb\\'B7}}\\fi-360\\li720 One\\par" +
        "{\\pntext\\f0\\'B7\\tab}Two\\par\\pard" +
        '{\\pntext\\f0 1.\\tab}{\\*\\pn\\pnlvlbody\\pnf0\\pnindent0\\pnstart1\\pndec{\\pntxta.}}\\fi-360\\li720 First\\par\\pard After\\par}',
    );
    expect(d.blocks.map((b) => b.type)).toEqual(['list', 'list', 'paragraph']);
    const [a, b] = d.blocks;
    expect(a.type === 'list' && !a.ordered && a.items.map((i) => blockText(i.blocks[0]))).toEqual(['One', 'Two']);
    expect(b.type === 'list' && b.ordered && b.items.map((i) => blockText(i.blocks[0]))).toEqual(['First']);
  });

  it('reads tables with header rows and merged cells', () => {
    const row = (def: string, cells: string[]): string => `\\trowd${def}${cells.map(() => '').join('')}` + cells.map((c, i) => `\\cellx${(i + 1) * 2000}`).join('') + `\\pard\\intbl ${cells.join('\\cell ')}\\cell\\row`;
    const d = read(
      '{\\rtf1 ' +
        row('\\trhdr', ['A', 'B', 'C']) +
        '\\trowd\\clmgf\\cellx2000\\clmrg\\cellx4000\\cellx6000\\pard\\intbl wide\\cell \\cell x\\cell\\row' +
        '\\pard After\\par}',
    );
    expect(d.blocks.map((b) => b.type)).toEqual(['table', 'paragraph']);
    const t = d.blocks[0];
    if (t.type !== 'table') return;
    expect(t.rows[0].cells.every((c) => c.header)).toBe(true);
    expect(t.rows[0].cells.map((c) => c.blocks.map(blockText).join())).toEqual(['A', 'B', 'C']);
    expect(t.rows[1].cells.map((c) => [c.blocks.map(blockText).join(), c.colspan ?? 1])).toEqual([
      ['wide', 2],
      ['x', 1],
    ]);
  });

  it('skips hidden text, unknown destinations and fallback pictures', () => {
    const d = read('{\\rtf1 A{\\v hidden}B{\\*\\unknowndest junk}C{\\nonshppict{\\pict\\wmetafile8 0102}}{\\*\\generator Foo;}\\par}');
    expect(blockText(d.blocks[0])).toBe('ABC');
    expect(d.blocks.length).toBe(1);
  });

  it('reads hyperlinks and PAGE fields, and binary \\bin pictures', () => {
    const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 3, 8, 2, 0, 0, 0]);
    const head = enc('{\\rtf1 {\\field{\\*\\fldinst{ HYPERLINK "https://ex.com/a?b=1" }}{\\fldrslt{\\ul link}}} p.{\\field{\\*\\fldinst PAGE}{\\fldrslt 7}}\\par{\\pict\\pngblip\\picwgoal300\\pichgoal450\\bin' + png.length + ' ');
    const tail = enc('}\\par}');
    const bytes = new Uint8Array(head.length + png.length + tail.length);
    bytes.set(head);
    bytes.set(png, head.length);
    bytes.set(tail, head.length + png.length);
    const d = rtfToDoc(bytes).doc;
    const runs = runsOf(d.blocks[0]);
    expect(runs.find((r) => r.link)).toMatchObject({ text: 'link', link: 'https://ex.com/a?b=1' });
    expect(runs.find((r) => r.field)).toMatchObject({ text: '7', field: 'page' });
    const img = d.blocks.find((b) => b.type === 'image');
    expect(img).toMatchObject({ mime: 'image/png', width: 20, height: 30 });
  });

  it('treats \\page, \\pagebb and section breaks as page breaks', () => {
    const d = read('{\\rtf1 A\\par\\page B\\par\\pard\\pagebb C\\par\\sect\\sectd D\\par\\sect\\sectd\\sbknone E\\par}');
    expect(d.blocks.map((b) => (b.type === 'pagebreak' ? '|' : blockText(b))).join('')).toBe('A|B|C|DE');
  });

  it('survives malformed input and rejects non-RTF', () => {
    expect(blockText(read('{\\rtf1 unbalanced {\\b bold').blocks[0])).toBe('unbalanced bold');
    expect(read('{\\rtf1 }}}} extra').blocks.length).toBe(0);
    expect(() => rtfToDoc(enc('hello'))).toThrow(ConvertError);
    const deep = '{\\rtf1 ' + '{'.repeat(5000) + 'x' + '}'.repeat(5000) + '}';
    expect(() => rtfToDoc(enc(deep))).toThrow(ConvertError);
  });

  it('parses a 5 MB document quickly', () => {
    const para = '\\pard\\plain\\f0\\fs22 Lorem ipsum {\\b dolor} sit amet, \\u1078?\\u1080?\\u1074?.\\par\n';
    const big = '{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}' + para.repeat(Math.ceil(5_000_000 / para.length)) + '}';
    const t0 = performance.now();
    const d = read(big);
    expect(performance.now() - t0).toBeLessThan(4000);
    expect(d.blocks.length).toBeGreaterThan(50_000);
  });
});
