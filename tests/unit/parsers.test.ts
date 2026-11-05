import { describe, expect, it } from 'vitest';
import { parseXml, find, textContent, childrenNamed, rootElement } from '../../src/core/xml';
import { htmlToDoc } from '../../src/core/readers/html';
import { markdownToDoc, parseInline } from '../../src/core/readers/markdown';
import { textToDoc } from '../../src/core/readers/text';
import { csvToBook, detectDelimiter, jsonToBook, parseCsv, xlsxToBook } from '../../src/core/readers/data';
import { bookToJson, bookToXlsx, sheetToCsv } from '../../src/core/writers/data';
import { detectFormat, targetsFor, pipelineFor } from '../../src/core/formats';
import { decodeText, imageSize, parseDataUri } from '../../src/core/util';
import { encodePng } from '../../src/core/png';
import { docToBook, runsText, type Block } from '../../src/core/model';
import { bytes } from './helpers';

describe('xml parser', () => {
  it('parses elements, attributes, entities and CDATA', () => {
    const doc = parseXml('<?xml version="1.0"?><!-- c --><a x="1 &amp; 2"><b>Привет &lt;мир&gt; &#x41;&#66;</b><c/><![CDATA[<raw>]]></a>');
    const a = rootElement(doc)!;
    expect(a.name).toBe('a');
    expect(a.attrs.x).toBe('1 & 2');
    expect(textContent(find(a, 'b'))).toBe('Привет <мир> AB');
    expect(childrenNamed(a, 'c')).toHaveLength(1);
    expect(a.children[a.children.length - 1]).toBe('<raw>');
  });

  it('handles namespaced names and single-quoted attributes', () => {
    const doc = parseXml("<w:p w:val='x'><w:r><w:t xml:space=\"preserve\"> a </w:t></w:r></w:p>");
    expect(textContent(find(doc, 'w:t'))).toBe(' a ');
    expect(rootElement(doc)!.attrs['w:val']).toBe('x');
  });
});

describe('html reader', () => {
  it('maps structure and inline formatting', () => {
    const doc = htmlToDoc(`<!doctype html><html><head><title>Тест</title><style>p{}</style></head><body>
      <h1 style="text-align:center">Заголовок</h1>
      <p>Обычный <b>жирный</b> <i>курсив</i> <a href="https://example.com">ссылка</a> &laquo;кавычки&raquo;</p>
      <ul><li>один<li>два<ul><li>вложенный</li></ul></ul>
      <ol start="3"><li>три</li></ol>
      <table><tr><th>Имя</th><th>Возраст</th></tr><tr><td>Аня</td><td>30</td></tr></table>
      <pre>code  line\n  2</pre>
      <blockquote><p>цитата</p></blockquote>
      <hr class="page-break"><p>после разрыва</p>
      <script>alert(1)</script>
    </body></html>`);
    expect(doc.title).toBe('Тест');
    const types = doc.blocks.map((b) => b.type);
    expect(types).toEqual(['heading', 'paragraph', 'list', 'list', 'table', 'code', 'quote', 'pagebreak', 'paragraph']);
    const h = doc.blocks[0] as Extract<Block, { type: 'heading' }>;
    expect(h.align).toBe('center');
    const p = doc.blocks[1] as Extract<Block, { type: 'paragraph' }>;
    expect(runsText(p.runs)).toBe('Обычный жирный курсив ссылка «кавычки»');
    expect(p.runs.find((r) => r.text === 'жирный')?.bold).toBe(true);
    expect(p.runs.find((r) => r.text === 'курсив')?.italic).toBe(true);
    expect(p.runs.find((r) => r.text === 'ссылка')?.link).toBe('https://example.com');
    const ul = doc.blocks[2] as Extract<Block, { type: 'list' }>;
    expect(ul.items).toHaveLength(2);
    expect(ul.items[1].blocks[1].type).toBe('list');
    const table = doc.blocks[4] as Extract<Block, { type: 'table' }>;
    expect(table.rows[0].cells[0].header).toBe(true);
    expect((doc.blocks[5] as { text: string }).text).toBe('code  line\n  2');
    expect(JSON.stringify(doc)).not.toContain('alert');
  });

  it('handles Google-Docs style spans and br', () => {
    const doc = htmlToDoc('<b style="font-weight:normal"><p><span style="font-weight:700">A</span><span style="font-style:italic;color:#ff0000">B</span><br>C</p></b>');
    const p = doc.blocks[0] as Extract<Block, { type: 'paragraph' }>;
    expect(p.runs[0]).toMatchObject({ text: 'A', bold: true });
    expect(p.runs[1]).toMatchObject({ text: 'B', italic: true, color: 'FF0000' });
    expect(runsText(p.runs)).toBe('AB\nC');
  });

  it('extracts data-URI images', () => {
    const png = encodePng(new Uint8Array(4 * 3 * 3).fill(200), 4, 3, 3);
    const b64 = Buffer.from(png).toString('base64');
    const doc = htmlToDoc(`<p>до <img src="data:image/png;base64,${b64}" alt="рис"> после</p>`);
    expect(doc.blocks.map((b) => b.type)).toEqual(['paragraph', 'image', 'paragraph']);
    const img = doc.blocks[1] as Extract<Block, { type: 'image' }>;
    expect([img.width, img.height, img.alt]).toEqual([4, 3, 'рис']);
  });
});

describe('markdown reader', () => {
  it('parses blocks', () => {
    const md = [
      '# Заголовок',
      '',
      'Текст с **жирным**, *курсивом*, `кодом` и [ссылкой](https://ya.ru).',
      'Вторая строка абзаца.',
      '',
      '- пункт 1',
      '- пункт 2',
      '  - вложенный',
      '',
      '3. третий',
      '4. четвёртый',
      '',
      '| A | B |',
      '|---|--:|',
      '| 1 | 2 |',
      '',
      '```js',
      'const x = 1;',
      '```',
      '',
      '> цитата',
      '',
      '---',
      '',
      'Подзаголовок',
      '------------',
    ].join('\n');
    const doc = markdownToDoc(md);
    expect(doc.title).toBe('Заголовок');
    expect(doc.blocks.map((b) => b.type)).toEqual(['heading', 'paragraph', 'list', 'list', 'table', 'code', 'quote', 'hr', 'heading']);
    const p = doc.blocks[1] as Extract<Block, { type: 'paragraph' }>;
    expect(runsText(p.runs)).toBe('Текст с жирным, курсивом, кодом и ссылкой. Вторая строка абзаца.');
    expect(p.runs.find((r) => r.text === 'жирным')?.bold).toBe(true);
    expect(p.runs.find((r) => r.text === 'кодом')?.code).toBe(true);
    const list = doc.blocks[2] as Extract<Block, { type: 'list' }>;
    expect(list.items[1].blocks[1]).toMatchObject({ type: 'list', ordered: false });
    const ol = doc.blocks[3] as Extract<Block, { type: 'list' }>;
    expect(ol.ordered).toBe(true);
    expect((doc.blocks[8] as Extract<Block, { type: 'heading' }>).level).toBe(2);
  });

  it('respects escapes, intraword underscores and bare URLs', () => {
    expect(runsText(parseInline('snake_case_name \\*не курсив\\*'))).toBe('snake_case_name *не курсив*');
    const runs = parseInline('см. https://example.com/a_b.');
    expect(runs.find((r) => r.link)?.text).toBe('https://example.com/a_b');
    const nested = parseInline('***оба*** ~~зачёркнуто~~');
    expect(nested[0]).toMatchObject({ text: 'оба', bold: true, italic: true });
    expect(nested.find((r) => r.strike)?.text).toBe('зачёркнуто');
  });
});

describe('text reader', () => {
  it('splits paragraphs on blank lines and keeps line breaks', () => {
    const doc = textToDoc('Строка 1\nСтрока 2\r\n\r\n\nАбзац 2\n');
    expect(doc.blocks).toHaveLength(2);
    expect(runsText((doc.blocks[0] as Extract<Block, { type: 'paragraph' }>).runs)).toBe('Строка 1\nСтрока 2');
  });

  it('decodes windows-1251 and UTF-16 text', () => {
    const cp1251 = new Uint8Array([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2]); // "Привет"
    expect(decodeText(cp1251)).toBe('Привет');
    const utf16 = new Uint8Array([0xff, 0xfe, 0x1f, 0x04, 0x40, 0x04]);
    expect(decodeText(utf16)).toBe('Пр');
  });
});

describe('csv / json / xlsx', () => {
  it('parses RFC4180 CSV with quotes and multiline fields', () => {
    const rows = parseCsv('a,b,c\n"x, y","he said ""hi""","line1\nline2"\n1,2,3\n');
    expect(rows).toEqual([
      ['a', 'b', 'c'],
      ['x, y', 'he said "hi"', 'line1\nline2'],
      ['1', '2', '3'],
    ]);
  });

  it('detects semicolon delimiter (Russian Excel exports)', () => {
    const text = '﻿Имя;Сумма\nАня;1,5\nБорис;2,75\n';
    expect(detectDelimiter(text)).toBe(';');
    const book = csvToBook(text);
    expect(book.sheets[0].rows[2]).toEqual(['Борис', '2,75']);
    expect(book.sheets[0].header).toBe(true);
  });

  it('writes CSV with proper quoting', () => {
    const csv = sheetToCsv({ name: 's', rows: [['a', 'b,c', 'q"x'], [1, null, true]] });
    expect(csv).toBe('a,"b,c","q""x"\r\n1,,true\r\n');
  });

  it('converts JSON arrays of objects and back', () => {
    const book = jsonToBook(JSON.stringify([{ name: 'Аня', age: 30 }, { name: 'Борис', city: 'Москва' }]));
    expect(book.sheets[0].rows).toEqual([
      ['name', 'age', 'city'],
      ['Аня', 30, null],
      ['Борис', null, 'Москва'],
    ]);
    const back = JSON.parse(bookToJson(book));
    expect(back[1]).toEqual({ name: 'Борис', age: null, city: 'Москва' });
  });

  it('rejects invalid JSON with a readable error', () => {
    expect(() => jsonToBook('{oops')).toThrow(/Некорректный JSON/);
  });

  it('round-trips XLSX with multiple sheets, numbers, booleans and unicode', () => {
    const book = {
      sheets: [
        { name: 'Данные', rows: [['Имя', 'Сумма', 'Флаг'], ['Аня', '12.5', true], ['Код', '007', false]], header: true },
        { name: 'Второй/лист?', rows: [['x'], [1e21]] },
      ],
    };
    const xlsx = bookToXlsx(book);
    const back = xlsxToBook(xlsx);
    expect(back.sheets.map((s) => s.name)).toEqual(['Данные', 'Второй лист']);
    expect(back.sheets[0].rows).toEqual([
      ['Имя', 'Сумма', 'Флаг'],
      ['Аня', 12.5, true],
      ['Код', '007', false],
    ]);
    expect(back.sheets[1].rows[1][0]).toBe(1e21);
  });
});

describe('formats', () => {
  it('detects by magic bytes and extension', () => {
    expect(detectFormat('x.bin', bytes('%PDF-1.7')).format).toBe('pdf');
    expect(detectFormat('a.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0])).format).toBe('png');
    expect(detectFormat('a.md', bytes('# hi')).format).toBe('md');
    expect(detectFormat('noext', bytes('<!DOCTYPE html><p>x')).format).toBe('html');
    expect(detectFormat('noext', bytes('[1,2]')).format).toBe('json');
    expect(detectFormat('old.doc', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0])).error).toMatch(/\.docx/);
    expect(detectFormat('x.xlsx', bookToXlsx({ sheets: [{ name: 'a', rows: [] }] })).format).toBe('xlsx');
    expect(detectFormat('x.svg', bytes('<?xml version="1.0"?><svg></svg>')).format).toBe('svg');
  });

  it('builds a consistent conversion graph', () => {
    expect(targetsFor('pdf')).toContain('docx');
    expect(targetsFor('docx')).toContain('pdf');
    expect(targetsFor('png')).not.toContain('png');
    expect(pipelineFor('pdf', 'jpg')).toBe('raster');
    expect(pipelineFor('jpg', 'pdf')).toBe('images-to-pdf');
    expect(pipelineFor('docx', 'md')).toBe('doc');
    expect(pipelineFor('txt', 'xlsx')).toBeNull();
  });
});

describe('utils', () => {
  it('encodes a valid PNG and reads its size', () => {
    const px = new Uint8Array(10 * 7 * 4).map((_, i) => (i * 37) & 255);
    const png = encodePng(px, 10, 7, 4);
    expect(imageSize(png)).toEqual({ width: 10, height: 7 });
  });

  it('parses data URIs', () => {
    expect(parseDataUri('data:text/plain;base64,0J/RgNC40LLQtdGC')).toMatchObject({ mime: 'text/plain' });
    expect(new TextDecoder().decode(parseDataUri('data:,a%20b')!.data)).toBe('a b');
    expect(parseDataUri('https://x')).toBeNull();
  });

  it('turns document tables into sheets', () => {
    const book = docToBook(htmlToDoc('<table><tr><th>a</th><th>b</th></tr><tr><td colspan="2">x</td></tr></table>'));
    expect(book.sheets[0].rows).toEqual([
      ['a', 'b'],
      ['x', ''],
    ]);
    expect(book.sheets[0].header).toBe(true);
  });
});
