#!/usr/bin/env node
/**
 * End-to-end tests against the built site served by `wrangler dev` (the real Worker +
 * static assets), driven by Playwright/Chromium.
 *
 *   npm run build && npm run test:e2e
 *   BASE_URL=https://salatic.example.workers.dev npm run test:e2e   # test a deployment
 */
import { spawn } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import { unzipSync } from 'fflate';
import { PDFDocument } from 'pdf-lib';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '../..');
const out = join(here, '.out');
const fx = join(out, 'fixtures');
rmSync(out, { recursive: true, force: true });
mkdirSync(fx, { recursive: true });

const PORT = 8790 + Math.floor(Math.random() * 100);
let BASE = process.env.BASE_URL?.replace(/\/$/, '');
let server = null;

// ---------------------------------------------------------------------------- harness
const tests = [];
const test = (name, fn) => tests.push({ name, fn });
class AssertionError extends Error {}
function expect(cond, msg) {
  if (!cond) throw new AssertionError(msg);
}
const td = new TextDecoder();
const text = (b) => td.decode(b);

async function pdfText(bytes) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), verbosity: 0 }).promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const tc = await (await doc.getPage(i)).getTextContent();
    pages.push(tc.items.map((it) => it.str).join(' '));
  }
  await doc.loadingTask.destroy();
  return pages;
}
const docxText = (bytes) => text(unzipSync(bytes)['word/document.xml']).replace(/<w:tab\/>/g, '\t').replace(/<[^>]+>/g, '');
function imageDims(b) {
  if (b[0] === 0x89 && b[1] === 0x50) return { type: 'png', w: (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19], h: (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23] };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let o = 2;
    while (o < b.length) {
      if (b[o] !== 0xff) {
        o++;
        continue;
      }
      const m = b[o + 1];
      const len = (b[o + 2] << 8) | b[o + 3];
      if (m >= 0xc0 && m <= 0xc3) return { type: 'jpg', h: (b[o + 5] << 8) | b[o + 6], w: (b[o + 7] << 8) | b[o + 8] };
      o += 2 + len;
    }
  }
  return { type: 'unknown', w: 0, h: 0 };
}

let browser;
/** Fresh browser context per test (clean storage); collects console errors. */
async function withPage(fn, viewport = { width: 1360, height: 900 }) {
  const ctx = await browser.newContext({ acceptDownloads: true, viewport });
  const page = await ctx.newPage();
  const errors = [];
  page.on('console', (m) => {
    if (m.type() === 'error') errors.push(m.text());
  });
  page.on('pageerror', (e) => errors.push(e.message));
  try {
    await fn(page);
    expect(!errors.length, `console errors: ${errors.join(' | ')}`);
  } finally {
    await ctx.close();
  }
}

async function downloadFrom(page, trigger) {
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 60_000 }), trigger()]);
  const path = join(out, `${Date.now()}-${dl.suggestedFilename()}`);
  await dl.saveAs(path);
  return { name: dl.suggestedFilename(), bytes: new Uint8Array(readFileSync(path)) };
}

async function api(path, init) {
  return fetch(`${BASE}${path}`, init);
}

// ---------------------------------------------------------------------------- fixtures
const SAMPLE_MD = `# Отчёт о продажах

Первый абзац с **жирным** и *курсивом* текстом. Ссылка: [пример](https://example.com).

## Итоги

- Яблоки
- Груши
  - Лесные

| Товар | Кол-во |
|-------|-------:|
| Яблоки | 10 |
| Груши | 3 |
`;
// Six chapters, each short enough for one page and separated by page breaks: page N = chapter N
const LONG_MD = Array.from({ length: 6 }, (_, i) => `# Глава ${i + 1}\n\n${'Абзац текста для многостраничного документа. '.repeat(30)}\n\n<hr class="page-break">\n`).join('\n');
const SAMPLE_CSV = 'Товар;Цена;Кол-во\nЯблоки;99,5;10\n"Груши, лесные";120;3\n';

async function makeFixtures() {
  const conv = async (body, to, name) => {
    const r = await api(`/api/convert?to=${to}&name=${encodeURIComponent(name)}`, { method: 'POST', body });
    if (!r.ok) throw new Error(`fixture ${name}→${to}: ${r.status} ${await r.text()}`);
    return new Uint8Array(await r.arrayBuffer());
  };
  writeFileSync(join(fx, 'sample.md'), SAMPLE_MD);
  writeFileSync(join(fx, 'long.md'), LONG_MD);
  writeFileSync(join(fx, 'sample.csv'), SAMPLE_CSV);
  writeFileSync(join(fx, 'sample.docx'), await conv(SAMPLE_MD, 'docx', 'sample.md'));
  writeFileSync(join(fx, 'sample.xlsx'), await conv(SAMPLE_CSV, 'xlsx', 'sample.csv'));
  writeFileSync(join(fx, 'old.doc'), new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]));
  const ctx = await browser.newContext({ viewport: { width: 640, height: 400 } });
  const page = await ctx.newPage();
  await page.setContent('<div style="height:400px;background:linear-gradient(90deg,#e11d48,#4f46e5,#10b981)"><h1 style="color:#fff;font:700 48px sans-serif;padding:40px">Salatic</h1></div>');
  writeFileSync(join(fx, 'photo.png'), await page.screenshot());
  await ctx.close();
}
const F = (n) => join(fx, n);

// ---------------------------------------------------------------------------- API & server
test('API: health and formats', async () => {
  const h = await (await api('/api/health')).json();
  expect(h.status === 'ok', 'health');
  const f = await (await api('/api/formats')).json();
  expect(f.conversions.browser.pdf.includes('docx'), 'pdf→docx listed');
  expect(f.conversions.server.docx.includes('md'), 'server docx→md listed');
  expect(!f.conversions.server.pdf, 'server does not claim pdf');
});

test('API: DOCX → Markdown keeps structure', async () => {
  const r = await api('/api/convert?to=md', { method: 'POST', body: readFileSync(F('sample.docx')), headers: { 'Content-Type': 'application/octet-stream' } });
  expect(r.status === 200, `status ${r.status}`);
  expect(r.headers.get('x-source-format') === 'docx', 'detected docx');
  const md = await r.text();
  for (const s of ['# Отчёт о продажах', '**жирным**', '- Яблоки', '| Яблоки | 10 |']) expect(md.includes(s), `md contains ${s}: ${md}`);
});

test('API: multipart CSV → XLSX and Cyrillic file name', async () => {
  const form = new FormData();
  form.append('file', new Blob([SAMPLE_CSV]), 'продажи.csv');
  const r = await api('/api/convert?to=xlsx', { method: 'POST', body: form });
  expect(r.status === 200, `status ${r.status}`);
  expect(r.headers.get('content-disposition').includes("filename*=UTF-8''%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B8.xlsx"), 'utf-8 filename');
  const files = unzipSync(new Uint8Array(await r.arrayBuffer()));
  expect(text(files['xl/worksheets/sheet1.xml']).includes('Груши, лесные'), 'cell value');
});

test('API: errors are explicit (400/405/413/422/404) and CORS works', async () => {
  expect((await api('/api/convert', { method: 'POST', body: 'x' })).status === 400, 'missing target → 400');
  expect((await api('/api/convert?to=docx')).status === 405, 'GET → 405');
  const pdfReq = await api('/api/convert?to=docx&from=pdf', { method: 'POST', body: '%PDF-1.4' });
  expect(pdfReq.status === 422 && (await pdfReq.json()).error.includes('браузере'), 'pdf on server → 422 with hint');
  expect((await api('/api/convert?to=xlsx&name=a.txt', { method: 'POST', body: 'hello' })).status === 422, 'txt→xlsx → 422');
  const big = await api('/api/convert?to=md&name=a.txt', { method: 'POST', body: new Uint8Array(11 * 1024 * 1024) });
  expect(big.status === 413, `too large → 413 (got ${big.status})`);
  expect((await api('/api/nope')).status === 404, 'unknown → 404');
  const pre = await api('/api/convert', { method: 'OPTIONS' });
  expect(pre.status === 204 && pre.headers.get('access-control-allow-origin') === '*', 'CORS preflight');
});

test('Static: security headers, SPA fallback, immutable assets', async () => {
  const r = await fetch(`${BASE}/pdf-editor`);
  expect(r.status === 200 && (await r.text()).includes('<div') , 'SPA route served');
  const csp = r.headers.get('content-security-policy') ?? '';
  expect(csp.includes("default-src 'self'") && csp.includes("frame-ancestors 'none'"), `CSP: ${csp}`);
  expect(r.headers.get('x-content-type-options') === 'nosniff', 'nosniff');
  const html = await (await fetch(`${BASE}/`)).text();
  const asset = /\/assets\/[^"']+\.js/.exec(html)?.[0];
  const a = await fetch(`${BASE}${asset}`);
  expect((a.headers.get('cache-control') ?? '').includes('immutable'), 'immutable cache on assets');
  expect((await fetch(`${BASE}/pdfjs/cmaps/Identity-H.bcmap`)).status === 200, 'pdf.js cmaps served');
});

// ---------------------------------------------------------------------------- converter
async function convertOne(page, files, target) {
  await page.goto(`${BASE}/`);
  await page.setInputFiles('.dropzone input[type=file]', files);
  await page.waitForSelector('.file .select');
  if (target) await page.selectOption('.file .select', target);
  await page.click('.file .file__actions .btn--primary');
  await page.waitForSelector('.file .note--ok, .file .note--error', { timeout: 90_000 });
  const err = await page.$('.file .note--error');
  if (err) throw new AssertionError(`conversion failed: ${await err.textContent()}`);
  return downloadFrom(page, () => page.click('.file .file__actions .btn--primary'));
}

test('Converter: DOCX → PDF (default target) with Cyrillic text', async () => {
  await withPage(async (page) => {
    const { name, bytes } = await convertOne(page, F('sample.docx'));
    expect(name === 'sample.pdf', name);
    const t = (await pdfText(bytes)).join('\n');
    for (const s of ['Отчёт о продажах', 'жирным', 'Яблоки', 'Лесные']) expect(t.includes(s), `pdf has ${s}`);
    writeFileSync(F('sample.pdf'), bytes);
  });
});

test('Converter: PDF → Word reconstructs headings, list and table', async () => {
  await withPage(async (page) => {
    const { name, bytes } = await convertOne(page, F('sample.pdf'));
    expect(name === 'sample.docx', name);
    const files = unzipSync(bytes);
    const xml = text(files['word/document.xml']);
    expect(xml.includes('w:val="Heading1"'), 'heading style');
    expect(xml.includes('<w:numPr>'), 'list numbering');
    expect(xml.includes('<w:tbl>'), 'table');
    expect(docxText(bytes).includes('Отчёт о продажах'), 'text');
  });
});

test('Converter: Markdown → PDF multi-page, then PDF → JPG pages as ZIP', async () => {
  await withPage(async (page) => {
    const pdf = await convertOne(page, F('long.md'), 'pdf');
    const doc = await PDFDocument.load(pdf.bytes);
    expect(doc.getPageCount() === 6, `pages: ${doc.getPageCount()}`);
    writeFileSync(F('long.pdf'), pdf.bytes);
    const zip = await convertOne(page, F('long.pdf'), 'jpg');
    expect(zip.name === 'long-jpg.zip', zip.name);
    const entries = unzipSync(zip.bytes);
    const names = Object.keys(entries);
    expect(names.length === doc.getPageCount(), `zip entries ${names.length}`);
    const dims = imageDims(entries[names[0]]);
    expect(dims.type === 'jpg' && dims.w > 1000, `jpg ${JSON.stringify(dims)}`);
  });
});

test('Converter: PNG → JPG, XLSX → CSV (semicolon option), CSV → XLSX', async () => {
  await withPage(async (page) => {
    const jpg = await convertOne(page, F('photo.png'), 'jpg');
    const d = imageDims(jpg.bytes);
    expect(d.type === 'jpg' && d.w === 640 && d.h === 400, JSON.stringify(d));

    await page.goto(`${BASE}/`);
    await page.setInputFiles('.dropzone input[type=file]', F('sample.xlsx'));
    await page.waitForSelector('.file .select');
    expect((await page.inputValue('.file .select')) === 'csv', 'xlsx defaults to csv');
    await page.click('button[aria-label="Параметры"]');
    await page.selectOption('.options select', ';');
    await page.click('.file .file__actions .btn--primary');
    await page.waitForSelector('.file .note--ok');
    const csv = await downloadFrom(page, () => page.click('.file .file__actions .btn--primary'));
    expect(text(csv.bytes).includes('Товар;Цена;Кол-во'), `csv: ${text(csv.bytes)}`);

    const xlsx = await convertOne(page, F('sample.csv'), 'xlsx');
    expect(text(unzipSync(xlsx.bytes)['xl/worksheets/sheet1.xml']).includes('Груши, лесные'), 'xlsx cell');
  });
});

test('Converter: batch, ZIP download, merge images, unsupported file message', async () => {
  await withPage(async (page) => {
    await page.goto(`${BASE}/`);
    await page.setInputFiles('.dropzone input[type=file]', [F('sample.md'), F('photo.png'), F('photo.png'), F('old.doc')]);
    await page.waitForFunction(() => document.querySelectorAll('.file').length === 4 && document.querySelectorAll('.file .select').length === 3);
    expect((await page.textContent('.file:nth-child(4) .note--error')).includes('.docx'), 'old .doc explains');
    await page.click('.bulk .btn--primary');
    await page.waitForFunction(() => document.querySelectorAll('.file .note--ok').length === 3, null, { timeout: 90_000 });
    const all = await downloadFrom(page, () => page.getByRole('button', { name: /Скачать все/ }).click());
    expect(Object.keys(unzipSync(all.bytes)).length === 3, 'zip of 3 results');
    const merged = await downloadFrom(page, () => page.getByRole('button', { name: /Объединить 2 изобр/ }).click());
    expect((await PDFDocument.load(merged.bytes)).getPageCount() === 2, 'merged pdf has 2 pages');
  });
});

// ---------------------------------------------------------------------------- document editor
test('Doc editor: open DOCX, edit, save as DOCX and PDF, draft survives reload', async () => {
  await withPage(async (page) => {
    await page.goto(`${BASE}/doc-editor`);
    const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.getByRole('button', { name: 'Открыть' }).click()]);
    await chooser.setFiles(F('sample.docx'));
    await page.waitForFunction(() => document.querySelector('.doc-page')?.textContent?.includes('Отчёт о продажах'));
    expect(await page.$('.doc-page h1'), 'heading rendered');
    expect(await page.$('.doc-page table'), 'table rendered');
    await page.click('.doc-page h1');
    await page.keyboard.press('End');
    await page.keyboard.type(' — правка E2E');
    await page.keyboard.press('Enter');
    await page.keyboard.press('Control+b');
    await page.keyboard.type('Жирная строка');
    await page.selectOption('select[aria-label="Формат сохранения"]', 'docx');
    const docx = await downloadFrom(page, () => page.getByRole('button', { name: 'Скачать' }).click());
    expect(docx.name === 'sample.docx', docx.name);
    const xml = text(unzipSync(docx.bytes)['word/document.xml']);
    expect(docxText(docx.bytes).includes('Отчёт о продажах — правка E2E'), 'edited heading saved');
    expect(/<w:b\/>.*?Жирная строка/s.test(xml), 'bold text saved');
    await page.selectOption('select[aria-label="Формат сохранения"]', 'pdf');
    const pdf = await downloadFrom(page, () => page.getByRole('button', { name: 'Скачать' }).click());
    expect((await pdfText(pdf.bytes)).join(' ').includes('правка E2E'), 'pdf has edit');
    await page.waitForTimeout(1000); // autosave debounce
    await page.reload();
    await page.waitForSelector('.note--ok');
    expect((await page.textContent('.doc-page')).includes('Жирная строка'), 'draft restored');
  });
});

test('Doc editor: font, size, highlight, footnote, footer with page numbers', async () => {
  await withPage(async (page) => {
    await page.goto(`${BASE}/doc-editor`);
    await page.click('.doc-page');
    await page.keyboard.type('Обычный текст ');
    await page.keyboard.type('крупный');
    // select the last word and make it Times New Roman 14 pt, highlighted
    await page.keyboard.down('Shift');
    for (let i = 0; i < 'крупный'.length; i++) await page.keyboard.press('ArrowLeft');
    await page.keyboard.up('Shift');
    await page.selectOption('select[aria-label="Шрифт"]', 'serif');
    await page.selectOption('select[aria-label="Размер шрифта"]', '14');
    await page.evaluate(() => {
      const m = document.querySelector('input[aria-label="Цвет выделения текста"]');
      m.value = '#ffff00';
      m.dispatchEvent(new Event('change', { bubbles: true }));
    });
    await page.keyboard.press('End');
    await page.click('button[aria-label="Вставить сноску"]');
    await page.keyboard.type('Пояснение в сноске');
    await page.click('button[aria-label^="Нижний колонтитул"]');
    await page.selectOption('select[aria-label="Формат сохранения"]', 'docx');
    const docx = await downloadFrom(page, () => page.getByRole('button', { name: 'Скачать' }).click());
    const files = unzipSync(docx.bytes);
    const xml = text(files['word/document.xml']);
    expect(/Times New Roman.*?<w:sz w:val="28"\/>.*?крупный/s.test(xml), `font+size: ${xml.slice(0, 600)}`);
    expect(/w:fill="FFFF00"/i.test(xml), 'highlight');
    expect(text(files['word/footnotes.xml'] ?? new Uint8Array()).includes('Пояснение в сноске'), 'footnote saved');
    expect(text(files['word/footer1.xml'] ?? new Uint8Array()).includes('PAGE'), 'page field in footer');
    await page.selectOption('select[aria-label="Формат сохранения"]', 'pdf');
    const pdf = await downloadFrom(page, () => page.getByRole('button', { name: 'Скачать' }).click());
    const t = (await pdfText(pdf.bytes)).join(' ');
    expect(t.includes('Пояснение в сноске'), `pdf footnote: ${t}`);
    expect(/Страница\s+1\s+из\s+1/.test(t), `pdf footer: ${t}`);
  });
});

test('Doc editor: paste is sanitised (no scripts or handlers)', async () => {
  await withPage(async (page) => {
    await page.goto(`${BASE}/doc-editor`);
    await page.click('.doc-page');
    await page.evaluate(() => {
      const dt = new DataTransfer();
      dt.setData('text/html', '<p onclick="alert(1)">Вставка <b>жирная</b><script>window.__x=1</script><img src=x onerror="window.__y=1"></p>');
      document.querySelector('.doc-page').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    });
    const html = await page.innerHTML('.doc-page');
    expect(html.includes('<strong>жирная</strong>') || html.includes('<b>жирная</b>'), html);
    expect(!/onclick|onerror|script/i.test(html), `sanitised: ${html}`);
    expect(!(await page.evaluate(() => window.__x || window.__y)), 'no script ran');
  });
});

// ---------------------------------------------------------------------------- image editor
test('Image editor: rotate, crop, resize, adjust, draw, undo, export JPG', async () => {
  await withPage(async (page) => {
    await page.goto(`${BASE}/image-editor`);
    await page.setInputFiles('.dropzone input[type=file]', F('photo.png'));
    const dims = async () => (await page.textContent('.toolbar span[style*="margin-right"]')).trim();
    await page.waitForFunction(() => document.querySelector('.img-stage canvas'));
    expect((await dims()) === '640 × 400', await dims());
    await page.click('button[aria-label="Повернуть вправо"]');
    expect((await dims()) === '400 × 640', `rotated ${await dims()}`);
    await page.click('button[aria-label="Обрезка"]');
    await page.locator('.panel__section', { hasText: 'Обрезка' }).getByRole('button', { name: 'Применить' }).click();
    // default crop box: equal 10% margin of the shorter side
    expect((await dims()) === '320 × 560', `cropped ${await dims()}`);
    await page.getByRole('button', { name: '50%' }).click();
    expect((await dims()) === '160 × 280', `resized ${await dims()}`);
    await page.evaluate(() => {
      const r = document.querySelector('.slider input[type=range]');
      r.value = '40';
      r.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await page.locator('.panel__section', { hasText: 'Коррекция' }).getByRole('button', { name: 'Применить' }).click();
    await page.click('button[aria-label="Рисование"]');
    const box = await page.locator('.img-stage canvas').boundingBox();
    await page.mouse.move(box.x + 10, box.y + 10);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width - 10, box.y + box.height - 10, { steps: 8 });
    await page.mouse.up();
    await page.click('button[aria-label="Отменить (Ctrl+Z)"]');
    await page.click('button[aria-label="Повторить (Ctrl+Y)"]');
    await page.selectOption('select[aria-label="Формат экспорта"]', 'jpg');
    const jpg = await downloadFrom(page, () => page.locator('.panel').getByRole('button', { name: 'Скачать' }).click());
    const d = imageDims(jpg.bytes);
    expect(d.type === 'jpg' && d.w === 160 && d.h === 280, JSON.stringify(d));
  });
});

// ---------------------------------------------------------------------------- PDF editor
test('PDF editor: merge PDF + image, rotate, delete, annotate, save', async () => {
  await withPage(async (page) => {
    await page.goto(`${BASE}/pdf-editor`);
    const srcPages = (await PDFDocument.load(readFileSync(F('long.pdf')))).getPageCount();
    await page.setInputFiles('.dropzone input[type=file]', [F('long.pdf'), F('photo.png')]);
    await page.waitForFunction((n) => document.querySelectorAll('.page-card').length === n, srcPages + 1, { timeout: 60_000 });
    await page.waitForSelector('.page-card canvas');
    const card = (i) => page.locator('.page-card').nth(i);
    await card(0).locator('button[aria-label="Повернуть вправо"]').click();
    await card(1).locator('button[aria-label="Удалить страницу"]').click();
    await page.waitForFunction((n) => document.querySelectorAll('.page-card').length === n, srcPages);
    // Move the image page (last) to the front
    const last = srcPages - 1;
    for (let i = last; i > 0; i--) await card(i).locator('button[aria-label="Переместить влево"]').click();
    await card(1).locator('button[aria-label="Редактировать страницу"]').click();
    await page.waitForSelector('.page-stage canvas');
    await page.click('button[aria-label="Текст"]');
    const stage = await page.locator('.page-stage').boundingBox();
    await page.mouse.click(stage.x + 80, stage.y + 120);
    await page.keyboard.type('Подпись E2E');
    await page.keyboard.press('Control+Enter');
    await page.click('button[aria-label="Закрасить (замазка)"]');
    await page.mouse.move(stage.x + 60, stage.y + 300);
    await page.mouse.down();
    await page.mouse.move(stage.x + 260, stage.y + 360, { steps: 4 });
    await page.mouse.up();
    expect((await page.$$('.page-stage svg g')).length === 2, 'two annotations visible');
    await page.getByRole('button', { name: 'К страницам' }).click();
    await page.waitForSelector('.page-card__badge');
    const saved = await downloadFrom(page, () => page.locator('.toolbar').getByRole('button', { name: 'Сохранить PDF' }).click());
    const doc = await PDFDocument.load(saved.bytes);
    expect(doc.getPageCount() === srcPages, `page count ${doc.getPageCount()}`);
    const first = doc.getPage(0);
    expect(Math.round(first.getWidth()) === 480 && Math.round(first.getHeight()) === 300, `image page first ${first.getWidth()}x${first.getHeight()}`);
    expect(doc.getPage(1).getRotation().angle === 90, `rotation ${doc.getPage(1).getRotation().angle}`);
    const texts = await pdfText(saved.bytes);
    expect(texts[1].includes('Подпись E2E'), `annotation text on page 2: ${texts[1].slice(0, 120)}`);
    // Expected order: image, chapter 1 (rotated, annotated), chapters 3..6 – chapter 2 was deleted
    const chapters = texts.map((t) => /Глава (\d)/.exec(t)?.[1] ?? '-').join('');
    expect(chapters === '-13456', `page order ${chapters}`);
  });
});

// ---------------------------------------------------------------------------- responsive
test('Mobile layout: no horizontal scrolling on any page', async () => {
  await withPage(
    async (page) => {
      for (const path of ['/', '/pdf-editor', '/image-editor', '/doc-editor']) {
        await page.goto(`${BASE}${path}`);
        await page.waitForSelector('.page-head h1');
        const sw = await page.evaluate(() => document.documentElement.scrollWidth);
        expect(sw <= 390, `${path}: scrollWidth ${sw}`);
      }
    },
    { width: 390, height: 844 },
  );
});

// ---------------------------------------------------------------------------- runner
async function startServer() {
  if (BASE) return;
  BASE = `http://127.0.0.1:${PORT}`;
  server = spawn('npx', ['wrangler', 'dev', '--port', String(PORT), '--ip', '127.0.0.1', '--log-level', 'warn'], { cwd: root, stdio: ['ignore', 'pipe', 'pipe'], detached: true, env: { ...process.env, WRANGLER_SEND_METRICS: 'false' } });
  let log = '';
  server.stdout.on('data', (d) => (log += d));
  server.stderr.on('data', (d) => (log += d));
  for (let i = 0; i < 90; i++) {
    try {
      if ((await fetch(`${BASE}/api/health`)).ok) return;
    } catch {
      /* not up yet */
    }
    await new Promise((r) => setTimeout(r, 1000));
  }
  throw new Error(`wrangler dev did not start:\n${log}`);
}

function stopServer() {
  if (server) {
    try {
      process.kill(-server.pid, 'SIGTERM');
    } catch {
      /* already gone */
    }
  }
}

let failed = 0;
try {
  await startServer();
  browser = await chromium.launch();
  await makeFixtures();
  const only = process.env.ONLY ? new RegExp(process.env.ONLY, 'i') : null;
  for (const t of tests) {
    if (only && !only.test(t.name)) continue;
    const t0 = Date.now();
    try {
      await t.fn();
      console.log(`  ✓ ${t.name} (${Date.now() - t0} ms)`);
    } catch (e) {
      failed++;
      console.log(`  ✗ ${t.name}\n      ${(e.stack || e.message || String(e)).split('\n').slice(0, 4).join('\n      ')}`);
    }
  }
  const run = only ? tests.filter((t) => only.test(t.name)).length : tests.length;
  console.log(`\n${run - failed} passed, ${failed} failed (${run} total) against ${BASE}`);
} catch (e) {
  failed++;
  console.error(e);
} finally {
  await browser?.close();
  stopServer();
}
process.exit(failed ? 1 : 0);
