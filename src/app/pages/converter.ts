import { acceptFor, detectFormat, DOC_READABLE, FORMATS, IMAGE_INPUTS, pipelineFor, targetLabel, targetsFor, type FormatId } from '../../core/formats';
import type { ConvertOptions } from '../../core/convert-lite';
import { engine, EngineError } from '../engine/client';
import type { ConvertResult } from '../engine/protocol';
import { convertImage, pdfToImages, supportsAvifEncoding, toPdfImages, zipFiles } from '../raster';
import { baseName, button, download, dropzone, formatBytes, h, handOff, icon, iconButton, note, onPasteFiles, prefs, readFile, toast } from '../ui';

interface Options {
  pageSize: 'auto' | 'a4' | 'letter';
  imagePage: 'fit' | 'a4' | 'letter';
  quality: number;
  dpi: number;
  csvDelimiter: ',' | ';';
  /** Text recognition for PDFs: only scanned pages, every page, or never. */
  ocr: 'auto' | 'on' | 'off';
  /** Add an invisible OCR text layer when images become a PDF. */
  imageOcr: boolean;
}

interface Item {
  id: number;
  file: File;
  bytes: Uint8Array | null;
  format: FormatId | null;
  error?: string;
  target: FormatId | null;
  status: 'idle' | 'working' | 'done' | 'error';
  progress: number;
  message?: string;
  result?: { data: Uint8Array; name: string; mime: string };
  warnings: string[];
  abort?: AbortController;
  opts: Options;
  showOptions: boolean;
  thumbUrl?: string;
  el: HTMLElement;
}

const POPULAR: { label: string; from: FormatId[]; to: FormatId }[] = [
  { label: 'PDF → Word', from: ['pdf'], to: 'docx' },
  { label: 'Word → PDF', from: ['docx'], to: 'pdf' },
  { label: 'JPG → PDF', from: ['jpg', 'png'], to: 'pdf' },
  { label: 'PDF → JPG', from: ['pdf'], to: 'jpg' },
  { label: 'PNG → JPG', from: ['png'], to: 'jpg' },
  { label: 'WebP → PNG', from: ['webp'], to: 'png' },
  { label: 'Excel → CSV', from: ['xlsx'], to: 'csv' },
  { label: 'CSV → Excel', from: ['csv'], to: 'xlsx' },
  { label: 'Word → Markdown', from: ['docx'], to: 'md' },
  { label: 'Markdown → PDF', from: ['md'], to: 'pdf' },
  { label: 'HTML → Word', from: ['html'], to: 'docx' },
  { label: 'JSON → Excel', from: ['json'], to: 'xlsx' },
  { label: 'DOC → DOCX', from: ['doc'], to: 'docx' },
  { label: 'XLS → XLSX', from: ['xls'], to: 'xlsx' },
  { label: 'ODT → Word', from: ['odt'], to: 'docx' },
  { label: 'RTF → Word', from: ['rtf'], to: 'docx' },
  { label: 'PowerPoint → PDF', from: ['pptx'], to: 'pdf' },
  { label: 'EPUB → PDF', from: ['epub'], to: 'pdf' },
  { label: 'Word → EPUB', from: ['docx'], to: 'epub' },
];

const DEFAULT_TARGET: Partial<Record<FormatId, FormatId>> = {
  pdf: 'docx',
  docx: 'pdf',
  doc: 'docx',
  odt: 'docx',
  rtf: 'docx',
  epub: 'pdf',
  pptx: 'pdf',
  ods: 'xlsx',
  xls: 'xlsx',
  txt: 'pdf',
  md: 'pdf',
  html: 'pdf',
  csv: 'xlsx',
  tsv: 'xlsx',
  json: 'xlsx',
  xlsx: 'csv',
  png: 'jpg',
  jpg: 'png',
  webp: 'png',
  gif: 'png',
  bmp: 'png',
  svg: 'png',
  ico: 'png',
  avif: 'png',
};

const ACCEPT = acceptFor([...DOC_READABLE, ...IMAGE_INPUTS]);
const FILE_LIMIT = 300 * 1024 * 1024;

const kindIcon = (f: FormatId | null): string => (!f ? 'file' : FORMATS[f].kind === 'image' ? 'image' : FORMATS[f].kind === 'data' ? 'table' : 'file-text');

export function mount(root: HTMLElement, params: URLSearchParams): () => void {
  let items: Item[] = [];
  let nextId = 1;
  let avif = false;
  void supportsAvifEncoding().then((v) => {
    avif = v;
    items.forEach(paint);
  });
  const savedOpts = prefs.get<Partial<Options>>('convert-options', {});
  const defaults = (): Options => ({ ocr: 'auto', imageOcr: false, pageSize: 'auto', imagePage: 'fit', quality: 0.9, dpi: 150, csvDelimiter: ',', ...savedOpts });
  const urlTo = params.get('to') as FormatId | null;
  let preferred: FormatId | null = urlTo && FORMATS[urlTo] ? urlTo : null;

  // ------------------------------------------------------------------ layout
  const chips = h(
    'div',
    { class: 'chips', role: 'group', 'aria-label': 'Популярные преобразования' },
    POPULAR.map((p) =>
      h(
        'button',
        {
          type: 'button',
          class: 'chip',
          'aria-pressed': String(preferred === p.to && params.get('from') === p.from[0]),
          onclick: (e: MouseEvent) => {
            const btn = e.currentTarget as HTMLButtonElement;
            const on = btn.getAttribute('aria-pressed') !== 'true';
            chips.querySelectorAll('.chip').forEach((c) => c.setAttribute('aria-pressed', 'false'));
            btn.setAttribute('aria-pressed', String(on));
            preferred = on ? p.to : null;
            history.replaceState(null, '', on ? `/?from=${p.from[0]}&to=${p.to}` : '/');
            // Apply to files that are waiting
            for (const it of items) {
              if (on && it.format && p.from.includes(it.format) && it.status !== 'working') setTarget(it, p.to);
            }
            if (on && !items.length) zone.click();
          },
        },
        p.label,
      ),
    ),
  );

  const zone = dropzone({
    accept: ACCEPT,
    title: 'Перетащите файлы сюда или нажмите, чтобы выбрать',
    hint: 'PDF, Word (DOCX/DOC), Excel (XLSX/XLS), PowerPoint, OpenDocument, RTF, EPUB, CSV, JSON, Markdown, HTML, TXT, JPG, PNG, WebP, GIF, SVG и другие · можно вставить из буфера (Ctrl+V)',
    onFiles: addFiles,
  });
  onPasteFiles(root, addFiles);

  const list = h('div', { class: 'files', 'aria-live': 'polite' });
  const bulk = h('div', { class: 'bulk card', hidden: true });

  root.append(
    h(
      'div',
      { class: 'page-head' },
      h('h1', null, 'Конвертер файлов'),
      h('p', null, 'Быстро и приватно: файлы обрабатываются прямо в вашем браузере и никуда не загружаются. PDF в Word, Word в PDF, изображения, таблицы и текстовые форматы.'),
    ),
    zone,
    chips,
    list,
    bulk,
    h(
      'section',
      { class: 'features', 'aria-label': 'Возможности' },
      feature('Редактор PDF', 'Объединяйте, переставляйте, поворачивайте и удаляйте страницы, добавляйте текст, подписи, выделения и пометки.', '/pdf-editor'),
      feature('Редактор изображений', 'Обрезка, поворот, изменение размера, яркость и контраст, рисование и текст, экспорт в JPG/PNG/WebP.', '/image-editor'),
      feature('Редактор документов', 'Откройте DOCX, PDF, Markdown или HTML, отредактируйте и сохраните в Word, PDF и другие форматы.', '/doc-editor'),
      h('div', { class: 'feature card' }, h('h3', null, 'Конфиденциально'), h('p', null, 'Преобразование выполняется локально на вашем устройстве. Файлы не отправляются на сервер и не сохраняются.')),
    ),
  );

  function feature(title: string, text: string, href: string): HTMLElement {
    return h('div', { class: 'feature card' }, h('h3', null, title), h('p', null, text, ' ', h('a', { href, 'data-link': true }, 'Открыть →')));
  }

  // ------------------------------------------------------------------ items
  async function addFiles(files: File[]): Promise<void> {
    for (const file of files) {
      const item: Item = {
        id: nextId++,
        file,
        bytes: null,
        format: null,
        target: null,
        status: 'idle',
        progress: 0,
        warnings: [],
        opts: defaults(),
        showOptions: false,
        el: h('div', { class: 'file card' }),
      };
      items.push(item);
      list.append(item.el);
      if (file.size > FILE_LIMIT) {
        item.error = `Файл слишком большой (${formatBytes(file.size)}). Максимум — ${formatBytes(FILE_LIMIT)}.`;
        item.status = 'error';
        paint(item);
        continue;
      }
      try {
        item.bytes = await readFile(file);
      } catch {
        item.error = 'Не удалось прочитать файл.';
        item.status = 'error';
        paint(item);
        continue;
      }
      const det = detectFormat(file.name, item.bytes);
      item.format = det.format;
      if (!det.format) {
        item.error = det.error;
        item.status = 'error';
      } else {
        const targets = targetsFor(det.format, { avif });
        item.target = preferred && targets.includes(preferred) ? preferred : (DEFAULT_TARGET[det.format] ?? targets[0] ?? null);
        // Browsers cannot show TIFF/HEIC/SVG thumbnails directly: those keep the generic icon
        if (FORMATS[det.format].kind === 'image' && !['svg', 'tiff', 'heic'].includes(det.format)) item.thumbUrl = URL.createObjectURL(file);
      }
      paint(item);
    }
    paintBulk();
  }

  function setTarget(item: Item, to: FormatId): void {
    item.target = to;
    resetResult(item);
    paint(item);
    paintBulk();
  }

  function resetResult(item: Item): void {
    if (item.status === 'done' || item.status === 'error') {
      item.status = 'idle';
      item.result = undefined;
      item.warnings = [];
      item.message = undefined;
    }
  }

  function removeItem(item: Item): void {
    item.abort?.abort();
    if (item.thumbUrl) URL.revokeObjectURL(item.thumbUrl);
    item.el.remove();
    items = items.filter((i) => i !== item);
    paintBulk();
  }

  function saveOpts(item: Item): void {
    prefs.set('convert-options', item.opts);
    resetResult(item);
    paint(item);
    paintBulk();
  }

  function optionsPanel(item: Item): HTMLElement | null {
    if (!item.format || !item.target) return null;
    const from = item.format;
    const to = item.target;
    const pipe = pipelineFor(from, to);
    const fields: HTMLElement[] = [];
    const select = (label: string, value: string, options: [string, string][], onChange: (v: string) => void): HTMLElement =>
      h(
        'label',
        { class: 'field' },
        h('span', null, label),
        h(
          'select',
          { class: 'select', onchange: (e: Event) => onChange((e.target as HTMLSelectElement).value) },
          options.map(([v, t]) => h('option', { value: v, selected: v === value }, t)),
        ),
      );
    if (to === 'pdf' && pipe === 'doc') {
      fields.push(select('Размер страницы', item.opts.pageSize, [['auto', 'Как в документе'], ['a4', 'A4'], ['letter', 'Letter']], (v) => ((item.opts.pageSize = v as Options['pageSize']), saveOpts(item))));
    }
    if (from === 'pdf' && pipe === 'doc') {
      fields.push(
        select('Распознавание текста (OCR)', item.opts.ocr, [['auto', 'Авто — только сканы'], ['on', 'Все страницы'], ['off', 'Выключено']], (v) => ((item.opts.ocr = v as Options['ocr']), saveOpts(item))),
      );
    }
    if (pipe === 'images-to-pdf') {
      fields.push(
        h(
          'label',
          { class: 'checkbox', style: 'align-self:center' },
          h('input', { type: 'checkbox', checked: item.opts.imageOcr, onchange: (e: Event) => ((item.opts.imageOcr = (e.target as HTMLInputElement).checked), saveOpts(item)) }),
          'Распознать текст (поиск в PDF)',
        ),
      );
    }
    if (pipe === 'images-to-pdf') {
      fields.push(
        select('Страница', item.opts.imagePage, [['fit', 'По размеру изображения'], ['a4', 'A4 с полями'], ['letter', 'Letter с полями']], (v) => ((item.opts.imagePage = v as Options['imagePage']), saveOpts(item))),
      );
    }
    if (from === 'pdf' && pipe === 'raster') {
      fields.push(select('Разрешение', String(item.opts.dpi), [['72', '72 DPI (экран)'], ['150', '150 DPI'], ['300', '300 DPI (печать)']], (v) => ((item.opts.dpi = +v), saveOpts(item))));
    }
    if (to === 'jpg' || to === 'webp' || to === 'avif') {
      const out = h('output', null, `${Math.round(item.opts.quality * 100)}%`);
      fields.push(
        h(
          'label',
          { class: 'field slider', style: 'min-width:180px' },
          h('span', null, 'Качество'),
          out,
          h('input', {
            type: 'range',
            min: 30,
            max: 100,
            value: Math.round(item.opts.quality * 100),
            oninput: (e: Event) => (out.textContent = `${(e.target as HTMLInputElement).value}%`),
            onchange: (e: Event) => ((item.opts.quality = +(e.target as HTMLInputElement).value / 100), saveOpts(item)),
          }),
        ),
      );
    }
    if (to === 'csv') {
      fields.push(select('Разделитель', item.opts.csvDelimiter, [[',', 'Запятая ,'], [';', 'Точка с запятой ; (Excel RU)']], (v) => ((item.opts.csvDelimiter = v as ',' | ';'), saveOpts(item))));
    }
    return fields.length ? h('div', { class: 'options' }, fields) : null;
  }

  function hasOptions(item: Item): boolean {
    if (!item.format || !item.target) return false;
    const pipe = pipelineFor(item.format, item.target);
    return (item.target === 'pdf' && pipe === 'doc') || (item.format === 'pdf' && pipe === 'doc') || pipe === 'images-to-pdf' || (item.format === 'pdf' && pipe === 'raster') || ['jpg', 'webp', 'avif', 'csv'].includes(item.target);
  }

  function paint(item: Item): void {
    const f = item.format;
    const thumb = h('div', { class: 'file__icon' }, item.thumbUrl ? h('img', { src: item.thumbUrl, alt: '' }) : icon(kindIcon(f)));
    const meta = h('div', { class: 'file__meta' }, f ? h('span', { class: `badge badge--${FORMATS[f].kind}` }, FORMATS[f].label) : null, h('span', null, formatBytes(item.file.size)));
    const actions = h('div', { class: 'file__actions' });
    if (f && item.target) {
      const targets = targetsFor(f, { avif });
      actions.append(
        h('span', { class: 'file__arrow', style: 'color:var(--muted);display:inline-flex' }, icon('arrow')),
        h(
          'select',
          {
            class: 'select',
            'aria-label': `Формат для ${item.file.name}`,
            disabled: item.status === 'working',
            onchange: (e: Event) => setTarget(item, (e.target as HTMLSelectElement).value as FormatId),
          },
          targets.map((t) => h('option', { value: t, selected: t === item.target }, targetLabel(f, t))),
        ),
      );
      if (hasOptions(item)) {
        actions.append(
          iconButton('settings', 'Параметры', () => {
            item.showOptions = !item.showOptions;
            paint(item);
          }, { 'aria-pressed': String(item.showOptions), 'aria-expanded': String(item.showOptions) }),
        );
      }
      if (item.status === 'working') actions.append(button('Отмена', () => item.abort?.abort(), { small: false }));
      else if (item.status === 'done' && item.result) {
        const r = item.result;
        actions.append(button('Скачать', () => download(r.data, r.name, r.mime), { icon: 'download', variant: 'primary' }));
      } else actions.append(button('Преобразовать', () => void run(item), { variant: 'primary' }));
    }
    actions.append(iconButton('x', 'Убрать файл', () => removeItem(item)));

    const status = h('div', { class: 'file__status' });
    if (item.showOptions) {
      const panel = optionsPanel(item);
      if (panel) status.append(panel);
    }
    if (item.status === 'working') {
      const pct = Math.round(item.progress * 100);
      status.append(
        h('div', { class: 'progress', role: 'progressbar', 'aria-valuenow': pct, 'aria-valuemin': 0, 'aria-valuemax': 100 }, h('div', { style: { width: `${Math.max(4, pct)}%` } })),
        h('div', { style: 'color:var(--muted);font-size:13px' }, item.message ?? 'Обработка…'),
      );
    } else if (item.status === 'error') status.append(note('error', item.error ?? 'Ошибка'));
    else if (item.status === 'done' && item.result) {
      const r = item.result;
      const editor = editorFor(r.name);
      status.append(
        h(
          'div',
          { class: 'row' },
          note('ok', `Готово: ${r.name} · ${formatBytes(r.data.length)}`),
          editor ? button(editor.label, () => handOff(new File([r.data as BlobPart], r.name, { type: r.mime }), editor.path), { small: true, variant: 'ghost', icon: 'pen' }) : null,
        ),
      );
      for (const w of item.warnings) status.append(note('warn', w));
    }
    item.el.replaceChildren(thumb, h('div', { class: 'file__main' }, h('div', { class: 'file__name', title: item.file.name }, item.file.name), meta), actions, status);
  }

  function editorFor(name: string): { label: string; path: string } | null {
    const ext = name.split('.').pop()!.toLowerCase();
    if (ext === 'pdf') return { label: 'Открыть в редакторе PDF', path: '/pdf-editor' };
    if (['docx', 'md', 'html', 'txt'].includes(ext)) return { label: 'Редактировать', path: '/doc-editor' };
    if (['png', 'jpg', 'webp', 'bmp'].includes(ext)) return { label: 'Редактировать', path: '/image-editor' };
    return null;
  }

  async function run(item: Item): Promise<void> {
    if (!item.bytes || !item.format || !item.target || item.status === 'working') return;
    const from = item.format;
    const to = item.target;
    const base = baseName(item.file.name);
    const ctrl = new AbortController();
    item.abort = ctrl;
    item.status = 'working';
    item.progress = 0;
    item.message = 'Обработка…';
    item.warnings = [];
    paint(item);
    paintBulk();
    const onProgress = (v: number): void => {
      item.progress = v;
      const bar = item.el.querySelector<HTMLElement>('.progress > div');
      if (bar) bar.style.width = `${Math.max(4, Math.round(v * 100))}%`;
    };
    const t0 = performance.now();
    try {
      const pipe = pipelineFor(from, to);
      const status = (v: number, label?: string): void => {
        onProgress(v);
        if (label && label !== item.message) {
          item.message = label;
          const msg = item.el.querySelector<HTMLElement>('.file__status .progress + div');
          if (msg) msg.textContent = label;
        }
      };
      if (pipe === 'doc' && from === 'pdf' && item.opts.ocr !== 'off') {
        const { pdfToDocWithOcr } = await import('../ocr-convert');
        const opts: ConvertOptions = { pageSize: item.opts.pageSize, csvDelimiter: item.opts.csvDelimiter };
        const r = await pdfToDocWithOcr(item.bytes, to, opts, item.opts.ocr, base, status, ctrl.signal);
        item.result = { data: r.data, name: `${base}.${r.ext}`, mime: r.mime };
        item.warnings = r.warnings;
      } else if (pipe === 'ocr') {
        const { imageToDocWithOcr } = await import('../ocr-convert');
        const r = await imageToDocWithOcr(item.file, from, to, base, status);
        item.result = { data: r.data, name: `${base}.${r.ext}`, mime: r.mime };
        item.warnings = r.warnings;
      } else if (pipe === 'ocr-pdf') {
        const { searchablePdf } = await import('../ocr-convert');
        const r = await searchablePdf(item.bytes, status, ctrl.signal);
        item.result = { data: r.data, name: `${base}-ocr.pdf`, mime: FORMATS.pdf.mime };
        item.warnings = r.warnings;
      } else if (pipe === 'images-to-pdf' && item.opts.imageOcr) {
        const { imagesToSearchablePdf } = await import('../ocr-convert');
        const data = await imagesToSearchablePdf([{ file: item.file, format: from }], item.opts.imagePage, status);
        item.result = { data, name: `${base}.pdf`, mime: FORMATS.pdf.mime };
      } else if (pipe === 'doc') {
        const copy = item.bytes.slice();
        const opts: ConvertOptions = { pageSize: item.opts.pageSize, csvDelimiter: item.opts.csvDelimiter };
        const r = await engine.run<ConvertResult>({ op: 'convert', bytes: copy, from, to, opts, title: base }, { transfer: [copy.buffer], onProgress, signal: ctrl.signal });
        item.result = { data: r.data, name: `${base}.${r.ext}`, mime: r.mime };
        item.warnings = r.warnings;
      } else if (pipe === 'raster' && from === 'pdf') {
        item.message = 'Отрисовка страниц…';
        paint(item);
        const { files } = await pdfToImages(item.bytes, to as 'png' | 'jpg', item.opts.dpi, item.opts.quality, base, onProgress, ctrl.signal);
        item.result = files.length === 1 ? { data: files[0].data, name: files[0].name, mime: FORMATS[to].mime } : { data: zipFiles(files), name: `${base}-${to}.zip`, mime: 'application/zip' };
        if (files.length > 1) item.warnings = [`${files.length} страниц сохранены отдельными файлами в ZIP-архиве.`];
      } else if (pipe === 'raster') {
        const pages = await convertImage(item.file, from, to, item.opts.quality);
        if (pages.length === 1) item.result = { data: pages[0], name: `${base}.${FORMATS[to].ext}`, mime: FORMATS[to].mime };
        else {
          const files = pages.map((data, i) => ({ name: `${base}-${String(i + 1).padStart(String(pages.length).length, '0')}.${FORMATS[to].ext}`, data }));
          item.result = { data: zipFiles(files), name: `${base}-${to}.zip`, mime: 'application/zip' };
          item.warnings = [`${pages.length} страниц сохранены отдельными файлами в ZIP-архиве.`];
        }
      } else if (pipe === 'images-to-pdf') {
        const imgs = await toPdfImages(item.file, from);
        const data = await engine.run<Uint8Array>({ op: 'imagesToPdf', images: imgs, opts: { pageSize: item.opts.imagePage } }, { transfer: imgs.map((i) => i.data.buffer), signal: ctrl.signal });
        item.result = { data, name: `${base}.pdf`, mime: FORMATS.pdf.mime };
      } else throw new Error('Неподдерживаемое преобразование.');
      item.status = 'done';
      item.el.dataset.ms = String(Math.round(performance.now() - t0));
    } catch (e) {
      if ((e as EngineError).code === 'cancelled' || ctrl.signal.aborted) {
        item.status = 'idle';
      } else {
        item.status = 'error';
        item.error = (e as Error).message || 'Неизвестная ошибка';
      }
    } finally {
      item.abort = undefined;
      if (items.includes(item)) {
        paint(item);
        paintBulk();
      }
    }
  }

  // ------------------------------------------------------------------ bulk actions
  function paintBulk(): void {
    bulk.hidden = items.length === 0;
    if (!items.length) return;
    const ready = items.filter((i) => i.status === 'idle' && i.target);
    const done = items.filter((i) => i.status === 'done' && i.result);
    const images = items.filter((i) => i.format && FORMATS[i.format].kind === 'image' && i.bytes);
    const working = items.some((i) => i.status === 'working');
    bulk.replaceChildren(
      h(
        'div',
        { class: 'bulk__group' },
        button(`Преобразовать все${ready.length ? ` (${ready.length})` : ''}`, () => ready.forEach((i) => void run(i)), { variant: 'primary', attrs: { disabled: !ready.length } }),
        done.length > 1 ? button(`Скачать все (${done.length}) ZIP`, () => download(zipFiles(done.map((d) => ({ name: d.result!.name, data: d.result!.data }))), 'salatic-files.zip', 'application/zip'), { icon: 'download' }) : null,
        images.length > 1 ? button(`Объединить ${images.length} изобр. в один PDF`, () => void mergeImages(images), { icon: 'file-plus', attrs: { disabled: working } }) : null,
      ),
      button('Очистить список', () => [...items].forEach(removeItem), { variant: 'ghost', icon: 'trash' }),
    );
  }

  async function mergeImages(images: Item[]): Promise<void> {
    try {
      toast('Создаю PDF из изображений…');
      if (images[0].opts.imageOcr) {
        const { imagesToSearchablePdf } = await import('../ocr-convert');
        download(await imagesToSearchablePdf(images.map((i) => ({ file: i.file, format: i.format! })), images[0].opts.imagePage, () => {}), 'images.pdf', 'application/pdf');
        return;
      }
      const inputs = [];
      for (const it of images) inputs.push(...(await toPdfImages(it.file, it.format!)));
      const data = await engine.run<Uint8Array>(
        { op: 'imagesToPdf', images: inputs, opts: { pageSize: images[0].opts.imagePage } },
        { transfer: inputs.map((i) => i.data.buffer) },
      );
      download(data, 'images.pdf', 'application/pdf');
    } catch (e) {
      toast((e as Error).message, 'error');
    }
  }

  return () => {
    items.forEach((i) => {
      i.abort?.abort();
      if (i.thumbUrl) URL.revokeObjectURL(i.thumbUrl);
    });
  };
}
