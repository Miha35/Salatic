import { acceptFor, detectFormat, DOC_READABLE, FORMATS, type FormatId } from '../../core/formats';
import { unlockPdf } from '../pdf-unlock';
import { htmlToDoc } from '../../core/readers/html';
import { blocksToHtml, CSS_FONT, docToHtmlFragment } from '../../core/writers/text';
import { familyOf, type FontFamily } from '../../core/model';
import { escapeHtml } from '../../core/util';
import { engine } from '../engine/client';
import type { ConvertResult, DocMeta, HtmlResult } from '../engine/protocol';
import { registerEditorFonts } from '../editor-fonts';
import { baseName, button, confirmDialog, download, formatBytes, h, iconButton, note, pickFiles, prefs, promptText, readFile, takeHandoff, toast } from '../ui';

const OPEN_ACCEPT = acceptFor(DOC_READABLE);
const SAVE_FORMATS: FormatId[] = ['docx', 'pdf', 'odt', 'rtf', 'html', 'md', 'txt'];
const DRAFT_KEY = 'doc-draft';
const MAX_DRAFT = 4_000_000;

/** Sanitise arbitrary HTML through the document model (drops scripts, handlers, unknown markup). */
const sanitize = (html: string): string => blocksToHtml(htmlToDoc(html).blocks, { sources: true });

const FONT_CHOICES: [FontFamily, string][] = [
  ['sans', 'Arial'],
  ['serif', 'Times New Roman'],
  ['calibri', 'Calibri'],
  ['mono', 'Courier New'],
];
const SIZE_CHOICES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

export function mount(root: HTMLElement): () => void {
  registerEditorFonts();
  const draft = prefs.get<{ title: string; html: string; at: number; meta?: DocMeta } | null>(DRAFT_KEY, null);
  let meta: DocMeta = draft?.meta ?? { font: 'serif', size: 12 };
  let format: FormatId = prefs.get<FormatId>('doc-save-format', 'docx');
  let dirty = false;
  let saveTimer = 0;

  const editor = h('div', {
    class: 'doc-page',
    contenteditable: 'true',
    spellcheck: 'true',
    role: 'textbox',
    'aria-multiline': 'true',
    'aria-label': 'Текст документа',
    'data-placeholder': 'Начните печатать или откройте файл (Word, PDF, ODT, RTF, EPUB, Markdown, HTML, TXT)…',
  });
  const titleInput = h('input', { class: 'input', value: draft?.title ?? 'Документ', placeholder: 'Название файла', 'aria-label': 'Название документа', style: 'min-width:0;flex:1;max-width:320px' });
  const statusLeft = h('span');
  const statusRight = h('span');
  const banner = h('div');

  // ------------------------------------------------------------------ commands
  const exec = (cmd: string, value?: string): void => {
    editor.focus();
    document.execCommand(cmd, false, value);
    changed();
    refreshState();
  };

  const blockSelect = h(
    'select',
    {
      class: 'select',
      'aria-label': 'Стиль абзаца',
      onchange: () => {
        exec('formatBlock', `<${blockSelect.value}>`);
      },
    },
    [
      ['p', 'Обычный текст'],
      ['h1', 'Заголовок 1'],
      ['h2', 'Заголовок 2'],
      ['h3', 'Заголовок 3'],
      ['h4', 'Заголовок 4'],
      ['blockquote', 'Цитата'],
      ['pre', 'Код'],
    ].map(([v, t]) => h('option', { value: v }, t)),
  );

  const stateButtons: [HTMLButtonElement, string][] = [];
  const cmdBtn = (name: string, label: string, cmd: string, value?: string, stateful = false): HTMLButtonElement => {
    const b = iconButton(name, label, () => exec(cmd, value));
    b.addEventListener('mousedown', (e) => e.preventDefault()); // keep the selection
    if (stateful) {
      b.setAttribute('aria-pressed', 'false');
      stateButtons.push([b, cmd]);
    }
    return b;
  };

  const color = h('input', { type: 'color', class: 'input', value: '#d92d3a', title: 'Цвет текста', 'aria-label': 'Цвет текста' });
  color.addEventListener('change', () => withCss(() => exec('foreColor', color.value)));
  const marker = h('input', { type: 'color', class: 'input', value: '#fff176', title: 'Цвет выделения', 'aria-label': 'Цвет выделения текста' });
  marker.addEventListener('change', () => withCss(() => exec('hiliteColor', marker.value)));

  /** Run a formatting command producing <span style> rather than <font>. */
  function withCss(fn: () => void): void {
    try {
      document.execCommand('styleWithCSS', false, 'true');
      fn();
    } finally {
      document.execCommand('styleWithCSS', false, 'false');
    }
  }

  const fontSelect = h(
    'select',
    {
      class: 'select',
      'aria-label': 'Шрифт',
      title: 'Шрифт',
      onchange: () => {
        withCss(() => exec('fontName', FONT_CHOICES.find(([f]) => f === fontSelect.value)![1]));
      },
    },
    FONT_CHOICES.map(([f, label]) => h('option', { value: f, style: `font-family:${CSS_FONT[f]}` }, label)),
  );
  const sizeSelect = h(
    'select',
    {
      class: 'select',
      'aria-label': 'Размер шрифта',
      title: 'Размер шрифта, пт',
      style: 'width:76px',
      onchange: () => applySize(+sizeSelect.value),
    },
    SIZE_CHOICES.map((n) => h('option', { value: n }, String(n))),
  );

  /** execCommand only knows sizes 1–7: apply 7, then replace it with the exact point size. */
  function applySize(pt: number): void {
    withCss(() => exec('fontSize', '7'));
    editor.querySelectorAll<HTMLElement>('[style*="xxx-large"], font[size="7"]').forEach((el) => {
      if (el.tagName === 'FONT') {
        const span = document.createElement('span');
        span.style.fontSize = `${pt}pt`;
        span.append(...el.childNodes);
        el.replaceWith(span);
      } else el.style.fontSize = `${pt}pt`;
    });
    changed();
  }

  function applyMeta(): void {
    editor.style.fontFamily = CSS_FONT[meta.font ?? 'sans'];
    editor.style.fontSize = `${meta.size ?? 12}pt`;
  }

  // ------------------------------------------------------------------ footnotes, header, footer
  function insertFootnote(): void {
    const ids = [...editor.querySelectorAll('li[id^="fn"]')].map((li) => parseInt(li.id.slice(2), 10)).filter(Number.isFinite);
    const n = (ids.length ? Math.max(...ids) : 0) + 1;
    exec('insertHTML', `<sup class="footnote-ref"><a href="#fn${n}" id="fnref${n}">${n}</a></sup>&#8203;`);
    let section = editor.querySelector('section.footnotes');
    if (!section) {
      section = h('section', { class: 'footnotes' }, h('hr'), h('ol'));
      const footer = editor.querySelector('footer.doc-footer');
      if (footer) editor.insertBefore(section, footer);
      else editor.append(section);
    }
    const li = h('li', { id: `fn${n}` }, h('p', null, 'Текст сноски'));
    section.querySelector('ol')!.append(li);
    // Put the caret into the new note so the user can type it right away
    const range = document.createRange();
    range.selectNodeContents(li.querySelector('p')!);
    const sel = document.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    changed();
  }

  function toggleRegion(kind: 'header' | 'footer'): void {
    const cls = kind === 'header' ? 'doc-header' : 'doc-footer';
    const existing = editor.querySelector(`${kind}.${cls}`);
    if (existing) {
      existing.remove();
    } else {
      const el = h(kind, { class: cls }, h('p', { style: kind === 'footer' ? 'text-align:center' : 'text-align:right' }, kind === 'footer' ? ['Страница ', fieldSpan('page'), ' из ', fieldSpan('pages')] : 'Верхний колонтитул'));
      if (kind === 'header') editor.prepend(el);
      else editor.append(el);
    }
    focusBody();
    changed();
  }

  /** Make sure the body has at least one paragraph and put the caret into it (not into a header/footer). */
  function focusBody(): void {
    const isRegion = (n: Element): boolean => n.matches('header.doc-header, footer.doc-footer, section.footnotes');
    let target = [...editor.children].find((n) => !isRegion(n));
    if (!target) {
      target = h('p', null, h('br'));
      const after = editor.querySelector('section.footnotes, footer.doc-footer');
      if (after) editor.insertBefore(target, after);
      else editor.append(target);
    }
    const sel = document.getSelection();
    const anchor = sel?.anchorNode ?? null;
    const inRegion = anchor ? !!(anchor.nodeType === 1 ? (anchor as Element) : anchor.parentElement)?.closest('header.doc-header, footer.doc-footer') : true;
    if (inRegion || !editor.contains(anchor)) {
      const range = document.createRange();
      range.selectNodeContents(target);
      range.collapse(false);
      sel?.removeAllRanges();
      sel?.addRange(range);
    }
    editor.focus();
  }

  function fieldSpan(field: 'page' | 'pages'): HTMLElement {
    return h('span', { 'data-field': field, contenteditable: 'false', title: field === 'page' ? 'Номер страницы' : 'Число страниц' }, '#');
  }

  const formatSelect = h(
    'select',
    {
      class: 'select',
      'aria-label': 'Формат сохранения',
      onchange: () => {
        format = formatSelect.value as FormatId;
        prefs.set('doc-save-format', format);
      },
    },
    SAVE_FORMATS.map((f) => h('option', { value: f, selected: f === format }, `${FORMATS[f].label} (.${FORMATS[f].ext})`)),
  );

  const toolbar = h(
    'div',
    { class: 'toolbar card', role: 'toolbar', 'aria-label': 'Форматирование' },
    button('Открыть', () => void openPicker(), { icon: 'folder', small: true }),
    button('Новый', () => void newDoc(), { icon: 'file-plus', small: true, variant: 'ghost' }),
    h('span', { class: 'toolbar__sep' }),
    cmdBtn('undo', 'Отменить (Ctrl+Z)', 'undo'),
    cmdBtn('redo', 'Повторить (Ctrl+Y)', 'redo'),
    h('span', { class: 'toolbar__sep' }),
    blockSelect,
    cmdBtn('bold', 'Жирный (Ctrl+B)', 'bold', undefined, true),
    cmdBtn('italic', 'Курсив (Ctrl+I)', 'italic', undefined, true),
    cmdBtn('underline', 'Подчёркнутый (Ctrl+U)', 'underline', undefined, true),
    cmdBtn('strike', 'Зачёркнутый', 'strikeThrough', undefined, true),
    color,
    marker,
    fontSelect,
    sizeSelect,
    h('span', { class: 'toolbar__sep' }),
    cmdBtn('list-ul', 'Маркированный список', 'insertUnorderedList', undefined, true),
    cmdBtn('list-ol', 'Нумерованный список', 'insertOrderedList', undefined, true),
    cmdBtn('align-left', 'По левому краю', 'justifyLeft', undefined, true),
    cmdBtn('align-center', 'По центру', 'justifyCenter', undefined, true),
    cmdBtn('align-right', 'По правому краю', 'justifyRight', undefined, true),
    cmdBtn('align-justify', 'По ширине', 'justifyFull', undefined, true),
    h('span', { class: 'toolbar__sep' }),
    withMouseGuard(iconButton('link', 'Ссылка', () => void insertLink())),
    withMouseGuard(iconButton('image-plus', 'Изображение', () => void insertImage())),
    withMouseGuard(iconButton('table', 'Таблица', () => void insertTable())),
    cmdBtn('hr', 'Горизонтальная линия', 'insertHorizontalRule'),
    withMouseGuard(iconButton('page-break', 'Разрыв страницы', () => exec('insertHTML', '<hr class="page-break"><p><br></p>'))),
    cmdBtn('clear-format', 'Очистить форматирование', 'removeFormat'),
    h('span', { class: 'toolbar__sep' }),
    withMouseGuard(iconButton('footnote', 'Вставить сноску', insertFootnote)),
    withMouseGuard(iconButton('header', 'Верхний колонтитул (добавить/убрать)', () => toggleRegion('header'))),
    withMouseGuard(iconButton('footer', 'Нижний колонтитул с номером страницы (добавить/убрать)', () => toggleRegion('footer'))),
    withMouseGuard(iconButton('hash', 'Вставить номер страницы', () => exec('insertHTML', '<span data-field="page" contenteditable="false">#</span>&#8203;'))),
    h('span', { class: 'toolbar__spacer' }),
    formatSelect,
    button('Скачать', () => void save(), { icon: 'download', variant: 'primary', small: true }),
  );

  function withMouseGuard(b: HTMLButtonElement): HTMLButtonElement {
    b.addEventListener('mousedown', (e) => e.preventDefault());
    return b;
  }

  root.append(
    h('div', { class: 'page-head' }, h('h1', null, 'Редактор документов'), h('p', null, 'Откройте Word (DOCX/DOC), PDF, ODT, RTF, EPUB, презентацию, Markdown или HTML, отредактируйте и сохраните в нужном формате. Черновик сохраняется автоматически в этом браузере.')),
    toolbar,
    banner,
    h('div', { class: 'row', style: 'margin-top:12px' }, h('label', { class: 'sr-only', for: 'doc-title' }, 'Название'), Object.assign(titleInput, { id: 'doc-title' })),
    h('div', { class: 'doc-shell' }, editor),
    h('div', { class: 'doc-status' }, statusLeft, statusRight),
  );

  try {
    document.execCommand('defaultParagraphSeparator', false, 'p');
    document.execCommand('styleWithCSS', false, 'false');
  } catch {
    /* not supported – harmless */
  }

  // ------------------------------------------------------------------ state
  function refreshState(): void {
    for (const [b, cmd] of stateButtons) {
      let on = false;
      try {
        on = document.queryCommandState(cmd);
      } catch {
        /* ignore */
      }
      b.setAttribute('aria-pressed', String(on));
    }
    // Font and size of the text at the caret
    const node = document.getSelection()?.anchorNode;
    const el = node ? (node.nodeType === 1 ? (node as Element) : node.parentElement) : null;
    if (el && editor.contains(el)) {
      const cs = getComputedStyle(el);
      const fam = familyOf(cs.fontFamily) ?? 'sans';
      fontSelect.value = fam;
      const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
      const nearest = SIZE_CHOICES.reduce((a, b) => (Math.abs(b - pt) < Math.abs(a - pt) ? b : a));
      sizeSelect.value = String(nearest);
    }
    try {
      const block = String(document.queryCommandValue('formatBlock') || 'p').toLowerCase();
      blockSelect.value = ['h1', 'h2', 'h3', 'h4', 'blockquote', 'pre'].includes(block) ? block : 'p';
    } catch {
      /* ignore */
    }
  }

  function updateCounts(): void {
    const text = editor.innerText;
    const words = text.match(/[\p{L}\p{N}]+(?:[-'’][\p{L}\p{N}]+)*/gu)?.length ?? 0;
    statusLeft.textContent = `Слов: ${words} · Символов: ${text.replace(/\n/g, '').length}`;
  }

  function changed(): void {
    dirty = true;
    updateCounts();
    clearTimeout(saveTimer);
    saveTimer = window.setTimeout(persist, 800);
  }

  function persist(): void {
    const html = editor.innerHTML;
    if (html.length > MAX_DRAFT) {
      statusRight.textContent = 'Черновик слишком большой для автосохранения';
      return;
    }
    prefs.set(DRAFT_KEY, { title: titleInput.value, html, at: Date.now(), meta });
    statusRight.textContent = `Черновик сохранён ${new Date().toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' })}`;
  }

  const onSelection = (): void => {
    if (editor.contains(document.getSelection()?.anchorNode ?? null)) refreshState();
  };
  document.addEventListener('selectionchange', onSelection);
  editor.addEventListener('input', changed);
  titleInput.addEventListener('input', changed);

  // Paste & drop: sanitise HTML, embed images
  editor.addEventListener('paste', (e) => {
    const dt = e.clipboardData;
    if (!dt) return;
    const images = [...dt.files].filter((f) => f.type.startsWith('image/'));
    if (images.length) {
      e.preventDefault();
      void Promise.all(images.map(imageHtml)).then((parts) => exec('insertHTML', parts.join('')));
      return;
    }
    const html = dt.getData('text/html');
    if (html) {
      e.preventDefault();
      exec('insertHTML', sanitize(html));
      return;
    }
    const text = dt.getData('text/plain');
    if (text) {
      e.preventDefault();
      exec('insertText', text);
    }
  });
  editor.addEventListener('drop', (e) => {
    const files = [...(e.dataTransfer?.files ?? [])];
    if (!files.length) return;
    e.preventDefault();
    const imgs = files.filter((f) => f.type.startsWith('image/'));
    if (imgs.length) void Promise.all(imgs.map(imageHtml)).then((parts) => exec('insertHTML', parts.join('')));
    else void openFile(files[0]);
  });

  editor.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      void save();
    }
  });
  const beforeUnload = (): void => {
    if (dirty) persist();
  };
  window.addEventListener('beforeunload', beforeUnload);

  // ------------------------------------------------------------------ actions
  async function imageHtml(file: File): Promise<string> {
    const url = await downscaleToDataUrl(file, 1600);
    return `<img src="${url}" alt="${escapeHtml(baseName(file.name))}">`;
  }

  async function insertImage(): Promise<void> {
    const [file] = await pickFiles('image/*', false);
    if (!file) return;
    try {
      exec('insertHTML', await imageHtml(file));
    } catch (e) {
      toast((e as Error).message, 'error');
    }
  }

  async function insertLink(): Promise<void> {
    const sel = document.getSelection();
    const range = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const url = (await promptText('Вставить ссылку', 'Адрес (https://…)', 'https://'))?.trim();
    if (!url) return;
    if (!/^(https?:\/\/|mailto:)/i.test(url)) {
      toast('Ссылка должна начинаться с http://, https:// или mailto:', 'error');
      return;
    }
    editor.focus();
    if (range) {
      sel!.removeAllRanges();
      sel!.addRange(range);
    }
    if (!range || range.collapsed) exec('insertHTML', `<a href="${escapeHtml(url)}">${escapeHtml(url)}</a>`);
    else exec('createLink', url);
  }

  async function insertTable(): Promise<void> {
    const sel = document.getSelection();
    const range = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const spec = await promptText('Вставить таблицу', 'Строки × столбцы', '3x3');
    const m = /^\s*(\d{1,2})\s*[x×хX*]\s*(\d{1,2})\s*$/.exec(spec ?? '');
    if (!m) return;
    const rows = Math.min(50, Math.max(1, +m[1]));
    const cols = Math.min(12, Math.max(1, +m[2]));
    editor.focus();
    if (range) {
      sel!.removeAllRanges();
      sel!.addRange(range);
    }
    let html = '<table><tr>' + Array.from({ length: cols }, (_, i) => `<th>Столбец ${i + 1}</th>`).join('') + '</tr>';
    for (let r = 1; r < rows; r++) html += '<tr>' + '<td><br></td>'.repeat(cols) + '</tr>';
    exec('insertHTML', html + '</table><p><br></p>');
  }

  async function newDoc(): Promise<void> {
    if (editor.textContent?.trim() && !(await confirmDialog('Новый документ', 'Текущий текст будет удалён. Продолжить?', 'Создать'))) return;
    editor.innerHTML = '';
    meta = { font: 'serif', size: 12 };
    applyMeta();
    titleInput.value = 'Документ';
    banner.replaceChildren();
    dirty = false;
    prefs.remove(DRAFT_KEY);
    updateCounts();
    statusRight.textContent = '';
    editor.focus();
  }

  async function openPicker(): Promise<void> {
    const [file] = await pickFiles(OPEN_ACCEPT, false);
    if (file) await openFile(file);
  }

  async function openFile(file: File): Promise<void> {
    if (editor.textContent?.trim() && dirty && !(await confirmDialog('Открыть файл', 'Несохранённые изменения будут заменены содержимым файла. Продолжить?', 'Открыть'))) return;
    let bytes = await readFile(file);
    const det = detectFormat(file.name, bytes);
    if (det.format === 'pdf') {
      try {
        bytes = await unlockPdf(bytes, file.name);
      } catch (e) {
        toast((e as Error).message, 'error');
        return;
      }
    }
    if (!det.format || FORMATS[det.format].kind === 'image') {
      toast(det.error ?? 'Этот формат нельзя открыть в редакторе документов. Изображения откройте в редакторе изображений.', 'error');
      return;
    }
    banner.replaceChildren(h('div', { class: 'row', style: 'margin-top:10px;color:var(--muted)' }, h('span', { class: 'spinner' }), `Открываю ${file.name} (${formatBytes(file.size)})…`));
    try {
      const r = await engine.run<HtmlResult>({ op: 'toHtml', bytes, from: det.format, title: baseName(file.name) }, { transfer: [bytes.buffer] });
      editor.innerHTML = r.html;
      meta = r.meta;
      applyMeta();
      titleInput.value = baseName(file.name);
      banner.replaceChildren(...r.warnings.map((w) => h('div', { style: 'margin-top:10px' }, note('warn', w))));
      if (det.format === 'pdf') banner.append(h('div', { style: 'margin-top:10px' }, note('ok', 'PDF преобразован в редактируемый текст. Для правки страниц «как есть» используйте редактор PDF.')));
      changed();
      editor.focus();
    } catch (e) {
      banner.replaceChildren(h('div', { style: 'margin-top:10px' }, note('error', (e as Error).message)));
    }
  }

  async function save(): Promise<void> {
    const title = titleInput.value.trim() || 'Документ';
    const html = editor.innerHTML;
    if (!editor.textContent?.trim() && !editor.querySelector('img, table, hr')) {
      toast('Документ пуст.', 'error');
      return;
    }
    try {
      toast(`Сохраняю в ${FORMATS[format].label}…`, 'info', 2000);
      const r = await engine.run<ConvertResult>({ op: 'fromHtml', html, to: format, opts: {}, title, meta });
      download(r.data, `${title.replace(/[\\/:*?"<>|]+/g, '_')}.${r.ext}`, r.mime);
      for (const w of r.warnings) toast(w);
      dirty = false;
    } catch (e) {
      toast((e as Error).message, 'error');
    }
  }

  // ------------------------------------------------------------------ init
  const handed = takeHandoff('/doc-editor');
  if (handed) void openFile(handed);
  else if (draft?.html) {
    editor.innerHTML = docToHtmlFragment(htmlToDoc(draft.html), { sources: true });
    banner.replaceChildren(
      h(
        'div',
        { class: 'row', style: 'margin-top:10px' },
        note('ok', `Восстановлен черновик от ${new Date(draft.at).toLocaleString('ru-RU')}`),
        button('Начать заново', () => void newDoc(), { small: true, variant: 'ghost' }),
      ),
    );
  }
  applyMeta();
  updateCounts();
  if (!handed) setTimeout(() => editor.focus(), 0);

  return () => {
    clearTimeout(saveTimer);
    if (dirty) persist();
    document.removeEventListener('selectionchange', onSelection);
    window.removeEventListener('beforeunload', beforeUnload);
  };
}

/** Read an image file as a data URL, downscaling very large photos. */
async function downscaleToDataUrl(file: File, maxSide: number): Promise<string> {
  const bmp = await createImageBitmap(file).catch(() => null);
  if (!bmp) throw new Error('Не удалось прочитать изображение.');
  const k = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(bmp.width * k);
  canvas.height = Math.round(bmp.height * k);
  canvas.getContext('2d')!.drawImage(bmp, 0, 0, canvas.width, canvas.height);
  bmp.close();
  const png = file.type === 'image/png' || file.type === 'image/gif';
  return canvas.toDataURL(png ? 'image/png' : 'image/jpeg', 0.9);
}
