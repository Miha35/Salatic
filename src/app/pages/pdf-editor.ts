import { detectFormat, FORMATS } from '../../core/formats';
import type { Annotation, PageSpec } from '../../core/pdf-tools';
import type { FontFamily } from '../../core/model';
import { registerEditorFonts } from '../editor-fonts';
import { buildSegments, familyFor, type TextSeg } from '../../core/text-lines';
import { bytesToBase64 } from '../../core/util';
import { engine } from '../engine/client';
import { openPdf, pdfErrorMessage, renderPage, type PdfDoc } from '../pdfjs';
import { toPdfImage, toPdfImages, zipFiles } from '../raster';
import { unlockPdf } from '../pdf-unlock';
import { baseName, button, confirmDialog, download, dropzone, h, iconButton, note, onPasteFiles, pickFiles, readFile, takeHandoff, toast } from '../ui';

// ---------------------------------------------------------------------------
// Coordinate system (same maths as pdf.js PageViewport)

type Rot = 0 | 90 | 180 | 270;

class Viewport {
  readonly transform: number[];
  readonly width: number;
  readonly height: number;
  private inv: number[];
  constructor(
    readonly viewBox: number[],
    readonly scale: number,
    readonly rotation: Rot,
  ) {
    const [x0, y0, x1, y1] = viewBox;
    const cx = (x0 + x1) / 2;
    const cy = (y0 + y1) / 2;
    const [a, b, c, d] = { 0: [1, 0, 0, -1], 90: [0, 1, 1, 0], 180: [-1, 0, 0, 1], 270: [0, -1, -1, 0] }[rotation];
    let ox: number;
    let oy: number;
    if (a === 0) {
      ox = Math.abs(cy - y0) * scale;
      oy = Math.abs(cx - x0) * scale;
      this.width = (y1 - y0) * scale;
      this.height = (x1 - x0) * scale;
    } else {
      ox = Math.abs(cx - x0) * scale;
      oy = Math.abs(cy - y0) * scale;
      this.width = (x1 - x0) * scale;
      this.height = (y1 - y0) * scale;
    }
    this.transform = [a * scale, b * scale, c * scale, d * scale, ox - a * scale * cx - c * scale * cy, oy - b * scale * cx - d * scale * cy];
    const [A, B, C, D, E, F] = this.transform;
    const det = A * D - B * C;
    this.inv = [D / det, -B / det, -C / det, A / det, (C * F - D * E) / det, (B * E - A * F) / det];
  }
  toView(x: number, y: number): [number, number] {
    const [a, b, c, d, e, f] = this.transform;
    return [a * x + c * y + e, b * x + d * y + f];
  }
  toPdf(x: number, y: number): [number, number] {
    const [a, b, c, d, e, f] = this.inv;
    return [a * x + c * y + e, b * x + d * y + f];
  }
}

const normRot = (r: number): Rot => ((((Math.round(r / 90) * 90) % 360) + 360) % 360) as Rot;

// ---------------------------------------------------------------------------
// Model

interface Source {
  name: string;
  bytes: Uint8Array;
  pdf: PdfDoc;
}

interface PageItem {
  key: number;
  src: number; // -1 = blank page
  index: number;
  view: number[]; // page box in user space
  baseRotate: number; // /Rotate from the file
  rotate: number; // user rotation (clockwise)
  annotations: Annotation[];
  selected: boolean;
  thumb?: HTMLCanvasElement;
  dirty: boolean;
  /** Ids of original text lines replaced with the "edit text" tool. */
  replaced?: string[];
}

const CSS_FAMILY: Record<FontFamily, string> = {
  sans: '"Liberation Sans", Arial, sans-serif',
  serif: '"Liberation Serif", "Times New Roman", serif',
  mono: '"Liberation Mono", "Courier New", monospace',
  calibri: 'Carlito, Calibri, sans-serif',
};
const cssFont = (a: { font?: FontFamily }): string => CSS_FAMILY[a.font ?? 'sans'];



type Tool = 'select' | 'edittext' | 'text' | 'whiteout' | 'highlight' | 'pen' | 'image';

const A4: [number, number] = [595.28, 841.89];

function hexToRgb(hex: string): [number, number, number] {
  const n = parseInt(hex.slice(1), 16);
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255];
}
const rgbCss = (c: [number, number, number]): string => `rgb(${c.map((v) => Math.round(v * 255)).join(',')})`;

const cloneAnn = (a: Annotation): Annotation => (a.type === 'path' ? { ...a, points: a.points.map((p) => [...p] as [number, number]) } : { ...a });

let fontsLoaded = false;
function loadDisplayFonts(): void {
  if (fontsLoaded || typeof FontFace === 'undefined') return;
  fontsLoaded = true;
  // Same faces that are embedded into the saved PDF, so the preview matches the output
  for (const [url, weight] of [
    ['/pdfjs/standard_fonts/LiberationSans-Regular.ttf', '400'],
    ['/pdfjs/standard_fonts/LiberationSans-Bold.ttf', '700'],
  ] as const) {
    const face = new FontFace('Liberation Sans', `url(${url})`, { weight });
    face.load().then((f) => document.fonts.add(f)).catch(() => undefined);
  }
}

/** Draw annotations onto a 2D canvas (thumbnails). */
function paintAnnotations(ctx: CanvasRenderingContext2D, vp: Viewport, anns: Annotation[], images: Map<Uint8Array, HTMLImageElement>): void {
  for (const a of anns) {
    ctx.save();
    if (a.type === 'path') {
      ctx.globalAlpha = a.opacity;
      ctx.strokeStyle = rgbCss(a.color);
      ctx.lineWidth = a.width * vp.scale;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.beginPath();
      a.points.forEach(([x, y], i) => {
        const [vx, vy] = vp.toView(x, y);
        if (i) ctx.lineTo(vx, vy);
        else ctx.moveTo(vx, vy);
      });
      ctx.stroke();
    } else {
      const [px, py] = vp.toView(a.x, a.y);
      ctx.translate(px, py);
      ctx.rotate(((vp.rotation - a.rotate) * Math.PI) / 180);
      const s = vp.scale;
      if (a.type === 'text') {
        ctx.fillStyle = rgbCss(a.color);
        ctx.font = `${a.italic ? 'italic ' : ''}${a.bold ? 700 : 400} ${a.size * s}px ${cssFont(a)}`;
        a.text.split('\n').forEach((line, i) => ctx.fillText(line, 0, i * a.size * 1.25 * s));
      } else if (a.type === 'rect') {
        ctx.globalAlpha = a.opacity;
        ctx.fillStyle = rgbCss(a.color);
        ctx.fillRect(0, -a.h * s, a.w * s, a.h * s);
      } else if (a.type === 'image') {
        const img = images.get(a.data);
        if (img?.complete) ctx.drawImage(img, 0, -a.h * s, a.w * s, a.h * s);
      }
    }
    ctx.restore();
  }
}

// ---------------------------------------------------------------------------

export function mount(root: HTMLElement): () => void {
  const sources: Source[] = [];
  let pages: PageItem[] = [];
  let nextKey = 1;
  let fileName = 'document';
  let editing: PageItem | null = null;
  const imageCache = new Map<Uint8Array, HTMLImageElement>();
  let busy = false;

  const empty = h(
    'div',
    { class: 'empty' },
    dropzone({ accept: '.pdf,image/*', title: 'Откройте PDF', hint: 'Перетащите один или несколько PDF (и изображения) — они будут объединены. Можно вставить из буфера (Ctrl+V).', onFiles: (f) => void addFiles(f) }),
    h(
      'div',
      { class: 'features' },
      info('Страницы', 'Переставляйте перетаскиванием, поворачивайте, удаляйте, дублируйте и извлекайте.'),
      info('Объединение', 'Добавляйте другие PDF и изображения в любом порядке, вставляйте пустые страницы.'),
      info('Пометки', 'Текст, закрашивание (замазка), выделение маркером, рисование, подписи и картинки.'),
    ),
  );
  function info(title: string, text: string): HTMLElement {
    return h('div', { class: 'feature card' }, h('h3', null, title), h('p', null, text));
  }
  const toolbar = h('div', { class: 'toolbar card', role: 'toolbar', 'aria-label': 'Действия со страницами' });
  const grid = h('div', { class: 'pages', role: 'list', 'aria-label': 'Страницы документа' });
  const gridView = h('div', { hidden: true }, toolbar, grid);
  const pageView = h('div', { class: 'page-view', hidden: true });

  root.append(
    h('div', { class: 'page-head' }, h('h1', null, 'Редактор PDF'), h('p', null, 'Объединение, порядок и поворот страниц, удаление и извлечение, текст, закрашивание, выделение, подписи. Обработка выполняется в браузере.')),
    empty,
    gridView,
    pageView,
  );
  onPasteFiles(root, (f) => void addFiles(f));

  // ------------------------------------------------------------------ loading
  async function addFiles(files: File[]): Promise<void> {
    if (busy) return;
    busy = true;
    const loading = h('div', { class: 'row', style: 'margin:12px 0;color:var(--muted)' }, h('span', { class: 'spinner' }), 'Открываю файлы…');
    root.insertBefore(loading, gridView);
    try {
      const images: { file: File; format: ReturnType<typeof detectFormat>['format'] }[] = [];
      for (const file of files) {
        const bytes = await readFile(file);
        const det = detectFormat(file.name, bytes);
        if (det.format === 'pdf') {
          if (!sources.length) fileName = baseName(file.name);
          const plain = await unlockPdf(bytes, file.name);
          if (plain !== bytes) toast(`${file.name}: защита снята — сохранённый файл будет без пароля.`);
          await addPdf(file.name, plain);
        } else if (det.format && FORMATS[det.format].kind === 'image') images.push({ file, format: det.format });
        else toast(`${file.name}: ${det.error ?? 'поддерживаются только PDF и изображения'}`, 'error');
      }
      if (images.length) {
        const inputs = [];
        for (const im of images) inputs.push(...(await toPdfImages(im.file, im.format!)));
        const pdfBytes = await engine.run<Uint8Array>({ op: 'imagesToPdf', images: inputs, opts: { pageSize: 'fit' } }, { transfer: inputs.map((i) => i.data.buffer) });
        if (!sources.length) fileName = baseName(images[0].file.name);
        await addPdf('images.pdf', pdfBytes);
      }
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      loading.remove();
      busy = false;
      if (pages.length) {
        empty.hidden = true;
        if (!editing) gridView.hidden = false;
      }
      renderToolbar();
      renderGrid();
    }
  }

  async function addPdf(name: string, bytes: Uint8Array): Promise<void> {
    let pdf: PdfDoc;
    try {
      pdf = await openPdf(bytes);
    } catch (e) {
      const msg = (e as Error)?.name === 'PasswordException' ? `${name}: PDF защищён паролем.` : `${name}: ${pdfErrorMessage(e)}`;
      throw new Error(msg);
    }
    const src = sources.push({ name, bytes, pdf }) - 1;
    for (let i = 0; i < pdf.numPages; i++) {
      const page = await pdf.getPage(i + 1);
      pages.push({ key: nextKey++, src, index: i, view: [...page.view], baseRotate: page.rotate, rotate: 0, annotations: [], selected: false, dirty: true });
    }
  }

  const totalRotation = (p: PageItem): Rot => normRot(p.baseRotate + p.rotate);

  // ------------------------------------------------------------------ thumbnails
  const observer = new IntersectionObserver(
    (entries) => {
      for (const e of entries) {
        if (!e.isIntersecting) continue;
        const key = Number((e.target as HTMLElement).dataset.key);
        const p = pages.find((x) => x.key === key);
        if (p?.dirty) void renderThumb(p, e.target as HTMLElement);
      }
    },
    { rootMargin: '300px' },
  );

  async function renderThumb(p: PageItem, card: HTMLElement): Promise<void> {
    p.dirty = false;
    const rot = totalRotation(p);
    const w = p.view[2] - p.view[0];
    const hh = p.view[3] - p.view[1];
    const longSide = rot % 180 ? w : hh;
    const scale = (Math.min(320, 220 * (window.devicePixelRatio || 1)) / Math.max(longSide, rot % 180 ? hh : w)) * 1;
    const canvas = document.createElement('canvas');
    try {
      if (p.src >= 0) {
        const page = await sources[p.src].pdf.getPage(p.index + 1);
        await renderPage(page, canvas, scale, rot);
      } else {
        const vp = new Viewport(p.view, scale, rot);
        canvas.width = Math.ceil(vp.width);
        canvas.height = Math.ceil(vp.height);
        const ctx = canvas.getContext('2d')!;
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }
      if (p.annotations.length) {
        await ensureImages(p.annotations);
        paintAnnotations(canvas.getContext('2d')!, new Viewport(p.view, scale, rot), p.annotations, imageCache);
      }
      p.thumb = canvas;
      canvas.style.width = `${canvas.width / (window.devicePixelRatio || 1)}px`;
      const holder = card.querySelector('.page-card__thumb');
      if (holder) holder.replaceChildren(canvas);
    } catch {
      p.dirty = true;
    }
  }

  async function ensureImages(anns: Annotation[]): Promise<void> {
    await Promise.all(
      anns.map((a) => {
        if (a.type !== 'image' || imageCache.has(a.data)) return;
        const img = new Image();
        img.src = `data:${a.mime};base64,${bytesToBase64(a.data)}`;
        imageCache.set(a.data, img);
        return img.decode().catch(() => undefined);
      }),
    );
  }

  // ------------------------------------------------------------------ grid
  let dragKey: number | null = null;

  function renderGrid(): void {
    observer.disconnect();
    grid.replaceChildren(
      ...pages.map((p, i) => {
        const check = h('input', {
          type: 'checkbox',
          class: 'page-card__check',
          checked: p.selected,
          'aria-label': `Выбрать страницу ${i + 1}`,
          onchange: (e: Event) => {
            p.selected = (e.target as HTMLInputElement).checked;
            card.classList.toggle('is-selected', p.selected);
            renderToolbar();
          },
        });
        const card = h(
          'div',
          { class: `page-card${p.selected ? ' is-selected' : ''}`, role: 'listitem', draggable: 'true', 'data-key': p.key, 'aria-label': `Страница ${i + 1}` },
          check,
          p.annotations.length ? h('span', { class: 'badge page-card__badge', title: 'Есть пометки' }, `✎ ${p.annotations.length}`) : null,
          h('div', { class: 'page-card__thumb', ondblclick: () => openEditor(p), title: 'Двойной щелчок — редактировать' }, p.thumb && !p.dirty ? p.thumb : h('span', { class: 'spinner', style: 'color:var(--muted)' })),
          h(
            'div',
            { class: 'page-card__move' },
            iconButton('chevron-left', 'Переместить влево', () => move(i, i - 1), { disabled: i === 0 }),
            iconButton('chevron-right', 'Переместить вправо', () => move(i, i + 1), { disabled: i === pages.length - 1 }),
          ),
          h(
            'div',
            { class: 'page-card__bar' },
            h('span', { class: 'page-card__num' }, `${i + 1}`),
            iconButton('rotate-ccw', 'Повернуть влево', () => rotatePages([p], -90)),
            iconButton('rotate-cw', 'Повернуть вправо', () => rotatePages([p], 90)),
            iconButton('pen', 'Редактировать страницу', () => openEditor(p)),
            iconButton('copy', 'Дублировать', () => duplicate(p)),
            iconButton('trash', 'Удалить страницу', () => removePages([p])),
          ),
        );
        card.addEventListener('dragstart', (e) => {
          dragKey = p.key;
          card.classList.add('is-dragging');
          e.dataTransfer?.setData('text/plain', String(p.key));
          if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
        });
        card.addEventListener('dragend', () => {
          dragKey = null;
          card.classList.remove('is-dragging');
          grid.querySelectorAll('.is-drop-before, .is-drop-after').forEach((el) => el.classList.remove('is-drop-before', 'is-drop-after'));
        });
        card.addEventListener('dragover', (e) => {
          if (dragKey === null) return;
          e.preventDefault();
          const r = card.getBoundingClientRect();
          const after = e.clientX > r.left + r.width / 2;
          card.classList.toggle('is-drop-after', after);
          card.classList.toggle('is-drop-before', !after);
        });
        card.addEventListener('dragleave', () => card.classList.remove('is-drop-before', 'is-drop-after'));
        card.addEventListener('drop', (e) => {
          e.preventDefault();
          if (dragKey === null || dragKey === p.key) return;
          const from = pages.findIndex((x) => x.key === dragKey);
          const r = card.getBoundingClientRect();
          let to = pages.indexOf(p) + (e.clientX > r.left + r.width / 2 ? 1 : 0);
          if (from < to) to--;
          move(from, to);
        });
        observer.observe(card);
        return card;
      }),
    );
  }

  function move(from: number, to: number): void {
    if (to < 0 || to >= pages.length || from === to) return;
    const [p] = pages.splice(from, 1);
    pages.splice(to, 0, p);
    renderGrid();
  }

  function rotatePages(list: PageItem[], deg: number): void {
    for (const p of list) {
      p.rotate = (p.rotate + deg + 360) % 360;
      p.dirty = true;
    }
    renderGrid();
  }

  function duplicate(p: PageItem): void {
    const copy: PageItem = { ...p, key: nextKey++, annotations: p.annotations.map(cloneAnn), selected: false, dirty: true, thumb: undefined };
    pages.splice(pages.indexOf(p) + 1, 0, copy);
    renderGrid();
  }

  async function removePages(list: PageItem[]): Promise<void> {
    if (list.length === pages.length) {
      toast('Нельзя удалить все страницы документа.', 'error');
      return;
    }
    if (list.length > 1 && !(await confirmDialog('Удалить страницы', `Удалить выбранные страницы (${list.length})?`, 'Удалить'))) return;
    pages = pages.filter((p) => !list.includes(p));
    renderToolbar();
    renderGrid();
  }

  function insertBlank(): void {
    const sel = pages.filter((p) => p.selected);
    const at = sel.length ? pages.indexOf(sel[sel.length - 1]) + 1 : pages.length;
    const ref = sel[sel.length - 1] ?? pages[pages.length - 1];
    const size: [number, number] = ref ? [ref.view[2] - ref.view[0], ref.view[3] - ref.view[1]] : A4;
    pages.splice(at, 0, { key: nextKey++, src: -1, index: 0, view: [0, 0, size[0], size[1]], baseRotate: 0, rotate: ref ? totalRotation(ref) : 0, annotations: [], selected: false, dirty: true });
    renderGrid();
  }

  // ------------------------------------------------------------------ saving
  function specs(list: PageItem[]): PageSpec[] {
    return list.map((p) => ({
      src: p.src,
      index: p.index,
      rotate: p.rotate,
      size: p.src < 0 ? ([p.view[2] - p.view[0], p.view[3] - p.view[1]] as [number, number]) : undefined,
      annotations: p.annotations,
    }));
  }

  async function compose(list: PageItem[]): Promise<Uint8Array> {
    const copies = sources.map((s) => s.bytes.slice());
    return engine.run<Uint8Array>({ op: 'composePdf', sources: copies, pages: specs(list) }, { transfer: copies.map((c) => c.buffer) });
  }

  async function save(list: PageItem[], suffix: string): Promise<void> {
    if (busy) return;
    busy = true;
    renderToolbar();
    try {
      const data = await compose(list);
      download(data, `${fileName}${suffix}.pdf`, 'application/pdf');
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      busy = false;
      renderToolbar();
    }
  }

  async function splitAll(): Promise<void> {
    if (busy) return;
    busy = true;
    renderToolbar();
    try {
      const files = [];
      for (let i = 0; i < pages.length; i++) files.push({ name: `${fileName}-${String(i + 1).padStart(String(pages.length).length, '0')}.pdf`, data: await compose([pages[i]]) });
      download(zipFiles(files), `${fileName}-pages.zip`, 'application/zip');
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      busy = false;
      renderToolbar();
    }
  }

  function renderToolbar(): void {
    const sel = pages.filter((p) => p.selected);
    toolbar.replaceChildren(
      button('Добавить файлы', async () => void addFiles(await pickFiles('.pdf,image/*')), { icon: 'file-plus', small: true }),
      button('Пустая страница', insertBlank, { icon: 'plus', small: true, variant: 'ghost' }),
      h('span', { class: 'toolbar__sep' }),
      button(sel.length === pages.length && pages.length ? 'Снять выделение' : 'Выделить все', () => {
        const all = !(sel.length === pages.length);
        pages.forEach((p) => (p.selected = all));
        renderToolbar();
        renderGrid();
      }, { small: true, variant: 'ghost' }),
      h('span', { style: 'color:var(--muted);font-size:13px' }, sel.length ? `Выбрано: ${sel.length}` : ''),
      iconButton('rotate-ccw', 'Повернуть выбранные влево', () => rotatePages(sel, -90), { disabled: !sel.length }),
      iconButton('rotate-cw', 'Повернуть выбранные вправо', () => rotatePages(sel, 90), { disabled: !sel.length }),
      iconButton('trash', 'Удалить выбранные', () => void removePages(sel), { disabled: !sel.length }),
      button('Извлечь выбранные', () => void save(sel, '-extract'), { icon: 'scissors', small: true, variant: 'ghost', attrs: { disabled: !sel.length || busy } }),
      button('Разделить на страницы', () => void splitAll(), { small: true, variant: 'ghost', attrs: { disabled: pages.length < 2 || busy } }),
      h('span', { class: 'toolbar__spacer' }),
      h('span', { style: 'color:var(--muted);font-size:13px' }, `Страниц: ${pages.length}`),
      button(busy ? 'Сохраняю…' : 'Сохранить PDF', () => void save(pages, '-edited'), { icon: 'download', variant: 'primary', small: true, attrs: { disabled: busy || !pages.length } }),
    );
  }

  // ------------------------------------------------------------------ page editor
  let tool: Tool = 'select';
  let color = '#d92d3a';
  let fontSize = 14;
  let strokeWidth = 2;
  let zoom = 1.4;
  let undoStack: { annotations: Annotation[]; replaced: string[] }[] = [];
  let selectedIdx = -1;
  let viewport: Viewport | null = null;
  let svg: SVGSVGElement | null = null;
  let stage: HTMLElement | null = null;
  let textEditor: HTMLTextAreaElement | null = null;
  let pageCanvas: HTMLCanvasElement | null = null;
  let segments: TextSeg[] = [];

  function openEditor(p: PageItem): void {
    loadDisplayFonts();
    registerEditorFonts();
    editing = p;
    undoStack = [];
    selectedIdx = -1;
    gridView.hidden = true;
    pageView.hidden = false;
    void renderEditor();
  }

  function closeEditor(): void {
    commitText();
    if (editing) editing.dirty = true;
    editing = null;
    pageView.hidden = true;
    gridView.hidden = false;
    pageView.replaceChildren();
    renderGrid();
  }

  function snapshot(): void {
    if (!editing) return;
    undoStack.push({ annotations: editing.annotations.map(cloneAnn), replaced: [...(editing.replaced ?? [])] });
    if (undoStack.length > 100) undoStack.shift();
  }

  async function renderEditor(): Promise<void> {
    const p = editing;
    if (!p) return;
    const idx = pages.indexOf(p);
    const rot = totalRotation(p);
    const toolBtn = (t: Tool, icn: string, label: string): HTMLButtonElement =>
      iconButton(icn, label, () => {
        commitText();
        tool = t;
        selectedIdx = -1;
        void renderEditor();
      }, { 'aria-pressed': String(tool === t) });
    const colorIn = h('input', { type: 'color', class: 'input', value: color, title: 'Цвет', 'aria-label': 'Цвет' });
    colorIn.addEventListener('input', () => {
      color = colorIn.value;
      const a = editing?.annotations[selectedIdx];
      if (a && a.type !== 'image' && !(a.type === 'rect' && a.opacity === 1 && tool === 'whiteout')) {
        snapshot();
        a.color = hexToRgb(color);
        drawOverlay();
      }
    });
    const sizeIn = h('input', { type: 'number', class: 'input', min: 4, max: 200, value: tool === 'pen' ? strokeWidth : fontSize, style: 'width:72px', title: tool === 'pen' ? 'Толщина линии' : 'Размер шрифта', 'aria-label': tool === 'pen' ? 'Толщина линии' : 'Размер шрифта' });
    sizeIn.addEventListener('change', () => {
      const v = Math.min(200, Math.max(1, +sizeIn.value || 1));
      if (tool === 'pen') strokeWidth = v;
      else fontSize = v;
      const a = editing?.annotations[selectedIdx];
      if (a?.type === 'text') {
        snapshot();
        a.size = v;
        drawOverlay();
      }
    });
    const bar = h(
      'div',
      { class: 'toolbar card', role: 'toolbar', 'aria-label': 'Инструменты страницы' },
      button('К страницам', closeEditor, { icon: 'chevron-left', small: true }),
      iconButton('chevron-left', 'Предыдущая страница', () => switchPage(idx - 1), { disabled: idx <= 0 }),
      h('span', { style: 'font-size:13px;color:var(--muted)' }, `${idx + 1} / ${pages.length}`),
      iconButton('chevron-right', 'Следующая страница', () => switchPage(idx + 1), { disabled: idx >= pages.length - 1 }),
      h('span', { class: 'toolbar__sep' }),
      toolBtn('select', 'pointer', 'Выбор и перемещение'),
      toolBtn('edittext', 'edit-text', 'Править текст страницы'),
      toolBtn('text', 'type', 'Добавить текст'),
      toolBtn('whiteout', 'eraser', 'Закрасить (замазка)'),
      toolBtn('highlight', 'highlighter', 'Выделить маркером'),
      toolBtn('pen', 'pen', 'Рисовать'),
      iconButton('image-plus', 'Вставить изображение или подпись', () => void addImage()),
      h('span', { class: 'toolbar__sep' }),
      colorIn,
      sizeIn,
      iconButton('trash', 'Удалить выбранную пометку (Delete)', deleteSelected, { disabled: selectedIdx < 0 }),
      iconButton('undo', 'Отменить (Ctrl+Z)', undo, { disabled: !undoStack.length }),
      h('span', { class: 'toolbar__sep' }),
      iconButton('zoom-out', 'Уменьшить', () => setZoom(zoom / 1.2)),
      iconButton('zoom-in', 'Увеличить', () => setZoom(zoom * 1.2)),
      h('span', { class: 'toolbar__spacer' }),
      button('Сохранить PDF', () => void save(pages, '-edited'), { icon: 'download', variant: 'primary', small: true }),
    );
    const hint = {
      select: 'Нажмите на пометку, чтобы выбрать; перетащите для перемещения; двойной щелчок по тексту — правка; угол — изменение размера.',
      edittext:
        p.src < 0
          ? 'На пустой странице нет текста для правки — используйте «Добавить текст».'
          : rot !== 0
            ? 'Правка текста доступна на неповёрнутых страницах: верните поворот страницы к 0°.'
            : 'Нажмите на строку текста страницы, исправьте её и нажмите Ctrl+Enter или щёлкните вне поля. Старая строка закрывается цветом фона, новая подбирает шрифт, размер и цвет.',
      text: 'Нажмите на страницу, введите текст. Ctrl+Enter или щелчок вне поля — готово.',
      whiteout: 'Протяните прямоугольник, чтобы закрыть часть страницы белым (или выбранным) цветом.',
      highlight: 'Протяните прямоугольник поверх текста, чтобы выделить его маркером.',
      pen: 'Рисуйте мышью, пером или пальцем.',
      image: '',
    }[tool];

    stage = h('div', { class: 'page-stage' });
    const canvas = document.createElement('canvas');
    pageCanvas = canvas;
    segments = [];
    viewport = new Viewport(p.view, zoom, rot);
    stage.style.width = `${viewport.width}px`;
    stage.style.height = `${viewport.height}px`;
    svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.classList.add('overlay');
    svg.setAttribute('viewBox', `0 0 ${viewport.width} ${viewport.height}`);
    stage.append(canvas, svg);
    pageView.replaceChildren(bar, hint ? note('ok', hint) : h('span'), h('div', { class: 'stage-wrap' }, stage));
    attachStageEvents();
    try {
      const dpr = window.devicePixelRatio || 1;
      if (p.src >= 0) {
        const page = await sources[p.src].pdf.getPage(p.index + 1);
        await renderPage(page, canvas, zoom * dpr, rot);
      } else {
        canvas.width = Math.ceil(viewport.width * dpr);
        canvas.height = Math.ceil(viewport.height * dpr);
        const ctx = canvas.getContext('2d')!;
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }
      canvas.style.width = `${viewport.width}px`;
      canvas.style.height = `${viewport.height}px`;
      if (tool === 'edittext' && p.src >= 0 && rot === 0) segments = await loadSegments(p);
    } catch (e) {
      toast(pdfErrorMessage(e), 'error');
    }
    await ensureImages(p.annotations);
    drawOverlay();
  }

  /** Lines of the page's own text (pdf.js text layer), with fonts identified from the rendered page. */
  async function loadSegments(p: PageItem): Promise<TextSeg[]> {
    const page = await sources[p.src].pdf.getPage(p.index + 1);
    const tc = await page.getTextContent();
    const objs = page.commonObjs as { has(id: string): boolean; get(id: string): { name?: string; bold?: boolean; italic?: boolean; isSerifFont?: boolean; isMonospace?: boolean } };
    const styles = tc.styles as Record<string, { fontFamily?: string }>;
    const info = (name: string): { family: FontFamily; bold: boolean; italic: boolean } => {
      let f: { name?: string; bold?: boolean; italic?: boolean; isSerifFont?: boolean; isMonospace?: boolean } = {};
      try {
        if (objs.has(name)) f = objs.get(name) ?? {};
      } catch {
        f = {};
      }
      const fname = f.name ?? '';
      const generic = styles[name]?.fontFamily ?? '';
      return {
        family: familyFor(fname, f.isSerifFont || generic === 'serif', f.isMonospace || generic === 'monospace'),
        bold: !!f.bold || /bold|black|heavy|semibold|demi/i.test(fname),
        italic: !!f.italic || /italic|oblique/i.test(fname),
      };
    };
    const items = tc.items.filter((i): i is (typeof tc.items)[number] & { str: string; transform: number[]; width: number; fontName: string; hasEOL: boolean } => 'str' in i);
    return buildSegments(items, info).filter((s) => !p.replaced?.includes(s.id));
  }

  /** Background and text colours of a page area, sampled from the rendered canvas. */
  function sampleColors(box: { x: number; y: number; w: number; h: number }): { bg: [number, number, number]; fg: [number, number, number] } {
    const white: [number, number, number] = [1, 1, 1];
    const canvas = pageCanvas;
    if (!canvas || !viewport) return { bg: white, fg: [0, 0, 0] };
    const k = canvas.width / viewport.width;
    const x0 = Math.max(0, Math.floor((box.x - 3) * k));
    const y0 = Math.max(0, Math.floor((box.y - 3) * k));
    const x1 = Math.min(canvas.width, Math.ceil((box.x + box.w + 3) * k));
    const y1 = Math.min(canvas.height, Math.ceil((box.y + box.h + 3) * k));
    if (x1 <= x0 || y1 <= y0) return { bg: white, fg: [0, 0, 0] };
    const data = canvas.getContext('2d')!.getImageData(x0, y0, x1 - x0, y1 - y0).data;
    const w = x1 - x0;
    const hgt = y1 - y0;
    // Background: median of the border ring
    const ring: number[][] = [];
    for (let x = 0; x < w; x++) for (const y of [0, hgt - 1]) ring.push([...data.subarray((y * w + x) * 4, (y * w + x) * 4 + 3)]);
    for (let y = 0; y < hgt; y++) for (const x of [0, w - 1]) ring.push([...data.subarray((y * w + x) * 4, (y * w + x) * 4 + 3)]);
    const med = (ch: number): number => ring.map((p) => p[ch]).sort((a, b) => a - b)[ring.length >> 1] / 255;
    const bg: [number, number, number] = [med(0), med(1), med(2)];
    // Text: the pixels farthest from the background
    const px: { d: number; c: number[] }[] = [];
    for (let i = 0; i < data.length; i += 4) {
      const c = [data[i] / 255, data[i + 1] / 255, data[i + 2] / 255];
      px.push({ d: Math.abs(c[0] - bg[0]) + Math.abs(c[1] - bg[1]) + Math.abs(c[2] - bg[2]), c });
    }
    px.sort((a, b) => b.d - a.d);
    const top = px.slice(0, Math.max(1, Math.floor(px.length * 0.05)));
    const fg = [0, 1, 2].map((ch) => top.reduce((s, p) => s + p.c[ch], 0) / top.length) as [number, number, number];
    return { bg, fg: top[0].d < 0.15 ? [0, 0, 0] : fg };
  }

  function segBox(seg: TextSeg): { x: number; y: number; w: number; h: number } {
    const s = viewport!.scale;
    const [vx, vy] = viewport!.toView(seg.x, seg.y);
    return { x: vx - 1, y: vy - seg.size * 0.92 * s, w: seg.w * s + 2, h: seg.size * 1.18 * s };
  }

  function editSegment(seg: TextSeg): void {
    if (!viewport || !editing) return;
    const box = segBox(seg);
    const { bg, fg } = sampleColors(box);
    const s = viewport.scale;
    const [vx, vy] = viewport.toView(seg.x, seg.y);
    const pseudo: Extract<Annotation, { type: 'text' }> = { type: 'text', x: seg.x, y: seg.y, size: Math.round(seg.size * 10) / 10, color: fg, text: seg.text, rotate: viewport.rotation, bold: seg.bold, italic: seg.italic, font: seg.font };
    startText([vx - 1, vy - seg.size * s * 0.95], pseudo, { seg, bg });
  }

  function switchPage(i: number): void {
    if (i < 0 || i >= pages.length || !editing) return;
    commitText();
    editing.dirty = true;
    editing = pages[i];
    undoStack = [];
    selectedIdx = -1;
    void renderEditor();
  }

  function setZoom(z: number): void {
    commitText();
    zoom = Math.min(4, Math.max(0.4, z));
    void renderEditor();
  }

  function undo(): void {
    const prev = undoStack.pop();
    if (!prev || !editing) return;
    editing.annotations = prev.annotations;
    editing.replaced = prev.replaced;
    selectedIdx = -1;
    void renderEditor();
  }

  function deleteSelected(): void {
    if (!editing || selectedIdx < 0) return;
    snapshot();
    editing.annotations.splice(selectedIdx, 1);
    selectedIdx = -1;
    void renderEditor();
  }

  const SVGNS = 'http://www.w3.org/2000/svg';
  const svgEl = <K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number>): SVGElementTagNameMap[K] => {
    const el = document.createElementNS(SVGNS, tag);
    for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
    return el;
  };

  /** Bounding box (display coords) of an annotation that is axis-aligned on screen. */
  function screenBox(a: Annotation): { x: number; y: number; w: number; h: number } | null {
    if (!viewport) return null;
    if (a.type === 'path') {
      const pts = a.points.map(([x, y]) => viewport!.toView(x, y));
      const xs = pts.map((q) => q[0]);
      const ys = pts.map((q) => q[1]);
      const pad = (a.width * viewport.scale) / 2 + 3;
      return { x: Math.min(...xs) - pad, y: Math.min(...ys) - pad, w: Math.max(...xs) - Math.min(...xs) + 2 * pad, h: Math.max(...ys) - Math.min(...ys) + 2 * pad };
    }
    const [px, py] = viewport.toView(a.x, a.y);
    const s = viewport.scale;
    if (a.type === 'text') {
      const lines = a.text.split('\n');
      const w = Math.max(...lines.map((l) => l.length)) * a.size * 0.55 * s;
      return { x: px, y: py - a.size * s, w: Math.max(w, 10), h: a.size * s * (1.25 * (lines.length - 1) + 1.25) };
    }
    return { x: px, y: py - a.h * s, w: a.w * s, h: a.h * s };
  }

  function drawOverlay(): void {
    if (!svg || !viewport || !editing) return;
    svg.replaceChildren();
    const vp = viewport;
    editing.annotations.forEach((a, i) => {
      let el: SVGElement;
      if (a.type === 'path') {
        el = svgEl('polyline', {
          points: a.points.map(([x, y]) => vp.toView(x, y).join(',')).join(' '),
          fill: 'none',
          stroke: rgbCss(a.color),
          'stroke-width': a.width * vp.scale,
          'stroke-linecap': 'round',
          'stroke-linejoin': 'round',
          opacity: a.opacity,
        });
      } else {
        const [px, py] = vp.toView(a.x, a.y);
        const g = svgEl('g', { transform: `translate(${px} ${py}) rotate(${vp.rotation - a.rotate})` });
        const s = vp.scale;
        if (a.type === 'text') {
          const t = svgEl('text', { x: 0, y: 0, fill: rgbCss(a.color), 'font-size': a.size * s, 'font-family': cssFont(a), 'font-weight': a.bold ? 700 : 400, 'font-style': a.italic ? 'italic' : 'normal', 'xml:space': 'preserve' });
          a.text.split('\n').forEach((line, k) => {
            const span = svgEl('tspan', { x: 0, dy: k ? `${1.25}em` : 0 });
            span.textContent = line || ' ';
            t.append(span);
          });
          g.append(t);
        } else if (a.type === 'rect') g.append(svgEl('rect', { x: 0, y: -a.h * s, width: a.w * s, height: a.h * s, fill: rgbCss(a.color), opacity: a.opacity }));
        else if (a.type === 'image') {
          const img = imageCache.get(a.data);
          g.append(svgEl('image', { x: 0, y: -a.h * s, width: a.w * s, height: a.h * s, href: img?.src ?? '', preserveAspectRatio: 'none' }));
        }
        el = g;
      }
      el.dataset.index = String(i);
      el.style.cursor = tool === 'select' ? 'move' : '';
      el.style.pointerEvents = tool === 'select' ? 'all' : 'none';
      svg!.append(el);
    });
    if (tool === 'edittext') {
      for (const seg of segments) {
        const b = segBox(seg);
        const r = svgEl('rect', { x: b.x, y: b.y, width: Math.max(4, b.w), height: b.h, class: 'text-seg', 'data-seg': seg.id });
        r.append(svgEl('title', {}));
        r.firstChild!.textContent = 'Править: ' + seg.text;
        svg.append(r);
      }
    }
    // Selection outline + resize handle
    const sel = editing.annotations[selectedIdx];
    if (sel && tool === 'select') {
      const box = screenBox(sel);
      if (box) {
        svg.append(svgEl('rect', { x: box.x - 3, y: box.y - 3, width: box.w + 6, height: box.h + 6, fill: 'none', stroke: '#4f46e5', 'stroke-width': 1.5, 'stroke-dasharray': '5 4', 'pointer-events': 'none' }));
        const aligned = sel.type !== 'path' && normRot(vp.rotation - sel.rotate) === 0;
        if (aligned && (sel.type === 'rect' || sel.type === 'image')) {
          const handle = svgEl('rect', { x: box.x + box.w - 6, y: box.y + box.h - 6, width: 12, height: 12, fill: '#fff', stroke: '#4f46e5', 'stroke-width': 1.5, 'data-handle': 'se' });
          handle.style.cursor = 'nwse-resize';
          svg.append(handle);
        }
      }
    }
  }

  function localPoint(e: PointerEvent): [number, number] {
    const r = stage!.getBoundingClientRect();
    return [e.clientX - r.left, e.clientY - r.top];
  }

  function attachStageEvents(): void {
    if (!stage) return;
    let action: { kind: 'rect' | 'pen' | 'move' | 'resize'; start: [number, number]; temp?: SVGElement; orig?: Annotation; points?: [number, number][] } | null = null;

    stage.addEventListener('pointerdown', (e) => {
      if (!editing || !viewport || e.button !== 0) return;
      if ((e.target as HTMLElement).tagName === 'TEXTAREA') return;
      const pt = localPoint(e);
      if (tool === 'edittext') {
        if (textEditor) {
          commitText();
          return;
        }
        const id = ((e.target as SVGElement).closest?.('[data-seg]') as SVGElement | null)?.dataset.seg;
        const seg = segments.find((x) => x.id === id);
        if (seg) {
          e.preventDefault();
          editSegment(seg);
        }
        return;
      }
      if (tool === 'text') {
        if (textEditor) {
          commitText();
          return;
        }
        e.preventDefault();
        startText(pt);
        return;
      }
      if (tool === 'select') {
        const target = e.target as SVGElement;
        if (target.dataset?.handle === 'se' && selectedIdx >= 0) {
          action = { kind: 'resize', start: pt, orig: cloneAnn(editing.annotations[selectedIdx]) };
          snapshot();
        } else {
          const hit = (target.closest?.('[data-index]') as SVGElement | null)?.dataset.index;
          selectedIdx = hit !== undefined ? Number(hit) : -1;
          if (selectedIdx >= 0) {
            snapshot();
            action = { kind: 'move', start: pt, orig: cloneAnn(editing.annotations[selectedIdx]) };
          }
          drawOverlay();
          refreshToolbarState();
        }
      } else if (tool === 'whiteout' || tool === 'highlight') {
        const temp = svgEl('rect', { x: pt[0], y: pt[1], width: 0, height: 0, fill: tool === 'whiteout' ? color : '#ffe94d', opacity: tool === 'whiteout' ? 1 : 0.4, stroke: '#4f46e5', 'stroke-dasharray': '4 3' });
        svg!.append(temp);
        action = { kind: 'rect', start: pt, temp };
      } else if (tool === 'pen') {
        const temp = svgEl('polyline', { points: pt.join(','), fill: 'none', stroke: color, 'stroke-width': strokeWidth * viewport.scale, 'stroke-linecap': 'round', 'stroke-linejoin': 'round' });
        svg!.append(temp);
        action = { kind: 'pen', start: pt, temp, points: [pt] };
      } else return;
      stage!.setPointerCapture(e.pointerId);
    });

    stage.addEventListener('pointermove', (e) => {
      if (!action || !editing || !viewport) return;
      const pt = localPoint(e);
      const dx = pt[0] - action.start[0];
      const dy = pt[1] - action.start[1];
      if (action.kind === 'rect' && action.temp) {
        action.temp.setAttribute('x', String(Math.min(pt[0], action.start[0])));
        action.temp.setAttribute('y', String(Math.min(pt[1], action.start[1])));
        action.temp.setAttribute('width', String(Math.abs(dx)));
        action.temp.setAttribute('height', String(Math.abs(dy)));
      } else if (action.kind === 'pen' && action.temp && action.points) {
        action.points.push(pt);
        action.temp.setAttribute('points', action.points.map((q) => q.join(',')).join(' '));
      } else if (action.kind === 'move' && action.orig) {
        editing.annotations[selectedIdx] = shifted(action.orig, dx, dy);
        drawOverlay();
      } else if (action.kind === 'resize' && action.orig && (action.orig.type === 'rect' || action.orig.type === 'image')) {
        const o = action.orig;
        const s = viewport.scale;
        const [px, py] = viewport.toView(o.x, o.y);
        const top = py - o.h * s;
        let w = Math.max(8, o.w * s + dx);
        let hh = Math.max(8, o.h * s + dy);
        if (o.type === 'image' && !e.shiftKey) hh = w * (o.h / o.w); // keep aspect unless Shift
        if (o.type === 'image' && e.shiftKey) w = Math.max(8, w);
        const [nx, ny] = viewport.toPdf(px, top + hh);
        editing.annotations[selectedIdx] = { ...o, x: nx, y: ny, w: w / s, h: hh / s };
        drawOverlay();
      }
    });

    const finish = (e: PointerEvent): void => {
      if (!action || !editing || !viewport) return;
      const vp = viewport;
      const pt = localPoint(e);
      if (action.kind === 'rect') {
        const x = Math.min(pt[0], action.start[0]);
        const y = Math.min(pt[1], action.start[1]);
        const w = Math.abs(pt[0] - action.start[0]);
        const hh = Math.abs(pt[1] - action.start[1]);
        if (w > 3 && hh > 3) {
          snapshot();
          const [px, py] = vp.toPdf(x, y + hh);
          const highlight = tool === 'highlight';
          editing.annotations.push({ type: 'rect', x: px, y: py, w: w / vp.scale, h: hh / vp.scale, color: highlight ? hexToRgb('#ffe94d') : hexToRgb(color), opacity: highlight ? 0.4 : 1, rotate: vp.rotation });
        }
      } else if (action.kind === 'pen' && action.points && action.points.length > 1) {
        snapshot();
        editing.annotations.push({ type: 'path', points: simplify(action.points).map(([x, y]) => vp.toPdf(x, y)), color: hexToRgb(color), width: strokeWidth, opacity: 1 });
      }
      action = null;
      drawOverlay();
      refreshToolbarState();
    };
    stage.addEventListener('pointerup', finish);
    stage.addEventListener('pointercancel', finish);

    stage.addEventListener('dblclick', (e) => {
      if (tool !== 'select' || !editing) return;
      const hit = ((e.target as SVGElement).closest?.('[data-index]') as SVGElement | null)?.dataset.index;
      const a = hit !== undefined ? editing.annotations[Number(hit)] : undefined;
      if (a?.type === 'text' && viewport) {
        const box = screenBox(a)!;
        snapshot();
        editing.annotations.splice(Number(hit), 1);
        selectedIdx = -1;
        drawOverlay();
        startText([box.x, box.y], a);
      }
    });
  }

  function refreshToolbarState(): void {
    const del = pageView.querySelector<HTMLButtonElement>('button[aria-label^="Удалить выбранную"]');
    if (del) del.disabled = selectedIdx < 0;
    const und = pageView.querySelector<HTMLButtonElement>('button[aria-label^="Отменить"]');
    if (und) und.disabled = !undoStack.length;
  }

  /** Move an annotation by a display-space delta. */
  function shifted(a: Annotation, dx: number, dy: number): Annotation {
    const vp = viewport!;
    if (a.type === 'path') return { ...a, points: a.points.map(([x, y]) => { const [vx, vy] = vp.toView(x, y); return vp.toPdf(vx + dx, vy + dy); }) };
    const [vx, vy] = vp.toView(a.x, a.y);
    const [nx, ny] = vp.toPdf(vx + dx, vy + dy);
    return { ...a, x: nx, y: ny };
  }

  /** Drop points closer than 1.5px to keep paths small. */
  function simplify(points: [number, number][]): [number, number][] {
    const out: [number, number][] = [points[0]];
    for (const p of points.slice(1)) {
      const last = out[out.length - 1];
      if (Math.hypot(p[0] - last[0], p[1] - last[1]) >= 1.5) out.push(p);
    }
    if (out.length === 1) out.push([points[0][0] + 0.5, points[0][1] + 0.5]);
    return out;
  }

  function startText(pt: [number, number], existing?: Extract<Annotation, { type: 'text' }>, replace?: { seg: TextSeg; bg: [number, number, number] }): void {
    if (!stage || !viewport) return;
    const s = viewport.scale;
    const size = existing?.size ?? fontSize;
    const ta = h('textarea', { class: 'text-edit', 'aria-label': replace ? 'Исправленный текст строки' : 'Текст пометки', rows: 1 });
    ta.value = existing?.text ?? '';
    Object.assign(ta.style, {
      left: `${pt[0]}px`,
      top: `${pt[1]}px`,
      fontSize: `${size * s}px`,
      color: existing ? rgbCss(existing.color) : color,
      fontWeight: existing?.bold ? '700' : '400',
      fontStyle: existing?.italic ? 'italic' : 'normal',
      fontFamily: cssFont(existing ?? {}),
      background: replace ? rgbCss(replace.bg) : '',
    });
    const fit = (): void => {
      ta.style.height = 'auto';
      ta.style.height = `${ta.scrollHeight}px`;
      ta.style.width = `${Math.max(120, ...ta.value.split('\n').map((l) => l.length * size * s * 0.6 + 20))}px`;
    };
    ta.addEventListener('input', fit);
    ta.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) commitText();
      if (e.key === 'Escape') {
        ta.value = existing?.text ?? '';
        commitText();
      }
      e.stopPropagation();
    });
    ta.addEventListener('blur', () => setTimeout(commitText, 0));
    (ta as HTMLTextAreaElement & { _meta?: unknown })._meta = { pt, size, color: existing?.color ?? hexToRgb(color), bold: existing?.bold, italic: existing?.italic, font: existing?.font, existing: !!existing, replace };
    stage.append(ta);
    textEditor = ta;
    fit();
    ta.focus();
  }

  function commitText(): void {
    const ta = textEditor;
    if (!ta || !editing || !viewport) return;
    textEditor = null;
    type Meta = { pt: [number, number]; size: number; color: [number, number, number]; bold?: boolean; italic?: boolean; font?: FontFamily; existing: boolean; replace?: { seg: TextSeg; bg: [number, number, number] } };
    const meta = (ta as HTMLTextAreaElement & { _meta?: Meta })._meta!;
    const text = ta.value.replace(/\s+$/, '');
    ta.remove();
    if (meta.replace) {
      // Editing the page's own text: cover the old line, write the new one in its place
      const seg = meta.replace.seg;
      if (text !== seg.text) {
        snapshot();
        const pad = seg.size * 0.08;
        editing.annotations.push({ type: 'rect', x: seg.x - pad, y: seg.y - seg.size * 0.26, w: seg.w + 2 * pad, h: seg.size * 1.2, color: meta.replace.bg, opacity: 1, rotate: viewport.rotation });
        if (text) editing.annotations.push({ type: 'text', x: seg.x, y: seg.y, size: meta.size, color: meta.color, text, rotate: viewport.rotation, bold: meta.bold, italic: meta.italic, font: meta.font });
        // The page is not reflowed: warn when the new text runs into the text that follows it
        const ctx = document.createElement('canvas').getContext('2d');
        if (ctx && text) {
          ctx.font = `${meta.italic ? 'italic ' : ''}${meta.bold ? 700 : 400} ${meta.size}px ${CSS_FAMILY[meta.font ?? 'sans']}`;
          const neighbour = segments.some((x) => x !== seg && Math.abs(x.y - seg.y) < seg.size * 0.3 && x.x > seg.x && x.x < seg.x + ctx.measureText(text).width - 0.5);
          if (neighbour) toast('Новый текст длиннее исходного и заходит на соседний фрагмент строки. Уменьшите размер шрифта (поле «размер») или исправьте и соседний фрагмент.');
        }
        editing.replaced = [...(editing.replaced ?? []), seg.id];
        segments = segments.filter((x) => x !== seg);
      }
      drawOverlay();
      refreshToolbarState();
      return;
    }
    if (text) {
      if (!meta.existing) snapshot();
      const s = viewport.scale;
      // Baseline of the first line ≈ top + ascent
      const [x, y] = viewport.toPdf(meta.pt[0] + 1, meta.pt[1] + meta.size * s * 0.95);
      editing.annotations.push({ type: 'text', x, y, size: meta.size, color: meta.color, text, rotate: viewport.rotation, bold: meta.bold, italic: meta.italic, font: meta.font });
    }
    drawOverlay();
    refreshToolbarState();
  }

  async function addImage(): Promise<void> {
    const [file] = await pickFiles('image/*', false);
    if (!file || !editing || !viewport) return;
    try {
      const bytes = await readFile(file);
      const det = detectFormat(file.name, bytes);
      if (!det.format || FORMATS[det.format].kind !== 'image') throw new Error('Это не изображение.');
      const img = await toPdfImage(file, det.format);
      await ensureImages([{ type: 'image', x: 0, y: 0, w: 1, h: 1, data: img.data, mime: img.mime, rotate: 0 }]);
      const vp = viewport;
      const maxW = vp.width * 0.4;
      const k = Math.min(1, maxW / img.width, (vp.height * 0.4) / img.height);
      const w = img.width * k;
      const hh = img.height * k;
      const left = (vp.width - w) / 2;
      const top = (vp.height - hh) / 2;
      const [x, y] = vp.toPdf(left, top + hh);
      snapshot();
      editing.annotations.push({ type: 'image', x, y, w: w / vp.scale, h: hh / vp.scale, data: img.data, mime: img.mime, rotate: vp.rotation });
      tool = 'select';
      selectedIdx = editing.annotations.length - 1;
      await renderEditor();
      toast('Изображение добавлено: перетащите его в нужное место, размер меняется за правый нижний угол.');
    } catch (e) {
      toast((e as Error).message, 'error');
    }
  }

  const onKey = (e: KeyboardEvent): void => {
    if (!editing || textEditor || (e.target as HTMLElement).closest('input, textarea, select')) return;
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedIdx >= 0) {
      e.preventDefault();
      deleteSelected();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      undo();
    } else if (e.key === 'Escape') closeEditor();
  };
  document.addEventListener('keydown', onKey);

  renderToolbar();
  const handed = takeHandoff('/pdf-editor');
  if (handed) void addFiles([handed]);

  return () => {
    observer.disconnect();
    document.removeEventListener('keydown', onKey);
    for (const s of sources) void s.pdf.loadingTask.destroy();
  };
}
