import { detectFormat, FORMATS, IMAGE_OUTPUTS, type FormatId } from '../../core/formats';
import { decodeToCanvas, encodeCanvas, supportsAvifEncoding } from '../raster';
import { baseName, button, download, dropzone, formatBytes, h, iconButton, onPasteFiles, pickFiles, prefs, promptText, readFile, takeHandoff, toast } from '../ui';

type Tool = 'none' | 'crop' | 'draw' | 'text';

interface Adjust {
  brightness: number;
  contrast: number;
  saturation: number;
  grayscale: boolean;
  sepia: boolean;
  invert: boolean;
}

const NO_ADJUST: Adjust = { brightness: 0, contrast: 0, saturation: 0, grayscale: false, sepia: false, invert: false };
const HISTORY_PIXELS = 200_000_000;

function cloneCanvas(src: HTMLCanvasElement): HTMLCanvasElement {
  const c = document.createElement('canvas');
  c.width = src.width;
  c.height = src.height;
  c.getContext('2d')!.drawImage(src, 0, 0);
  return c;
}

/** Per-pixel colour adjustments (works in every browser, unlike ctx.filter). */
export function applyAdjust(data: Uint8ClampedArray, a: Adjust): void {
  const b = a.brightness * 2.55;
  const c = a.contrast * 2.55;
  const f = (259 * (c + 255)) / (255 * (259 - c));
  const s = 1 + a.saturation / 100;
  for (let i = 0; i < data.length; i += 4) {
    let r = data[i];
    let g = data[i + 1];
    let bl = data[i + 2];
    if (b) {
      r += b;
      g += b;
      bl += b;
    }
    if (c) {
      r = f * (r - 128) + 128;
      g = f * (g - 128) + 128;
      bl = f * (bl - 128) + 128;
    }
    if (s !== 1 || a.grayscale) {
      const l = 0.2126 * r + 0.7152 * g + 0.0722 * bl;
      const k = a.grayscale ? 0 : s;
      r = l + (r - l) * k;
      g = l + (g - l) * k;
      bl = l + (bl - l) * k;
    }
    if (a.sepia) {
      const nr = 0.393 * r + 0.769 * g + 0.189 * bl;
      const ng = 0.349 * r + 0.686 * g + 0.168 * bl;
      const nb = 0.272 * r + 0.534 * g + 0.131 * bl;
      r = nr;
      g = ng;
      bl = nb;
    }
    if (a.invert) {
      r = 255 - r;
      g = 255 - g;
      bl = 255 - bl;
    }
    data[i] = r;
    data[i + 1] = g;
    data[i + 2] = bl;
  }
}

export function mount(root: HTMLElement): () => void {
  let canvas: HTMLCanvasElement | null = null;
  let fileName = 'image';
  let history: HTMLCanvasElement[] = [];
  let future: HTMLCanvasElement[] = [];
  let tool: Tool = 'none';
  let adjust: Adjust = { ...NO_ADJUST };
  let previewBase: HTMLCanvasElement | null = null;
  let crop = { x: 0, y: 0, w: 0, h: 0 };
  let cropRatio = 0;
  let exportFormat = prefs.get<FormatId>('img-format', 'png');
  let quality = prefs.get<number>('img-quality', 0.9);
  let penColor = '#e11d48';
  let penSize = 8;
  let textSize = 48;
  let avif = false;

  const stageInner = h('div', { class: 'img-stage__inner' });
  const stage = h('div', { class: 'img-stage' }, stageInner);
  const panel = h('aside', { class: 'panel card', 'aria-label': 'Параметры' });
  const toolbar = h('div', { class: 'toolbar card', role: 'toolbar', 'aria-label': 'Инструменты' });
  const empty = h(
    'div',
    { class: 'empty' },
    dropzone({ accept: 'image/*,.svg,.ico,.avif,.tif,.tiff,.heic,.heif', multiple: false, title: 'Откройте изображение', hint: 'JPG, PNG, WebP, GIF, BMP, SVG, ICO, AVIF, TIFF, HEIC · перетащите файл или вставьте из буфера (Ctrl+V)', onFiles: (f) => void open(f[0]) }),
  );
  const workspace = h('div', { class: 'workspace', hidden: true }, toolbar, h('div', { class: 'img-layout' }, stage, panel));

  root.append(
    h('div', { class: 'page-head' }, h('h1', null, 'Редактор изображений'), h('p', null, 'Обрезка, поворот, изменение размера, цветокоррекция, рисование и надписи. Всё выполняется в браузере.')),
    empty,
    workspace,
  );
  onPasteFiles(root, (files) => {
    const img = files.find((f) => f.type.startsWith('image/'));
    if (img) void open(img);
  });
  void supportsAvifEncoding().then((v) => {
    avif = v;
    if (canvas) renderPanel();
  });

  // ------------------------------------------------------------------ open / history
  async function open(file: File | undefined): Promise<void> {
    if (!file) return;
    try {
      const bytes = await readFile(file);
      const det = detectFormat(file.name, bytes);
      if (!det.format || FORMATS[det.format].kind !== 'image') throw new Error(det.error ?? 'Это не изображение.');
      canvas = await decodeToCanvas(file, det.format);
      fileName = baseName(file.name);
      history = [];
      future = [];
      adjust = { ...NO_ADJUST };
      previewBase = null;
      setTool('none');
      empty.hidden = true;
      workspace.hidden = false;
      show(canvas);
      renderToolbar();
      renderPanel();
    } catch (e) {
      toast((e as Error).message, 'error');
    }
  }

  function historyPixels(): number {
    return history.reduce((s, c) => s + c.width * c.height, 0);
  }

  /** Record the current state, then replace the working canvas. */
  function commit(next: HTMLCanvasElement): void {
    if (!canvas) return;
    history.push(canvas);
    while (history.length > 1 && historyPixels() > HISTORY_PIXELS) history.shift();
    future = [];
    canvas = next;
    show(canvas);
    renderToolbar();
    renderPanel();
  }

  function undo(): void {
    const prev = history.pop();
    if (!prev || !canvas) return;
    future.push(canvas);
    canvas = prev;
    previewBase = null;
    show(canvas);
    renderToolbar();
    renderPanel();
  }

  function redo(): void {
    const next = future.pop();
    if (!next || !canvas) return;
    history.push(canvas);
    canvas = next;
    show(canvas);
    renderToolbar();
    renderPanel();
  }

  function show(c: HTMLCanvasElement): void {
    stageInner.replaceChildren(c);
    c.setAttribute('aria-label', `Изображение ${fileName}, ${c.width}×${c.height}`);
    c.setAttribute('role', 'img');
    if (tool === 'crop') renderCrop();
  }

  // ------------------------------------------------------------------ transforms
  function transformed(w: number, h2: number, draw: (ctx: CanvasRenderingContext2D, src: HTMLCanvasElement) => void): HTMLCanvasElement {
    const c = document.createElement('canvas');
    c.width = w;
    c.height = h2;
    const ctx = c.getContext('2d')!;
    ctx.imageSmoothingQuality = 'high';
    draw(ctx, canvas!);
    return c;
  }

  /** Apply a pending colour adjustment preview before any other operation. */
  function settle(): void {
    if (isAdjusted()) applyAdjustments();
  }

  function rotate(dir: 1 | -1): void {
    if (!canvas) return;
    settle();
    const { width: w, height: hh } = canvas;
    commit(
      transformed(hh, w, (ctx, src) => {
        ctx.translate(dir === 1 ? hh : 0, dir === 1 ? 0 : w);
        ctx.rotate((dir * Math.PI) / 2);
        ctx.drawImage(src, 0, 0);
      }),
    );
  }

  function flip(axis: 'h' | 'v'): void {
    if (!canvas) return;
    settle();
    const { width: w, height: hh } = canvas;
    commit(
      transformed(w, hh, (ctx, src) => {
        ctx.translate(axis === 'h' ? w : 0, axis === 'v' ? hh : 0);
        ctx.scale(axis === 'h' ? -1 : 1, axis === 'v' ? -1 : 1);
        ctx.drawImage(src, 0, 0);
      }),
    );
  }

  function resize(w: number, hh: number): void {
    if (!canvas) return;
    settle();
    w = Math.round(w);
    hh = Math.round(hh);
    if (w < 1 || hh < 1 || w > 16384 || hh > 16384) {
      toast('Размер должен быть от 1 до 16384 пикселей.', 'error');
      return;
    }
    // Step down by halves for good downscaling quality
    let src = canvas;
    while (src.width / 2 >= w && src.height / 2 >= hh) {
      const half = document.createElement('canvas');
      half.width = Math.round(src.width / 2);
      half.height = Math.round(src.height / 2);
      const ctx = half.getContext('2d')!;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(src, 0, 0, half.width, half.height);
      src = half;
    }
    const finalSrc = src;
    commit(transformed(w, hh, (ctx) => ctx.drawImage(finalSrc, 0, 0, w, hh)));
  }

  // ------------------------------------------------------------------ adjustments (live preview)
  function previewAdjust(): void {
    if (!canvas) return;
    if (!previewBase) {
      const k = Math.min(1, 1200 / Math.max(canvas.width, canvas.height));
      previewBase = document.createElement('canvas');
      previewBase.width = Math.max(1, Math.round(canvas.width * k));
      previewBase.height = Math.max(1, Math.round(canvas.height * k));
      previewBase.getContext('2d')!.drawImage(canvas, 0, 0, previewBase.width, previewBase.height);
    }
    const view = cloneCanvas(previewBase);
    const ctx = view.getContext('2d')!;
    const img = ctx.getImageData(0, 0, view.width, view.height);
    applyAdjust(img.data, adjust);
    ctx.putImageData(img, 0, 0);
    view.style.width = `${canvas.width}px`;
    view.style.aspectRatio = `${canvas.width} / ${canvas.height}`;
    stageInner.replaceChildren(view);
  }

  function applyAdjustments(): void {
    if (!canvas) return;
    const out = cloneCanvas(canvas);
    const ctx = out.getContext('2d')!;
    const img = ctx.getImageData(0, 0, out.width, out.height);
    applyAdjust(img.data, adjust);
    ctx.putImageData(img, 0, 0);
    adjust = { ...NO_ADJUST };
    previewBase = null;
    commit(out);
  }

  function resetAdjustments(): void {
    adjust = { ...NO_ADJUST };
    previewBase = null;
    if (canvas) show(canvas);
    renderPanel();
  }

  const isAdjusted = (): boolean => JSON.stringify(adjust) !== JSON.stringify(NO_ADJUST);

  // ------------------------------------------------------------------ pointer tools
  function toImage(e: PointerEvent): { x: number; y: number } {
    const r = canvas!.getBoundingClientRect();
    return { x: ((e.clientX - r.left) * canvas!.width) / r.width, y: ((e.clientY - r.top) * canvas!.height) / r.height };
  }

  let drawing: { last: { x: number; y: number }; work: HTMLCanvasElement } | null = null;

  stage.addEventListener('pointerdown', (e) => {
    if (!canvas || e.button !== 0) return;
    if (tool === 'draw' && e.target === canvas) {
      const work = cloneCanvas(canvas);
      history.push(canvas);
      future = [];
      canvas = work;
      show(canvas);
      drawing = { last: toImage(e), work };
      stage.setPointerCapture(e.pointerId);
      dot(drawing.last);
    } else if (tool === 'text' && e.target === canvas) {
      void placeText(toImage(e));
    }
  });
  stage.addEventListener('pointermove', (e) => {
    if (!drawing || !canvas) return;
    const p = toImage(e);
    const ctx = canvas.getContext('2d')!;
    ctx.strokeStyle = penColor;
    ctx.lineWidth = penSize;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.beginPath();
    ctx.moveTo(drawing.last.x, drawing.last.y);
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
    drawing.last = p;
  });
  const endDraw = (): void => {
    if (!drawing) return;
    drawing = null;
    renderToolbar();
  };
  stage.addEventListener('pointerup', endDraw);
  stage.addEventListener('pointercancel', endDraw);

  function dot(p: { x: number; y: number }): void {
    const ctx = canvas!.getContext('2d')!;
    ctx.fillStyle = penColor;
    ctx.beginPath();
    ctx.arc(p.x, p.y, penSize / 2, 0, Math.PI * 2);
    ctx.fill();
  }

  async function placeText(p: { x: number; y: number }): Promise<void> {
    const text = await promptText('Добавить надпись', 'Текст', '');
    if (!text || !canvas) return;
    const out = cloneCanvas(canvas);
    const ctx = out.getContext('2d')!;
    ctx.fillStyle = penColor;
    ctx.font = `600 ${textSize}px system-ui, "Segoe UI", Arial, sans-serif`;
    ctx.textBaseline = 'top';
    text.split('\n').forEach((line, i) => ctx.fillText(line, p.x, p.y + i * textSize * 1.2));
    commit(out);
  }

  // ------------------------------------------------------------------ crop overlay
  let cropEl: HTMLElement | null = null;

  function startCrop(): void {
    if (!canvas) return;
    const m = Math.round(Math.min(canvas.width, canvas.height) * 0.1);
    crop = { x: m, y: m, w: canvas.width - 2 * m, h: canvas.height - 2 * m };
    if (cropRatio) fitRatio();
    renderCrop();
  }

  function fitRatio(): void {
    if (!canvas || !cropRatio) return;
    let w = crop.w;
    let hh = w / cropRatio;
    if (hh > canvas.height) {
      hh = canvas.height;
      w = hh * cropRatio;
    }
    if (crop.x + w > canvas.width) crop.x = Math.max(0, canvas.width - w);
    if (crop.y + hh > canvas.height) crop.y = Math.max(0, canvas.height - hh);
    crop.w = w;
    crop.h = hh;
  }

  function renderCrop(): void {
    cropEl?.remove();
    if (!canvas || tool !== 'crop') return;
    const r = canvas.getBoundingClientRect();
    const k = r.width / canvas.width || 1;
    cropEl = h('div', { class: 'crop-box', style: { left: `${crop.x * k}px`, top: `${crop.y * k}px`, width: `${crop.w * k}px`, height: `${crop.h * k}px` }, 'aria-label': 'Область обрезки' });
    const handles: [string, number, number][] = [
      ['nw', 0, 0],
      ['ne', 1, 0],
      ['sw', 0, 1],
      ['se', 1, 1],
    ];
    for (const [name, hx, hy] of handles) {
      const hd = h('div', { class: 'crop-box__handle', style: { left: `calc(${hx * 100}% - 7px)`, top: `calc(${hy * 100}% - 7px)`, cursor: `${name}-resize` } });
      hd.dataset.handle = name;
      cropEl.append(hd);
    }
    cropEl.addEventListener('pointerdown', (e) => {
      e.stopPropagation();
      const handle = (e.target as HTMLElement).dataset.handle;
      const start = { x: e.clientX, y: e.clientY, crop: { ...crop } };
      const kk = canvas!.width / canvas!.getBoundingClientRect().width;
      const move = (ev: PointerEvent): void => {
        const dx = (ev.clientX - start.x) * kk;
        const dy = (ev.clientY - start.y) * kk;
        const c0 = start.crop;
        const W = canvas!.width;
        const H = canvas!.height;
        if (!handle) {
          crop.x = Math.min(Math.max(0, c0.x + dx), W - c0.w);
          crop.y = Math.min(Math.max(0, c0.y + dy), H - c0.h);
        } else {
          let x0 = c0.x;
          let y0 = c0.y;
          let x1 = c0.x + c0.w;
          let y1 = c0.y + c0.h;
          if (handle.includes('w')) x0 = Math.min(Math.max(0, c0.x + dx), x1 - 10);
          if (handle.includes('e')) x1 = Math.max(Math.min(W, x1 + dx), x0 + 10);
          if (handle.includes('n')) y0 = Math.min(Math.max(0, c0.y + dy), y1 - 10);
          if (handle.includes('s')) y1 = Math.max(Math.min(H, y1 + dy), y0 + 10);
          if (cropRatio) {
            const w = x1 - x0;
            const hh = w / cropRatio;
            if (handle.includes('n')) y0 = y1 - hh;
            else y1 = y0 + hh;
            if (y0 < 0 || y1 > H) return;
          }
          crop = { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
        }
        const k2 = canvas!.getBoundingClientRect().width / canvas!.width;
        Object.assign(cropEl!.style, { left: `${crop.x * k2}px`, top: `${crop.y * k2}px`, width: `${crop.w * k2}px`, height: `${crop.h * k2}px` });
        renderCropInfo();
      };
      const up = (): void => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
      };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
    });
    stageInner.append(cropEl);
  }

  function applyCrop(): void {
    if (!canvas) return;
    const x = Math.round(crop.x);
    const y = Math.round(crop.y);
    const w = Math.max(1, Math.round(crop.w));
    const hh = Math.max(1, Math.round(crop.h));
    setTool('none');
    commit(transformed(w, hh, (ctx, src) => ctx.drawImage(src, x, y, w, hh, 0, 0, w, hh)));
  }

  const cropInfo = h('div', { style: 'font-size:13px;color:var(--muted)' });
  function renderCropInfo(): void {
    cropInfo.textContent = `Область: ${Math.round(crop.w)} × ${Math.round(crop.h)} px`;
  }

  const onResize = (): void => renderCrop();
  window.addEventListener('resize', onResize);

  function setTool(t: Tool): void {
    if (t !== 'none') settle();
    tool = tool === t ? 'none' : t;
    stage.style.cursor = tool === 'draw' ? 'crosshair' : tool === 'text' ? 'text' : '';
    if (tool === 'crop') startCrop();
    else {
      cropEl?.remove();
      cropEl = null;
    }
    if (canvas) {
      renderToolbar();
      renderPanel();
    }
  }

  // ------------------------------------------------------------------ export
  async function exportImage(): Promise<void> {
    if (!canvas) return;
    try {
      if (isAdjusted()) applyAdjustments();
      const data = await encodeCanvas(canvas, exportFormat, quality);
      download(data, `${fileName}-edited.${FORMATS[exportFormat].ext}`, FORMATS[exportFormat].mime);
      toast(`Сохранено: ${formatBytes(data.length)}`);
    } catch (e) {
      toast((e as Error).message, 'error');
    }
  }

  // ------------------------------------------------------------------ rendering of controls
  function renderToolbar(): void {
    const toolBtn = (t: Tool, icn: string, label: string): HTMLButtonElement => iconButton(icn, label, () => setTool(t), { 'aria-pressed': String(tool === t) });
    toolbar.replaceChildren(
      button('Открыть', async () => void open((await pickFiles('image/*,.svg,.ico,.avif,.tif,.tiff,.heic,.heif', false))[0]), { icon: 'folder', small: true }),
      h('span', { class: 'toolbar__sep' }),
      iconButton('undo', 'Отменить (Ctrl+Z)', undo, { disabled: !history.length }),
      iconButton('redo', 'Повторить (Ctrl+Y)', redo, { disabled: !future.length }),
      h('span', { class: 'toolbar__sep' }),
      iconButton('rotate-ccw', 'Повернуть влево', () => rotate(-1)),
      iconButton('rotate-cw', 'Повернуть вправо', () => rotate(1)),
      iconButton('flip-h', 'Отразить по горизонтали', () => flip('h')),
      iconButton('flip-v', 'Отразить по вертикали', () => flip('v')),
      h('span', { class: 'toolbar__sep' }),
      toolBtn('crop', 'crop', 'Обрезка'),
      toolBtn('draw', 'pen', 'Рисование'),
      toolBtn('text', 'type', 'Надпись'),
      h('span', { class: 'toolbar__spacer' }),
      h('span', { style: 'color:var(--muted);font-size:13px;margin-right:6px' }, canvas ? `${canvas.width} × ${canvas.height}` : ''),
      button('Скачать', () => void exportImage(), { icon: 'download', variant: 'primary', small: true }),
    );
  }

  function slider(label: string, key: 'brightness' | 'contrast' | 'saturation'): HTMLElement {
    const out = h('output', null, String(adjust[key]));
    return h(
      'label',
      { class: 'slider' },
      h('span', null, label),
      out,
      h('input', {
        type: 'range',
        min: -100,
        max: 100,
        value: adjust[key],
        oninput: (e: Event) => {
          adjust[key] = +(e.target as HTMLInputElement).value;
          out.textContent = String(adjust[key]);
          previewAdjust();
        },
      }),
    );
  }

  function toggle(label: string, key: 'grayscale' | 'sepia' | 'invert'): HTMLElement {
    return h(
      'label',
      { class: 'checkbox' },
      h('input', {
        type: 'checkbox',
        checked: adjust[key],
        onchange: (e: Event) => {
          adjust[key] = (e.target as HTMLInputElement).checked;
          previewAdjust();
        },
      }),
      label,
    );
  }

  function renderPanel(): void {
    if (!canvas) return;
    const sections: HTMLElement[] = [];
    if (tool === 'crop') {
      renderCropInfo();
      const ratios: [string, number][] = [
        ['Свободно', 0],
        ['1:1', 1],
        ['4:3', 4 / 3],
        ['3:2', 3 / 2],
        ['16:9', 16 / 9],
        ['9:16', 9 / 16],
      ];
      sections.push(
        h(
          'div',
          { class: 'panel__section' },
          h('h3', null, 'Обрезка'),
          h(
            'div',
            { class: 'row' },
            ratios.map(([label, r]) =>
              h(
                'button',
                {
                  type: 'button',
                  class: 'chip',
                  'aria-pressed': String(cropRatio === r),
                  onclick: () => {
                    cropRatio = r;
                    fitRatio();
                    renderCrop();
                    renderPanel();
                  },
                },
                label,
              ),
            ),
          ),
          cropInfo,
          h('div', { class: 'row' }, button('Применить', applyCrop, { variant: 'primary', small: true, icon: 'check' }), button('Отмена', () => setTool('none'), { small: true, variant: 'ghost' })),
        ),
      );
    }
    if (tool === 'draw' || tool === 'text') {
      const sizeOut = h('output', null, String(tool === 'draw' ? penSize : textSize));
      sections.push(
        h(
          'div',
          { class: 'panel__section' },
          h('h3', null, tool === 'draw' ? 'Кисть' : 'Надпись'),
          h('p', { style: 'margin:0;font-size:13px;color:var(--muted)' }, tool === 'draw' ? 'Рисуйте мышью или пальцем по изображению.' : 'Нажмите на изображение в месте, где нужна надпись.'),
          h(
            'label',
            { class: 'row' },
            'Цвет',
            h('input', { type: 'color', class: 'input', value: penColor, oninput: (e: Event) => (penColor = (e.target as HTMLInputElement).value) }),
          ),
          h(
            'label',
            { class: 'slider' },
            h('span', null, tool === 'draw' ? 'Толщина' : 'Размер шрифта'),
            sizeOut,
            h('input', {
              type: 'range',
              min: tool === 'draw' ? 1 : 8,
              max: tool === 'draw' ? 120 : 400,
              value: tool === 'draw' ? penSize : textSize,
              oninput: (e: Event) => {
                const v = +(e.target as HTMLInputElement).value;
                if (tool === 'draw') penSize = v;
                else textSize = v;
                sizeOut.textContent = String(v);
              },
            }),
          ),
        ),
      );
    }
    // Resize
    const wIn = h('input', { class: 'input', type: 'number', min: 1, max: 16384, value: canvas.width, 'aria-label': 'Ширина', style: 'width:90px' });
    const hIn = h('input', { class: 'input', type: 'number', min: 1, max: 16384, value: canvas.height, 'aria-label': 'Высота', style: 'width:90px' });
    const lock = h('input', { type: 'checkbox', checked: true });
    const ratio = canvas.width / canvas.height;
    wIn.addEventListener('input', () => {
      if (lock.checked) hIn.value = String(Math.max(1, Math.round(+wIn.value / ratio)));
    });
    hIn.addEventListener('input', () => {
      if (lock.checked) wIn.value = String(Math.max(1, Math.round(+hIn.value * ratio)));
    });
    sections.push(
      h(
        'div',
        { class: 'panel__section' },
        h('h3', null, 'Размер'),
        h('div', { class: 'row' }, wIn, '×', hIn, 'px'),
        h('label', { class: 'checkbox' }, lock, 'Сохранять пропорции'),
        h(
          'div',
          { class: 'row' },
          button('Изменить', () => resize(+wIn.value, +hIn.value), { small: true }),
          button('50%', () => resize(canvas!.width / 2, canvas!.height / 2), { small: true, variant: 'ghost' }),
        ),
      ),
    );
    // Adjustments
    sections.push(
      h(
        'div',
        { class: 'panel__section' },
        h('h3', null, 'Коррекция'),
        slider('Яркость', 'brightness'),
        slider('Контраст', 'contrast'),
        slider('Насыщенность', 'saturation'),
        h('div', { class: 'row' }, toggle('Ч/Б', 'grayscale'), toggle('Сепия', 'sepia'), toggle('Негатив', 'invert')),
        h('div', { class: 'row' }, button('Применить', applyAdjustments, { small: true, variant: 'primary' }), button('Сбросить', resetAdjustments, { small: true, variant: 'ghost' })),
      ),
    );
    // Export
    const outs = [...IMAGE_OUTPUTS];
    if (avif) outs.splice(3, 0, 'avif');
    if (!outs.includes(exportFormat)) exportFormat = 'png';
    const qOut = h('output', null, `${Math.round(quality * 100)}%`);
    sections.push(
      h(
        'div',
        { class: 'panel__section' },
        h('h3', null, 'Экспорт'),
        h(
          'select',
          {
            class: 'select',
            'aria-label': 'Формат экспорта',
            onchange: (e: Event) => {
              exportFormat = (e.target as HTMLSelectElement).value as FormatId;
              prefs.set('img-format', exportFormat);
              renderPanel();
            },
          },
          outs.map((f) => h('option', { value: f, selected: f === exportFormat }, FORMATS[f].label)),
        ),
        ['jpg', 'webp', 'avif'].includes(exportFormat)
          ? h(
              'label',
              { class: 'slider' },
              h('span', null, 'Качество'),
              qOut,
              h('input', {
                type: 'range',
                min: 30,
                max: 100,
                value: Math.round(quality * 100),
                oninput: (e: Event) => {
                  quality = +(e.target as HTMLInputElement).value / 100;
                  qOut.textContent = `${Math.round(quality * 100)}%`;
                  prefs.set('img-quality', quality);
                },
              }),
            )
          : null,
        button('Скачать', () => void exportImage(), { icon: 'download', variant: 'primary' }),
      ),
    );
    panel.replaceChildren(...sections);
  }

  const onKey = (e: KeyboardEvent): void => {
    if (!canvas || (e.target as HTMLElement).closest('input, textarea, select, [contenteditable="true"]')) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 'z' && !e.shiftKey) {
      e.preventDefault();
      undo();
    } else if (mod && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) {
      e.preventDefault();
      redo();
    } else if (e.key === 'Enter' && tool === 'crop') applyCrop();
    else if (e.key === 'Escape' && tool !== 'none') setTool('none');
  };
  document.addEventListener('keydown', onKey);

  const handed = takeHandoff('/image-editor');
  if (handed) void open(handed);

  return () => {
    document.removeEventListener('keydown', onKey);
    window.removeEventListener('resize', onResize);
  };
}
