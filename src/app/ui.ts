/** Tiny DOM toolkit: element builder, icons, toasts, dialogs, downloads, file pickers. */

type Child = Node | string | number | null | undefined | false;
type Attrs = Record<string, unknown> & { class?: string; style?: Partial<CSSStyleDeclaration> | string };

export function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Attrs | null = null, ...children: (Child | Child[])[]): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null || v === false) continue;
      if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v as EventListener);
      else if (k === 'class') el.className = String(v);
      else if (k in el && typeof v !== 'string') (el as unknown as Record<string, unknown>)[k] = v;
      else el.setAttribute(k, v === true ? '' : String(v));
    }
  }
  append(el, children);
  return el;
}

function append(el: Element, children: (Child | Child[])[]): void {
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c instanceof Node ? c : String(c));
  }
}

const ICONS: Record<string, string> = {
  upload: '<path d="M12 16V4m0 0-4.5 4.5M12 4l4.5 4.5"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>',
  file: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
  'file-text': '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h6M9 9h1"/>',
  table: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/>',
  image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="m21 16-5-5L5 20"/>',
  download: '<path d="M12 4v12m0 0 4.5-4.5M12 16l-4.5-4.5"/><path d="M4 18v1a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-1"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12M9 7V4h6v3"/>',
  'rotate-cw': '<path d="M20 12a8 8 0 1 1-2.34-5.66L20 8.5"/><path d="M20 3.5v5h-5"/>',
  'rotate-ccw': '<path d="M4 12a8 8 0 1 0 2.34-5.66L4 8.5"/><path d="M4 3.5v5h5"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
  check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
  alert: '<path d="M12 3 2 20h20z"/><path d="M12 10v4M12 17.5v.5"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8v.5"/>',
  arrow: '<path d="M5 12h14m0 0-5-5m5 5-5 5"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  move: '<path d="M12 3v18M3 12h18M12 3l-3 3m3-3 3 3m-3 15-3-3m3 3 3-3M3 12l3-3m-3 3 3 3m15-3-3-3m3 3-3 3"/>',
  pointer: '<path d="M5 3l14 7-6 2-2 6z"/>',
  type: '<path d="M5 6V4h14v2M12 4v16M9 20h6"/>',
  'edit-text': '<path d="M4 7V5h11v2M9.5 5v12M7 17h5"/><path d="m14 20 1-3 5-5a1.4 1.4 0 0 1 2 2l-5 5z"/>',
  square: '<rect x="4" y="4" width="16" height="16" rx="1"/>',
  eraser: '<path d="M8 20h12M5.5 13.5l7-7a2 2 0 0 1 2.8 0l2.2 2.2a2 2 0 0 1 0 2.8L11 18H8.5z"/>',
  highlighter: '<path d="m9 11 6 6M4 20h6l9.5-9.5a2.1 2.1 0 0 0-3-3L7 17z"/>',
  pen: '<path d="M4 20l4-1 11-11a2.1 2.1 0 0 0-3-3L5 16z"/>',
  crop: '<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M18 22V8a2 2 0 0 0-2-2H2"/>',
  'flip-h': '<path d="M12 3v18M8 7 3 12l5 5zM16 7l5 5-5 5z"/>',
  'flip-v': '<path d="M3 12h18M7 8l5-5 5 5zM7 16l5 5 5-5z"/>',
  undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
  redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
  resize: '<path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/>',
  sliders: '<path d="M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12M20 18h0"/><circle cx="16" cy="6" r="2"/><circle cx="10" cy="12" r="2"/><circle cx="18" cy="18" r="2"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
  moon: '<path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5z"/>',
  'image-plus': '<path d="M21 12v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h7"/><circle cx="8.5" cy="9.5" r="1.5"/><path d="m21 15-5-5L5 20M16 5h6M19 2v6"/>',
  copy: '<rect x="8" y="8" width="13" height="13" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/>',
  'chevron-left': '<path d="m15 18-6-6 6-6"/>',
  'chevron-right': '<path d="m9 18 6-6-6-6"/>',
  save: '<path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M7 3v5h8M7 21v-7h10v7"/>',
  'zoom-in': '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5M8 11h6M11 8v6"/>',
  'zoom-out': '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5M8 11h6"/>',
  bold: '<path d="M7 4h6a4 4 0 0 1 0 8H7zM7 12h7a4 4 0 0 1 0 8H7z"/>',
  italic: '<path d="M19 4h-9M14 20H5M15 4 9 20"/>',
  underline: '<path d="M6 4v6a6 6 0 0 0 12 0V4M4 20h16"/>',
  strike: '<path d="M16 4H9a3 3 0 0 0-2.8 4M4 12h16M15 12a3.5 3.5 0 0 1-3 8H7"/>',
  'list-ul': '<path d="M9 6h11M9 12h11M9 18h11"/><circle cx="4.5" cy="6" r="1"/><circle cx="4.5" cy="12" r="1"/><circle cx="4.5" cy="18" r="1"/>',
  'list-ol': '<path d="M10 6h10M10 12h10M10 18h10M4 4h1v4M4 8h2M4 14h2l-2 3h2"/>',
  link: '<path d="M10 14a4.5 4.5 0 0 0 6.4 0l3-3a4.5 4.5 0 0 0-6.4-6.4l-1 1"/><path d="M14 10a4.5 4.5 0 0 0-6.4 0l-3 3a4.5 4.5 0 0 0 6.4 6.4l1-1"/>',
  quote: '<path d="M9 7H5v6h4v-2c0-1-1-2-2-2M19 7h-4v6h4v-2c0-1-1-2-2-2"/>',
  code: '<path d="m8 8-4 4 4 4M16 8l4 4-4 4M14 5l-4 14"/>',
  'align-left': '<path d="M4 6h16M4 12h10M4 18h14"/>',
  'align-center': '<path d="M4 6h16M7 12h10M5 18h14"/>',
  'align-right': '<path d="M4 6h16M10 12h10M6 18h14"/>',
  'align-justify': '<path d="M4 6h16M4 12h16M4 18h16"/>',
  hr: '<path d="M3 12h18"/>',
  'page-break': '<path d="M4 4v5h16V4M4 20v-5h16v5M3 12h2M8 12h3M13 12h3M19 12h2"/>',
  eraser2: '<path d="M4 20h16"/>',
  folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
  scissors: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M8.1 8.1 20 20M8.1 15.9 20 4"/>',
  'file-plus': '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M12 11v6M9 14h6"/>',
  'clear-format': '<path d="M5 6V4h14v2M12 4l-3 16M15 15l5 5M20 15l-5 5"/>',
  footnote: '<path d="M4 6h10M4 10h10M4 14h6M4 20h9"/><path d="M17 4h3v6"/><path d="M16 10h5"/>',
  header: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 8h18M7 5.5h6"/>',
  footer: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 16h18M10 18.5h4"/>',
  hash: '<path d="M5 9h14M5 15h14M10 4 8 20M16 4l-2 16"/>',
};

export function icon(name: string, label?: string): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.9');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  if (label) {
    svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', label);
  } else svg.setAttribute('aria-hidden', 'true');
  svg.innerHTML = ICONS[name] ?? ICONS.file;
  return svg;
}

export function iconButton(name: string, label: string, onClick: (e: MouseEvent) => void, extra: Attrs = {}): HTMLButtonElement {
  return h('button', { type: 'button', class: 'icon-btn', 'aria-label': label, title: label, onclick: onClick, ...extra }, icon(name));
}

export function button(text: string, onClick: (e: MouseEvent) => void, opts: { icon?: string; variant?: 'primary' | 'danger' | 'ghost'; small?: boolean; attrs?: Attrs } = {}): HTMLButtonElement {
  const cls = ['btn', opts.variant ? `btn--${opts.variant}` : '', opts.small ? 'btn--sm' : ''].filter(Boolean).join(' ');
  return h('button', { type: 'button', class: cls, onclick: onClick, ...opts.attrs }, opts.icon ? icon(opts.icon) : null, h('span', null, text));
}

// ---------------------------------------------------------------- Toasts

export function toast(message: string, kind: 'info' | 'error' = 'info', ms = 4500): void {
  const host = document.querySelector('.toasts');
  if (!host) return;
  const el = h('div', { class: `toast${kind === 'error' ? ' toast--error' : ''}`, role: kind === 'error' ? 'alert' : 'status' }, icon(kind === 'error' ? 'alert' : 'info'), h('span', null, message));
  host.append(el);
  setTimeout(() => el.remove(), ms);
}

export function note(kind: 'error' | 'warn' | 'ok', text: string): HTMLElement {
  return h('div', { class: `note note--${kind}`, role: kind === 'error' ? 'alert' : undefined }, icon(kind === 'error' ? 'alert' : kind === 'warn' ? 'info' : 'check'), h('span', null, text));
}

// ---------------------------------------------------------------- Dialogs

export function openDialog(build: (close: (value?: unknown) => void) => Node): Promise<unknown> {
  return new Promise((resolve) => {
    const dlg = h('dialog');
    let result: unknown;
    const close = (value?: unknown): void => {
      result = value;
      dlg.close();
    };
    dlg.append(build(close));
    dlg.addEventListener('close', () => {
      dlg.remove();
      resolve(result);
    });
    document.body.append(dlg);
    dlg.showModal();
  });
}

export async function promptText(title: string, label: string, value = '', placeholder = '', type: 'text' | 'password' = 'text'): Promise<string | null> {
  const res = await openDialog((close) => {
    const input = h('input', { class: 'input', type, value, placeholder, 'aria-label': label, autocomplete: type === 'password' ? 'current-password' : 'off' });
    const form = h(
      'form',
      {
        class: 'dialog__body',
        method: 'dialog',
        onsubmit: (e: Event) => {
          e.preventDefault();
          close(input.value);
        },
      },
      h('h2', null, title),
      h('label', { class: 'field' }, h('span', null, label), input),
      h('div', { class: 'dialog__actions' }, button('Отмена', () => close(null), { variant: 'ghost' }), h('button', { class: 'btn btn--primary', type: 'submit' }, 'OK')),
    );
    setTimeout(() => input.focus(), 0);
    return form;
  });
  return typeof res === 'string' ? res : null;
}

export async function confirmDialog(title: string, text: string, ok = 'Да'): Promise<boolean> {
  const res = await openDialog((close) =>
    h(
      'div',
      { class: 'dialog__body' },
      h('h2', null, title),
      h('p', { style: 'margin:0;color:var(--muted)' }, text),
      h('div', { class: 'dialog__actions' }, button('Отмена', () => close(false), { variant: 'ghost' }), button(ok, () => close(true), { variant: 'primary' })),
    ),
  );
  return res === true;
}

// ---------------------------------------------------------------- Files

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} Б`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(n < 10240 ? 1 : 0)} КБ`;
  return `${(n / 1024 / 1024).toFixed(n < 10 * 1024 * 1024 ? 1 : 0)} МБ`;
}

export function download(data: Blob | Uint8Array, name: string, mime = 'application/octet-stream'): void {
  const blob = data instanceof Blob ? data : new Blob([data as BlobPart], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = h('a', { href: url, download: name, style: 'display:none' });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

export function pickFiles(accept: string, multiple = true): Promise<File[]> {
  return new Promise((resolve) => {
    const input = h('input', { type: 'file', accept, multiple, style: 'display:none' });
    input.addEventListener('change', () => {
      resolve([...(input.files ?? [])]);
      input.remove();
    });
    input.addEventListener('cancel', () => {
      resolve([]);
      input.remove();
    });
    document.body.append(input);
    input.click();
  });
}

/** Drop zone with click-to-browse, drag & drop and paste support. */
export function dropzone(opts: { accept: string; multiple?: boolean; title: string; hint: string; onFiles: (files: File[]) => void }): HTMLElement {
  const input = h('input', { type: 'file', accept: opts.accept, multiple: opts.multiple ?? true, tabindex: -1, 'aria-hidden': 'true' });
  const zone = h(
    'div',
    { class: 'dropzone', role: 'button', tabindex: 0, 'aria-label': opts.title },
    h('div', { class: 'dropzone__icon' }, icon('upload')),
    h('div', { class: 'dropzone__title' }, opts.title),
    h('div', { class: 'dropzone__hint' }, opts.hint),
    input,
  );
  const deliver = (list: FileList | File[] | null | undefined): void => {
    const files = [...(list ?? [])];
    if (files.length) opts.onFiles(opts.multiple === false ? files.slice(0, 1) : files);
  };
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      input.click();
    }
  });
  input.addEventListener('change', () => {
    deliver(input.files);
    input.value = '';
  });
  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('is-over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('is-over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('is-over');
    deliver(e.dataTransfer?.files);
  });
  return zone;
}

/** Listen for files pasted anywhere on the page while `owner` is attached. */
export function onPasteFiles(owner: HTMLElement, cb: (files: File[]) => void): void {
  const handler = (e: ClipboardEvent): void => {
    if (!owner.isConnected) {
      document.removeEventListener('paste', handler);
      return;
    }
    const target = e.target as HTMLElement | null;
    if (target?.closest('input, textarea, [contenteditable="true"]')) return;
    const files = [...(e.clipboardData?.files ?? [])];
    if (files.length) {
      e.preventDefault();
      cb(files);
    }
  };
  document.addEventListener('paste', handler);
}

export async function readFile(file: Blob): Promise<Uint8Array> {
  return new Uint8Array(await file.arrayBuffer());
}

export const baseName = (name: string): string => {
  const dot = name.lastIndexOf('.');
  return (dot > 0 ? name.slice(0, dot) : name) || 'file';
};

/** Remember small per-viewer preferences; storage may be unavailable. */
export const prefs = {
  get<T>(key: string, fallback: T): T {
    try {
      const v = localStorage.getItem(`salatic:${key}`);
      return v === null ? fallback : (JSON.parse(v) as T);
    } catch {
      return fallback;
    }
  },
  set(key: string, value: unknown): void {
    try {
      localStorage.setItem(`salatic:${key}`, JSON.stringify(value));
    } catch {
      /* storage unavailable (private mode) – ignore */
    }
  },
  remove(key: string): void {
    try {
      localStorage.removeItem(`salatic:${key}`);
    } catch {
      /* ignore */
    }
  },
};

/** Client-side navigation without a full reload. */
export function go(url: string): void {
  window.dispatchEvent(new CustomEvent('salatic:navigate', { detail: url }));
}

/** Hand a file over to another page (e.g. "open in editor" after conversion). */
let handoff: { file: File; target: string } | null = null;
export function handOff(file: File, target: string): void {
  handoff = { file, target };
  go(target);
}
export function takeHandoff(target: string): File | null {
  if (handoff?.target !== target) return null;
  const f = handoff.file;
  handoff = null;
  return f;
}
