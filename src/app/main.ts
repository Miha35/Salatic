import './styles.css';
import { icon } from './ui';

type PageModule = { mount(root: HTMLElement, params: URLSearchParams): void | (() => void) };

const ROUTES: Record<string, { nav: string; title: string; load: () => Promise<PageModule> }> = {
  '/': { nav: 'convert', title: 'Конвертер файлов', load: () => import('./pages/converter') },
  '/pdf-editor': { nav: 'pdf', title: 'Редактор PDF', load: () => import('./pages/pdf-editor') },
  '/image-editor': { nav: 'image', title: 'Редактор изображений', load: () => import('./pages/image-editor') },
  '/doc-editor': { nav: 'doc', title: 'Редактор документов', load: () => import('./pages/doc-editor') },
};

const main = document.getElementById('main')!;
let unmount: (() => void) | void;
let navToken = 0;

async function render(): Promise<void> {
  const token = ++navToken;
  const path = location.pathname.replace(/\/+$/, '') || '/';
  const route = ROUTES[path] ?? ROUTES['/'];
  document.querySelectorAll<HTMLAnchorElement>('[data-nav]').forEach((a) => {
    if (a.dataset.nav === route.nav) a.setAttribute('aria-current', 'page');
    else a.removeAttribute('aria-current');
  });
  document.title = `${route.title} — Salatic`;
  let mod: PageModule;
  try {
    mod = await route.load();
  } catch {
    // A new deployment replaced hashed chunks: reload to get fresh ones
    location.reload();
    return;
  }
  if (token !== navToken) return;
  if (typeof unmount === 'function') unmount();
  main.replaceChildren();
  unmount = mod.mount(main, new URLSearchParams(location.search));
}

function navigate(url: string): void {
  history.pushState(null, '', url);
  void render();
  window.scrollTo(0, 0);
}

document.addEventListener('click', (e) => {
  const a = (e.target as HTMLElement).closest<HTMLAnchorElement>('a[data-link]');
  if (!a || e.metaKey || e.ctrlKey || e.shiftKey || e.button !== 0) return;
  e.preventDefault();
  if (a.getAttribute('href') !== location.pathname + location.search) navigate(a.getAttribute('href')!);
});
window.addEventListener('popstate', () => void render());
// Programmatic navigation from pages (see ui.go)
window.addEventListener('salatic:navigate', (e) => navigate((e as CustomEvent<string>).detail));

// Theme toggle
const themeBtn = document.querySelector<HTMLButtonElement>('.theme-toggle')!;
const isDark = (): boolean => {
  const t = document.documentElement.dataset.theme;
  return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
};
const paintTheme = (): void => {
  themeBtn.replaceChildren(icon(isDark() ? 'sun' : 'moon'));
};
themeBtn.addEventListener('click', () => {
  const next = isDark() ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  try {
    localStorage.setItem('salatic:theme', next);
  } catch {
    /* ignore */
  }
  paintTheme();
});
paintTheme();

void render();
