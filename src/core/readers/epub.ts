/** EPUB 2/3 e-book → Doc: the spine's XHTML documents in reading order. */
import { unzipSync } from 'fflate';
import type { Block, Doc } from '../model';
import { find, findAll, parseXml, textContent } from '../xml';
import { ConvertError, sniffImageMime, unzipOffice } from '../util';
import { htmlToDoc } from './html';

const decoder = new TextDecoder('utf-8');
/** Font obfuscation is not DRM: the text stays readable. */
const FONT_OBFUSCATION = new Set(['http://www.idpf.org/2008/embedding', 'http://ns.adobe.com/pdf/enc#RC']);

/** Resolve `href` relative to the directory of `base` inside the package (handles ./, ../ and %-escapes). */
function resolve(base: string, href: string): string {
  const clean = href.split('#')[0].split('?')[0];
  let decoded = clean;
  try {
    decoded = decodeURIComponent(clean);
  } catch {
    /* keep as is */
  }
  if (decoded.startsWith('/')) return decoded.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of decoded.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

export function epubToDoc(bytes: Uint8Array): { doc: Doc; warnings: string[] } {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipOffice(bytes, unzipSync);
  } catch (e) {
    if (e instanceof ConvertError) throw e;
    throw new ConvertError('Файл EPUB повреждён.');
  }
  const warnings = new Set<string>();
  const encryption = files['META-INF/encryption.xml'];
  if (encryption) {
    const tree = parseXml(decoder.decode(encryption));
    const encMethods = [...findAll(tree, 'EncryptionMethod'), ...findAll(tree, 'enc:EncryptionMethod')].map((m) => m.attrs.Algorithm ?? '');
    if (encMethods.some((a) => a && !FONT_OBFUSCATION.has(a))) throw new ConvertError('Книга защищена DRM: её содержимое зашифровано и не может быть преобразовано.', 'password');
  }
  const container = files['META-INF/container.xml'];
  const rootPath = container ? find(parseXml(decoder.decode(container)), 'rootfile')?.attrs['full-path'] : Object.keys(files).find((p) => p.endsWith('.opf'));
  const opfBytes = rootPath ? files[rootPath] : undefined;
  if (!rootPath || !opfBytes) throw new ConvertError('В EPUB не найден файл описания книги (OPF).');
  const opf = parseXml(decoder.decode(opfBytes));
  const local = (n: string): string => n.replace(/^[a-z]+:/i, '');
  const all = (name: string) => findAll(opf, name).concat(findAll(opf, `opf:${name}`));
  const manifest = new Map<string, { href: string; type: string; props: string }>();
  for (const item of all('item')) manifest.set(item.attrs.id, { href: resolve(rootPath, item.attrs.href ?? ''), type: item.attrs['media-type'] ?? '', props: item.attrs.properties ?? '' });
  const titleNode = findAll(opf, 'dc:title')[0] ?? all('title')[0];
  const title = textContent(titleNode).trim() || undefined;
  const spine = (find(opf, 'spine') ?? find(opf, 'opf:spine'))?.children.filter((c): c is Exclude<typeof c, string> => typeof c !== 'string' && local(c.name) === 'itemref') ?? [];
  const resolver =
    (base: string) =>
    (src: string): { mime: string; data: Uint8Array } | null => {
      if (/^[a-z]+:/i.test(src)) return null;
      const path = resolve(base, src);
      const data = files[path];
      if (!data) return null;
      const mime = sniffImageMime(data) ?? [...manifest.values()].find((m) => m.href === path)?.type ?? '';
      return mime.startsWith('image/') ? { mime, data } : null;
    };
  const blocks: Block[] = [];
  let font: Doc['font'];
  let size: number | undefined;
  for (const ref of spine) {
    const item = manifest.get(ref.attrs.idref);
    if (!item || item.props.includes('nav')) continue;
    if (!/x?html|xml/.test(item.type) && !/\.x?html?$/i.test(item.href)) continue;
    const data = files[item.href];
    if (!data) {
      warnings.add('Некоторые главы книги отсутствуют в архиве и пропущены.');
      continue;
    }
    let html = decoder.decode(data);
    // SVG-wrapped pictures (typical for covers) become plain images
    html = html.replace(/<svg\b[^>]*>[\s\S]*?<image\b([^>]*?)\/?>[\s\S]*?<\/svg>/gi, (_m, attrs: string) => {
      const href = /(?:xlink:)?href\s*=\s*["']([^"']+)["']/i.exec(attrs)?.[1];
      return href ? `<img src="${href}"/>` : '';
    });
    const chapter = htmlToDoc(html, { resolveImage: resolver(item.href) });
    if (!chapter.blocks.length) continue;
    if (blocks.length) blocks.push({ type: 'pagebreak' });
    blocks.push(...chapter.blocks);
    font ??= chapter.font;
    size ??= chapter.size;
  }
  if (!blocks.length) throw new ConvertError('В книге не найдено текста.');
  const doc: Doc = { title, blocks };
  if (font) doc.font = font;
  if (size) doc.size = size;
  return { doc, warnings: [...warnings] };
}
