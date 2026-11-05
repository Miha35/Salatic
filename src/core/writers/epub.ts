/**
 * Doc → EPUB 3 e-book (with an EPUB 2 NCX for older readers). The book is split
 * into chapters at the top heading level; images become package resources.
 */
import { zipSync, type Zippable } from 'fflate';
import { blockText, type Block, type Doc } from '../model';
import { escapeXml, utf8 } from '../util';
import { blocksToXhtml, CSS_FONT } from './text';
import { metafileToSvgBytes } from '../metafile';

const IMAGE_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/svg+xml': 'svg', 'image/webp': 'webp' };
const MAX_CHAPTER_CHARS = 120_000;

interface Chapter {
  title: string;
  blocks: Block[];
}

/** Split at the highest heading level that occurs at least twice (or once, past the start). */
function splitChapters(doc: Doc): Chapter[] {
  const levels = doc.blocks.filter((b) => b.type === 'heading').map((b) => (b.type === 'heading' ? b.level : 9));
  let level = 0;
  for (let l = 1; l <= 3; l++) {
    if (levels.filter((x) => x === l).length >= 2) {
      level = l;
      break;
    }
  }
  const chapters: Chapter[] = [];
  let cur: Chapter = { title: '', blocks: [] };
  let size = 0;
  const flush = (): void => {
    while (cur.blocks.length && cur.blocks[cur.blocks.length - 1].type === 'pagebreak') cur.blocks.pop();
    if (cur.blocks.length) chapters.push(cur);
  };
  for (const b of doc.blocks) {
    const startsChapter = level > 0 && b.type === 'heading' && b.level <= level;
    if ((startsChapter && cur.blocks.some((x) => x.type !== 'pagebreak')) || size > MAX_CHAPTER_CHARS) {
      flush();
      cur = { title: '', blocks: [] };
      size = 0;
    }
    if (b.type === 'pagebreak' && !cur.blocks.length) continue;
    if (b.type === 'heading' && !cur.title) cur.title = blockText(b).trim();
    cur.blocks.push(b);
    size += b.type === 'image' ? 2000 : blockText(b).length + 20;
  }
  flush();
  chapters.forEach((c, i) => {
    if (!c.title) c.title = chapters.length === 1 ? (doc.title ?? 'Текст') : `Часть ${i + 1}`;
  });
  return chapters.length ? chapters : [{ title: doc.title ?? 'Текст', blocks: [{ type: 'paragraph', runs: [] }] }];
}

function uuid(): string {
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const h = [...b].map((x) => x.toString(16).padStart(2, '0')).join('');
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
}

export function docToEpub(doc: Doc): { data: Uint8Array; warnings: string[] } {
  const warnings = new Set<string>();
  const title = doc.title?.trim() || 'Документ';
  const chapters = splitChapters(doc);
  const images: { path: string; mime: string; data: Uint8Array }[] = [];
  const imageIds = new Map<Uint8Array, string>();
  const imageSrc = (b: Extract<Block, { type: 'image' }>): string => {
    let path = imageIds.get(b.data);
    if (!path && (b.mime === 'image/emf' || b.mime === 'image/wmf')) {
      const svg = metafileToSvgBytes(b.data);
      if (svg) {
        path = `images/img${images.length + 1}.svg`;
        images.push({ path, mime: 'image/svg+xml', data: svg });
        imageIds.set(b.data, path);
      }
    }
    if (!path) {
      const ext = IMAGE_EXT[b.mime];
      if (!ext) {
        warnings.add('Некоторые изображения сохранены как есть: их формат может не поддерживаться читалкой.');
      }
      path = `images/img${images.length + 1}.${ext ?? 'bin'}`;
      images.push({ path, mime: b.mime, data: b.data });
      imageIds.set(b.data, path);
    }
    return `../${path}`;
  };
  const xhtml = (heading: string, body: string): string =>
    '<?xml version="1.0" encoding="utf-8"?>\n<!DOCTYPE html>\n' +
    '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="ru" xml:lang="ru">\n' +
    `<head><meta charset="utf-8"/><title>${escapeXml(heading)}</title><link rel="stylesheet" type="text/css" href="../style.css"/></head>\n<body>\n${body}</body>\n</html>\n`;
  const files: Zippable = {};
  files['mimetype'] = [utf8('application/epub+zip'), { level: 0 }];
  files['META-INF/container.xml'] = utf8(
    '<?xml version="1.0" encoding="UTF-8"?>\n<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>',
  );
  const chapterFiles = chapters.map((c, i) => {
    const name = `text/ch${String(i + 1).padStart(3, '0')}.xhtml`;
    files[`OEBPS/${name}`] = utf8(xhtml(c.title, blocksToXhtml(c.blocks, imageSrc)));
    return name;
  });
  const font = CSS_FONT[doc.font ?? 'serif'];
  files['OEBPS/style.css'] = utf8(
    `body{font-family:${font};line-height:1.45;margin:0 4%}\n` +
      'h1,h2,h3,h4,h5,h6{line-height:1.25;margin:1.2em 0 .5em;page-break-after:avoid}\n' +
      'p{margin:.4em 0}\nimg{max-width:100%;height:auto}\n' +
      'table{border-collapse:collapse;margin:1em 0}td,th{border:1px solid #999;padding:.2em .5em;vertical-align:top}\n' +
      'pre{white-space:pre-wrap;font-family:monospace;font-size:.9em;background:#f4f4f4;padding:.5em}\n' +
      'blockquote{margin:1em 1.5em;font-style:italic}\n.footnotes{font-size:.9em}\nhr.page-break{border:0;page-break-after:always}\n',
  );
  const toc = chapters.map((c, i) => `<li><a href="${chapterFiles[i]}">${escapeXml(c.title)}</a></li>`).join('');
  files['OEBPS/nav.xhtml'] = utf8(
    '<?xml version="1.0" encoding="utf-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="ru" xml:lang="ru">' +
      `<head><meta charset="utf-8"/><title>${escapeXml(title)}</title></head><body><nav epub:type="toc" id="toc"><h1>Содержание</h1><ol>${toc}</ol></nav></body></html>`,
  );
  const id = `urn:uuid:${uuid()}`;
  files['OEBPS/toc.ncx'] = utf8(
    '<?xml version="1.0" encoding="UTF-8"?>\n<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">' +
      `<head><meta name="dtb:uid" content="${id}"/><meta name="dtb:depth" content="1"/><meta name="dtb:totalPageCount" content="0"/><meta name="dtb:maxPageNumber" content="0"/></head>` +
      `<docTitle><text>${escapeXml(title)}</text></docTitle><navMap>` +
      chapters.map((c, i) => `<navPoint id="np${i + 1}" playOrder="${i + 1}"><navLabel><text>${escapeXml(c.title)}</text></navLabel><content src="${chapterFiles[i]}"/></navPoint>`).join('') +
      '</navMap></ncx>',
  );
  for (const img of images) files[`OEBPS/${img.path}`] = [img.data, { level: 0 }];
  const modified = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const manifest =
    '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>' +
    '<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>' +
    '<item id="css" href="style.css" media-type="text/css"/>' +
    chapterFiles.map((f, i) => `<item id="ch${i + 1}" href="${f}" media-type="application/xhtml+xml"/>`).join('') +
    images.map((img, i) => `<item id="img${i + 1}" href="${img.path}" media-type="${img.mime}"/>`).join('');
  files['OEBPS/content.opf'] = utf8(
    '<?xml version="1.0" encoding="UTF-8"?>\n<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid" xml:lang="ru">' +
      '<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">' +
      `<dc:identifier id="bookid">${id}</dc:identifier><dc:title>${escapeXml(title)}</dc:title><dc:language>ru</dc:language>` +
      `<meta property="dcterms:modified">${modified}</meta><dc:creator>Salatic</dc:creator></metadata>` +
      `<manifest>${manifest}</manifest>` +
      `<spine toc="ncx">${chapterFiles.map((_, i) => `<itemref idref="ch${i + 1}"/>`).join('')}</spine></package>`,
  );
  return { data: zipSync(files, { level: 6 }), warnings: [...warnings] };
}
