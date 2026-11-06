/** PDF assembly: images → PDF, and page composition for the PDF editor (merge, reorder, rotate, annotate). */
import { redactPage, type Rect, type RedactOptions } from './pdf-redact';
import { concatTransformationMatrix, degrees, drawObject, LineCapStyle, PDFDocument, popGraphicsState, pushGraphicsState, rgb, type PDFImage, type PDFPage } from 'pdf-lib';
import { FontManager, styleOf, type FontKey, type FontLoader } from './pdf-fonts';
import type { FontFamily } from './model';
import type { Subsetter } from './hb-subset';
import { ConvertError } from './util';

export interface ImageInput {
  data: Uint8Array;
  mime: 'image/png' | 'image/jpeg';
  /** Pixel size of the image. */
  width: number;
  height: number;
}

export interface ImagesToPdfOptions {
  pageSize?: 'fit' | 'a4' | 'letter';
  margin?: number; // points, for a4/letter
}

const PAGE = { a4: [595.28, 841.89], letter: [612, 792] } as const;

async function embed(pdf: PDFDocument, img: { data: Uint8Array; mime: string }): Promise<PDFImage> {
  try {
    return img.mime === 'image/png' ? await pdf.embedPng(img.data) : await pdf.embedJpg(img.data);
  } catch {
    throw new ConvertError('Не удалось встроить изображение в PDF: файл повреждён.');
  }
}

/** Where each image was drawn: page index and rectangle in PDF points (origin bottom-left). */
export interface ImagePlacement {
  page: number;
  x: number;
  y: number;
  w: number;
  h: number;
}

export async function imagesToPdf(images: ImageInput[], opts: ImagesToPdfOptions = {}): Promise<Uint8Array> {
  return (await imagesToPdfLayout(images, opts)).pdf;
}

export async function imagesToPdfLayout(images: ImageInput[], opts: ImagesToPdfOptions = {}): Promise<{ pdf: Uint8Array; placements: ImagePlacement[] }> {
  if (!images.length) throw new ConvertError('Нет изображений для создания PDF.');
  const placements: ImagePlacement[] = [];
  const pdf = await PDFDocument.create();
  pdf.setCreator('Salatic');
  pdf.setProducer('Salatic');
  for (const im of images) {
    const img = await embed(pdf, im);
    const wPt = (im.width * 72) / 96;
    const hPt = (im.height * 72) / 96;
    if (!opts.pageSize || opts.pageSize === 'fit') {
      const page = pdf.addPage([wPt, hPt]);
      page.drawImage(img, { x: 0, y: 0, width: wPt, height: hPt });
      placements.push({ page: pdf.getPageCount() - 1, x: 0, y: 0, w: wPt, h: hPt });
      continue;
    }
    const [pw0, ph0] = PAGE[opts.pageSize];
    const landscape = im.width > im.height;
    const [pw, ph] = landscape ? [ph0, pw0] : [pw0, ph0];
    const m = opts.margin ?? 28.35;
    const k = Math.min((pw - 2 * m) / wPt, (ph - 2 * m) / hPt, 1);
    const w = wPt * k;
    const h = hPt * k;
    const page = pdf.addPage([pw, ph]);
    page.drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
    placements.push({ page: pdf.getPageCount() - 1, x: (pw - w) / 2, y: (ph - h) / 2, w, h });
  }
  return { pdf: await pdf.save({ useObjectStreams: true }), placements };
}

export interface TextLayer {
  /** Page index in the base document. */
  page: number;
  /** One-page PDF with invisible text (Tesseract "text only" output). */
  pdf: Uint8Array;
  /** Canvas pixel → base page user space (y of pixels grows downwards). */
  toUser: [number, number, number, number, number, number];
  /** Size of the recognised image in pixels. */
  width: number;
  height: number;
}

/**
 * Make a PDF searchable: draw OCR text layers (invisible text) over the original
 * pages, mapped exactly onto the area that was recognised.
 */
export async function overlayTextLayers(base: Uint8Array, layers: TextLayer[]): Promise<Uint8Array> {
  let doc: PDFDocument;
  try {
    doc = await PDFDocument.load(base, { ignoreEncryption: true, updateMetadata: false });
  } catch {
    throw new ConvertError('PDF повреждён и не может быть обработан.');
  }
  if (doc.isEncrypted) throw new ConvertError('PDF зашифрован — добавить текстовый слой нельзя.', 'encrypted');
  for (const l of layers) {
    const src = await PDFDocument.load(l.pdf);
    const srcPage = src.getPage(0);
    const { width: tw, height: th } = srcPage.getSize();
    const [emb] = await doc.embedPdf(src, [0]);
    const page = doc.getPage(l.page);
    // text space (y up, tw×th) → pixels (y down) → base user space
    const s: [number, number, number, number, number, number] = [l.width / tw, 0, 0, -l.height / th, 0, l.height];
    const [a, b, c, d, e, f] = l.toUser;
    const m = [a * s[0] + c * s[1], b * s[0] + d * s[1], a * s[2] + c * s[3], b * s[2] + d * s[3], a * s[4] + c * s[5] + e, b * s[4] + d * s[5] + f];
    const name = page.node.newXObject('OcrText', emb.ref);
    page.pushOperators(pushGraphicsState(), concatTransformationMatrix(m[0], m[1], m[2], m[3], m[4], m[5]), drawObject(name), popGraphicsState());
  }
  return doc.save({ useObjectStreams: true });
}

// ---------------------------------------------------------------------------
// Page composition

/** Annotation coordinates are in PDF user space of the target page (already converted by the editor). */
export type Annotation =
  | { type: 'text'; x: number; y: number; size: number; color: [number, number, number]; text: string; rotate: number; bold?: boolean; italic?: boolean; font?: FontFamily }
  | { type: 'rect'; x: number; y: number; w: number; h: number; color: [number, number, number]; opacity: number; rotate: number; /** also remove the page's text underneath (whiteout) */ erase?: boolean }
  /** Redaction: text, graphics, image pixels and annotations underneath are removed, then the area is filled. */
  | { type: 'redact'; x: number; y: number; w: number; h: number; color: [number, number, number]; rotate: number }
  /** Remove the page's text in the area (editing existing text); `fallback` covers it if no text could be removed. */
  | { type: 'erase'; x: number; y: number; w: number; h: number; rotate: number; fallback?: [number, number, number] }
  | { type: 'image'; x: number; y: number; w: number; h: number; data: Uint8Array; mime: 'image/png' | 'image/jpeg'; rotate: number }
  | { type: 'path'; points: [number, number][]; color: [number, number, number]; width: number; opacity: number };

export interface PageSpec {
  /** Source document index, or -1 for a blank page. */
  src: number;
  index: number;
  /** Additional clockwise rotation in degrees (multiple of 90). */
  rotate: number;
  /** Size for blank pages (points). */
  size?: [number, number];
  annotations?: Annotation[];
}

/** Font of a text annotation: family (sans by default) and style. */
const textKey = (a: { font?: FontFamily; bold?: boolean; italic?: boolean }): FontKey => `${a.font ?? 'sans'}-${styleOf(!!a.bold, !!a.italic)}`;

/** Axis-aligned area (user space) of a rectangle drawn at x, y rotated by `rotate` degrees around (x, y). */
function areaOf(a: { x: number; y: number; w: number; h: number; rotate: number }): Rect {
  const rad = (a.rotate * Math.PI) / 180;
  const pts = [
    [0, 0],
    [a.w, 0],
    [a.w, a.h],
    [0, a.h],
  ].map(([x, y]) => [a.x + x * Math.cos(rad) - y * Math.sin(rad), a.y + x * Math.sin(rad) + y * Math.cos(rad)]);
  return { x0: Math.min(...pts.map((p) => p[0])), y0: Math.min(...pts.map((p) => p[1])), x1: Math.max(...pts.map((p) => p[0])), y1: Math.max(...pts.map((p) => p[1])) };
}

export interface ComposeOptions {
  /** JPEG → RGBA decoder for redacting photos (browser: createImageBitmap). */
  decodeJpeg?: RedactOptions['decodeJpeg'];
  /** DCT/JPX/JBIG2/CCITT decoder for redacting images (pdf-image-filters.ts). */
  decodeImage?: RedactOptions['decodeImage'];
  /** Collects warnings from redaction. */
  warnings?: string[];
}

export async function composePdf(sources: Uint8Array[], pages: PageSpec[], fonts?: FontLoader, subsetter?: Subsetter, opts: ComposeOptions = {}): Promise<Uint8Array> {
  if (!pages.length) throw new ConvertError('В документе не осталось страниц.');
  const out = await PDFDocument.create();
  // Text annotations: fonts (with per-character fallback) are prepared and subset up front
  const texts = pages.flatMap((p) => (p.annotations ?? []).filter((a): a is Extract<Annotation, { type: 'text' }> => a.type === 'text'));
  let fm: FontManager | null = null;
  const segmentsOf = new Map<string, { key: FontKey; text: string }[]>();
  if (texts.length) {
    if (!fonts) throw new ConvertError('Шрифты недоступны для добавления текста.');
    out.registerFontkit((await import('@pdf-lib/fontkit')).default);
    fm = new FontManager(fonts, subsetter);
    for (const a of texts) {
      const primary = textKey(a);
      await fm.load([primary]);
      const need = new Set<FontKey>();
      for (const ch of a.text) for (const k of fm.fallbacksFor(primary, ch.codePointAt(0)!)) need.add(k);
      await fm.load(need);
      for (const line of a.text.split('\n')) {
        const segs: { key: FontKey; text: string }[] = [];
        for (const ch of line) {
          const key = fm.resolve(primary, ch.codePointAt(0)!);
          const c = key ? ch : '?';
          const k = key ?? primary;
          fm.markUsed(k, c);
          if (segs.length && segs[segs.length - 1].key === k) segs[segs.length - 1].text += c;
          else segs.push({ key: k, text: c });
        }
        segmentsOf.set(`${primary}|${line}`, segs);
      }
    }
    await fm.embedAll(out);
  }
  out.setCreator('Salatic');
  out.setProducer('Salatic');
  const loaded: PDFDocument[] = [];
  for (const s of sources) {
    let doc: PDFDocument;
    try {
      doc = await PDFDocument.load(s, { ignoreEncryption: true, updateMetadata: false });
    } catch {
      throw new ConvertError('Один из PDF-файлов повреждён и не может быть обработан.');
    }
    if (doc.isEncrypted) {
      // Permission-only protection (empty user password) is removed transparently
      const { decryptPdf } = await import('./pdf-decrypt');
      try {
        doc = await PDFDocument.load(await decryptPdf(s, ''), { updateMetadata: false });
      } catch (e) {
        if (e instanceof ConvertError && e.code === 'password_required') throw new ConvertError('PDF защищён паролем: откройте его заново и введите пароль.', 'password_required');
        throw e;
      }
    }
    loaded.push(doc);
  }
  // a redacted file does not carry the title: it often names what was hidden ("Contract with …")
  const redacting = pages.some((p) => p.annotations?.some((a) => a.type === 'redact'));
  if (loaded.length === 1 && !out.getTitle() && !redacting) {
    const t = loaded[0].getTitle();
    if (t) out.setTitle(t);
  }
  // Copy all needed pages per source in one call (shared resources are deduplicated by pdf-lib)
  const copies = new Map<number, PDFPage[]>();
  for (let s = 0; s < loaded.length; s++) {
    const idx = pages.filter((p) => p.src === s).map((p) => p.index);
    if (idx.length) copies.set(s, await out.copyPages(loaded[s], idx));
  }
  const cursor = new Map<number, number>();
  const imageCache = new Map<Uint8Array, PDFImage>();

  for (const spec of pages) {
    let page: PDFPage;
    if (spec.src < 0) page = out.addPage(spec.size ?? [595.28, 841.89]);
    else {
      const list = copies.get(spec.src)!;
      const k = cursor.get(spec.src) ?? 0;
      cursor.set(spec.src, k + 1);
      page = out.addPage(list[k]);
    }
    if (spec.rotate) page.setRotation(degrees((((page.getRotation().angle + spec.rotate) % 360) + 360) % 360));
    // Content removal first: redaction areas, then text under whiteout / edited text
    const anns = spec.annotations ?? [];
    const redactAreas = anns.filter((a): a is Extract<Annotation, { type: 'redact' }> => a.type === 'redact').map(areaOf);
    const eraseAnns = anns.filter((a): a is Extract<Annotation, { type: 'erase' }> => a.type === 'erase');
    const eraseAreas = [...eraseAnns, ...anns.filter((a): a is Extract<Annotation, { type: 'rect' }> => a.type === 'rect' && !!a.erase)].map(areaOf);
    const removedIn = new Map<Extract<Annotation, { type: 'erase' }>, number>();
    if (spec.src >= 0 && redactAreas.length) {
      const r = await redactPage(out, page, redactAreas, { redact: true, decodeJpeg: opts.decodeJpeg, decodeImage: opts.decodeImage });
      opts.warnings?.push(...r.warnings);
    }
    if (spec.src >= 0 && eraseAreas.length) {
      // one pass per edited-text area, so a failed removal can fall back to covering
      for (const a of eraseAnns) {
        const r = await redactPage(out, page, [areaOf(a)]);
        removedIn.set(a, r.glyphs);
        opts.warnings?.push(...r.warnings);
      }
      const rest = anns.filter((a): a is Extract<Annotation, { type: 'rect' }> => a.type === 'rect' && !!a.erase).map(areaOf);
      if (rest.length) {
        const r = await redactPage(out, page, rest);
        opts.warnings?.push(...r.warnings);
      }
    }
    for (const a of anns) {
      switch (a.type) {
        case 'redact':
          page.drawRectangle({ x: a.x, y: a.y, width: a.w, height: a.h, color: rgb(...a.color), rotate: degrees(a.rotate) });
          break;
        case 'erase':
          if (a.fallback && !removedIn.get(a)) page.drawRectangle({ x: a.x, y: a.y, width: a.w, height: a.h, color: rgb(...a.fallback), rotate: degrees(a.rotate) });
          break;
        case 'text': {
          const rad = (a.rotate * Math.PI) / 180;
          a.text.split('\n').forEach((line, i) => {
            // Successive lines move "down" in the rotated text frame; segments advance along it
            const dy = i * a.size * 1.25;
            let adv = 0;
            for (const seg of segmentsOf.get(`${textKey(a)}|${line}`) ?? []) {
              page.drawText(seg.text, {
                x: a.x + Math.sin(rad) * dy + Math.cos(rad) * adv,
                y: a.y - Math.cos(rad) * dy + Math.sin(rad) * adv,
                size: a.size,
                font: fm!.pdfFont(seg.key),
                color: rgb(...a.color),
                rotate: degrees(a.rotate),
              });
              adv += fm!.width(seg.key, seg.text, a.size);
            }
          });
          break;
        }
        case 'rect':
          page.drawRectangle({ x: a.x, y: a.y, width: a.w, height: a.h, color: rgb(...a.color), opacity: a.opacity, rotate: degrees(a.rotate) });
          break;
        case 'image': {
          let img = imageCache.get(a.data);
          if (!img) {
            img = await embed(out, a);
            imageCache.set(a.data, img);
          }
          page.drawImage(img, { x: a.x, y: a.y, width: a.w, height: a.h, rotate: degrees(a.rotate) });
          break;
        }
        case 'path':
          for (let i = 1; i < a.points.length; i++) {
            const [x1, y1] = a.points[i - 1];
            const [x2, y2] = a.points[i];
            page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: a.width, color: rgb(...a.color), opacity: a.opacity, lineCap: LineCapStyle.Round });
          }
          break;
      }
    }
  }
  return out.save({ useObjectStreams: true });
}

/** Page count and sizes – used by the editor before rendering. */
export async function pdfInfo(data: Uint8Array): Promise<{ pages: number; encrypted: boolean }> {
  const doc = await PDFDocument.load(data, { ignoreEncryption: true, updateMetadata: false });
  return { pages: doc.getPageCount(), encrypted: doc.isEncrypted };
}
