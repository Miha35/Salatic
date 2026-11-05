/** PDF assembly: images → PDF, and page composition for the PDF editor (merge, reorder, rotate, annotate). */
import { degrees, LineCapStyle, PDFDocument, rgb, type PDFFont, type PDFImage, type PDFPage } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import type { FontLoader } from './writers/pdf';
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

export async function imagesToPdf(images: ImageInput[], opts: ImagesToPdfOptions = {}): Promise<Uint8Array> {
  if (!images.length) throw new ConvertError('Нет изображений для создания PDF.');
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
  }
  return pdf.save({ useObjectStreams: true });
}

// ---------------------------------------------------------------------------
// Page composition

/** Annotation coordinates are in PDF user space of the target page (already converted by the editor). */
export type Annotation =
  | { type: 'text'; x: number; y: number; size: number; color: [number, number, number]; text: string; rotate: number; bold?: boolean }
  | { type: 'rect'; x: number; y: number; w: number; h: number; color: [number, number, number]; opacity: number; rotate: number }
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

export async function composePdf(sources: Uint8Array[], pages: PageSpec[], fonts?: FontLoader): Promise<Uint8Array> {
  if (!pages.length) throw new ConvertError('В документе не осталось страниц.');
  const out = await PDFDocument.create();
  out.registerFontkit(fontkit);
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
    if (doc.isEncrypted) throw new ConvertError('PDF зашифрован (защищён паролем или ограничениями). Редактирование такого файла невозможно.', 'encrypted');
    loaded.push(doc);
  }
  if (loaded.length === 1 && !out.getTitle()) {
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
  let font: PDFFont | null = null;
  let boldFont: PDFFont | null = null;
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
    for (const a of spec.annotations ?? []) {
      switch (a.type) {
        case 'text': {
          if (!fonts) throw new ConvertError('Шрифты недоступны для добавления текста.');
          const f = a.bold ? (boldFont ??= await out.embedFont(await fonts('bold'), { subset: true })) : (font ??= await out.embedFont(await fonts('regular'), { subset: true }));
          const lines = a.text.split('\n');
          lines.forEach((line, i) => {
            // Successive lines move "down" in the rotated text frame
            const dy = i * a.size * 1.25;
            const rad = (a.rotate * Math.PI) / 180;
            page.drawText(line, {
              x: a.x + Math.sin(rad) * dy,
              y: a.y - Math.cos(rad) * dy,
              size: a.size,
              font: f,
              color: rgb(...a.color),
              rotate: degrees(a.rotate),
            });
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
