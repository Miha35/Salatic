/** DOCX (WordprocessingML) → Doc. */
import { unzipSync } from 'fflate';
import type { Align, Block, Doc, ListItem, TableCell, TableRow, TextRun } from '../model';
import { normalizeRuns } from '../model';
import { child, childrenNamed, elements, find, parseXml, rootElement, textContent, type XNode } from '../xml';
import { ConvertError, imageSize, sniffImageMime, unzipOffice } from '../util';

type RStyle = Omit<TextRun, 'text'>;

interface StyleDef {
  name: string;
  basedOn?: string;
  rPr?: XNode;
  pPr?: XNode;
}

interface Rel {
  target: string;
  external: boolean;
}

const decoder = new TextDecoder('utf-8');

function onOff(el: XNode | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = el.attrs['w:val'];
  return v === undefined || !['0', 'false', 'off', 'none'].includes(v);
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

class DocxReader {
  private files: Record<string, Uint8Array>;
  private rels = new Map<string, Rel>();
  private styles = new Map<string, StyleDef>();
  private numFormats = new Map<string, Map<number, string>>(); // numId → ilvl → numFmt
  private mainPath = 'word/document.xml';
  private fieldStack: { phase: 'instr' | 'result'; instr: string; link?: string }[] = [];
  warnings: string[] = [];

  constructor(bytes: Uint8Array) {
    try {
      this.files = unzipOffice(bytes, unzipSync);
    } catch (e) {
      if (e instanceof ConvertError) throw e;
      throw new ConvertError('Файл DOCX повреждён или не является архивом Office.');
    }
    const rootRels = this.xml('_rels/.rels');
    if (rootRels) {
      for (const r of childrenNamed(rootElement(rootRels), 'Relationship')) {
        if (r.attrs.Type?.endsWith('/officeDocument')) this.mainPath = r.attrs.Target.replace(/^\//, '');
      }
    }
    if (!this.files[this.mainPath]) throw new ConvertError('В архиве нет основного документа Word (word/document.xml).');
    this.loadRels();
    this.loadStyles();
    this.loadNumbering();
  }

  private xml(path: string): XNode | undefined {
    const f = this.files[path];
    return f ? parseXml(decoder.decode(f)) : undefined;
  }

  private loadRels(): void {
    const dir = this.mainPath.slice(0, this.mainPath.lastIndexOf('/') + 1);
    const name = this.mainPath.slice(dir.length);
    const rels = this.xml(`${dir}_rels/${name}.rels`);
    for (const r of childrenNamed(rootElement(rels ?? { name: '', attrs: {}, children: [] }), 'Relationship')) {
      const external = r.attrs.TargetMode === 'External';
      this.rels.set(r.attrs.Id, { target: external ? r.attrs.Target : resolvePath(this.mainPath, r.attrs.Target), external });
    }
  }

  private loadStyles(): void {
    const styles = this.xml(resolvePath(this.mainPath, 'styles.xml'));
    const root = styles && rootElement(styles);
    for (const s of childrenNamed(root, 'w:style')) {
      const id = s.attrs['w:styleId'];
      if (!id) continue;
      this.styles.set(id, {
        name: (child(s, 'w:name')?.attrs['w:val'] ?? id).toLowerCase(),
        basedOn: child(s, 'w:basedOn')?.attrs['w:val'],
        rPr: child(s, 'w:rPr'),
        pPr: child(s, 'w:pPr'),
      });
    }
  }

  private loadNumbering(): void {
    const num = this.xml(resolvePath(this.mainPath, 'numbering.xml'));
    const root = num && rootElement(num);
    if (!root) return;
    const abstract = new Map<string, Map<number, string>>();
    for (const a of childrenNamed(root, 'w:abstractNum')) {
      const lvls = new Map<number, string>();
      for (const l of childrenNamed(a, 'w:lvl')) lvls.set(+l.attrs['w:ilvl'], child(l, 'w:numFmt')?.attrs['w:val'] ?? 'decimal');
      abstract.set(a.attrs['w:abstractNumId'], lvls);
    }
    for (const n of childrenNamed(root, 'w:num')) {
      const absId = child(n, 'w:abstractNumId')?.attrs['w:val'];
      const lvls = new Map(abstract.get(absId ?? '') ?? []);
      for (const o of childrenNamed(n, 'w:lvlOverride')) {
        const fmt = find(o, 'w:numFmt')?.attrs['w:val'];
        if (fmt) lvls.set(+o.attrs['w:ilvl'], fmt);
      }
      this.numFormats.set(n.attrs['w:numId'], lvls);
    }
  }

  private styleChain(id: string | undefined): StyleDef[] {
    const chain: StyleDef[] = [];
    const seen = new Set<string>();
    while (id && !seen.has(id)) {
      seen.add(id);
      const s = this.styles.get(id);
      if (!s) break;
      chain.unshift(s);
      id = s.basedOn;
    }
    return chain;
  }

  private applyRPr(style: RStyle, rPr: XNode | undefined): RStyle {
    if (!rPr) return style;
    const s = { ...style };
    const b = onOff(child(rPr, 'w:b'));
    if (b !== undefined) s.bold = b;
    const i = onOff(child(rPr, 'w:i'));
    if (i !== undefined) s.italic = i;
    const u = child(rPr, 'w:u');
    if (u) s.underline = u.attrs['w:val'] !== 'none';
    const st = onOff(child(rPr, 'w:strike')) ?? onOff(child(rPr, 'w:dstrike'));
    if (st !== undefined) s.strike = st;
    const va = child(rPr, 'w:vertAlign')?.attrs['w:val'];
    if (va === 'superscript') s.sup = true;
    else if (va === 'subscript') s.sub = true;
    const color = child(rPr, 'w:color')?.attrs['w:val'];
    if (color && /^[0-9a-fA-F]{6}$/.test(color) && color !== '000000') s.color = color.toUpperCase();
    else if (color === 'auto' || color === '000000') delete s.color;
    const fonts = child(rPr, 'w:rFonts');
    const face = fonts?.attrs['w:ascii'] ?? fonts?.attrs['w:hAnsi'] ?? '';
    if (/courier|consolas|mono|menlo/i.test(face)) s.code = true;
    if (onOff(child(rPr, 'w:vanish'))) (s as RStyle & { hidden?: boolean }).hidden = true;
    return s;
  }

  private styleRStyle(styleId: string | undefined): RStyle {
    let s: RStyle = {};
    for (const def of this.styleChain(styleId)) s = this.applyRPr(s, def.rPr);
    return s;
  }

  private headingLevel(styleId: string | undefined, pPr: XNode | undefined): number {
    const outline = child(pPr, 'w:outlineLvl')?.attrs['w:val'];
    if (outline !== undefined && +outline < 6) return +outline + 1;
    for (const def of this.styleChain(styleId).reverse()) {
      const m = /^heading\s*(\d)$/.exec(def.name);
      if (m) return Math.min(6, Math.max(1, +m[1]));
      if (def.name === 'title') return 1;
      if (def.name === 'subtitle') return 2;
      const ol = child(def.pPr, 'w:outlineLvl')?.attrs['w:val'];
      if (ol !== undefined && +ol < 6) return +ol + 1;
    }
    return 0;
  }

  private numInfo(styleId: string | undefined, pPr: XNode | undefined): { numId: string; ilvl: number } | null {
    let numPr = child(pPr, 'w:numPr');
    if (!numPr) {
      for (const def of this.styleChain(styleId).reverse()) {
        numPr = child(def.pPr, 'w:numPr');
        if (numPr) break;
      }
    }
    const numId = child(numPr, 'w:numId')?.attrs['w:val'];
    if (!numId || numId === '0') return null;
    return { numId, ilvl: +(child(numPr, 'w:ilvl')?.attrs['w:val'] ?? 0) };
  }

  private image(el: XNode): Block | null {
    const blip = find(el, 'a:blip') ?? find(el, 'v:imagedata');
    const rid = blip?.attrs['r:embed'] ?? blip?.attrs['r:id'];
    const rel = rid ? this.rels.get(rid) : undefined;
    const data = rel && !rel.external ? this.files[rel.target] : undefined;
    if (!data) return null;
    const mime = sniffImageMime(data);
    if (!mime) {
      this.warnings.push('Некоторые изображения (EMF/WMF/TIFF) пропущены: формат не поддерживается.');
      return null;
    }
    const extent = find(el, 'wp:extent');
    let width = extent ? +extent.attrs.cx / 9525 : 0;
    let height = extent ? +extent.attrs.cy / 9525 : 0;
    if (!width || !height) {
      const style = find(el, 'v:shape')?.attrs.style ?? '';
      const w = /width:([\d.]+)pt/.exec(style);
      const h = /height:([\d.]+)pt/.exec(style);
      if (w && h) {
        width = (+w[1] * 96) / 72;
        height = (+h[1] * 96) / 72;
      } else {
        const sz = imageSize(data) ?? { width: 300, height: 200 };
        width = sz.width;
        height = sz.height;
      }
    }
    const alt = find(el, 'wp:docPr')?.attrs.descr || undefined;
    return { type: 'image', data, mime, width: Math.round(width), height: Math.round(height), alt };
  }

  /** Collect runs of a paragraph-level container. Images and page breaks are emitted as separators. */
  private inline(nodes: (XNode | string)[], base: RStyle, link: string | undefined, out: (TextRun | Block)[], extra: Block[]): void {
    for (const n of nodes) {
      if (typeof n === 'string') continue;
      switch (n.name) {
        case 'w:r':
          this.runContent(n, base, link, out, extra);
          break;
        case 'w:hyperlink': {
          const rel = n.attrs['r:id'] ? this.rels.get(n.attrs['r:id']) : undefined;
          const href = rel?.external && /^(https?:|mailto:)/i.test(rel.target) ? rel.target : link;
          this.inline(n.children, base, href, out, extra);
          break;
        }
        case 'w:fldSimple': {
          const m = /HYPERLINK\s+"([^"]+)"/.exec(n.attrs['w:instr'] ?? '');
          this.inline(n.children, base, m ? m[1] : link, out, extra);
          break;
        }
        case 'w:ins':
        case 'w:moveTo':
        case 'w:smartTag':
        case 'w:customXml':
        case 'w:dir':
        case 'w:bdo':
          this.inline(n.children, base, link, out, extra);
          break;
        case 'w:sdt':
          this.inline(child(n, 'w:sdtContent')?.children ?? [], base, link, out, extra);
          break;
        default:
          break; // w:del, w:moveFrom, bookmarks, proofErr, etc.
      }
    }
  }

  private runContent(r: XNode, base: RStyle, link: string | undefined, out: (TextRun | Block)[], extra: Block[]): void {
    const rPr = child(r, 'w:rPr');
    let style = this.applyRPr({ ...base, ...this.styleRStyle(child(rPr, 'w:rStyle')?.attrs['w:val']) }, rPr);
    if ((style as { hidden?: boolean }).hidden) return;
    const field = this.fieldStack[this.fieldStack.length - 1];
    const fieldLink = this.fieldStack.find((f) => f.link)?.link;
    if (link ?? fieldLink) style = { ...style, link: link ?? fieldLink };
    const push = (text: string): void => {
      if (field?.phase === 'instr') return;
      out.push({ text, ...style });
    };
    for (const c of r.children) {
      if (typeof c === 'string') continue;
      switch (c.name) {
        case 'w:t':
          push(textContent(c));
          break;
        case 'w:instrText':
          if (field?.phase === 'instr') field.instr += textContent(c);
          break;
        case 'w:fldChar': {
          const t = c.attrs['w:fldCharType'];
          if (t === 'begin') this.fieldStack.push({ phase: 'instr', instr: '' });
          else if (t === 'separate' && field) {
            field.phase = 'result';
            const m = /HYPERLINK\s+"([^"]+)"/.exec(field.instr);
            if (m && /^(https?:|mailto:)/i.test(m[1])) field.link = m[1];
          } else if (t === 'end') this.fieldStack.pop();
          break;
        }
        case 'w:tab':
        case 'w:ptab':
          push('\t');
          break;
        case 'w:br':
        case 'w:cr':
          if (c.attrs['w:type'] === 'page') out.push({ type: 'pagebreak' });
          else push('\n');
          break;
        case 'w:noBreakHyphen':
          push('-');
          break;
        case 'w:sym': {
          const code = parseInt(c.attrs['w:char'] ?? '', 16);
          if (code >= 0xf020 && code <= 0xf0ff && /symbol|wingdings/i.test(c.attrs['w:font'] ?? '')) push('•');
          else if (Number.isFinite(code)) push(String.fromCharCode(code));
          break;
        }
        case 'w:drawing':
        case 'w:pict': {
          const img = this.image(c);
          if (img) out.push(img);
          this.textBoxes(c, extra);
          break;
        }
        case 'mc:AlternateContent': {
          const choice = child(c, 'mc:Choice');
          const fallback = child(c, 'mc:Fallback');
          const drawing = (choice && find(choice, 'w:drawing')) ?? (fallback && (find(fallback, 'w:pict') ?? find(fallback, 'w:drawing')));
          if (drawing) {
            const img = this.image(drawing) ?? (fallback ? this.image(fallback) : null);
            if (img) out.push(img);
            this.textBoxes(drawing, extra);
          }
          break;
        }
        default:
          break;
      }
    }
  }

  private textBoxes(el: XNode, extra: Block[]): void {
    const box = find(el, 'w:txbxContent');
    if (box) extra.push(...this.body(elements(box)));
  }

  private paragraph(p: XNode): { blocks: Block[]; list?: { numId: string; ilvl: number; ordered: boolean } } {
    const pPr = child(p, 'w:pPr');
    const styleId = child(pPr, 'w:pStyle')?.attrs['w:val'];
    let base: RStyle = {};
    for (const def of this.styleChain(styleId)) base = this.applyRPr(base, def.rPr);
    // Paragraph mark properties (w:pPr/w:rPr) apply to the mark only, not the runs → ignored.
    const level = this.headingLevel(styleId, pPr);
    if (level) delete base.bold; // headings are bold by definition; avoid redundant flags
    const jcRaw = child(pPr, 'w:jc')?.attrs['w:val'];
    const align: Align | undefined =
      jcRaw === 'center' ? 'center' : jcRaw === 'right' || jcRaw === 'end' ? 'right' : jcRaw === 'both' || jcRaw === 'distribute' ? 'justify' : undefined;
    const items: (TextRun | Block)[] = [];
    const extra: Block[] = [];
    this.inline(p.children, base, undefined, items, extra);
    const blocks: Block[] = [];
    if (onOff(child(pPr, 'w:pageBreakBefore'))) blocks.push({ type: 'pagebreak' });
    const num = level ? null : this.numInfo(styleId, pPr);
    const isQuote = this.styleChain(styleId).some((d) => d.name.includes('quote'));
    let runs: TextRun[] = [];
    let emittedText = false;
    const flush = (force: boolean): void => {
      const r = normalizeRuns(runs);
      runs = [];
      const hasText = r.some((x) => x.text.trim());
      if (!hasText && !force) return;
      let b: Block;
      if (level && hasText) b = { type: 'heading', level: level as 1, runs: r };
      else b = { type: 'paragraph', runs: hasText ? r : [] };
      if (align) (b as { align?: Align }).align = align;
      blocks.push(isQuote && hasText ? { type: 'quote', blocks: [b] } : b);
      emittedText = true;
    };
    for (const it of items) {
      if ('type' in it) {
        flush(false);
        blocks.push(it);
      } else runs.push(it);
    }
    flush(!emittedText && !blocks.some((b) => b.type === 'image' || b.type === 'pagebreak'));
    blocks.push(...extra);
    if (num) {
      const fmt = this.numFormats.get(num.numId)?.get(num.ilvl) ?? 'bullet';
      return { blocks, list: { ...num, ordered: fmt !== 'bullet' && fmt !== 'none' } };
    }
    return { blocks };
  }

  private table(tbl: XNode): Block {
    const rows: TableRow[] = [];
    for (const tr of childrenNamed(tbl, 'w:tr')) {
      const header = onOff(child(child(tr, 'w:trPr'), 'w:tblHeader')) ?? false;
      const cells: TableCell[] = [];
      const tcs = [...childrenNamed(tr, 'w:tc'), ...childrenNamed(tr, 'w:sdt').flatMap((s) => childrenNamed(child(s, 'w:sdtContent'), 'w:tc'))];
      for (const tc of tcs) {
        const tcPr = child(tc, 'w:tcPr');
        const span = +(child(tcPr, 'w:gridSpan')?.attrs['w:val'] ?? 1);
        const vMerge = child(tcPr, 'w:vMerge');
        const continued = vMerge && vMerge.attrs['w:val'] !== 'restart';
        const blocks = continued ? [] : this.body(elements(tc).filter((e) => e.name !== 'w:tcPr'));
        const cell: TableCell = { blocks: blocks.length ? blocks : [{ type: 'paragraph', runs: [] }] };
        if (header) cell.header = true;
        if (span > 1) cell.colspan = span;
        cells.push(cell);
      }
      if (cells.length) rows.push({ cells });
    }
    return { type: 'table', rows };
  }

  body(nodes: XNode[]): Block[] {
    const out: Block[] = [];
    // stack of open lists by level
    let stack: { list: Extract<Block, { type: 'list' }>; numId: string }[] = [];
    const endLists = (): void => {
      stack = [];
    };
    const visit = (n: XNode): void => {
      switch (n.name) {
        case 'w:p': {
          const { blocks, list } = this.paragraph(n);
          if (!list) {
            endLists();
            out.push(...blocks);
            return;
          }
          const level = Math.min(list.ilvl, 8);
          if (stack.length && level === 0 && stack[0].numId !== list.numId) endLists();
          while (stack.length > level + 1) stack.pop();
          if (stack.length === level + 1 && stack[level].list.ordered !== list.ordered) stack.pop();
          while (stack.length < level + 1) {
            const fresh: Extract<Block, { type: 'list' }> = { type: 'list', ordered: list.ordered, items: [] };
            if (!stack.length) out.push(fresh);
            else {
              const parent = stack[stack.length - 1].list;
              if (!parent.items.length) parent.items.push({ blocks: [] });
              parent.items[parent.items.length - 1].blocks.push(fresh);
            }
            stack.push({ list: fresh, numId: list.numId });
          }
          const item: ListItem = { blocks: blocks.length ? blocks : [{ type: 'paragraph', runs: [] }] };
          stack[level].list.items.push(item);
          return;
        }
        case 'w:tbl':
          endLists();
          out.push(this.table(n));
          return;
        case 'w:sdt':
          for (const c of elements(child(n, 'w:sdtContent'))) visit(c);
          return;
        case 'w:customXml':
        case 'w:ins':
          for (const c of elements(n)) visit(c);
          return;
        default:
          return; // sectPr, bookmarks, etc.
      }
    };
    for (const n of nodes) visit(n);
    return out;
  }

  read(): Doc {
    const doc = this.xml(this.mainPath)!;
    const body = find(doc, 'w:body');
    if (!body) throw new ConvertError('Документ Word не содержит тела (w:body).');
    const blocks = this.body(elements(body));
    let title: string | undefined;
    const core = this.xml('docProps/core.xml');
    if (core) title = textContent(find(core, 'dc:title')).trim() || undefined;
    return { title, blocks };
  }
}

export function docxToDoc(bytes: Uint8Array): { doc: Doc; warnings: string[] } {
  const r = new DocxReader(bytes);
  const doc = r.read();
  return { doc, warnings: [...new Set(r.warnings)] };
}
