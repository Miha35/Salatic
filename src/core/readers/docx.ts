/** DOCX (WordprocessingML) → Doc. */
import { unzipSync } from 'fflate';
import type { Align, Block, Doc, FontFamily, ListItem, PageSetup, ParaFormat, TableCell, TableRow, TextRun } from '../model';
import { familyOf, normalizeRuns, normalizeSections, trimRuns } from '../model';
import { child, childrenNamed, elements, find, findAll, findBy, parseXml, rootElement, textContent, type XNode } from '../xml';
import { ConvertError, imageSize, sniffImageMime, unzipOffice } from '../util';
import { chartToSvg } from '../chart';

type RStyle = Omit<TextRun, 'text'>;

interface StyleDef {
  type: string;
  isDefault: boolean;
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

const HIGHLIGHT: Record<string, string> = {
  yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080',
  darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
};

const twip = (v: string | undefined): number | undefined => (v === undefined || v === '' || !Number.isFinite(+v) ? undefined : +v / 20);

class DocxReader {
  private files: Record<string, Uint8Array>;
  private rels = new Map<string, Rel>();
  private themeFonts = { major: '', minor: '' };
  private defFont: FontFamily = 'calibri';
  private defSize = 11;
  private defPara: ParaFormat = {};
  private footnotes = new Map<string, XNode>();
  private notesPath = '';
  private styles = new Map<string, StyleDef>();
  private numFormats = new Map<string, Map<number, string>>(); // numId → ilvl → numFmt
  private mainPath = 'word/document.xml';
  private fieldStack: { phase: 'instr' | 'result'; instr: string; link?: string; page?: 'page' | 'pages'; emitted?: boolean }[] = [];
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
    this.rels = this.loadRels(this.mainPath);
    this.loadTheme();
    this.loadStyles();
    this.loadNumbering();
    this.loadDefaults();
  }

  private xml(path: string): XNode | undefined {
    const f = this.files[path];
    return f ? parseXml(decoder.decode(f)) : undefined;
  }

  private loadRels(partPath: string): Map<string, Rel> {
    const map = new Map<string, Rel>();
    const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
    const name = partPath.slice(dir.length);
    const rels = this.xml(`${dir}_rels/${name}.rels`);
    for (const r of childrenNamed(rootElement(rels ?? { name: '', attrs: {}, children: [] }), 'Relationship')) {
      const external = r.attrs.TargetMode === 'External';
      map.set(r.attrs.Id, { target: external ? r.attrs.Target : resolvePath(partPath, r.attrs.Target), external });
      (map as Map<string, Rel> & { types?: Map<string, string> }).types ??= new Map();
      (map as Map<string, Rel> & { types: Map<string, string> }).types.set(r.attrs.Id, r.attrs.Type ?? '');
    }
    return map;
  }

  private relOfType(type: string): Rel | undefined {
    const types = (this.rels as Map<string, Rel> & { types?: Map<string, string> }).types;
    for (const [id, t] of types ?? []) if (t.endsWith(`/${type}`)) return this.rels.get(id);
    return undefined;
  }

  /** Run `fn` with the relationships of another part (header, footer, footnotes). */
  private inPart<T>(path: string, fn: () => T): T {
    const saved = this.rels;
    this.rels = this.loadRels(path);
    try {
      return fn();
    } finally {
      this.rels = saved;
    }
  }

  private loadTheme(): void {
    const rel = this.relOfType('theme');
    const theme = rel ? this.xml(rel.target) : undefined;
    if (!theme) return;
    this.themeFonts.major = find(find(theme, 'a:majorFont'), 'a:latin')?.attrs.typeface ?? '';
    this.themeFonts.minor = find(find(theme, 'a:minorFont'), 'a:latin')?.attrs.typeface ?? '';
    for (const c of elements(find(theme, 'a:clrScheme'))) {
      const v = find(c, 'a:srgbClr')?.attrs.val ?? find(c, 'a:sysClr')?.attrs.lastClr;
      if (v) this.themeColors[c.name.replace(/^a:/, '')] = v.toUpperCase();
    }
  }

  private themeColors: Record<string, string> = {};

  /** Chart parts are drawn from their cached values (SVG, vector in PDF). */
  private chart(el: XNode): Block | null {
    const ref = find(el, 'c:chart');
    const rid = ref?.attrs['r:id'];
    const rel = rid ? this.rels.get(rid) : undefined;
    const data = rel && !rel.external ? this.files[rel.target] : undefined;
    if (!data) return null;
    const extent = find(el, 'wp:extent');
    const w = extent ? +extent.attrs.cx / 12700 : 360;
    const h = extent ? +extent.attrs.cy / 12700 : 216;
    const chartXml = decoder.decode(data);
    const svg = chartToSvg(chartXml, { colors: this.themeColors, width: w, height: h });
    if (!svg) {
      this.warnings.push('Некоторые диаграммы не удалось прочитать: в файле нет сохранённых значений.');
      return null;
    }
    return { type: 'image', data: new TextEncoder().encode(svg), mime: 'image/svg+xml', width: Math.round((w * 96) / 72), height: Math.round((h * 96) / 72), alt: 'Диаграмма', chart: { xml: chartXml, colors: this.themeColors } };
  }

  private fontName(rFonts: XNode | undefined): string | undefined {
    if (!rFonts) return undefined;
    const a = rFonts.attrs;
    const theme = a['w:asciiTheme'] ?? a['w:hAnsiTheme'];
    if (theme) return /major/i.test(theme) ? this.themeFonts.major || 'Calibri Light' : this.themeFonts.minor || 'Calibri';
    return a['w:ascii'] ?? a['w:hAnsi'] ?? a['w:cs'] ?? a['w:eastAsia'];
  }

  /** Document defaults: docDefaults + the default paragraph style. */
  private loadDefaults(): void {
    const styles = this.xml(resolvePath(this.mainPath, 'styles.xml'));
    const dd = find(styles, 'w:docDefaults');
    const rPr = find(find(dd, 'w:rPrDefault'), 'w:rPr');
    const pPr = find(find(dd, 'w:pPrDefault'), 'w:pPr');
    let font = this.fontName(child(rPr, 'w:rFonts')) ?? (this.themeFonts.minor || 'Calibri');
    let size = child(rPr, 'w:sz')?.attrs['w:val'];
    let para = this.paraFormat(pPr, {});
    const normal = [...this.styles.entries()].find(([, d]) => d.isDefault && d.type === 'paragraph')?.[1];
    if (normal) {
      font = this.fontName(child(normal.rPr, 'w:rFonts')) ?? font;
      size = child(normal.rPr, 'w:sz')?.attrs['w:val'] ?? size;
      para = this.paraFormat(normal.pPr, para);
    }
    this.defFont = familyOf(font) ?? 'calibri';
    this.defSize = size ? +size / 2 : 10;
    delete para.align;
    this.defPara = para;
  }

  /** Paragraph formatting from a w:pPr, on top of `base`. */
  private paraFormat(pPr: XNode | undefined, base: ParaFormat, size = this.defSize): ParaFormat {
    const f: ParaFormat = { ...base };
    if (!pPr) return f;
    const sp = child(pPr, 'w:spacing');
    if (sp) {
      const before = twip(sp.attrs['w:before']);
      const after = twip(sp.attrs['w:after']);
      if (before !== undefined && sp.attrs['w:beforeAutospacing'] !== '1') f.spaceBefore = before;
      if (after !== undefined && sp.attrs['w:afterAutospacing'] !== '1') f.spaceAfter = after;
      const line = sp.attrs['w:line'];
      if (line && Number.isFinite(+line)) {
        const rule = sp.attrs['w:lineRule'] ?? 'auto';
        f.lineHeight = rule === 'auto' ? +line / 240 : Math.max(0.6, +line / 20 / (size * 1.15));
      }
    }
    const ind = child(pPr, 'w:ind');
    if (ind) {
      const left = twip(ind.attrs['w:left'] ?? ind.attrs['w:start']);
      if (left !== undefined) f.indent = left;
      const first = twip(ind.attrs['w:firstLine']);
      const hanging = twip(ind.attrs['w:hanging']);
      if (hanging) f.firstLine = -hanging;
      else if (first !== undefined) f.firstLine = first;
    }
    const jc = child(pPr, 'w:jc')?.attrs['w:val'];
    if (jc) f.align = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc === 'both' || jc === 'distribute' ? 'justify' : 'left';
    return f;
  }

  private loadStyles(): void {
    const styles = this.xml(resolvePath(this.mainPath, 'styles.xml'));
    const root = styles && rootElement(styles);
    for (const s of childrenNamed(root, 'w:style')) {
      const id = s.attrs['w:styleId'];
      if (!id) continue;
      this.styles.set(id, {
        type: s.attrs['w:type'] ?? 'paragraph',
        isDefault: s.attrs['w:default'] === '1',
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
    const face = this.fontName(child(rPr, 'w:rFonts'));
    if (face) {
      s.font = familyOf(face);
      if (s.font === 'mono') s.code = true;
    }
    const sz = child(rPr, 'w:sz')?.attrs['w:val'];
    if (sz && Number.isFinite(+sz)) s.size = +sz / 2;
    const hl = child(rPr, 'w:highlight')?.attrs['w:val'];
    const shd = child(rPr, 'w:shd')?.attrs['w:fill'];
    if (hl && HIGHLIGHT[hl]) s.highlight = HIGHLIGHT[hl];
    else if (hl === 'none') delete s.highlight;
    else if (shd && /^[0-9a-fA-F]{6}$/.test(shd) && shd.toUpperCase() !== 'FFFFFF') s.highlight = shd.toUpperCase();
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
    if (find(el, 'c:chart')) return this.chart(el);
    const blip = find(el, 'a:blip') ?? find(el, 'v:imagedata');
    const rid = blip?.attrs['r:embed'] ?? blip?.attrs['r:id'];
    // Office 2016+ stores SVG pictures as an svgBlip extension next to a raster/EMF fallback
    const svgRid = blip && findBy(blip, (n) => n.name.endsWith(':svgBlip'))?.attrs['r:embed'];
    const svgRel = svgRid ? this.rels.get(svgRid) : undefined;
    const svgData = svgRel && !svgRel.external ? this.files[svgRel.target] : undefined;
    const rel = rid ? this.rels.get(rid) : undefined;
    const data = svgData ?? (rel && !rel.external ? this.files[rel.target] : undefined);
    if (!data) return null;
    const mime = svgData ? 'image/svg+xml' : sniffImageMime(data);
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
          const instr = n.attrs['w:instr'] ?? '';
          const pf = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/.exec(instr);
          if (pf) {
            const r0 = find(n, 'w:r');
            const st = this.applyRPr({ ...base }, child(r0, 'w:rPr'));
            out.push({ text: '', ...this.clean(st), field: pf[1] === 'PAGE' ? 'page' : 'pages' });
            break;
          }
          const m = /HYPERLINK\s+"([^"]+)"/.exec(instr);
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
      // Result of a PAGE/NUMPAGES field: replaced by a live field (drawn per page)
      if (field?.phase === 'result' && field.page) {
        if (!field.emitted) out.push({ text: '', ...this.clean(style), field: field.page });
        field.emitted = true;
        return;
      }
      out.push({ text, ...this.clean(style) });
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
            const pf = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/.exec(field.instr);
            if (pf) field.page = pf[1] === 'PAGE' ? 'page' : 'pages';
          } else if (t === 'end') this.fieldStack.pop();
          break;
        }
        case 'w:tab':
        case 'w:ptab':
          push('\t');
          break;
        case 'w:footnoteRef':
        case 'w:endnoteRef':
          break; // the number mark inside a note body
        case 'w:br':
        case 'w:cr':
          if (c.attrs['w:type'] === 'page') out.push({ type: 'pagebreak' });
          else push('\n');
          break;
        case 'w:noBreakHyphen':
          push('-');
          break;
        case 'w:footnoteReference':
        case 'w:endnoteReference': {
          const note = this.footnote(c.attrs['w:id'], c.name === 'w:endnoteReference');
          if (note) out.push({ text: '', ...this.clean(style), note });
          break;
        }
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
    // Paragraph formatting: style chain, then direct formatting; keep only what differs from defaults
    let fmt: ParaFormat = { ...this.defPara };
    for (const def of this.styleChain(styleId)) fmt = this.paraFormat(def.pPr, fmt);
    fmt = this.paraFormat(pPr, fmt);
    const own: ParaFormat = {};
    for (const k of ['indent', 'firstLine', 'spaceBefore', 'spaceAfter', 'lineHeight'] as const) {
      const v = fmt[k];
      if (v !== undefined && Math.abs(v - (this.defPara[k] ?? 0)) > 0.01 && !(v === 0 && this.defPara[k] === undefined)) own[k] = Math.round(v * 100) / 100;
    }
    const align: Align | undefined = fmt.align && fmt.align !== 'left' ? fmt.align : undefined;
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
      const hasText = r.some((x) => x.text.trim() || x.note || x.field);
      if (!hasText && !force) return;
      let b: Block;
      if (level && hasText) b = { type: 'heading', level: level as 1, runs: r };
      else b = { type: 'paragraph', runs: hasText ? r : [], ...(num ? {} : own) };
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
    // Sections: a paragraph carrying a sectPr ends a section; the body's last sectPr ends the final one
    const finalSect = childrenNamed(body, 'w:sectPr').pop();
    const groups: { nodes: XNode[]; sect?: XNode }[] = [];
    let cur: XNode[] = [];
    for (const n of elements(body)) {
      if (n.name === 'w:sectPr') continue;
      cur.push(n);
      const sp = n.name === 'w:p' ? child(child(n, 'w:pPr'), 'w:sectPr') : undefined;
      if (sp) {
        groups.push({ nodes: cur, sect: sp });
        cur = [];
      }
    }
    groups.push({ nodes: cur, sect: finalSect });
    const blocks: Block[] = [];
    groups.forEach((g, i) => {
      const cols = child(g.sect, 'w:cols');
      const listed = childrenNamed(cols, 'w:col');
      const columns = Math.max(1, Math.min(12, listed.length || +(cols?.attrs['w:num'] ?? 1) || 1));
      const gapTw = listed.length > 1 ? listed.slice(0, -1).reduce((sum, c) => sum + (+(c.attrs['w:space'] ?? 0) || 0), 0) / (listed.length - 1) : +(cols?.attrs['w:space'] ?? 720);
      const type = child(g.sect, 'w:type')?.attrs['w:val'] ?? 'nextPage';
      const inner = this.body(g.nodes);
      if (!inner.length && !g.sect) return;
      if (i > 0 || columns > 1)
        blocks.push({ type: 'section', columns, gap: (Number.isFinite(gapTw) ? gapTw : 720) / 20, ...(['1', 'true', 'on'].includes(cols?.attrs['w:sep'] ?? '') ? { separator: true } : {}), start: i > 0 && type !== 'continuous' && type !== 'nextColumn' ? 'page' : 'continuous' });
      blocks.push(...inner);
    });
    let title: string | undefined;
    const core = this.xml('docProps/core.xml');
    if (core) title = textContent(find(core, 'dc:title')).trim() || undefined;
    const result: Doc = { title, blocks: normalizeSections(blocks), font: this.defFont, size: this.defSize, defaults: this.defPara };
    const sect = childrenNamed(body, 'w:sectPr').pop() ?? findAll(body, 'w:sectPr').pop();
    if (sect) {
      const pg = child(sect, 'w:pgSz');
      const mar = child(sect, 'w:pgMar');
      const w = twip(pg?.attrs['w:w']);
      const h = twip(pg?.attrs['w:h']);
      if (w && h) {
        const m = (k: string, d: number): number => Math.abs(twip(mar?.attrs[`w:${k}`]) ?? d);
        result.page = { width: w, height: h, margins: { top: m('top', 72), right: m('right', 72), bottom: m('bottom', 72), left: m('left', 72) } } satisfies PageSetup;
      }
      const part = (tag: 'w:headerReference' | 'w:footerReference'): Block[] | undefined => {
        const refs = childrenNamed(sect, tag);
        const ref = refs.find((r) => r.attrs['w:type'] === 'default') ?? refs[0];
        const rel = ref ? this.rels.get(ref.attrs['r:id']) : undefined;
        const xml = rel ? this.xml(rel.target) : undefined;
        if (!rel || !xml) return undefined;
        const blocksOf = this.inPart(rel.target, () => this.body(elements(rootElement(xml))));
        return blocksOf.some((b) => b.type !== 'paragraph' || b.runs.length) ? blocksOf : undefined;
      };
      result.header = part('w:headerReference');
      result.footer = part('w:footerReference');
    }
    return result;
  }

  /** Drop properties equal to the document defaults so runs stay compact. */
  private clean(style: RStyle): RStyle {
    const s = { ...style };
    if (s.font === this.defFont) delete s.font;
    if (s.size !== undefined && Math.abs(s.size - this.defSize) < 0.01) delete s.size;
    delete (s as { hidden?: boolean }).hidden;
    return s;
  }

  private footnote(id: string | undefined, endnote: boolean): Block[] | null {
    if (!id) return null;
    const kind = endnote ? 'endnotes' : 'footnotes';
    const key = `${kind}:${id}`;
    if (!this.footnotes.has(key)) {
      const rel = this.relOfType(kind);
      const xml = rel ? this.xml(rel.target) : undefined;
      for (const n of childrenNamed(rootElement(xml ?? { name: '', attrs: {}, children: [] }), endnote ? 'w:endnote' : 'w:footnote')) this.footnotes.set(`${kind}:${n.attrs['w:id']}`, n);
      this.notesPath = rel?.target ?? '';
      if (!this.footnotes.has(key)) return null;
    }
    const node = this.footnotes.get(key)!;
    const parse = (): Block[] => this.body(elements(node));
    const blocks = this.notesPath ? this.inPart(this.notesPath, parse) : parse();
    // Word puts a space after the number mark; trim it
    return blocks.map((b) => (b.type === 'paragraph' ? { ...b, runs: trimRuns(b.runs) } : b)).filter((b) => b.type !== 'paragraph' || b.runs.length);
  }
}

export function docxToDoc(bytes: Uint8Array): { doc: Doc; warnings: string[] } {
  const r = new DocxReader(bytes);
  const doc = r.read();
  return { doc, warnings: [...new Set(r.warnings)] };
}
