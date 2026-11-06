/**
 * OpenDocument readers (ODF 1.2/1.3): text documents (.odt) → Doc and
 * spreadsheets (.ods) → Workbook. Styles are resolved through the parent chain
 * (default style → named styles → automatic styles) the way office suites do.
 */
import { unzipSync } from 'fflate';
import { familyOf, normalizeRuns, normalizeSections, trimRuns, type Align, type Block, type CellValue, type Doc, type FontFamily, type ListItem, type PageSetup, type ParaFormat, type Sheet, type TableRow, type TextRun, type Workbook } from '../model';
import { childrenNamed, elements, find, parseXml, textContent, type XNode } from '../xml';
import { base64ToBytes, ConvertError, imageSize, sniffImageMime, unzipOffice } from '../util';

const decoder = new TextDecoder('utf-8');

/** ODF length ("2cm", "0.5in", "12pt", "3mm", "10px") → points. */
export function odfLength(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  if (!Number.isFinite(n)) return undefined;
  const k: Record<string, number> = { cm: 28.3465, mm: 2.83465, in: 72, pt: 1, pc: 12, px: 0.75 };
  return Math.round(n * (k[m[2] ?? 'pt'] ?? 1) * 100) / 100;
}

const KNOWN_SANS = /arial|helvet|verdana|tahoma|segoe|trebuchet|roboto|ubuntu|cantarell|\binter\b|lato|montserrat|franklin|gill sans|futura|century gothic|lucida sans|freesans|nimbus sans|arimo|carlito|calibri/i;

/**
 * Font family of an ODF font face. The family name wins over the generic class:
 * converters often label Arial "roman", so the class only decides for unknown names.
 */
function fontFamilyFor(name: string, generic?: string, pitch?: string): FontFamily {
  if (/calibri|carlito/i.test(name)) return 'calibri';
  if (KNOWN_SANS.test(name) || /sans/i.test(name)) return familyOf(name) ?? 'sans';
  const fam = familyOf(name);
  if (fam && fam !== 'sans') return fam;
  if (generic === 'modern' || pitch === 'fixed') return 'mono';
  if (generic === 'roman') return 'serif';
  return 'sans';
}

const hex = (v: string | undefined): string | undefined => (v && /^#[0-9a-f]{6}$/i.test(v) ? v.slice(1).toUpperCase() : undefined);

interface TextProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sup?: boolean;
  sub?: boolean;
  color?: string;
  highlight?: string;
  font?: FontFamily;
  size?: number;
}

interface OdfStyle {
  name: string;
  display: string;
  family: string;
  parent?: string;
  text: TextProps;
  para: ParaFormat;
  breakBefore?: boolean;
  breakAfter?: boolean;
  outlineLevel?: number;
  listStyle?: string;
  pageLayout?: string;
  masterPage?: string;
  /** Section columns (style:section-properties) */
  columns?: { count: number; gap: number; sep: boolean; rel?: number[]; starts?: number[]; ends?: number[] };
}

interface ListLevel {
  ordered: boolean;
  start: number;
}

class OdfStyles {
  readonly styles = new Map<string, OdfStyle>();
  readonly defaults = new Map<string, OdfStyle>();
  readonly lists = new Map<string, ListLevel[]>();
  readonly fonts = new Map<string, FontFamily>();
  readonly pageLayouts = new Map<string, PageSetup & { headerExtent: number; footerExtent: number }>();

  addFontFaces(root: XNode | undefined): void {
    for (const f of childrenNamed(find(root, 'office:font-face-decls'), 'style:font-face')) {
      const family = (f.attrs['svg:font-family'] ?? f.attrs['style:name'] ?? '').replace(/^['"]|['"]$/g, '');
      this.fonts.set(f.attrs['style:name'], fontFamilyFor(family, f.attrs['style:font-family-generic'], f.attrs['style:font-pitch']));
    }
  }

  textProps(el: XNode | undefined): TextProps {
    const t: TextProps = {};
    if (!el) return t;
    const a = el.attrs;
    const w = a['fo:font-weight'];
    if (w) t.bold = w === 'bold' || (/^\d+$/.test(w) && +w >= 600);
    const s = a['fo:font-style'];
    if (s) t.italic = s === 'italic' || s === 'oblique';
    const u = a['style:text-underline-style'];
    if (u) t.underline = u !== 'none';
    const lt = a['style:text-line-through-style'];
    if (lt) t.strike = lt !== 'none';
    const pos = a['style:text-position'];
    if (pos) {
      const first = pos.trim().split(/\s+/)[0];
      t.sup = first === 'super' || (first.endsWith('%') && parseFloat(first) > 0);
      t.sub = first === 'sub' || (first.endsWith('%') && parseFloat(first) < 0);
    }
    const color = hex(a['fo:color']);
    if (color) t.color = color;
    const bg = a['fo:background-color'];
    if (bg) t.highlight = bg === 'transparent' ? undefined : hex(bg);
    const fontName = a['style:font-name'] ?? a['fo:font-family'];
    if (fontName) t.font = this.fonts.get(fontName) ?? fontFamilyFor(fontName.replace(/^['"]|['"]$/g, ''));
    const size = a['fo:font-size'];
    if (size && !size.endsWith('%')) t.size = odfLength(size);
    return t;
  }

  private paraProps(el: XNode | undefined, st: OdfStyle): void {
    if (!el) return;
    const a = el.attrs;
    const align = a['fo:text-align'];
    const map: Record<string, Align> = { start: 'left', left: 'left', center: 'center', end: 'right', right: 'right', justify: 'justify' };
    if (align && map[align]) st.para.align = map[align];
    const ml = odfLength(a['fo:margin-left'] ?? a['fo:margin-start']);
    if (ml !== undefined) st.para.indent = ml;
    const ti = odfLength(a['fo:text-indent']);
    if (ti !== undefined) st.para.firstLine = ti;
    const mt = odfLength(a['fo:margin-top']);
    if (mt !== undefined) st.para.spaceBefore = mt;
    const mb = odfLength(a['fo:margin-bottom']);
    if (mb !== undefined) st.para.spaceAfter = mb;
    const lh = a['fo:line-height'];
    if (lh?.endsWith('%')) st.para.lineHeight = parseFloat(lh) / 100;
    if (a['fo:break-before'] === 'page') st.breakBefore = true;
    if (a['fo:break-after'] === 'page') st.breakAfter = true;
  }

  add(container: XNode | undefined): void {
    for (const el of elements(container)) {
      if (el.name === 'style:style' || el.name === 'style:default-style') {
        const st: OdfStyle = {
          name: el.attrs['style:name'] ?? '',
          display: el.attrs['style:display-name'] ?? (el.attrs['style:name'] ?? '').replace(/_20_/g, ' '),
          family: el.attrs['style:family'] ?? 'paragraph',
          parent: el.attrs['style:parent-style-name'],
          text: this.textProps(find(el, 'style:text-properties')),
          para: {},
          outlineLevel: el.attrs['style:default-outline-level'] ? +el.attrs['style:default-outline-level'] : undefined,
          listStyle: el.attrs['style:list-style-name'],
          masterPage: el.attrs['style:master-page-name'] || undefined,
        };
        const cols = find(find(el, 'style:section-properties'), 'style:columns');
        const count = +(cols?.attrs['fo:column-count'] ?? 1);
        if (count > 1) {
          // fo:column-gap, or the spacing of the individual style:column elements
          const listed = elements(cols).filter((c) => c.name === 'style:column');
          const gap = odfLength(cols?.attrs['fo:column-gap']) ?? (listed.length > 1 ? (odfLength(listed[0].attrs['fo:end-indent']) ?? 0) + (odfLength(listed[1].attrs['fo:start-indent']) ?? 0) : 0);
          st.columns = { count: Math.min(12, count), gap: gap || 18, sep: !!find(cols, 'style:column-sep') };
          if (listed.length === count) {
            st.columns.rel = listed.map((c) => parseFloat(c.attrs['style:rel-width'] ?? '1') || 1);
            st.columns.starts = listed.map((c) => odfLength(c.attrs['fo:start-indent']) ?? 0);
            st.columns.ends = listed.map((c) => odfLength(c.attrs['fo:end-indent']) ?? 0);
          }
        }
        this.paraProps(find(el, 'style:paragraph-properties'), st);
        if (el.name === 'style:default-style') this.defaults.set(st.family, st);
        else this.styles.set(st.name, st);
      } else if (el.name === 'text:list-style') {
        const levels: ListLevel[] = [];
        for (const lv of elements(el)) {
          const level = +(lv.attrs['text:level'] ?? 1) - 1;
          if (level < 0 || level > 9) continue;
          const ordered = lv.name === 'text:list-level-style-number' && !!(lv.attrs['style:num-format'] ?? '1');
          levels[level] = { ordered, start: +(lv.attrs['text:start-value'] ?? 1) || 1 };
        }
        this.lists.set(el.attrs['style:name'], levels);
      } else if (el.name === 'style:page-layout') {
        const p = find(el, 'style:page-layout-properties');
        const w = odfLength(p?.attrs['fo:page-width']);
        const h = odfLength(p?.attrs['fo:page-height']);
        if (p && w && h) {
          const m = (k: string): number => odfLength(p.attrs[`fo:margin-${k}`]) ?? odfLength(p.attrs['fo:margin']) ?? 56.7;
          // The header/footer sit inside the page margins: body text starts after them
          const extent = (k: 'header' | 'footer'): number => {
            const hp = find(find(el, `style:${k}-style`), 'style:header-footer-properties');
            if (!hp) return 0;
            const height = odfLength(hp.attrs['svg:height']) ?? odfLength(hp.attrs['fo:min-height']) ?? 0;
            return height + (odfLength(hp.attrs[k === 'header' ? 'fo:margin-bottom' : 'fo:margin-top']) ?? 0);
          };
          this.pageLayouts.set(el.attrs['style:name'], { width: w, height: h, margins: { top: m('top'), right: m('right'), bottom: m('bottom'), left: m('left') }, headerExtent: extent('header'), footerExtent: extent('footer') });
        }
      }
    }
  }

  /** Style chain from the most generic (family default) to the given style. */
  chain(name: string | undefined, family: string): OdfStyle[] {
    const out: OdfStyle[] = [];
    let cur = name ? this.styles.get(name) : undefined;
    let guard = 0;
    while (cur && guard++ < 30) {
      out.unshift(cur);
      cur = cur.parent ? this.styles.get(cur.parent) : undefined;
    }
    const def = this.defaults.get(family);
    if (def) out.unshift(def);
    return out;
  }

  text(name: string | undefined, family: string): TextProps {
    return Object.assign({}, ...this.chain(name, family).map((s) => s.text)) as TextProps;
  }

  para(name: string | undefined): ParaFormat & { breakBefore?: boolean; breakAfter?: boolean; outlineLevel?: number; display?: string; listStyle?: string } {
    const r: ParaFormat & { breakBefore?: boolean; breakAfter?: boolean; outlineLevel?: number; display?: string; listStyle?: string } = {};
    const chain = this.chain(name, 'paragraph');
    for (const s of chain) {
      Object.assign(r, s.para);
      if (s.outlineLevel !== undefined) r.outlineLevel = s.outlineLevel;
      if (s.listStyle) r.listStyle = s.listStyle;
    }
    const own = name ? this.styles.get(name) : undefined;
    r.breakBefore = own?.breakBefore;
    r.breakAfter = own?.breakAfter;
    // The display name of the nearest named (non-automatic) ancestor identifies the role
    r.display = chain.length ? chain[chain.length - 1].display : undefined;
    for (let i = chain.length - 1; i >= 0; i--) {
      if (/^(Heading|Title|Subtitle|Quotations|Quote|Preformatted Text|Code|Horizontal Line|Source Text|Table Heading|Footnote|Endnote)/i.test(chain[i].display)) {
        r.display = chain[i].display;
        break;
      }
    }
    return r;
  }
}

type Inline = TextRun | Block;
const isBlock = (x: Inline): x is Block => 'type' in x;

class OdtReader {
  warnings = new Set<string>();
  private styles = new OdfStyles();
  private docFont?: FontFamily;
  private docSize?: number;
  private std: ParaFormat = {};
  private inHeaderRow = false;

  constructor(private files: Record<string, Uint8Array>) {}

  /** Page of each master page (by name), the document's master, the one in use, and the content width. */
  private masterPages = new Map<string, PageSetup>();
  private docMaster = 'Standard';
  private currentMaster = 'Standard';
  private contentW = 481.9;
  /** Parsing the body (not headers, footers or notes), and whether a body block was seen yet. */
  private inBody = false;
  private bodyStarted = false;

  private xml(path: string): XNode | undefined {
    const f = this.files[path];
    return f ? parseXml(decoder.decode(f)) : undefined;
  }

  read(): Doc {
    const content = this.xml('content.xml');
    if (!content) throw new ConvertError('В архиве нет содержимого документа OpenDocument (content.xml).');
    const stylesXml = this.xml('styles.xml');
    this.styles.addFontFaces(stylesXml);
    this.styles.addFontFaces(content);
    this.styles.add(find(stylesXml, 'office:styles'));
    this.styles.add(find(stylesXml, 'office:automatic-styles'));
    this.styles.add(find(content, 'office:automatic-styles'));
    // master pages: a paragraph switching to another one starts a section with that page
    const masterEls = childrenNamed(find(stylesXml, 'office:master-styles'), 'style:master-page');
    for (const mp of masterEls) {
      const layout = this.styles.pageLayouts.get(mp.attrs['style:page-layout-name']);
      if (!layout) continue;
      const hdr = !!find(mp, 'style:header');
      const ftr = !!find(mp, 'style:footer');
      const r = (v: number): number => Math.round(v * 100) / 100;
      this.masterPages.set(mp.attrs['style:name'] ?? '', {
        width: layout.width,
        height: layout.height,
        margins: { ...layout.margins, top: r(layout.margins.top + (hdr ? layout.headerExtent : 0)), bottom: r(layout.margins.bottom + (ftr ? layout.footerExtent : 0)) },
      });
    }
    const std0 = masterEls.find((m) => m.attrs['style:name'] === 'Standard' || m.attrs['style:name'] === 'Default') ?? masterEls[0];
    this.docMaster = this.currentMaster = std0?.attrs['style:name'] ?? 'Standard';
    const docPage = this.masterPages.get(this.docMaster);
    if (docPage) this.contentW = docPage.width - docPage.margins.left - docPage.margins.right;
    const base = this.styles.text('Standard', 'paragraph');
    this.docFont = base.font;
    this.docSize = base.size;
    this.std = this.styles.para('Standard');
    const body = find(content, 'office:text');
    if (!body) {
      if (find(content, 'office:spreadsheet')) throw new ConvertError('Это таблица OpenDocument (.ods), а не текстовый документ.');
      throw new ConvertError('В документе OpenDocument нет текста (office:text).');
    }
    this.inBody = true;
    const blocks = normalizeSections(this.blocks(body.children));
    this.inBody = false;
    const doc: Doc = { blocks };
    if (this.docFont) doc.font = this.docFont;
    if (this.docSize) doc.size = this.docSize;
    const std = this.std;
    const defaults: ParaFormat = {};
    if (std.spaceAfter !== undefined) defaults.spaceAfter = std.spaceAfter;
    if (std.spaceBefore !== undefined) defaults.spaceBefore = std.spaceBefore;
    if (std.lineHeight !== undefined) defaults.lineHeight = std.lineHeight;
    if (Object.keys(defaults).length) doc.defaults = defaults;
    // Master page: header, footer and page geometry
    const masters = childrenNamed(find(stylesXml, 'office:master-styles'), 'style:master-page');
    // the master of the first body paragraph (Standard unless it names another)
    const master = masters.find((m) => m.attrs['style:name'] === this.docMaster) ?? masters.find((m) => m.attrs['style:name'] === 'Standard' || m.attrs['style:name'] === 'Default') ?? masters[0];
    if (master) {
      const layout = this.styles.pageLayouts.get(master.attrs['style:page-layout-name']);
      const header = find(master, 'style:header');
      const footer = find(master, 'style:footer');
      if (header && header.attrs['style:display'] !== 'false') {
        const h = this.blocks(header.children);
        if (h.length) doc.header = h;
      }
      if (footer && footer.attrs['style:display'] !== 'false') {
        const f = this.blocks(footer.children);
        if (f.length) doc.footer = f;
      }
      if (layout) {
        const r = (v: number): number => Math.round(v * 100) / 100;
        const mg = layout.margins;
        doc.page = {
          width: layout.width,
          height: layout.height,
          margins: { ...mg, top: r(mg.top + (doc.header ? layout.headerExtent : 0)), bottom: r(mg.bottom + (doc.footer ? layout.footerExtent : 0)) },
        };
      }
    }
    const meta = this.xml('meta.xml');
    const title = textContent(find(meta, 'dc:title')).trim();
    if (title) doc.title = title;
    return doc;
  }

  /** Runs inherit only what differs from the document defaults. */
  private style(t: TextProps): Omit<TextRun, 'text'> {
    const s: Omit<TextRun, 'text'> = {};
    if (t.bold) s.bold = true;
    if (t.italic) s.italic = true;
    if (t.underline) s.underline = true;
    if (t.strike) s.strike = true;
    if (t.sup) s.sup = true;
    else if (t.sub) s.sub = true;
    if (t.color && t.color !== '000000') s.color = t.color;
    if (t.highlight && t.highlight !== 'FFFFFF') s.highlight = t.highlight;
    if (t.font && t.font !== this.docFont) s.font = t.font;
    if (t.size && this.docSize && Math.abs(t.size - this.docSize) > 0.01) s.size = t.size;
    return s;
  }

  private blocks(nodes: (XNode | string)[], listDepth = 0, listStyle?: string): Block[] {
    const out: Block[] = [];
    for (const n of nodes) {
      if (typeof n === 'string') continue;
      switch (n.name) {
        case 'text:p':
        case 'text:h':
          this.paragraph(n, out);
          break;
        case 'text:list':
          for (const b of this.list(n, listDepth, listStyle)) {
            const prev = out[out.length - 1];
            // LibreOffice writes a nested list after its parent as an item that holds only the sub-list
            if (b.type === 'list' && prev?.type === 'list' && b.items[0]?.blocks.every((x) => x.type === 'list')) {
              prev.items[prev.items.length - 1].blocks.push(...b.items[0].blocks);
              prev.items.push(...b.items.slice(1));
            } else out.push(b);
          }
          break;
        case 'table:table':
          this.switchMaster(n.attrs['table:style-name'], out);
          out.push(this.table(n));
          break;
        case 'text:section': {
          const cols = this.styles.styles.get(n.attrs['text:style-name'] ?? '')?.columns;
          if (cols) {
            // a page break before the section is written on its first paragraph
            const firstP = elements(n).find((c) => c.name === 'text:p' || c.name === 'text:h');
            let pageStart = !!firstP && !!this.styles.para(firstP.attrs['text:style-name'] ?? '').breakBefore;
            // a master page switched to by the first paragraph belongs to this section's marker
            const inner = this.blocks(n.children, listDepth, listStyle);
            const lead = inner[0];
            let page: PageSetup | undefined;
            if (lead?.type === 'section' && lead.columns <= 1 && lead.page) {
              page = lead.page;
              pageStart = true;
              inner.shift();
            }
            // unequal columns (in the content width of the section's page): relative widths (indents included) minus the indents, in a content width
            let widths: number[] | undefined;
            let gaps: number[] | undefined;
            if (cols.rel && cols.starts && cols.ends) {
              const perRel = this.contentW / cols.rel.reduce((a, b) => a + b, 0);
              widths = cols.rel.map((r, i) => Math.max(1, r * perRel - cols.starts![i] - cols.ends![i]));
              gaps = cols.rel.slice(0, -1).map((_, i) => cols.ends![i] + cols.starts![i + 1]);
              // equal columns are written with rel-widths that differ by rounding (32767*, 32768*)
              if (Math.max(...widths) - Math.min(...widths) < 1 && gaps.every((g) => Math.abs(g - gaps![0]) < 0.5)) widths = gaps = undefined;
            }
            out.push({ type: 'section', columns: cols.count, gap: cols.gap, ...(cols.sep ? { separator: true } : {}), start: pageStart && out.length ? 'page' : 'continuous', ...(widths ? { widths, gaps } : {}), ...(page ? { page } : {}) });
            out.push(...inner);
            out.push({ type: 'section', columns: 1, gap: 36, start: 'continuous' });
          } else out.push(...this.blocks(n.children, listDepth, listStyle));
          break;
        }
        case 'text:index-body':
        case 'text:table-of-content':
        case 'text:alphabetical-index':
        case 'text:illustration-index':
        case 'text:table-index':
        case 'text:object-index':
        case 'text:user-index':
        case 'text:bibliography':
        case 'text:index-title':
        case 'office:text':
        case 'text:deletion':
          if (n.name !== 'text:deletion') out.push(...this.blocks(n.children, listDepth, listStyle));
          break;
        case 'draw:frame':
        case 'draw:a':
        case 'draw:g':
        case 'draw:custom-shape':
        case 'draw:rect':
          for (const b of this.drawing(n)) out.push(b);
          break;
        case 'text:soft-page-break':
        case 'text:sequence-decls':
        case 'text:variable-decls':
        case 'text:user-field-decls':
        case 'text:tracked-changes':
        case 'office:forms':
        case 'table:shapes':
          break;
        default:
          break;
      }
    }
    return out;
  }

  /**
   * A body paragraph or table whose style names a master page: the first one sets the document's
   * page; a later change starts a new page of that kind (a section with its own page).
   */
  private switchMaster(styleName: string | undefined, out: Block[]): void {
    if (!this.inBody) return;
    const master = styleName ? this.styles.styles.get(styleName)?.masterPage : undefined;
    const started = this.bodyStarted;
    this.bodyStarted = true;
    if (!master || master === this.currentMaster || !this.masterPages.has(master)) return;
    this.currentMaster = master;
    const pg = this.masterPages.get(master)!;
    this.contentW = pg.width - pg.margins.left - pg.margins.right;
    if (!started) this.docMaster = master;
    else out.push({ type: 'section', columns: 1, gap: 36, start: 'page', page: pg });
  }

  private paragraph(el: XNode, out: Block[]): void {
    const styleName = el.attrs['text:style-name'];
    const p = this.styles.para(styleName);
    this.switchMaster(styleName, out);
    const display = p.display ?? '';
    const isHeading = el.name === 'text:h' || /^(Title|Heading)/i.test(display);
    const isNote = /^(Footnote|Endnote)/i.test(display);
    const base = this.styles.text(styleName, 'paragraph');
    // The look of headings, quotes and notes belongs to their styles, not to the runs
    let runBase: TextProps = base;
    if (isHeading) runBase = { ...base, bold: undefined, size: this.docSize, font: this.docFont, color: undefined };
    else if (/^(Quotations|Quote)/i.test(display)) runBase = { ...base, italic: undefined, color: undefined };
    else if (isNote) runBase = { ...base, size: this.docSize };
    else if (/^Table Heading/i.test(display) && this.inHeaderRow) runBase = { ...base, bold: undefined };
    const items: Inline[] = [];
    this.inline(el.children, runBase, undefined, items);
    if (p.breakBefore && out.length) out.push({ type: 'pagebreak' });
    const fmt: ParaFormat = {};
    for (const k of ['align', 'indent', 'firstLine', 'spaceBefore', 'spaceAfter', 'lineHeight'] as const) if (p[k] !== undefined && p[k] !== this.std[k]) (fmt as Record<string, unknown>)[k] = p[k];
    if (fmt.align === 'left') delete fmt.align;
    if (isNote) for (const k of Object.keys(fmt)) delete (fmt as Record<string, unknown>)[k];
    let level: number | undefined;
    if (el.name === 'text:h') level = +(el.attrs['text:outline-level'] ?? p.outlineLevel ?? 1) || 1;
    else if (/^Title$/i.test(display)) level = 1;
    else if (/^Heading (\d)/i.test(display)) level = +/^Heading (\d)/i.exec(display)![1];
    if (/^Horizontal Line/i.test(display) && !items.some((x) => !isBlock(x) && x.text.trim())) {
      out.push({ type: 'hr' });
      return;
    }
    if (/^(Preformatted Text|Source Text|Code)$/i.test(display)) {
      const raw = items.map((r) => (isBlock(r) ? '' : r.text)).join('');
      const prev = out[out.length - 1];
      if (prev?.type === 'code') prev.text += '\n' + raw;
      else out.push({ type: 'code', text: raw });
      return;
    }
    const quote = /^(Quotations|Quote)$/i.test(display);
    let runs: TextRun[] = [];
    let emitted = false;
    const flush = (force: boolean): void => {
      const r = trimRuns(runs);
      runs = [];
      const hasText = r.some((x) => x.text.trim() || x.note || x.field);
      if (!hasText && !force) return;
      emitted = true;
      let b: Block;
      if (level !== undefined && hasText) {
        const { indent: _i, firstLine: _f, ...hf } = fmt;
        b = { type: 'heading', level: Math.min(6, Math.max(1, level)) as 1, runs: r, ...hf };
      } else b = { type: 'paragraph', runs: hasText ? r : [], ...fmt };
      if (quote && hasText) {
        const prev = out[out.length - 1];
        if (prev?.type === 'quote') prev.blocks.push(b);
        else out.push({ type: 'quote', blocks: [b] });
      } else out.push(b);
    };
    let hadBlock = false;
    for (const it of items) {
      if (isBlock(it)) {
        flush(false);
        out.push(it);
        hadBlock = true;
      } else runs.push(it);
    }
    flush(!emitted && !hadBlock);
    if (p.breakAfter) out.push({ type: 'pagebreak' });
  }

  private inline(nodes: (XNode | string)[], base: TextProps, link: string | undefined, out: Inline[]): void {
    for (const n of nodes) {
      if (typeof n === 'string') {
        const t = n.replace(/[ \t\r\n]+/g, ' ');
        if (t) out.push({ text: t, ...this.style(base), ...(link ? { link } : {}) });
        continue;
      }
      switch (n.name) {
        case 'text:span':
          // Character styles inherit from the paragraph; only explicitly set properties override
          this.inline(n.children, this.mergeText(base, n.attrs['text:style-name']), link, out);
          break;
        case 'text:a': {
          const href = n.attrs['xlink:href'];
          const url = href && /^(https?:|mailto:)/i.test(href) ? href : undefined;
          this.inline(n.children, this.mergeText(base, n.attrs['text:style-name']), url ?? link, out);
          break;
        }
        case 'text:s':
          out.push({ text: ' '.repeat(Math.min(1000, Math.max(1, +(n.attrs['text:c'] ?? 1) || 1))), ...this.style(base) });
          break;
        case 'text:tab':
          out.push({ text: '\t', ...this.style(base) });
          break;
        case 'text:line-break':
          out.push({ text: '\n', ...this.style(base) });
          break;
        case 'text:page-number':
          out.push({ text: textContent(n) || '1', field: 'page', ...this.style(base) });
          break;
        case 'text:page-count':
          out.push({ text: textContent(n) || '1', field: 'pages', ...this.style(base) });
          break;
        case 'text:note': {
          const body = find(n, 'text:note-body');
          const note = this.blocks(body?.children ?? []);
          out.push({ text: '', note: note.length ? note : [{ type: 'paragraph', runs: [{ text: ' ' }] }], ...this.style(base) });
          break;
        }
        case 'draw:frame':
        case 'draw:a':
        case 'draw:g':
        case 'draw:custom-shape':
        case 'draw:rect':
          out.push(...this.drawing(n));
          break;
        case 'text:bookmark':
        case 'text:bookmark-start':
        case 'text:bookmark-end':
        case 'text:reference-mark':
        case 'text:reference-mark-start':
        case 'text:reference-mark-end':
        case 'text:change':
        case 'text:change-start':
        case 'text:change-end':
        case 'text:soft-page-break':
        case 'office:annotation':
        case 'office:annotation-end':
        case 'text:alphabetical-index-mark':
        case 'text:toc-mark':
          break;
        default:
          // Fields (dates, titles, references, …) and unknown inline wrappers: keep their visible text
          this.inline(n.children, base, link, out);
      }
    }
  }

  private mergeText(base: TextProps, styleName: string | undefined): TextProps {
    if (!styleName) return base;
    const def = this.styles.defaults.get('text');
    const own = Object.assign({}, ...this.styles.chain(styleName, 'text').filter((x) => x !== def).map((x) => x.text)) as TextProps;
    return { ...base, ...own };
  }

  private drawing(el: XNode): Block[] {
    const out: Block[] = [];
    if (el.name === 'draw:frame') {
      const img = childrenNamed(el, 'draw:image')[0];
      const textBox = childrenNamed(el, 'draw:text-box')[0];
      if (img) {
        const block = this.image(img, el);
        if (block) out.push(block);
      } else if (textBox) out.push(...this.blocks(textBox.children));
      return out;
    }
    if (el.name === 'draw:custom-shape' || el.name === 'draw:rect') {
      out.push(...this.blocks(el.children));
      return out;
    }
    for (const c of elements(el)) out.push(...this.drawing(c));
    return out;
  }

  private image(img: XNode, frame: XNode): Block | null {
    let data: Uint8Array | undefined;
    const href = img.attrs['xlink:href'];
    if (href && !/^[a-z]+:/i.test(href)) data = this.files[href.replace(/^\.\//, '')];
    const bin = childrenNamed(img, 'office:binary-data')[0];
    if (!data && bin) {
      try {
        data = base64ToBytes(textContent(bin).replace(/\s+/g, ''));
      } catch {
        data = undefined;
      }
    }
    if (!data) return null;
    let mime = sniffImageMime(data);
    if (!mime && /^\s*(<\?xml|<svg)/.test(decoder.decode(data.subarray(0, 200)))) mime = 'image/svg+xml';
    if (!mime) {
      // A frame may carry several renditions (e.g. SVG + PNG fallback): try the next one
      const alt = childrenNamed(frame, 'draw:image').find((x) => x !== img);
      if (alt) return this.image(alt, { ...frame, children: frame.children.filter((c) => c !== img) });
      this.warnings.add('Некоторые изображения (EMF/WMF/TIFF) пропущены: формат не поддерживается.');
      return null;
    }
    const wPt = odfLength(frame.attrs['svg:width']);
    const hPt = odfLength(frame.attrs['svg:height']);
    let width: number;
    let height: number;
    if (wPt && hPt) {
      width = (wPt * 96) / 72;
      height = (hPt * 96) / 72;
    } else {
      const sz = imageSize(data) ?? { width: 300, height: 200 };
      width = sz.width;
      height = sz.height;
    }
    const alt = textContent(childrenNamed(frame, 'svg:title')[0] ?? childrenNamed(frame, 'svg:desc')[0]).trim() || frame.attrs['draw:name'] || undefined;
    return { type: 'image', data, mime, width: Math.round(width), height: Math.round(height), alt };
  }

  private list(el: XNode, depth: number, inherited?: string): Block[] {
    const styleName = el.attrs['text:style-name'] ?? inherited;
    const levels = styleName ? this.styles.lists.get(styleName) : undefined;
    const lv = levels?.[Math.min(depth, 9)];
    const items: ListItem[] = [];
    let start: number | undefined;
    for (const it of elements(el)) {
      if (it.name !== 'text:list-item' && it.name !== 'text:list-header') continue;
      if (start === undefined && it.attrs['text:start-value']) start = +it.attrs['text:start-value'];
      const blocks = this.blocks(it.children, depth + 1, styleName);
      const prev = items[items.length - 1];
      if (prev && blocks.length && blocks.every((x) => x.type === 'list')) prev.blocks.push(...blocks);
      else items.push({ blocks: blocks.length ? blocks : [{ type: 'paragraph', runs: [] }] });
    }
    if (!items.length) return [];
    // A "list" of a single heading is how ODF numbers chapters: unwrap it
    if (items.every((i) => i.blocks.length === 1 && i.blocks[0].type === 'heading')) return items.map((i) => i.blocks[0]);
    const ordered = lv?.ordered ?? false;
    const list: Block = { type: 'list', ordered, items };
    const s = start ?? lv?.start ?? 1;
    if (ordered && s !== 1) list.start = s;
    return [list];
  }

  private table(el: XNode): Block {
    const rows: TableRow[] = [];
    const visit = (node: XNode, header: boolean): void => {
      for (const c of elements(node)) {
        if (c.name === 'table:table-header-rows') visit(c, true);
        else if (c.name === 'table:table-rows' || c.name === 'table:table-row-group') visit(c, header);
        else if (c.name === 'table:table-row') {
          const repeat = Math.min(100, Math.max(1, +(c.attrs['table:number-rows-repeated'] ?? 1) || 1));
          const cells = [];
          for (const cell of elements(c)) {
            if (cell.name !== 'table:table-cell') continue; // covered cells belong to a span
            const span = +(cell.attrs['table:number-columns-spanned'] ?? 1) || 1;
            const rep = Math.min(50, Math.max(1, +(cell.attrs['table:number-columns-repeated'] ?? 1) || 1));
            for (let k = 0; k < rep; k++) {
              this.inHeaderRow = header;
              const blocks = this.blocks(cell.children);
              this.inHeaderRow = false;
              const tc: TableRow['cells'][number] = { blocks };
              if (header) tc.header = true;
              if (span > 1) tc.colspan = span;
              cells.push(tc);
            }
          }
          for (let k = 0; k < repeat; k++) rows.push({ cells: k ? cells.map((x) => ({ ...x })) : cells });
        }
      }
    };
    visit(el, false);
    // Drop trailing empty columns/rows that come from repeated empty cells
    while (rows.length && rows[rows.length - 1].cells.every((c) => !c.blocks.length)) rows.pop();
    return { type: 'table', rows };
  }
}

function openPackage(bytes: Uint8Array): Record<string, Uint8Array> {
  try {
    return unzipOffice(bytes, unzipSync);
  } catch (e) {
    if (e instanceof ConvertError) throw e;
    throw new ConvertError('Файл OpenDocument повреждён.');
  }
}

function checkEncryption(files: Record<string, Uint8Array>): void {
  const manifest = files['META-INF/manifest.xml'];
  if (manifest && decoder.decode(manifest).includes('encryption-data')) throw new ConvertError('Документ OpenDocument защищён паролем. Снимите защиту и попробуйте снова.', 'password');
}

export function odtToDoc(bytes: Uint8Array): { doc: Doc; warnings: string[] } {
  const files = openPackage(bytes);
  checkEncryption(files);
  const r = new OdtReader(files);
  const doc = r.read();
  return { doc, warnings: [...r.warnings] };
}

// ---------------------------------------------------------------------------
// Spreadsheets

const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
const MAX_CELLS = 5_000_000;

function cellText(cell: XNode): string {
  const paras = childrenNamed(cell, 'text:p');
  const one = (n: XNode | string): string => {
    if (typeof n === 'string') return n.replace(/[ \t\r\n]+/g, ' ');
    if (n.name === 'text:s') return ' '.repeat(Math.min(1000, +(n.attrs['text:c'] ?? 1) || 1));
    if (n.name === 'text:tab') return '\t';
    if (n.name === 'text:line-break') return '\n';
    if (n.name === 'office:annotation') return '';
    return n.children.map(one).join('');
  };
  return paras.map((p) => p.children.map(one).join('')).join('\n');
}

function isoDate(v: string): string {
  const m = /^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2}:\d{2}))?/.exec(v);
  if (!m) return v;
  return m[2] && m[2] !== '00:00:00' ? `${m[1]} ${m[2]}` : m[1];
}

function isoTime(v: string): string {
  const m = /^-?PT(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(v);
  if (!m) return v;
  const pad = (n: number): string => String(n).padStart(2, '0');
  return `${pad(+(m[1] ?? 0))}:${pad(+(m[2] ?? 0))}:${pad(Math.round(+(m[3] ?? 0)))}`;
}

function cellValue(cell: XNode): CellValue {
  const a = cell.attrs;
  const type = a['office:value-type'] ?? a['calcext:value-type'];
  switch (type) {
    case 'float':
    case 'percentage':
    case 'currency': {
      const n = Number(a['office:value']);
      return Number.isFinite(n) ? n : cellText(cell);
    }
    case 'date':
      return a['office:date-value'] ? isoDate(a['office:date-value']) : cellText(cell);
    case 'time':
      return a['office:time-value'] ? isoTime(a['office:time-value']) : cellText(cell);
    case 'boolean':
      return a['office:boolean-value'] === 'true';
    case 'string':
      return a['office:string-value'] ?? cellText(cell);
    default: {
      const t = cellText(cell);
      return t ? t : null;
    }
  }
}

export function odsToBook(bytes: Uint8Array): { book: Workbook; warnings: string[] } {
  const files = openPackage(bytes);
  checkEncryption(files);
  const content = files['content.xml'] ? parseXml(decoder.decode(files['content.xml'])) : undefined;
  const sheetRoot = find(content, 'office:spreadsheet');
  if (!sheetRoot) {
    if (find(content, 'office:text')) throw new ConvertError('Это текстовый документ OpenDocument (.odt), а не таблица.');
    throw new ConvertError('В файле нет таблицы OpenDocument (office:spreadsheet).');
  }
  const warnings: string[] = [];
  const sheets: Sheet[] = [];
  let cellsTotal = 0;
  for (const t of childrenNamed(sheetRoot, 'table:table')) {
    const rows: CellValue[][] = [];
    let header = false;
    let pendingEmptyRows = 0;
    let truncated = false;
    const visit = (node: XNode, inHeader: boolean): void => {
      for (const r of elements(node)) {
        if (truncated) return;
        if (r.name === 'table:table-header-rows') {
          header = true;
          visit(r, true);
          continue;
        }
        if (r.name === 'table:table-rows' || r.name === 'table:table-row-group') {
          visit(r, inHeader);
          continue;
        }
        if (r.name !== 'table:table-row') continue;
        const row: CellValue[] = [];
        let pendingEmpty = 0;
        for (const c of elements(r)) {
          if (c.name !== 'table:table-cell' && c.name !== 'table:covered-table-cell') continue;
          const rep = Math.max(1, +(c.attrs['table:number-columns-repeated'] ?? 1) || 1);
          const v = c.name === 'table:covered-table-cell' ? null : cellValue(c);
          if (v === null || v === '') {
            pendingEmpty += rep;
            continue;
          }
          for (; pendingEmpty > 0 && row.length < MAX_COLS; pendingEmpty--) row.push(null);
          pendingEmpty = 0;
          for (let k = 0; k < rep && row.length < MAX_COLS; k++) row.push(v);
        }
        const rep = Math.max(1, +(r.attrs['table:number-rows-repeated'] ?? 1) || 1);
        if (!row.length) {
          pendingEmptyRows += rep;
          continue;
        }
        for (; pendingEmptyRows > 0 && rows.length < MAX_ROWS; pendingEmptyRows--) rows.push([]);
        pendingEmptyRows = 0;
        for (let k = 0; k < rep && rows.length < MAX_ROWS; k++) {
          cellsTotal += row.length;
          if (cellsTotal > MAX_CELLS) {
            truncated = true;
            warnings.push('Таблица слишком большая: прочитаны первые 5 млн ячеек.');
            return;
          }
          rows.push(k ? row.slice() : row);
        }
      }
    };
    visit(t, false);
    sheets.push({ name: t.attrs['table:name'] || `Лист${sheets.length + 1}`, rows, ...(header ? { header: true } : {}) });
    if (truncated) break;
  }
  if (!sheets.length) throw new ConvertError('В таблице OpenDocument не найдено листов.');
  return { book: { sheets }, warnings };
}

/** Kind of an OpenDocument package from its `mimetype` entry. */
export function odfKind(files: Record<string, Uint8Array>): 'odt' | 'ods' | 'odp' | null {
  const m = files['mimetype'] ? decoder.decode(files['mimetype']).trim() : '';
  if (m.startsWith('application/vnd.oasis.opendocument.text')) return 'odt';
  if (m.startsWith('application/vnd.oasis.opendocument.spreadsheet')) return 'ods';
  if (m.startsWith('application/vnd.oasis.opendocument.presentation')) return 'odp';
  return null;
}
