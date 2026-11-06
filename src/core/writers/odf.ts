/**
 * OpenDocument writers: Doc → .odt and Workbook → .ods (ODF 1.3, as produced by
 * LibreOffice). Formatting goes to automatic styles generated on demand and
 * de-duplicated by their properties.
 */
import { zipSync, type Zippable } from 'fflate';
import type { Block, CellValue, Doc, FontFamily, PageSetup, ParaFormat, Sheet, TableCell, TextRun, Workbook } from '../model';
import { escapeXml, maxOf, utf8 } from '../util';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8"?>\n';
const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ' +
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"';
const FONT_NAME: Record<FontFamily, string> = { sans: 'Arial', serif: 'Times New Roman', mono: 'Courier New', calibri: 'Calibri' };
const FONT_DECLS =
  '<office:font-face-decls>' +
  '<style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss" style:font-pitch="variable"/>' +
  `<style:font-face style:name="Times New Roman" svg:font-family="'Times New Roman'" style:font-family-generic="roman" style:font-pitch="variable"/>` +
  `<style:font-face style:name="Courier New" svg:font-family="'Courier New'" style:font-family-generic="modern" style:font-pitch="fixed"/>` +
  '<style:font-face style:name="Calibri" svg:font-family="Calibri" style:font-family-generic="swiss" style:font-pitch="variable"/>' +
  '</office:font-face-decls>';
const IMAGE_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/emf': 'emf', 'image/wmf': 'wmf' };
const BULLETS = ['•', '◦', '▪'];

const pt = (v: number): string => `${Math.round(v * 100) / 100}pt`;
const cm = (v: number): string => `${Math.round((v / 28.3465) * 1000) / 1000}cm`;

/** Generates de-duplicated automatic styles: identical properties share one name. */
class AutoStyles {
  private byKey = new Map<string, string>();
  private xml: string[] = [];
  private counters: Record<string, number> = {};
  constructor(private prefix: string) {}
  get(family: string, letter: string, body: string, parent?: string, attrs = ''): string {
    const key = `${family}|${parent ?? ''}|${attrs}|${body}`;
    let name = this.byKey.get(key);
    if (!name) {
      this.counters[letter] = (this.counters[letter] ?? 0) + 1;
      name = `${this.prefix}${letter}${this.counters[letter]}`;
      this.byKey.set(key, name);
      this.xml.push(`<style:style style:name="${name}" style:family="${family}"${parent ? ` style:parent-style-name="${parent}"` : ''}${attrs}>${body}</style:style>`);
    }
    return name;
  }
  add(raw: string): void {
    this.xml.push(raw);
  }
  toString(): string {
    return this.xml.join('');
  }
}

function textProps(r: Omit<TextRun, 'text'>): string {
  let a = '';
  const family: FontFamily | undefined = r.code && !r.font ? 'mono' : r.font;
  if (family) a += ` style:font-name="${FONT_NAME[family]}"`;
  if (r.size) a += ` fo:font-size="${pt(r.size)}"`;
  if (r.bold) a += ' fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"';
  if (r.italic) a += ' fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"';
  if (r.underline) a += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
  if (r.strike) a += ' style:text-line-through-style="solid" style:text-line-through-type="single"';
  if (r.sup) a += ' style:text-position="super 58%"';
  else if (r.sub) a += ' style:text-position="sub 58%"';
  if (r.color) a += ` fo:color="#${r.color}"`;
  if (r.highlight) a += ` fo:background-color="#${r.highlight}"`;
  return a ? `<style:text-properties${a}/>` : '';
}

function paraProps(f: ParaFormat, extra = ''): string {
  let a = '';
  if (f.align) a += ` fo:text-align="${f.align === 'left' ? 'start' : f.align === 'right' ? 'end' : f.align}"`;
  if (f.indent !== undefined) a += ` fo:margin-left="${cm(f.indent)}"`;
  if (f.firstLine !== undefined) a += ` fo:text-indent="${cm(f.firstLine)}"`;
  if (f.spaceBefore !== undefined) a += ` fo:margin-top="${cm(f.spaceBefore)}"`;
  if (f.spaceAfter !== undefined) a += ` fo:margin-bottom="${cm(f.spaceAfter)}"`;
  if (f.lineHeight !== undefined) a += ` fo:line-height="${Math.round(f.lineHeight * 100)}%"`;
  a += extra;
  return a ? `<style:paragraph-properties${a}/>` : '';
}

class OdtWriter {
  warnings = new Set<string>();
  private pictures: Record<string, Uint8Array> = {};
  private imageCount = 0;
  private tableCount = 0;
  private noteCount = 0;
  private listStyles = new Map<string, string>();
  private pendingBreak = false;
  /** A section with its own page: the master page its first paragraph switches to */
  private pendingMaster?: string;
  private masters: { name: string; page: PageSetup }[] = [];
  private inMaster = false;
  private auto: AutoStyles;
  readonly pageW: number;
  readonly pageH: number;
  readonly margins: { top: number; right: number; bottom: number; left: number };

  constructor(
    private doc: Doc,
    opts: { pageSize?: 'auto' | 'a4' | 'letter' },
  ) {
    const auto = (opts.pageSize ?? 'auto') === 'auto';
    if (auto && doc.page) {
      this.pageW = doc.page.width;
      this.pageH = doc.page.height;
    } else {
      const letter = opts.pageSize === 'letter';
      this.pageW = letter ? 612 : 595.3;
      this.pageH = letter ? 792 : 841.9;
    }
    this.margins = doc.page?.margins ?? { top: 56.7, right: 56.7, bottom: 56.7, left: 56.7 };
    this.auto = new AutoStyles('');
  }

  get contentW(): number {
    return this.pageW - this.margins.left - this.margins.right;
  }

  // -- inline -------------------------------------------------------------

  private textXml(text: string, state: { space: boolean }): string {
    let s = '';
    for (const ch of text) {
      if (ch === ' ') {
        s += state.space ? '<text:s/>' : ' ';
        state.space = true;
      } else if (ch === '\t') {
        s += '<text:tab/>';
        state.space = true;
      } else if (ch === '\n') {
        s += '<text:line-break/>';
        state.space = true;
      } else {
        s += ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '&' ? '&amp;' : ch === '"' ? '&quot;' : ch;
        state.space = false;
      }
    }
    return s.replace(/(<text:s\/>)+/g, (m) => {
      const n = m.length / 9;
      return n > 1 ? `<text:s text:c="${n}"/>` : m;
    });
  }

  private run(r: TextRun, state: { space: boolean }): string {
    const { text: _t, link: _l, note: _n, field: _f, ...style } = r;
    const props = textProps(style);
    const wrap = (inner: string): string => (props ? `<text:span text:style-name="${this.auto.get('text', 'T', props)}">${inner}</text:span>` : inner);
    if (r.field) {
      state.space = false;
      return wrap(r.field === 'page' ? '<text:page-number text:select-page="current">1</text:page-number>' : '<text:page-count>1</text:page-count>');
    }
    if (r.note && this.inMaster) {
      this.warnings.add('Сноски в колонтитулах не поддерживаются OpenDocument и пропущены.');
      return '';
    }
    if (r.note) {
      const id = ++this.noteCount;
      state.space = false;
      const saved = this.pendingBreak;
      this.pendingBreak = false;
      const body = this.blocks(r.note.length ? r.note : [{ type: 'paragraph', runs: [] }], 0, this.contentW, 'Footnote');
      this.pendingBreak = saved;
      return `<text:note text:id="ftn${id}" text:note-class="footnote"><text:note-citation>${id}</text:note-citation><text:note-body>${body}</text:note-body></text:note>`;
    }
    return r.text ? wrap(this.textXml(r.text, state)) : '';
  }

  private runs(runs: TextRun[]): string {
    const state = { space: true };
    let s = '';
    let i = 0;
    while (i < runs.length) {
      const link = runs[i].link;
      if (link && !runs[i].note && !runs[i].field) {
        let inner = '';
        while (i < runs.length && runs[i].link === link && !runs[i].note) inner += this.run(runs[i++], state);
        s += `<text:a xlink:type="simple" xlink:href="${escapeXml(link)}" text:style-name="Internet_20_link" text:visited-style-name="Visited_20_Internet_20_Link">${inner}</text:a>`;
      } else s += this.run(runs[i++], state);
    }
    return s;
  }

  // -- blocks -------------------------------------------------------------

  /** Paragraph style: the named parent plus own formatting (and a pending page break). */
  private pStyle(parent: string, f: ParaFormat = {}, indent = 0): string {
    const fmt: ParaFormat = { align: f.align, firstLine: f.firstLine, spaceBefore: f.spaceBefore, spaceAfter: f.spaceAfter, lineHeight: f.lineHeight };
    if (f.indent !== undefined || indent) fmt.indent = (f.indent ?? 0) + indent;
    let extra = '';
    if (this.pendingBreak) {
      extra = ' fo:break-before="page"';
      this.pendingBreak = false;
    }
    // switching the master page starts a new page with that page's size
    const master = this.pendingMaster ? ` style:master-page-name="${this.pendingMaster}"` : '';
    this.pendingMaster = undefined;
    const props = paraProps(fmt, extra);
    if (master) return this.auto.get('paragraph', 'P', props || '<style:paragraph-properties/>', parent, master);
    return props ? this.auto.get('paragraph', 'P', props, parent) : parent;
  }

  private listStyle(ordered: boolean, start: number): string {
    const key = `${ordered}|${start}`;
    let name = this.listStyles.get(key);
    if (name) return name;
    name = `L${this.listStyles.size + 1}`;
    this.listStyles.set(key, name);
    let levels = '';
    for (let lv = 1; lv <= 10; lv++) {
      const indent = cm(18 * lv);
      const props = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${indent}" fo:text-indent="-${cm(18)}" fo:margin-left="${indent}"/></style:list-level-properties>`;
      if (ordered) levels += `<text:list-level-style-number text:level="${lv}" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1"${lv === 1 && start !== 1 ? ` text:start-value="${start}"` : ''}>${props}</text:list-level-style-number>`;
      else levels += `<text:list-level-style-bullet text:level="${lv}" text:style-name="Bullet_20_Symbols" text:bullet-char="${BULLETS[(lv - 1) % 3]}">${props}</text:list-level-style-bullet>`;
    }
    this.auto.add(`<text:list-style style:name="${name}">${levels}</text:list-style>`);
    return name;
  }

  private image(b: Extract<Block, { type: 'image' }>, maxW: number): string {
    const ext = IMAGE_EXT[b.mime];
    if (!ext) {
      this.warnings.add('Некоторые изображения пропущены: формат не поддерживается OpenDocument.');
      return '';
    }
    const n = ++this.imageCount;
    const path = `Pictures/image${n}.${ext}`;
    this.pictures[path] = b.data;
    const maxH = this.pageH - this.margins.top - this.margins.bottom - 30;
    let w = (Math.max(1, b.width) * 72) / 96;
    let h = (Math.max(1, b.height) * 72) / 96;
    const k = Math.min(1, maxW / w, maxH / h);
    w *= k;
    h *= k;
    const look = 'style:mirror="none" fo:clip="rect(0cm, 0cm, 0cm, 0cm)" draw:luminance="0%" draw:contrast="0%" draw:red="0%" draw:green="0%" draw:blue="0%" draw:gamma="100%" draw:color-inversion="false" draw:image-opacity="100%" draw:color-mode="standard"';
    const title = b.alt ? `<svg:title>${escapeXml(b.alt)}</svg:title>` : '';
    let fr: string;
    let anchor = 'text:anchor-type="as-char"';
    const f = b.float;
    if (f) {
      // a floating frame anchored to the paragraph: position, wrapping and distance in its graphic style
      const H: Record<string, string> = { page: 'page', margin: 'page-content', column: 'paragraph', character: 'char', leftMargin: 'page-start-margin', rightMargin: 'page-end-margin' };
      const V: Record<string, string> = { page: 'page', margin: 'page-content', paragraph: 'paragraph', line: 'paragraph', topMargin: 'page', bottomMargin: 'page' };
      const HA: Record<string, string> = { left: 'left', center: 'center', right: 'right', inside: 'inside', outside: 'outside' };
      const VA: Record<string, string> = { top: 'top', center: 'middle', bottom: 'bottom', inside: 'top', outside: 'bottom' };
      const hp = f.h.align ? HA[f.h.align] : 'from-left';
      const vp = f.v.align && f.v.from !== 'paragraph' && f.v.from !== 'line' ? VA[f.v.align] : 'from-top';
      const wrap = f.wrap === 'square' ? ({ left: 'left', right: 'right', largest: 'dynamic', both: 'parallel' } as const)[f.side ?? 'both'] : f.wrap === 'topBottom' ? 'none' : 'run-through';
      const [t, r, bt, l] = f.dist ?? [0, 9, 0, 9];
      const props =
        `<style:graphic-properties style:wrap="${wrap}"${f.wrap === 'none' ? ` style:run-through="${f.behind ? 'background' : 'foreground'}"` : ''} style:number-wrapped-paragraphs="no-limit" style:wrap-contour="false" ` +
        `fo:margin-top="${cm(t)}" fo:margin-right="${cm(r)}" fo:margin-bottom="${cm(bt)}" fo:margin-left="${cm(l)}" ` +
        `style:horizontal-pos="${hp}" style:horizontal-rel="${H[f.h.from] ?? 'paragraph'}" style:vertical-pos="${vp}" style:vertical-rel="${V[f.v.from] ?? 'paragraph'}" ${look}/>`;
      fr = this.auto.get('graphic', 'fr', props, 'Graphics');
      anchor = `text:anchor-type="paragraph"${hp === 'from-left' ? ` svg:x="${cm(f.h.offset ?? 0)}"` : ''}${vp === 'from-top' ? ` svg:y="${cm(f.v.offset ?? 0)}"` : ''}`;
    } else fr = this.auto.get('graphic', 'fr', `<style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline" ${look}/>`, 'Graphics');
    return (
      `<draw:frame draw:style-name="${fr}" draw:name="Image${n}" ${anchor} svg:width="${cm(w)}" svg:height="${cm(h)}" draw:z-index="${n}">` +
      `<draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="${b.mime === 'image/emf' || b.mime === 'image/wmf' ? b.mime.replace('/', '/x-') : b.mime}"/>${title}</draw:frame>`
    );
  }

  private table(b: Extract<Block, { type: 'table' }>, width: number, cellStyle: string): string {
    const cols = maxOf(b.rows.map((r) => r.cells.reduce((s, c) => s + (c.colspan ?? 1), 0)), 1);
    const weights = new Array(cols).fill(1);
    for (const r of b.rows) {
      let c = 0;
      for (const cell of r.cells) {
        const span = cell.colspan ?? 1;
        if (span === 1) {
          const len = cell.blocks.reduce((s, x) => s + (x.type === 'paragraph' || x.type === 'heading' ? x.runs.reduce((a, rr) => a + rr.text.length, 0) : 10), 0);
          weights[c] = Math.max(weights[c], Math.min(40, Math.max(4, len)));
        }
        c += span;
      }
    }
    const total = weights.reduce((a, w) => a + w, 0);
    const n = ++this.tableCount;
    const name = `Table${n}`;
    let extra = '';
    if (this.pendingBreak) {
      extra = ' fo:break-before="page"';
      this.pendingBreak = false;
    }
    const master = this.pendingMaster ? ` style:master-page-name="${this.pendingMaster}"` : '';
    this.pendingMaster = undefined;
    this.auto.add(`<style:style style:name="${name}" style:family="table"${master}><style:table-properties style:width="${cm(width)}" table:align="margins"${extra}/></style:style>`);
    let s = `<table:table table:name="${name}" table:style-name="${name}">`;
    weights.forEach((w, i) => {
      const col = `${name}.C${i + 1}`;
      this.auto.add(`<style:style style:name="${col}" style:family="table-column"><style:table-column-properties style:column-width="${cm((w / total) * width)}"/></style:style>`);
      s += `<table:table-column table:style-name="${col}"/>`;
    });
    const cellBase = '<style:table-cell-properties fo:padding="0.097cm" fo:border="0.5pt solid #000000"/>';
    const plain = this.auto.get('table-cell', 'TC', cellBase);
    const head = this.auto.get('table-cell', 'TC', cellBase.replace('<style:table-cell-properties', '<style:table-cell-properties fo:background-color="#f2f2f2"'));
    const headerRows = b.rows.filter((r, i) => i === 0 && r.cells.length > 0 && r.cells.every((c) => c.header));
    const row = (r: (typeof b.rows)[number]): string => {
      let x = '<table:table-row>';
      const cells: TableCell[] = [...r.cells];
      let used = cells.reduce((a, c) => a + (c.colspan ?? 1), 0);
      while (used < cols) {
        cells.push({ blocks: [] });
        used++;
      }
      let c = 0;
      for (const cell of cells) {
        const span = Math.min(cell.colspan ?? 1, cols - c);
        if (span <= 0) break;
        const w = weights.slice(c, c + span).reduce((a, v) => a + v, 0);
        c += span;
        const inner = this.blocks(cell.blocks.length ? cell.blocks : [{ type: 'paragraph', runs: [] }], 0, (w / total) * width - 8, cell.header ? 'Table_20_Heading' : cellStyle);
        x += `<table:table-cell table:style-name="${cell.header ? head : plain}" office:value-type="string"${span > 1 ? ` table:number-columns-spanned="${span}"` : ''}>${inner}</table:table-cell>`;
        for (let k = 1; k < span; k++) x += '<table:covered-table-cell/>';
      }
      return x + '</table:table-row>';
    };
    if (headerRows.length) s += `<table:table-header-rows>${headerRows.map(row).join('')}</table:table-header-rows>`;
    for (const r of b.rows.slice(headerRows.length)) s += row(r);
    return s + '</table:table>';
  }

  blocks(blocks: Block[], indent: number, width: number, paraStyle = 'Standard'): string {
    let s = '';
    // floating frames are anchored in the paragraph that follows them
    let floats = '';
    for (const b of blocks) {
      if (b.type === 'image' && b.float) {
        floats += this.image(b, width - indent);
        continue;
      }
      if (floats && b.type !== 'paragraph' && b.type !== 'heading') {
        s += `<text:p text:style-name="${this.pStyle(paraStyle, {}, indent)}">${floats}</text:p>`;
        floats = '';
      }
      switch (b.type) {
        case 'paragraph':
          s += `<text:p text:style-name="${this.pStyle(paraStyle, b, indent)}">${floats}${this.runs(b.runs)}</text:p>`;
          floats = '';
          break;
        case 'heading':
          s += `<text:h text:style-name="${this.pStyle(`Heading_20_${b.level}`, b, indent)}" text:outline-level="${b.level}">${floats}${this.runs(b.runs)}</text:h>`;
          floats = '';
          break;
        case 'list': {
          const style = this.listStyle(b.ordered, b.start ?? 1);
          s += `<text:list text:style-name="${style}">`;
          for (const it of b.items) {
            const inner = it.blocks.length ? it.blocks : [{ type: 'paragraph', runs: [] } as Block];
            // List paragraphs take their indentation from the list style
            s += `<text:list-item>${this.blocks(inner, 0, width - 36, paraStyle === 'Standard' ? 'List_20_Paragraph' : paraStyle)}</text:list-item>`;
          }
          s += '</text:list>';
          break;
        }
        case 'table':
          s += this.table(b, Math.max(72, width - indent), 'Table_20_Contents');
          break;
        case 'image': {
          const align: ParaFormat = { align: (b.width * 72) / 96 < width * 0.9 ? undefined : 'center' };
          const frame = this.image(b, width - indent);
          if (frame) s += `<text:p text:style-name="${this.pStyle(paraStyle === 'Standard' ? 'Standard' : paraStyle, align, indent)}">${frame}</text:p>`;
          break;
        }
        case 'code':
          for (const line of b.text.split('\n')) s += `<text:p text:style-name="${this.pStyle('Preformatted_20_Text', {}, indent)}">${this.textXml(line, { space: true })}</text:p>`;
          break;
        case 'quote':
          s += this.blocks(b.blocks, indent, width, 'Quotations');
          break;
        case 'hr':
          s += `<text:p text:style-name="${this.pStyle('Horizontal_20_Line', {}, indent)}"/>`;
          break;
        case 'pagebreak':
          this.pendingBreak = true;
          break;
      }
    }
    return s;
  }

  private stylesXml(header: string, footer: string, autoForMaster: string): string {
    const doc = this.doc;
    const font = FONT_NAME[doc.font ?? 'sans'];
    const size = doc.size ?? 11;
    const d = doc.defaults ?? {};
    const std = paraProps({ spaceBefore: d.spaceBefore ?? 0, spaceAfter: d.spaceAfter ?? 6, lineHeight: d.lineHeight ?? 1.15 });
    const heading = (lv: number, k: number, before: number): string =>
      `<style:style style:name="Heading_20_${lv}" style:display-name="Heading ${lv}" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Standard" style:default-outline-level="${lv}" style:class="text">` +
      `<style:paragraph-properties fo:margin-top="${cm(before)}" fo:margin-bottom="${cm(6)}" fo:keep-with-next="always"/>` +
      `<style:text-properties fo:font-size="${pt(size * k)}" fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold" fo:color="#1f2937"/></style:style>`;
    const m = this.margins;
    const hasHeader = !!header;
    const hasFooter = !!footer;
    // In ODF the header lives inside the page margin area: shrink the margin by the header block
    const hf = 0.8 * 28.3465;
    const layoutOf = (name: string, w: number, h: number, mg: { top: number; right: number; bottom: number; left: number }): string =>
      `<style:page-layout style:name="${name}"><style:page-layout-properties fo:page-width="${cm(w)}" fo:page-height="${cm(h)}" style:print-orientation="${w > h ? 'landscape' : 'portrait'}" ` +
      `fo:margin-top="${cm(hasHeader ? Math.max(14, mg.top - hf - 14) : mg.top)}" fo:margin-bottom="${cm(hasFooter ? Math.max(14, mg.bottom - hf - 14) : mg.bottom)}" fo:margin-left="${cm(mg.left)}" fo:margin-right="${cm(mg.right)}" style:writing-mode="lr-tb"/>` +
      `<style:header-style>${hasHeader ? `<style:header-footer-properties fo:min-height="${cm(hf)}" fo:margin-bottom="${cm(14)}" fo:margin-left="0cm" fo:margin-right="0cm"/>` : ''}</style:header-style>` +
      `<style:footer-style>${hasFooter ? `<style:header-footer-properties fo:min-height="${cm(hf)}" fo:margin-top="${cm(14)}" fo:margin-left="0cm" fo:margin-right="0cm"/>` : ''}</style:footer-style></style:page-layout>`;
    const layout = layoutOf('pm1', this.pageW, this.pageH, m) + this.masters.map((x, i) => layoutOf(`pm${i + 2}`, x.page.width, x.page.height, x.page.margins)).join('');
    const hf_ = `${hasHeader ? `<style:header>${header}</style:header>` : ''}${hasFooter ? `<style:footer>${footer}</style:footer>` : ''}`;
    const extraMasters = this.masters.map((x, i) => `<style:master-page style:name="${x.name}" style:page-layout-name="pm${i + 2}">${hf_}</style:master-page>`).join('');
    return (
      XML_HEAD +
      `<office:document-styles ${NS}>` +
      FONT_DECLS +
      '<office:styles>' +
      `<style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="1.25cm" style:writing-mode="page"/><style:text-properties style:font-name="${font}" fo:font-size="${pt(size)}" fo:language="ru" fo:country="RU" style:font-name-asian="${font}" style:font-size-asian="${pt(size)}" style:font-name-complex="${font}" style:font-size-complex="${pt(size)}"/></style:default-style>` +
      '<style:default-style style:family="graphic"><style:graphic-properties style:wrap="none"/></style:default-style>' +
      '<style:default-style style:family="table"><style:table-properties table:border-model="collapsing"/></style:default-style>' +
      `<style:style style:name="Standard" style:family="paragraph" style:class="text">${std}</style:style>` +
      '<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:class="text"><style:paragraph-properties fo:keep-with-next="always"/></style:style>' +
      heading(1, 2, 18) +
      heading(2, 1.64, 15) +
      heading(3, 1.36, 12) +
      heading(4, 1.18, 10) +
      heading(5, 1.09, 10) +
      heading(6, 1, 8) +
      '<style:style style:name="List_20_Paragraph" style:display-name="List Paragraph" style:family="paragraph" style:parent-style-name="Standard" style:class="list"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.071cm"/></style:style>' +
      '<style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-left="1cm" fo:padding-left="0.28cm" fo:border-left="1.5pt solid #c8c8cc"/><style:text-properties fo:font-style="italic" fo:color="#404040"/></style:style>' +
      `<style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:line-height="100%" fo:background-color="#f3f4f6"/><style:text-properties style:font-name="Courier New" fo:font-size="9.5pt"/></style:style>` +
      '<style:style style:name="Horizontal_20_Line" style:display-name="Horizontal Line" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.25cm" fo:border-bottom="0.75pt solid #a0a0a0" fo:padding="0cm"/><style:text-properties fo:font-size="6pt"/></style:style>' +
      '<style:style style:name="Table_20_Contents" style:display-name="Table Contents" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:line-height="100%"/></style:style>' +
      '<style:style style:name="Table_20_Heading" style:display-name="Table Heading" style:family="paragraph" style:parent-style-name="Table_20_Contents" style:class="extra"><style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>' +
      `<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-left="0.6cm" fo:text-indent="-0.6cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:line-height="100%"/><style:text-properties fo:font-size="${pt(size * 0.82)}"/></style:style>` +
      '<style:style style:name="Header" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/></style:style>' +
      '<style:style style:name="Footer" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/></style:style>' +
      '<style:style style:name="Graphics" style:family="graphic"><style:graphic-properties text:anchor-type="as-char" svg:y="0cm" style:vertical-pos="top" style:vertical-rel="baseline"/></style:style>' +
      '<style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text"><style:text-properties fo:color="#0563c1" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>' +
      '<style:style style:name="Visited_20_Internet_20_Link" style:display-name="Visited Internet Link" style:family="text"><style:text-properties fo:color="#954f72" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>' +
      '<style:style style:name="Footnote_20_Symbol" style:display-name="Footnote Symbol" style:family="text"/>' +
      '<style:style style:name="Footnote_20_anchor" style:display-name="Footnote anchor" style:family="text"><style:text-properties style:text-position="super 58%"/></style:style>' +
      '<style:style style:name="Bullet_20_Symbols" style:display-name="Bullet Symbols" style:family="text"/>' +
      '<style:style style:name="Numbering_20_Symbols" style:display-name="Numbering Symbols" style:family="text"/>' +
      '<text:notes-configuration text:note-class="footnote" text:citation-style-name="Footnote_20_Symbol" text:citation-body-style-name="Footnote_20_anchor" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/>' +
      '</office:styles>' +
      `<office:automatic-styles>${layout}${autoForMaster}</office:automatic-styles>` +
      `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">${hf_}</style:master-page>${extraMasters}</office:master-styles>` +
      '</office:document-styles>'
    );
  }

  build(): Uint8Array {
    const doc = this.doc;
    const blocks = [...doc.blocks];
    while (blocks.length && (blocks[blocks.length - 1].type === 'pagebreak' || blocks[blocks.length - 1].type === 'section')) blocks.pop();
    while (blocks.length && blocks[0].type === 'pagebreak') blocks.shift();
    // Sections with columns become text:section elements with a column style
    type Sect = Extract<Block, { type: 'section' }>;
    const groups: { sec?: Sect; blocks: Block[] }[] = [{ blocks: [] }];
    for (const b of blocks) {
      if (b.type === 'section') groups.push({ sec: b, blocks: [] });
      else groups[groups.length - 1].blocks.push(b);
    }
    let body = '';
    let contentW = this.contentW;
    const docKey = JSON.stringify({ width: this.pageW, height: this.pageH, margins: this.margins });
    groups.forEach((g, k) => {
      if (!g.blocks.length) return;
      if (g.sec?.page) {
        // a page of its own: a master page with that page layout (the document's is "Standard")
        const key = JSON.stringify(g.sec.page);
        let name = 'Standard';
        if (key !== docKey) {
          let found = this.masters.find((x) => JSON.stringify(x.page) === key);
          if (!found) {
            found = { name: `Section_${this.masters.length + 1}`, page: g.sec.page };
            this.masters.push(found);
          }
          name = found.name;
        }
        this.pendingMaster = name;
        contentW = g.sec.page.width - g.sec.page.margins.left - g.sec.page.margins.right;
      } else if (g.sec?.start === 'page' && body) this.pendingBreak = true;
      const n = Math.max(1, Math.round(g.sec?.columns ?? 1));
      if (n < 2) {
        body += this.blocks(g.blocks, 0, contentW);
        return;
      }
      const gap = g.sec!.gap;
      // unequal columns: style:column elements with relative widths; the gaps are their indents
      let colEls = '';
      let narrowest = (contentW - (n - 1) * gap) / n;
      if (g.sec!.widths && g.sec!.widths.length === n) {
        const gaps = Array.from({ length: n - 1 }, (_, i) => g.sec!.gaps?.[i] ?? gap);
        const free = contentW - gaps.reduce((a, b) => a + b, 0);
        const sum = g.sec!.widths.reduce((a, b) => a + b, 0);
        const ws = g.sec!.widths.map((v) => (free * v) / sum);
        narrowest = Math.min(...ws);
        // (no fo:column-gap then: with it LibreOffice makes the columns equal); rel-widths add up to 65535
        colEls = ws
          .map((w, i) => {
            const start = i > 0 ? gaps[i - 1] / 2 : 0;
            const end = i < n - 1 ? gaps[i] / 2 : 0;
            return `<style:column style:rel-width="${Math.round(((w + start + end) / contentW) * 65535)}*" fo:start-indent="${Math.round(start * 100) / 100}pt" fo:end-indent="${Math.round(end * 100) / 100}pt"/>`;
          })
          .join('');
      }
      // like Word: balanced only when a continuous section follows
      const balance = groups[k + 1]?.sec?.start === 'continuous';
      const style = this.auto.get(
        'section',
        'Sect',
        `<style:section-properties text:dont-balance-text-columns="${balance ? 'false' : 'true'}" style:editable="false"><style:columns fo:column-count="${n}"${colEls ? '' : ` fo:column-gap="${Math.round(gap * 100) / 100}pt"`}>` +
          (g.sec!.separator ? '<style:column-sep style:width="0.5pt" style:color="#000000" style:height="100%" style:vertical-align="top"/>' : '') +
          colEls +
          '</style:columns></style:section-properties>',
      );
      body += `<text:section text:style-name="${style}" text:name="Section${k}">${this.blocks(g.blocks, 0, narrowest)}</text:section>`;
    });
    if (!body) body = '<text:p text:style-name="Standard"/>';
    const contentAuto = this.auto.toString();
    // Header/footer automatic styles belong to styles.xml
    this.auto = new AutoStyles('M');
    this.listStyles.clear();
    this.pendingBreak = false;
    this.inMaster = true;
    const header = doc.header?.length ? this.blocks(doc.header, 0, this.contentW, 'Header') : '';
    const footer = doc.footer?.length ? this.blocks(doc.footer, 0, this.contentW, 'Footer') : '';
    const masterAuto = this.auto.toString();
    const content =
      XML_HEAD +
      `<office:document-content ${NS}>` +
      FONT_DECLS +
      `<office:automatic-styles>${contentAuto}</office:automatic-styles>` +
      '<office:body><office:text><text:sequence-decls><text:sequence-decl text:display-outline-level="0" text:name="Illustration"/><text:sequence-decl text:display-outline-level="0" text:name="Table"/></text:sequence-decls>' +
      body +
      '</office:text></office:body></office:document-content>';
    const now = new Date().toISOString().replace(/\.\d+Z$/, '');
    const meta =
      XML_HEAD +
      `<office:document-meta ${NS}><office:meta><meta:generator>Salatic</meta:generator>` +
      (doc.title ? `<dc:title>${escapeXml(doc.title)}</dc:title>` : '') +
      `<meta:creation-date>${now}</meta:creation-date><dc:date>${now}</dc:date></office:meta></office:document-meta>`;
    const files: Zippable = {};
    files['mimetype'] = [utf8('application/vnd.oasis.opendocument.text'), { level: 0 }];
    files['content.xml'] = utf8(content);
    files['styles.xml'] = utf8(this.stylesXml(header, footer, masterAuto));
    files['meta.xml'] = utf8(meta);
    for (const [path, data] of Object.entries(this.pictures)) files[path] = [data, { level: 0 }];
    files['META-INF/manifest.xml'] = utf8(manifest('application/vnd.oasis.opendocument.text', Object.keys(this.pictures)));
    return zipSync(files, { level: 6 });
  }
}

function manifest(mime: string, pictures: string[]): string {
  const pic = (p: string): string => {
    const ext = p.split('.').pop() ?? '';
    const type = ext === 'jpg' ? 'image/jpeg' : ext === 'svg' ? 'image/svg+xml' : ext === 'emf' || ext === 'wmf' ? `image/x-${ext}` : `image/${ext}`;
    return `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="${type}"/>`;
  };
  return (
    XML_HEAD +
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">' +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mime}"/>` +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
    pictures.map(pic).join('') +
    '</manifest:manifest>'
  );
}

export function docToOdt(doc: Doc, opts: { pageSize?: 'auto' | 'a4' | 'letter' } = {}): { data: Uint8Array; warnings: string[] } {
  const w = new OdtWriter(doc, opts);
  const data = w.build();
  return { data, warnings: [...w.warnings] };
}

// ---------------------------------------------------------------------------
// Spreadsheets

const NUMERIC = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}):(\d{2}))?$/;

function cellXml(v: CellValue, style: string): string {
  const st = style ? ` table:style-name="${style}"` : '';
  if (v === null || v === '') return `<table:table-cell${st}/>`;
  if (typeof v === 'number' && !Number.isFinite(v)) return cellXml(String(v), style);
  if (typeof v === 'number') return `<table:table-cell${st} office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
  if (typeof v === 'boolean') return `<table:table-cell${st} office:value-type="boolean" office:boolean-value="${v}"><text:p>${v ? 'TRUE' : 'FALSE'}</text:p></table:table-cell>`;
  const t = v.trim();
  if (t === v && NUMERIC.test(t) && !/^[-+]?0\d/.test(t) && t.replace(/\D/g, '').length <= 15) return cellXml(Number(t), style);
  const d = ISO_DATE.exec(v);
  if (d && +d[2] >= 1 && +d[2] <= 12 && +d[3] >= 1 && +d[3] <= 31) {
    const iso = d[4] ? `${d[1]}-${d[2]}-${d[3]}T${d[4]}:${d[5]}:${d[6]}` : `${d[1]}-${d[2]}-${d[3]}`;
    return `<table:table-cell table:style-name="${d[4] ? 'ceDT' : 'ceD'}" office:value-type="date" office:date-value="${iso}"><text:p>${escapeXml(v)}</text:p></table:table-cell>`;
  }
  const paras = v
    .split('\n')
    .map((line) => `<text:p>${escapeXml(line).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`)}</text:p>`)
    .join('');
  return `<table:table-cell${st} office:value-type="string">${paras}</table:table-cell>`;
}

function sheetNames(sheets: Sheet[]): string[] {
  const used = new Set<string>();
  return sheets.map((s, i) => {
    const base = (s.name || `Лист${i + 1}`).replace(/[[\]:*?/\\']/g, ' ').trim().slice(0, 100) || `Лист${i + 1}`;
    let name = base;
    let k = 2;
    while (used.has(name.toLowerCase())) name = `${base} (${k++})`;
    used.add(name.toLowerCase());
    return name;
  });
}

export function bookToOds(book: Workbook): Uint8Array {
  const sheets = book.sheets.length ? book.sheets : [{ name: 'Лист1', rows: [] }];
  const names = sheetNames(sheets);
  let colStyles = '';
  let tables = '';
  sheets.forEach((sheet, si) => {
    const width = Math.max(1, maxOf(sheet.rows.map((r) => r.length), 0));
    const widths = new Array(width).fill(8);
    for (const r of sheet.rows)
      r.forEach((v, ci) => {
        if (v === null) return;
        const len = typeof v === 'string' ? maxOf(v.split('\n').map((l) => l.length), 0) : String(v).length;
        widths[ci] = Math.min(60, Math.max(widths[ci], len + 2));
      });
    let cols = '';
    widths.forEach((w, ci) => {
      const name = `co${si + 1}_${ci + 1}`;
      colStyles += `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="${Math.round(w * 0.19 * 1000) / 1000}cm"/></style:style>`;
      cols += `<table:table-column table:style-name="${name}" table:default-cell-style-name="Default"/>`;
    });
    let rows = '';
    sheet.rows.forEach((r, ri) => {
      const style = ri === 0 && sheet.header ? 'ceH' : '';
      let cells = '';
      for (let ci = 0; ci < Math.max(1, r.length); ci++) cells += cellXml(r[ci] ?? null, style);
      const row = `<table:table-row table:style-name="ro1">${cells}</table:table-row>`;
      rows += ri === 0 && sheet.header ? `<table:table-header-rows>${row}</table:table-header-rows>` : row;
    });
    if (!rows) rows = '<table:table-row table:style-name="ro1"><table:table-cell/></table:table-row>';
    tables += `<table:table table:name="${escapeXml(names[si])}" table:style-name="ta1">${cols}${rows}</table:table>`;
  });
  const content =
    XML_HEAD +
    `<office:document-content ${NS}>` +
    FONT_DECLS +
    '<office:automatic-styles>' +
    colStyles +
    '<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:use-optimal-row-height="true"/></style:style>' +
    '<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>' +
    '<number:date-style style:name="N36"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>' +
    '<number:date-style style:name="N37"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/><number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:date-style>' +
    '<style:style style:name="ceD" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N36"/>' +
    '<style:style style:name="ceDT" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>' +
    '<style:style style:name="ceH" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>' +
    '</office:automatic-styles>' +
    `<office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`;
  const styles =
    XML_HEAD +
    `<office:document-styles ${NS}>` +
    FONT_DECLS +
    '<office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Arial" fo:font-size="10pt" fo:language="ru" fo:country="RU"/></style:default-style>' +
    '<style:style style:name="Default" style:family="table-cell"/></office:styles>' +
    '<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties style:writing-mode="lr-tb"/></style:page-layout></office:automatic-styles>' +
    '<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';
  const files: Zippable = {};
  files['mimetype'] = [utf8('application/vnd.oasis.opendocument.spreadsheet'), { level: 0 }];
  files['content.xml'] = utf8(content);
  files['styles.xml'] = utf8(styles);
  files['meta.xml'] = utf8(XML_HEAD + `<office:document-meta ${NS}><office:meta><meta:generator>Salatic</meta:generator></office:meta></office:document-meta>`);
  files['META-INF/manifest.xml'] = utf8(manifest('application/vnd.oasis.opendocument.spreadsheet', []));
  return zipSync(files, { level: 6 });
}
