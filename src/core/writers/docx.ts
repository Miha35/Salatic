/** Doc → DOCX (Office Open XML). Element order follows the ECMA-376 schema so Word opens files without repair. */
import { zipSync, type Zippable } from 'fflate';
import type { Block, Doc, FontFamily, ParaFormat, TableCell, TextRun } from '../model';
import { escapeXml, maxOf, utf8 } from '../util';

export interface DocxOptions {
  pageSize?: 'auto' | 'a4' | 'letter';
}

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_ALL = `xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const EMU_PER_PX = 9525;
const IMAGE_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/emf': 'emf', 'image/wmf': 'wmf' };
const BULLETS = ['•', '◦', '▪'];
const FONT_NAME: Record<FontFamily, string> = { sans: 'Arial', serif: 'Times New Roman', mono: 'Courier New', calibri: 'Calibri' };
const tw = (pt: number): number => Math.round(pt * 20);

/** A package part with its own relationships (document, header, footer, footnotes). */
class Part {
  rels: string[] = [];
  private next: number;
  private links = new Map<string, string>();
  constructor(firstId: number) {
    this.next = firstId;
  }
  add(type: string, target: string, external = false): string {
    const id = `rId${this.next++}`;
    this.rels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${escapeXml(target)}"${external ? ' TargetMode="External"' : ''}/>`);
    return id;
  }
  link(url: string): string {
    let id = this.links.get(url);
    if (!id) {
      id = this.add('hyperlink', url, true);
      this.links.set(url, id);
    }
    return id;
  }
  xml(): string {
    return XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + this.rels.join('') + '</Relationships>';
  }
}

class DocxWriter {
  private media: Record<string, Uint8Array> = {};
  private imageCount = 0;
  private nums: string[] = [];
  private numId = 0;
  private part: Part;
  private notes: string[] = [];
  warnings: string[] = [];
  readonly pageW: number;
  readonly pageH: number;
  readonly margins: { top: number; right: number; bottom: number; left: number };

  constructor(
    private doc: Doc,
    opts: DocxOptions,
  ) {
    const auto = (opts.pageSize ?? 'auto') === 'auto';
    if (auto && doc.page) {
      this.pageW = tw(doc.page.width);
      this.pageH = tw(doc.page.height);
    } else {
      const letter = opts.pageSize === 'letter';
      this.pageW = letter ? 12240 : 11906;
      this.pageH = letter ? 15840 : 16838;
    }
    const m = doc.page?.margins;
    this.margins = m ? { top: tw(m.top), right: tw(m.right), bottom: tw(m.bottom), left: tw(m.left) } : { top: 1134, right: 1134, bottom: 1134, left: 1134 };
    this.part = new Part(10); // rId1..9 reserved for fixed document relationships
  }

  /** Usable width in twips. */
  get contentW(): number {
    return this.pageW - this.margins.left - this.margins.right;
  }

  private newNum(ordered: boolean, level: number, start = 1): number {
    const id = ++this.numId;
    const override = ordered ? `<w:lvlOverride w:ilvl="${level}"><w:startOverride w:val="${start}"/></w:lvlOverride>` : '';
    this.nums.push(`<w:num w:numId="${id}"><w:abstractNumId w:val="${ordered ? 1 : 0}"/>${override}</w:num>`);
    return id;
  }

  private rPr(r: TextRun, style?: string): string {
    let s = '';
    if (style) s += `<w:rStyle w:val="${style}"/>`;
    else if (r.link) s += '<w:rStyle w:val="Hyperlink"/>';
    const family: FontFamily | undefined = r.code && !r.font ? 'mono' : r.font;
    if (family) {
      const f = escapeXml(FONT_NAME[family]);
      s += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}" w:eastAsia="${f}"/>`;
    }
    if (r.bold) s += '<w:b/><w:bCs/>';
    if (r.italic) s += '<w:i/><w:iCs/>';
    if (r.strike) s += '<w:strike/>';
    if (r.color) s += `<w:color w:val="${r.color}"/>`;
    if (r.size) s += `<w:sz w:val="${Math.round(r.size * 2)}"/><w:szCs w:val="${Math.round(r.size * 2)}"/>`;
    if (r.underline) s += '<w:u w:val="single"/>';
    if (r.highlight) s += `<w:shd w:val="clear" w:color="auto" w:fill="${r.highlight}"/>`;
    if (!style) {
      if (r.sup) s += '<w:vertAlign w:val="superscript"/>';
      else if (r.sub) s += '<w:vertAlign w:val="subscript"/>';
    }
    return s ? `<w:rPr>${s}</w:rPr>` : '';
  }

  private runText(text: string): string {
    let s = '';
    for (const p of text.split(/(\n|\t)/)) {
      if (p === '\n') s += '<w:br/>';
      else if (p === '\t') s += '<w:tab/>';
      else if (p) s += `<w:t xml:space="preserve">${escapeXml(p)}</w:t>`;
    }
    return s;
  }

  private run(r: TextRun): string {
    if (r.field) return `<w:fldSimple w:instr=" ${r.field === 'page' ? 'PAGE' : 'NUMPAGES'} "><w:r>${this.rPr(r)}<w:t>1</w:t></w:r></w:fldSimple>`;
    if (r.note) {
      const id = this.notes.length + 1;
      this.notes.push('');
      const saved = this.part;
      this.part = this.notesPart;
      const body = this.noteBody(r.note);
      this.part = saved;
      this.notes[id - 1] = `<w:footnote w:id="${id}">${body}</w:footnote>`;
      return `<w:r>${this.rPr(r, 'FootnoteReference')}<w:footnoteReference w:id="${id}"/></w:r>`;
    }
    return r.text ? `<w:r>${this.rPr(r)}${this.runText(r.text)}</w:r>` : '';
  }

  readonly notesPart = new Part(1);

  private noteBody(blocks: Block[]): string {
    const mark = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>';
    const list = blocks.length ? blocks : [{ type: 'paragraph', runs: [] } as Block];
    let s = '';
    list.forEach((b, i) => {
      if (b.type === 'paragraph') s += `<w:p>${this.pPr({ style: 'FootnoteText' }, b)}${i === 0 ? mark : ''}${this.runs(b.runs.filter((r) => !r.note))}</w:p>`;
      else s += (i === 0 ? `<w:p>${this.pPr({ style: 'FootnoteText' })}${mark}</w:p>` : '') + this.blocks([b], 0, this.contentW);
    });
    return s;
  }

  private runs(runs: TextRun[]): string {
    let s = '';
    let i = 0;
    while (i < runs.length) {
      const link = runs[i].link;
      if (link && !runs[i].note && !runs[i].field) {
        let inner = '';
        while (i < runs.length && runs[i].link === link) inner += this.run(runs[i++]);
        s += `<w:hyperlink r:id="${this.part.link(link)}" w:history="1">${inner}</w:hyperlink>`;
      } else s += this.run(runs[i++]);
    }
    return s;
  }

  private pPr(opts: { style?: string; num?: { id: number; level: number }; indent?: number; border?: boolean }, f?: ParaFormat): string {
    let s = '';
    if (opts.style) s += `<w:pStyle w:val="${opts.style}"/>`;
    if (opts.num) s += `<w:numPr><w:ilvl w:val="${opts.num.level}"/><w:numId w:val="${opts.num.id}"/></w:numPr>`;
    if (opts.border) s += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A0A0A0"/></w:pBdr>';
    if (f && (f.spaceBefore !== undefined || f.spaceAfter !== undefined || f.lineHeight !== undefined)) {
      let sp = '';
      if (f.spaceBefore !== undefined) sp += ` w:before="${tw(f.spaceBefore)}"`;
      if (f.spaceAfter !== undefined) sp += ` w:after="${tw(f.spaceAfter)}"`;
      if (f.lineHeight !== undefined) sp += ` w:line="${Math.round(f.lineHeight * 240)}" w:lineRule="auto"`;
      s += `<w:spacing${sp}/>`;
    }
    const left = (opts.indent ?? 0) + (f?.indent ? tw(f.indent) : 0);
    if (!opts.num && (left || f?.firstLine)) {
      let ind = left ? ` w:left="${left}"` : '';
      if (f?.firstLine) ind += f.firstLine > 0 ? ` w:firstLine="${tw(f.firstLine)}"` : ` w:hanging="${tw(-f.firstLine)}"`;
      s += `<w:ind${ind}/>`;
    }
    if (f?.align && f.align !== 'left') s += `<w:jc w:val="${f.align === 'justify' ? 'both' : f.align}"/>`;
    return s ? `<w:pPr>${s}</w:pPr>` : '';
  }

  private image(b: Extract<Block, { type: 'image' }>, maxWidthTw: number): string {
    const ext = IMAGE_EXT[b.mime];
    if (!ext) {
      this.warnings.push('Некоторые изображения пропущены: формат не поддерживается Word.');
      return '';
    }
    const n = ++this.imageCount;
    const name = `image${n}.${ext}`;
    this.media[`word/media/${name}`] = b.data;
    const rid = this.part.add('image', `media/${name}`);
    const maxW = maxWidthTw * 635;
    const maxH = (this.pageH - this.margins.top - this.margins.bottom - 600) * 635;
    let cx = Math.max(1, b.width) * EMU_PER_PX;
    let cy = Math.max(1, b.height) * EMU_PER_PX;
    const k = Math.min(1, maxW / cx, maxH / cy);
    cx = Math.round(cx * k);
    cy = Math.round(cy * k);
    const descr = b.alt ? ` descr="${escapeXml(b.alt)}"` : '';
    return (
      `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${n}" name="Picture ${n}"${descr}/>` +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
      `<pic:nvPicPr><pic:cNvPr id="${n}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    );
  }

  private table(b: Extract<Block, { type: 'table' }>, widthTw: number): string {
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
    const widths = weights.map((w) => Math.floor((w / total) * widthTw));
    let s = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
    s += `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
    for (const r of b.rows) {
      const header = r.cells.length > 0 && r.cells.every((c) => c.header);
      s += '<w:tr>' + (header ? '<w:trPr><w:tblHeader/></w:trPr>' : '');
      let c = 0;
      const cells: TableCell[] = [...r.cells];
      let used = cells.reduce((a, x) => a + (x.colspan ?? 1), 0);
      while (used < cols) {
        cells.push({ blocks: [] });
        used++;
      }
      for (const cell of cells) {
        const span = Math.min(cell.colspan ?? 1, cols - c);
        if (span <= 0) break;
        const w = widths.slice(c, c + span).reduce((a, x) => a + x, 0);
        c += span;
        let tcPr = `<w:tcW w:w="${w}" w:type="dxa"/>`;
        if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
        if (cell.header) tcPr += '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>';
        let inner = this.blocks(cell.blocks, 0, w - 216);
        if (!inner.endsWith('</w:p>')) inner += '<w:p/>';
        s += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`;
      }
      s += '</w:tr>';
    }
    return s + '</w:tbl>';
  }

  blocks(blocks: Block[], indent: number, widthTw: number, list?: { id: number; level: number }): string {
    let s = '';
    for (const b of blocks) {
      switch (b.type) {
        case 'paragraph':
          s += `<w:p>${this.pPr({ num: list, indent: list ? undefined : indent || undefined }, b)}${this.runs(b.runs)}</w:p>`;
          list = undefined; // only the first paragraph of a list item carries the number
          break;
        case 'heading':
          s += `<w:p>${this.pPr({ style: `Heading${b.level}`, indent: indent || undefined }, b)}${this.runs(b.runs)}</w:p>`;
          break;
        case 'list': {
          const level = Math.min(8, Math.round(indent / 720));
          const id = this.newNum(b.ordered, level, b.start ?? 1);
          for (const it of b.items) {
            const itemBlocks = it.blocks.length ? it.blocks : [{ type: 'paragraph', runs: [] } as Block];
            const first = itemBlocks[0];
            if (first.type !== 'paragraph') s += `<w:p>${this.pPr({ num: { id, level } })}</w:p>`;
            s += this.blocks(itemBlocks, indent + 720, widthTw - 720, first.type === 'paragraph' ? { id, level } : undefined);
          }
          break;
        }
        case 'table':
          s += this.table(b, widthTw);
          break;
        case 'image':
          s += `<w:p>${this.pPr({ indent: indent || undefined }, { align: b.width * 15 < widthTw * 0.9 ? undefined : 'center' })}${this.image(b, widthTw)}</w:p>`;
          break;
        case 'code':
          s += `<w:p>${this.pPr({ style: 'Code', indent: indent || undefined })}<w:r>${this.runText(b.text)}</w:r></w:p>`;
          break;
        case 'quote':
          for (const inner of b.blocks) {
            if (inner.type === 'paragraph') s += `<w:p>${this.pPr({ style: 'Quote' }, inner)}${this.runs(inner.runs)}</w:p>`;
            else s += this.blocks([inner], indent + 567, widthTw - 567);
          }
          break;
        case 'hr':
          s += `<w:p>${this.pPr({ border: true })}</w:p>`;
          break;
        case 'pagebreak':
          s += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
          break;
      }
    }
    return s;
  }

  private headerFooter(kind: 'hdr' | 'ftr', blocks: Block[]): { xml: string; rels: string } {
    const saved = this.part;
    this.part = new Part(1);
    let body = this.blocks(blocks, 0, this.contentW);
    if (!body.endsWith('</w:p>')) body += '<w:p/>';
    const rels = this.part.xml();
    this.part = saved;
    return { xml: `${XML_HEAD}<w:${kind} ${NS_ALL}>${body}</w:${kind}>`, rels };
  }

  build(): Uint8Array {
    const doc = this.doc;
    const blocks = [...doc.blocks];
    while (blocks.length && blocks[blocks.length - 1].type === 'pagebreak') blocks.pop();
    let body = this.blocks(blocks, 0, this.contentW);
    if (!body || body.endsWith('</w:tbl>')) body += '<w:p/>';
    const files: Zippable = {};
    const overrides: string[] = [];
    const docRels = [
      `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`,
      `<Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/>`,
      `<Relationship Id="rId3" Type="${REL}/settings" Target="settings.xml"/>`,
    ];
    let refs = '';
    if (doc.header?.length) {
      const h = this.headerFooter('hdr', doc.header);
      files['word/header1.xml'] = utf8(h.xml);
      files['word/_rels/header1.xml.rels'] = utf8(h.rels);
      docRels.push(`<Relationship Id="rId4" Type="${REL}/header" Target="header1.xml"/>`);
      overrides.push('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>');
      refs += '<w:headerReference w:type="default" r:id="rId4"/>';
    }
    if (doc.footer?.length) {
      const f = this.headerFooter('ftr', doc.footer);
      files['word/footer1.xml'] = utf8(f.xml);
      files['word/_rels/footer1.xml.rels'] = utf8(f.rels);
      docRels.push(`<Relationship Id="rId5" Type="${REL}/footer" Target="footer1.xml"/>`);
      overrides.push('<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>');
      refs += '<w:footerReference w:type="default" r:id="rId5"/>';
    }
    if (this.notes.length) {
      files['word/footnotes.xml'] = utf8(
        `${XML_HEAD}<w:footnotes ${NS_ALL}>` +
          '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
          '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
          this.notes.join('') +
          '</w:footnotes>',
      );
      files['word/_rels/footnotes.xml.rels'] = utf8(this.notesPart.xml());
      docRels.push(`<Relationship Id="rId6" Type="${REL}/footnotes" Target="footnotes.xml"/>`);
      overrides.push('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>');
    }
    const m = this.margins;
    const document =
      `${XML_HEAD}<w:document ${NS_ALL}><w:body>${body}<w:sectPr>${refs}` +
      `<w:pgSz w:w="${this.pageW}" w:h="${this.pageH}"${this.pageW > this.pageH ? ' w:orient="landscape"' : ''}/>` +
      `<w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>`;
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    Object.assign(files, {
      '[Content_Types].xml': utf8(contentTypes(overrides)),
      '_rels/.rels': utf8(
        XML_HEAD +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>` +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
          `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` +
          '</Relationships>',
      ),
      'docProps/core.xml': utf8(
        XML_HEAD +
          '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
          (doc.title ? `<dc:title>${escapeXml(doc.title)}</dc:title>` : '') +
          `<dc:creator>Salatic</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
          '</cp:coreProperties>',
      ),
      'docProps/app.xml': utf8(XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Salatic</Application></Properties>'),
      'word/document.xml': utf8(document),
      'word/styles.xml': utf8(styles(doc)),
      'word/numbering.xml': utf8(numbering(this.nums)),
      'word/settings.xml': utf8(
        XML_HEAD +
          `<w:settings xmlns:w="${NS_W}"><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/>` +
          (this.notes.length ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : '') +
          '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
      ),
      'word/_rels/document.xml.rels': utf8(XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + docRels.join('') + this.part.rels.join('') + '</Relationships>'),
    });
    for (const [path, data] of Object.entries(this.media)) files[path] = [data, { level: 0 }];
    return zipSync(files, { level: 6 });
  }
}

function contentTypes(overrides: string[]): string {
  return (
    XML_HEAD +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
    '<Default Extension="gif" ContentType="image/gif"/>' +
    '<Default Extension="bmp" ContentType="image/bmp"/>' +
    '<Default Extension="emf" ContentType="image/x-emf"/>' +
    '<Default Extension="wmf" ContentType="image/x-wmf"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    overrides.join('') +
    '</Types>'
  );
}

function numbering(nums: string[]): string {
  const lvls = (ordered: boolean): string =>
    Array.from({ length: 9 }, (_, i) => {
      const fmt = ordered ? (i % 3 === 1 ? 'lowerLetter' : i % 3 === 2 ? 'lowerRoman' : 'decimal') : 'bullet';
      const text = ordered ? `%${i + 1}.` : BULLETS[i % 3];
      return (
        `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`
      );
    }).join('');
  return (
    XML_HEAD +
    `<w:numbering xmlns:w="${NS_W}">` +
    `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvls(false)}</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvls(true)}</w:abstractNum>` +
    nums.join('') +
    '</w:numbering>'
  );
}

const heading = (level: number, size: number, before: number): string =>
  `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
  `<w:rPr><w:b/><w:bCs/><w:color w:val="1F2937"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;

function styles(doc: Doc): string {
  const font = escapeXml(FONT_NAME[doc.font ?? 'sans']);
  const sz = Math.round((doc.size ?? 11) * 2);
  const d = doc.defaults ?? {};
  const after = d.spaceAfter !== undefined ? tw(d.spaceAfter) : 120;
  const before = d.spaceBefore !== undefined ? tw(d.spaceBefore) : 0;
  const line = d.lineHeight !== undefined ? Math.round(d.lineHeight * 240) : 276;
  const hs = (k: number): number => Math.round(sz * k);
  return (
    XML_HEAD +
    `<w:styles xmlns:w="${NS_W}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:eastAsia="${font}" w:hAnsi="${font}" w:cs="${font}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    heading(1, hs(2), 360) +
    heading(2, hs(1.64), 300) +
    heading(3, hs(1.36), 240) +
    heading(4, hs(1.18), 200) +
    heading(5, hs(1.09), 200) +
    heading(6, sz, 160) +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="C8C8CC"/></w:pBdr><w:ind w:left="567"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr></w:style>' +
    `<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="${Math.round(sz * 0.82)}"/><w:szCs w:val="${Math.round(sz * 0.82)}"/></w:rPr></w:style>` +
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
    '</w:tblBorders></w:tblPr></w:style>' +
    '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    '</w:styles>'
  );
}

export function docToDocx(doc: Doc, opts: DocxOptions = {}): { data: Uint8Array; warnings: string[] } {
  const w = new DocxWriter(doc, opts);
  const data = w.build();
  return { data, warnings: [...new Set(w.warnings)] };
}
