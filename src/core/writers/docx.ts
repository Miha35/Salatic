/** Doc → DOCX (Office Open XML). Element order follows the ECMA-376 schema so Word opens files without repair. */
import { zipSync, type Zippable } from 'fflate';
import type { Align, Block, Doc, TableCell, TextRun } from '../model';
import { escapeXml, utf8 } from '../util';

export interface DocxOptions {
  pageSize?: 'a4' | 'letter';
}

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const EMU_PER_PX = 9525;
const IMAGE_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' };

const BULLETS = ['•', '◦', '▪'];

class DocxWriter {
  private rels: string[] = [];
  private relId = 3; // rId1..3 reserved for styles, numbering, settings
  private media: Record<string, Uint8Array> = {};
  private imageCount = 0;
  private nums: string[] = [];
  private numId = 0;
  private links = new Map<string, string>();
  warnings: string[] = [];
  readonly pageW: number;
  readonly pageH: number;
  readonly margin = 1134; // 2 cm in twips

  constructor(opts: DocxOptions) {
    const letter = opts.pageSize === 'letter';
    this.pageW = letter ? 12240 : 11906;
    this.pageH = letter ? 15840 : 16838;
  }

  /** Usable width in twips. */
  get contentW(): number {
    return this.pageW - 2 * this.margin;
  }

  private nextRel(type: string, target: string, external = false): string {
    const id = `rId${++this.relId}`;
    this.rels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${escapeXml(target)}"${external ? ' TargetMode="External"' : ''}/>`);
    return id;
  }

  private linkRel(url: string): string {
    let id = this.links.get(url);
    if (!id) {
      id = this.nextRel('hyperlink', url, true);
      this.links.set(url, id);
    }
    return id;
  }

  private newNum(ordered: boolean, level: number, start = 1): number {
    const id = ++this.numId;
    const override = ordered ? `<w:lvlOverride w:ilvl="${level}"><w:startOverride w:val="${start}"/></w:lvlOverride>` : '';
    this.nums.push(`<w:num w:numId="${id}"><w:abstractNumId w:val="${ordered ? 1 : 0}"/>${override}</w:num>`);
    return id;
  }

  private rPr(r: TextRun): string {
    let s = '';
    if (r.link) s += '<w:rStyle w:val="Hyperlink"/>';
    if (r.code) s += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>';
    if (r.bold) s += '<w:b/><w:bCs/>';
    if (r.italic) s += '<w:i/><w:iCs/>';
    if (r.strike) s += '<w:strike/>';
    if (r.color) s += `<w:color w:val="${r.color}"/>`;
    if (r.underline) s += '<w:u w:val="single"/>';
    if (r.sup) s += '<w:vertAlign w:val="superscript"/>';
    else if (r.sub) s += '<w:vertAlign w:val="subscript"/>';
    return s ? `<w:rPr>${s}</w:rPr>` : '';
  }

  private runText(text: string): string {
    let s = '';
    const parts = text.split(/(\n|\t)/);
    for (const p of parts) {
      if (p === '\n') s += '<w:br/>';
      else if (p === '\t') s += '<w:tab/>';
      else if (p) s += `<w:t xml:space="preserve">${escapeXml(p)}</w:t>`;
    }
    return s;
  }

  private runs(runs: TextRun[]): string {
    let s = '';
    let i = 0;
    while (i < runs.length) {
      const link = runs[i].link;
      if (link) {
        let inner = '';
        while (i < runs.length && runs[i].link === link) {
          inner += `<w:r>${this.rPr(runs[i])}${this.runText(runs[i].text)}</w:r>`;
          i++;
        }
        s += `<w:hyperlink r:id="${this.linkRel(link)}" w:history="1">${inner}</w:hyperlink>`;
      } else {
        if (runs[i].text) s += `<w:r>${this.rPr(runs[i])}${this.runText(runs[i].text)}</w:r>`;
        i++;
      }
    }
    return s;
  }

  private pPr(opts: { style?: string; num?: { id: number; level: number }; indent?: number; align?: Align; border?: boolean }): string {
    let s = '';
    if (opts.style) s += `<w:pStyle w:val="${opts.style}"/>`;
    if (opts.num) s += `<w:numPr><w:ilvl w:val="${opts.num.level}"/><w:numId w:val="${opts.num.id}"/></w:numPr>`;
    if (opts.border) s += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A0A0A0"/></w:pBdr>';
    if (opts.indent) s += `<w:ind w:left="${opts.indent}"/>`;
    if (opts.align && opts.align !== 'left') s += `<w:jc w:val="${opts.align === 'justify' ? 'both' : opts.align}"/>`;
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
    const rid = this.nextRel('image', `media/${name}`);
    // Fit into the available width/height (twips → EMU: ×635)
    const maxW = maxWidthTw * 635;
    const maxH = (this.pageH - 2 * this.margin - 600) * 635;
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
    const cols = Math.max(1, ...b.rows.map((r) => r.cells.reduce((s, c) => s + (c.colspan ?? 1), 0)));
    // Column widths proportional to content length (bounded), summing to the available width
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
          s += `<w:p>${this.pPr({ num: list, indent: list ? undefined : indent || undefined, align: b.align })}${this.runs(b.runs)}</w:p>`;
          list = undefined; // only the first paragraph of a list item carries the number
          break;
        case 'heading':
          s += `<w:p>${this.pPr({ style: `Heading${b.level}`, align: b.align, indent: indent || undefined })}${this.runs(b.runs)}</w:p>`;
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
          s += `<w:p>${this.pPr({ indent: indent || undefined, align: b.width * 15 < widthTw * 0.9 ? undefined : 'center' })}${this.image(b, widthTw)}</w:p>`;
          break;
        case 'code':
          s += `<w:p>${this.pPr({ style: 'Code', indent: indent || undefined })}<w:r>${this.runText(b.text)}</w:r></w:p>`;
          break;
        case 'quote':
          for (const inner of b.blocks) {
            if (inner.type === 'paragraph') s += `<w:p>${this.pPr({ style: 'Quote', align: inner.align })}${this.runs(inner.runs)}</w:p>`;
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

  build(doc: Doc): Uint8Array {
    let body = this.blocks(doc.blocks, 0, this.contentW);
    if (!body || body.endsWith('</w:tbl>')) body += '<w:p/>';
    const document =
      XML_HEAD +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<w:body>${body}<w:sectPr><w:pgSz w:w="${this.pageW}" w:h="${this.pageH}"/><w:pgMar w:top="${this.margin}" w:right="${this.margin}" w:bottom="${this.margin}" w:left="${this.margin}" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>`;
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const files: Zippable = {
      '[Content_Types].xml': utf8(contentTypes()),
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
      'docProps/app.xml': utf8(
        XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Salatic</Application></Properties>',
      ),
      'word/document.xml': utf8(document),
      'word/styles.xml': utf8(STYLES),
      'word/numbering.xml': utf8(numbering(this.nums)),
      'word/settings.xml': utf8(
        XML_HEAD +
          `<w:settings xmlns:w="${NS_W}"><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/>` +
          '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
      ),
      'word/_rels/document.xml.rels': utf8(
        XML_HEAD +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>` +
          `<Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/>` +
          `<Relationship Id="rId3" Type="${REL}/settings" Target="settings.xml"/>` +
          this.rels.join('') +
          '</Relationships>',
      ),
    };
    for (const [path, data] of Object.entries(this.media)) files[path] = [data, { level: 0 }];
    return zipSync(files, { level: 6 });
  }
}

function contentTypes(): string {
  return (
    XML_HEAD +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
    '<Default Extension="gif" ContentType="image/gif"/>' +
    '<Default Extension="bmp" ContentType="image/bmp"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
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

const STYLES =
  XML_HEAD +
  `<w:styles xmlns:w="${NS_W}">` +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  heading(1, 36, 360) +
  heading(2, 30, 300) +
  heading(3, 26, 240) +
  heading(4, 24, 200) +
  heading(5, 22, 200) +
  heading(6, 22, 160) +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="C8C8CC"/></w:pBdr><w:ind w:left="567"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
  '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
  '</w:tblBorders></w:tblPr></w:style>' +
  '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
  '</w:styles>';

export function docToDocx(doc: Doc, opts: DocxOptions = {}): { data: Uint8Array; warnings: string[] } {
  const w = new DocxWriter(opts);
  const data = w.build(doc);
  return { data, warnings: [...new Set(w.warnings)] };
}
