/**
 * A PowerPoint-structured test deck: theme, master with txStyles, layouts whose
 * placeholders carry the geometry (slides omit it), groups, cropping, a styled
 * table, notes, a hidden slide and a slide-number field.
 */
import { strToU8, zipSync } from 'fflate';

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (items: [string, string, string][]): string =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${target.startsWith('http') ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
const xfrm = (x: number, y: number, w: number, h: number): string => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
const ph = (type: string, idx?: number): string => `<p:nvPr><p:ph type="${type}"${idx !== undefined ? ` idx="${idx}"` : ''}/></p:nvPr>`;
const sp = (id: number, nvPr: string, spPr: string, body = ''): string => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/>${nvPr}</p:nvSpPr><p:spPr>${spPr}</p:spPr>${body}</p:sp>`;
const txBody = (paras: string, bodyPr = '<a:bodyPr/>'): string => `<p:txBody>${bodyPr}<a:lstStyle/>${paras}</p:txBody>`;
const para = (text: string, pPr = '', rPr = '<a:rPr lang="ru-RU"/>'): string => `<a:p>${pPr}<a:r>${rPr}<a:t>${text}</a:t></a:r></a:p>`;
const tree = (shapes: string): string => `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld>`;

/** 1 pt = 12700 EMU; the slide is 10 in × 7.5 in (720 × 540 pt). */
const E = (pt: number): number => Math.round(pt * 12700);

export function samplePng(w = 40, h = 20): Uint8Array {
  // Uncompressed (stored) deflate PNG, red left half, blue right half
  const raw: number[] = [];
  for (let y = 0; y < h; y++) {
    raw.push(0);
    for (let x = 0; x < w; x++) raw.push(...(x < w / 2 ? [220, 30, 30] : [30, 30, 220]));
  }
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (b: number[]): number => {
    let c = 0xffffffff;
    for (const v of b) c = crcTable[(c ^ v) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const u32 = (v: number): number[] => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
  const chunk = (type: string, data: number[]): number[] => {
    const t = [...type].map((c) => c.charCodeAt(0));
    return [...u32(data.length), ...t, ...data, ...u32(crc([...t, ...data]))];
  };
  // zlib stream with stored blocks
  const z: number[] = [0x78, 0x01];
  for (let i = 0; i < raw.length; i += 65535) {
    const part = raw.slice(i, i + 65535);
    const last = i + 65535 >= raw.length ? 1 : 0;
    z.push(last, part.length & 255, part.length >> 8, ~part.length & 255, (~part.length >> 8) & 255, ...part);
  }
  let a = 1;
  let b = 0;
  for (const v of raw) {
    a = (a + v) % 65521;
    b = (b + a) % 65521;
  }
  z.push(...u32((b << 16) | a));
  return Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, ...chunk('IHDR', [...u32(w), ...u32(h), 8, 2, 0, 0, 0]), ...chunk('IDAT', z), ...chunk('IEND', [])]);
}

export function makePptx(): Uint8Array {
  const theme =
    `<?xml version="1.0" encoding="UTF-8"?><a:theme ${NS} name="Office"><a:themeElements><a:clrScheme name="Office">` +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
    '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
    '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';
  const lvl = (n: number, attrs: string, inner: string): string => `<a:lvl${n}pPr ${attrs}>${inner}</a:lvl${n}pPr>`;
  const master =
    `<?xml version="1.0" encoding="UTF-8"?><p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    sp(2, ph('title'), xfrm(E(36), E(28.8), E(648), E(104.4)) + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', txBody('<a:p><a:endParaRPr lang="ru-RU"/></a:p>', '<a:bodyPr anchor="ctr"/>')) +
    sp(3, ph('body', 1), xfrm(E(36), E(144), E(648), E(342)) + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', txBody('<a:p><a:endParaRPr lang="ru-RU"/></a:p>')) +
    sp(4, ph('sldNum', 12), xfrm(E(516), E(500.5), E(168), E(28.75)) + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', txBody('<a:p><a:endParaRPr lang="ru-RU"/></a:p>', '<a:bodyPr anchor="ctr"/>')) +
    // decorative band (not a placeholder): drawn on every slide
    sp(5, '<p:nvPr/>', xfrm(0, E(530), E(720), E(10)) + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln>') +
    '</p:spTree></p:cSld>' +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>' +
    '<p:txStyles><p:titleStyle>' +
    lvl(1, 'algn="l"', '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr>') +
    '</p:titleStyle><p:bodyStyle>' +
    lvl(1, 'marL="228600" indent="-228600" algn="l"', '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr>') +
    lvl(2, 'marL="685800" indent="-228600" algn="l"', '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="–"/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr>') +
    '</p:bodyStyle><p:otherStyle>' +
    lvl(1, 'marL="0" algn="l"', '<a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr>') +
    '</p:otherStyle></p:txStyles></p:sldMaster>';
  const layout1 =
    `<?xml version="1.0" encoding="UTF-8"?><p:sldLayout ${NS} type="title" preserve="1">` +
    tree(
      sp(2, ph('ctrTitle'), xfrm(E(90), E(88.4), E(540), E(188)), txBody('<a:lstStyle/><a:p><a:endParaRPr/></a:p>'.replace('<a:lstStyle/>', ''), '<a:bodyPr anchor="b"/>').replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle>')) +
        sp(3, ph('subTitle', 1), xfrm(E(90), E(283.6), E(540), E(130.4)), txBody('<a:p><a:endParaRPr/></a:p>').replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>')),
    ) +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
  const layout2 = `<?xml version="1.0" encoding="UTF-8"?><p:sldLayout ${NS} type="obj" preserve="1">` + tree(sp(2, ph('title'), '') + sp(3, ph('body', 1), '')) + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
  const slide1 =
    `<?xml version="1.0" encoding="UTF-8"?><p:sld ${NS}>` +
    tree(sp(2, ph('ctrTitle'), '', txBody(para('Заголовок презентации'))) + sp(3, ph('subTitle', 1), '', txBody(para('Подзаголовок')))) +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const group =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    `<p:grpSpPr><a:xfrm><a:off x="${E(400)}" y="${E(300)}"/><a:ext cx="${E(200)}" cy="${E(100)}"/><a:chOff x="0" y="0"/><a:chExt cx="${E(400)}" cy="${E(200)}"/></a:xfrm></p:grpSpPr>` +
    sp(11, '<p:nvPr/>', xfrm(E(200), E(0), E(200), E(100)) + '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>', txBody(para('В группе', '<a:pPr algn="ctr"/>', '<a:rPr lang="ru-RU" sz="1400" b="1"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr>'), '<a:bodyPr anchor="ctr"/>')) +
    '</p:grpSp>';
  const pic =
    `<p:pic><p:nvPicPr><p:cNvPr id="20" name="pic" descr="Картинка"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/><a:srcRect r="50000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrm(E(400), E(150), E(100), E(100))}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const cell = (t: string, extra = ''): string => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ru-RU" sz="1400"/><a:t>${t}</a:t></a:r></a:p></a:txBody><a:tcPr${extra}/></a:tc>`;
  const table =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="30" name="t"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${E(36)}" y="${E(400)}"/><a:ext cx="${E(300)}" cy="${E(60)}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>' +
    `<a:tblGrid><a:gridCol w="${E(150)}"/><a:gridCol w="${E(150)}"/></a:tblGrid>` +
    `<a:tr h="${E(20)}">${cell('Город')}${cell('Продажи')}</a:tr><a:tr h="${E(20)}">${cell('Москва')}${cell('1 200')}</a:tr><a:tr h="${E(20)}">${cell('Широкая ячейка', ' gridSpan="2"').replace('<a:tc>', '<a:tc gridSpan="2">').replace(' gridSpan="2"/>', '/>')}<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>` +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  const line = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="40" name="l"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm(E(36), E(390), E(300), 0)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="${E(2)}"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`;
  const body =
    para('Первый пункт') +
    para('Вложенный пункт', '<a:pPr lvl="1"/>') +
    para('Нумерованный один', '<a:pPr><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>') +
    para('Нумерованный два', '<a:pPr><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>') +
    `<a:p><a:r><a:rPr lang="ru-RU"><a:hlinkClick r:id="rId4"/></a:rPr><a:t>Ссылка</a:t></a:r></a:p>`;
  const slide2 =
    `<?xml version="1.0" encoding="UTF-8"?><p:sld ${NS}>` +
    tree(
      sp(2, ph('title'), '', txBody(para('Список и объекты'))) +
        sp(3, ph('body', 1), xfrm(E(36), E(144), E(340), E(240)), txBody(body, '<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="10000"/></a:bodyPr>')) +
        group +
        pic +
        line +
        table,
    ) +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const slide3 = `<?xml version="1.0" encoding="UTF-8"?><p:sld ${NS} show="0">` + tree(sp(2, ph('title'), '', txBody(para('Скрытый слайд')))) + '</p:sld>';
  const slide4 =
    `<?xml version="1.0" encoding="UTF-8"?><p:sld ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    sp(2, ph('title'), '', txBody(para('Тёмный слайд', '', '<a:rPr lang="ru-RU"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr>'))) +
    sp(3, ph('sldNum', 12), '', txBody('<a:p><a:pPr algn="r"/><a:fld id="{1}" type="slidenum"><a:rPr lang="ru-RU"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>‹#›</a:t></a:fld></a:p>')) +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const notes =
    `<?xml version="1.0" encoding="UTF-8"?><p:notes ${NS}>` + tree(sp(2, ph('body', 1), '', txBody(para('Не забыть про сроки')))) + '</p:notes>';
  const pres =
    `<?xml version="1.0" encoding="UTF-8"?><p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/><p:sldId id="259" r:id="rId5"/></p:sldIdLst>' +
    `<p:sldSz cx="${E(720)}" cy="${E(540)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const ct =
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
    [1, 2].map((i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`).join('') +
    [1, 2, 3, 4].map((i) => `<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') +
    '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' +
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>';
  return zipSync({
    '[Content_Types].xml': strToU8(ct),
    '_rels/.rels': strToU8(rels([['rId1', 'officeDocument', 'ppt/presentation.xml']])),
    'ppt/presentation.xml': strToU8(pres),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels([
        ['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'],
        ['rId2', 'slide', 'slides/slide1.xml'],
        ['rId3', 'slide', 'slides/slide2.xml'],
        ['rId4', 'slide', 'slides/slide3.xml'],
        ['rId5', 'slide', 'slides/slide4.xml'],
        ['rId6', 'theme', 'theme/theme1.xml'],
      ]),
    ),
    'ppt/theme/theme1.xml': strToU8(theme),
    'ppt/slideMasters/slideMaster1.xml': strToU8(master),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': strToU8(
      rels([
        ['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'],
        ['rId2', 'slideLayout', '../slideLayouts/slideLayout2.xml'],
        ['rId3', 'theme', '../theme/theme1.xml'],
      ]),
    ),
    'ppt/slideLayouts/slideLayout1.xml': strToU8(layout1),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']])),
    'ppt/slideLayouts/slideLayout2.xml': strToU8(layout2),
    'ppt/slideLayouts/_rels/slideLayout2.xml.rels': strToU8(rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']])),
    'ppt/slides/slide1.xml': strToU8(slide1),
    'ppt/slides/_rels/slide1.xml.rels': strToU8(rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']])),
    'ppt/slides/slide2.xml': strToU8(slide2),
    'ppt/slides/_rels/slide2.xml.rels': strToU8(
      rels([
        ['rId1', 'slideLayout', '../slideLayouts/slideLayout2.xml'],
        ['rId2', 'notesSlide', '../notesSlides/notesSlide1.xml'],
        ['rId3', 'image', '../media/image1.png'],
        ['rId4', 'hyperlink', 'https://example.com/deck'],
      ]),
    ),
    'ppt/slides/slide3.xml': strToU8(slide3),
    'ppt/slides/_rels/slide3.xml.rels': strToU8(rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout2.xml']])),
    'ppt/slides/slide4.xml': strToU8(slide4),
    'ppt/slides/_rels/slide4.xml.rels': strToU8(rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout2.xml']])),
    'ppt/notesSlides/notesSlide1.xml': strToU8(notes),
    'ppt/media/image1.png': samplePng(),
    'docProps/core.xml': strToU8('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Тестовая презентация</dc:title></cp:coreProperties>'),
  });
}
