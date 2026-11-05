/**
 * Tolerant HTML tokenizer + tree builder that works without a DOM
 * (Web Workers, Cloudflare Workers, Node). Handles void elements, raw-text
 * elements, implicit closing of p/li/td/tr, and named/numeric entities.
 */

export interface HNode {
  tag: string; // lowercase; '#text' for text nodes
  attrs: Record<string, string>;
  children: HNode[];
  text?: string;
  parent?: HNode;
}

const VOID = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'source', 'track', 'wbr', 'param']);
const RAW = new Set(['script', 'style', 'textarea', 'title', 'noscript', 'template', 'xmp']);
const CLOSES_P = new Set([
  'address', 'article', 'aside', 'blockquote', 'details', 'div', 'dl', 'fieldset', 'figcaption', 'figure', 'footer', 'form',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hr', 'main', 'nav', 'ol', 'p', 'pre', 'section', 'table', 'ul',
]);

const ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', ensp: ' ', emsp: ' ', thinsp: ' ',
  mdash: '—', ndash: '–', laquo: '«', raquo: '»', ldquo: '“', rdquo: '”', bdquo: '„', lsquo: '‘', rsquo: '’', sbquo: '‚',
  hellip: '…', copy: '©', reg: '®', trade: '™', euro: '€', pound: '£', yen: '¥', cent: '¢', sect: '§', para: '¶',
  bull: '•', middot: '·', times: '×', divide: '÷', deg: '°', plusmn: '±', minus: '−', le: '≤', ge: '≥', ne: '≠',
  asymp: '≈', infin: '∞', larr: '←', rarr: '→', uarr: '↑', darr: '↓', harr: '↔', shy: '­', zwj: '‍',
  zwnj: '‌', frac12: '½', frac14: '¼', frac34: '¾', sup1: '¹', sup2: '²', sup3: '³', micro: 'µ', alpha: 'α',
  beta: 'β', gamma: 'γ', delta: 'δ', pi: 'π', sigma: 'σ', omega: 'ω', laquo2: '«', iexcl: '¡', iquest: '¿',
  acute: '´', uml: '¨', ordf: 'ª', ordm: 'º', not: '¬', macr: '¯', check: '✓', hearts: '♥', numero: '№', prime: '′',
  Prime: '″', dagger: '†', Dagger: '‡', permil: '‰', lsaquo: '‹', rsaquo: '›', oline: '‾', star: '☆',
};

export function decodeHtmlEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#[xX][0-9a-fA-F]+|#[0-9]+|[a-zA-Z][a-zA-Z0-9]*);?/g, (m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      if (!Number.isFinite(cp) || cp <= 0 || cp > 0x10ffff) return '�';
      // Windows-1252 remapping that browsers apply to C1 controls
      const C1: Record<number, number> = { 0x80: 0x20ac, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x96: 0x2013, 0x97: 0x2014, 0x85: 0x2026 };
      return String.fromCodePoint(C1[cp] ?? cp);
    }
    const v = ENTITIES[e] ?? ENTITIES[e.toLowerCase()];
    return v ?? m;
  });
}

const isWs = (c: number): boolean => c === 32 || c === 9 || c === 10 || c === 13 || c === 12;

function hasInScope(stack: HNode[], tag: string, boundaries: Set<string>): number {
  for (let k = stack.length - 1; k > 0; k--) {
    if (stack[k].tag === tag) return k;
    if (boundaries.has(stack[k].tag)) return -1;
  }
  return -1;
}

const SCOPE_BOUNDARY = new Set(['table', 'td', 'th', 'caption', 'html', 'template']);
const LIST_BOUNDARY = new Set(['ul', 'ol', 'table', 'td', 'th']);
const ROW_BOUNDARY = new Set(['table', 'thead', 'tbody', 'tfoot']);
const CELL_BOUNDARY = new Set(['tr', 'table']);

export function parseHtml(src: string): HNode {
  const root: HNode = { tag: '#document', attrs: {}, children: [] };
  const stack: HNode[] = [root];
  const n = src.length;
  let i = 0;

  const current = (): HNode => stack[stack.length - 1];
  const append = (node: HNode): void => {
    node.parent = current();
    current().children.push(node);
  };
  const addText = (raw: string): void => {
    if (!raw) return;
    const text = decodeHtmlEntities(raw);
    const cur = current();
    const last = cur.children[cur.children.length - 1];
    if (last && last.tag === '#text') last.text += text;
    else append({ tag: '#text', attrs: {}, children: [], text });
  };
  const closeTo = (index: number): void => {
    if (index > 0) stack.length = index;
  };
  const implicitClose = (tag: string): void => {
    if (CLOSES_P.has(tag)) {
      const k = hasInScope(stack, 'p', SCOPE_BOUNDARY);
      if (k > 0) closeTo(k);
    }
    if (tag === 'li') {
      const k = hasInScope(stack, 'li', LIST_BOUNDARY);
      if (k > 0) closeTo(k);
    } else if (tag === 'dt' || tag === 'dd') {
      const k = Math.max(hasInScope(stack, 'dt', LIST_BOUNDARY), hasInScope(stack, 'dd', LIST_BOUNDARY));
      if (k > 0) closeTo(k);
    } else if (tag === 'tr') {
      const k = hasInScope(stack, 'tr', ROW_BOUNDARY);
      if (k > 0) closeTo(k);
    } else if (tag === 'td' || tag === 'th') {
      const k = Math.max(hasInScope(stack, 'td', CELL_BOUNDARY), hasInScope(stack, 'th', CELL_BOUNDARY));
      if (k > 0) closeTo(k);
    } else if (tag === 'thead' || tag === 'tbody' || tag === 'tfoot') {
      for (const t of ['thead', 'tbody', 'tfoot']) {
        const k = hasInScope(stack, t, new Set(['table']));
        if (k > 0) closeTo(k);
      }
    } else if (tag === 'option') {
      const k = hasInScope(stack, 'option', new Set(['select']));
      if (k > 0) closeTo(k);
    }
  };

  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) {
      addText(src.slice(i));
      break;
    }
    const next = src.charCodeAt(lt + 1);
    const isTagStart = (next >= 65 && next <= 90) || (next >= 97 && next <= 122) || next === 47 || next === 33 || next === 63;
    if (!isTagStart) {
      addText(src.slice(i, lt + 1));
      i = lt + 1;
      continue;
    }
    if (lt > i) addText(src.slice(i, lt));
    if (next === 33 /* ! */ || next === 63 /* ? */) {
      if (src.startsWith('<!--', lt)) {
        const end = src.indexOf('-->', lt + 4);
        i = end < 0 ? n : end + 3;
      } else if (src.startsWith('<![CDATA[', lt)) {
        const end = src.indexOf(']]>', lt);
        addText(src.slice(lt + 9, end < 0 ? n : end));
        i = end < 0 ? n : end + 3;
      } else {
        const end = src.indexOf('>', lt);
        i = end < 0 ? n : end + 1;
      }
      continue;
    }
    if (next === 47 /* / */) {
      const end = src.indexOf('>', lt);
      const tag = src
        .slice(lt + 2, end < 0 ? n : end)
        .trim()
        .split(/\s/)[0]
        .toLowerCase();
      i = end < 0 ? n : end + 1;
      if (tag === 'p' && hasInScope(stack, 'p', SCOPE_BOUNDARY) < 0) {
        // </p> without an open <p> creates an empty paragraph per the HTML spec
        append({ tag: 'p', attrs: {}, children: [] });
        continue;
      }
      if (tag === 'br') {
        append({ tag: 'br', attrs: {}, children: [] });
        continue;
      }
      for (let k = stack.length - 1; k > 0; k--) {
        if (stack[k].tag === tag) {
          stack.length = k;
          break;
        }
        if (tag !== 'table' && (stack[k].tag === 'table' || stack[k].tag === 'td' || stack[k].tag === 'th') && !['td', 'th', 'tr', 'tbody', 'thead', 'tfoot'].includes(tag)) break;
      }
      continue;
    }
    // Start tag
    let j = lt + 1;
    while (j < n && !isWs(src.charCodeAt(j)) && src.charCodeAt(j) !== 62 && src.charCodeAt(j) !== 47) j++;
    const tag = src.slice(lt + 1, j).toLowerCase();
    const attrs: Record<string, string> = {};
    let selfClose = false;
    while (j < n) {
      while (j < n && isWs(src.charCodeAt(j))) j++;
      const ch = src.charCodeAt(j);
      if (ch === 62) {
        j++;
        break;
      }
      if (ch === 47) {
        selfClose = true;
        j++;
        continue;
      }
      const ns = j;
      while (j < n) {
        const c = src.charCodeAt(j);
        if (c === 61 || c === 62 || c === 47 || isWs(c)) break;
        j++;
      }
      const an = src.slice(ns, j).toLowerCase();
      while (j < n && isWs(src.charCodeAt(j))) j++;
      let value = '';
      if (src.charCodeAt(j) === 61) {
        j++;
        while (j < n && isWs(src.charCodeAt(j))) j++;
        const q = src.charCodeAt(j);
        if (q === 34 || q === 39) {
          const end = src.indexOf(String.fromCharCode(q), j + 1);
          value = src.slice(j + 1, end < 0 ? n : end);
          j = end < 0 ? n : end + 1;
        } else {
          const vs = j;
          while (j < n && !isWs(src.charCodeAt(j)) && src.charCodeAt(j) !== 62) j++;
          value = src.slice(vs, j);
        }
      }
      if (an) attrs[an] = decodeHtmlEntities(value);
      else j++;
    }
    i = j;
    implicitClose(tag);
    const node: HNode = { tag, attrs, children: [] };
    append(node);
    if (VOID.has(tag) || selfClose) continue;
    if (RAW.has(tag)) {
      const close = src.toLowerCase().indexOf(`</${tag}`, i);
      const end = close < 0 ? n : close;
      if (tag === 'textarea' || tag === 'title') node.children.push({ tag: '#text', attrs: {}, children: [], text: decodeHtmlEntities(src.slice(i, end)), parent: node });
      const gt = close < 0 ? n : src.indexOf('>', close);
      i = gt < 0 ? n : gt + 1;
      continue;
    }
    stack.push(node);
  }
  return root;
}

export function findTag(node: HNode, tag: string): HNode | undefined {
  for (const c of node.children) {
    if (c.tag === tag) return c;
    const f = findTag(c, tag);
    if (f) return f;
  }
  return undefined;
}

export function htmlText(node: HNode): string {
  if (node.tag === '#text') return node.text ?? '';
  if (node.tag === 'br') return '\n';
  let s = '';
  for (const c of node.children) s += htmlText(c);
  return s;
}

/** Parse an inline `style` attribute into a lowercase map. */
export function parseStyle(style: string | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (!style) return out;
  for (const decl of style.split(';')) {
    const k = decl.indexOf(':');
    if (k < 0) continue;
    out[decl.slice(0, k).trim().toLowerCase()] = decl.slice(k + 1).trim().toLowerCase();
  }
  return out;
}
