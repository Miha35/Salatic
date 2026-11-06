/**
 * Minimal, fast, non-validating XML parser (sufficient for OOXML parts).
 * Element names keep their namespace prefix (e.g. "w:p").
 */

export interface XNode {
  name: string;
  attrs: Record<string, string>;
  children: (XNode | string)[];
}

const NAMED: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeXmlEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '';
    }
    return NAMED[e] ?? m;
  });
}

const isNameEnd = (c: number): boolean => c === 32 || c === 9 || c === 10 || c === 13 || c === 47 || c === 62;
const isWs = (c: number): boolean => c === 32 || c === 9 || c === 10 || c === 13;

/** Deeper nesting is flattened (no real OOXML part comes close) to protect recursive helpers. */
const MAX_DEPTH = 1000;

export function parseXml(src: string): XNode {
  const root: XNode = { name: '#root', attrs: {}, children: [] };
  const stack: XNode[] = [root];
  const n = src.length;
  let i = 0;
  const pushText = (t: string): void => {
    if (t) stack[stack.length - 1].children.push(decodeXmlEntities(t));
  };
  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) {
      pushText(src.slice(i));
      break;
    }
    if (lt > i) pushText(src.slice(i, lt));
    const c = src.charCodeAt(lt + 1);
    if (c === 33 /* ! */) {
      if (src.startsWith('<!--', lt)) {
        const end = src.indexOf('-->', lt + 4);
        i = end < 0 ? n : end + 3;
      } else if (src.startsWith('<![CDATA[', lt)) {
        const end = src.indexOf(']]>', lt + 9);
        stack[stack.length - 1].children.push(src.slice(lt + 9, end < 0 ? n : end));
        i = end < 0 ? n : end + 3;
      } else {
        // DOCTYPE (possibly with an internal subset)
        const bracket = src.indexOf('[', lt);
        const gt = src.indexOf('>', lt);
        if (bracket >= 0 && bracket < gt) {
          const end = src.indexOf(']>', bracket);
          i = end < 0 ? n : end + 2;
        } else i = gt < 0 ? n : gt + 1;
      }
      continue;
    }
    if (c === 63 /* ? */) {
      const end = src.indexOf('?>', lt);
      i = end < 0 ? n : end + 2;
      continue;
    }
    if (c === 47 /* / */) {
      const end = src.indexOf('>', lt);
      const name = src.slice(lt + 2, end < 0 ? n : end).trim();
      for (let k = stack.length - 1; k > 0; k--) {
        if (stack[k].name === name) {
          stack.length = k;
          break;
        }
      }
      i = end < 0 ? n : end + 1;
      continue;
    }
    // Start tag
    let j = lt + 1;
    while (j < n && !isNameEnd(src.charCodeAt(j))) j++;
    const node: XNode = { name: src.slice(lt + 1, j), attrs: {}, children: [] };
    let selfClosing = false;
    while (j < n) {
      while (j < n && isWs(src.charCodeAt(j))) j++;
      const ch = src.charCodeAt(j);
      if (ch === 62 /* > */) {
        j++;
        break;
      }
      if (ch === 47 /* / */) {
        selfClosing = true;
        j++;
        continue;
      }
      const ns = j;
      while (j < n) {
        const cc = src.charCodeAt(j);
        if (cc === 61 || isWs(cc) || cc === 62 || cc === 47) break;
        j++;
      }
      const an = src.slice(ns, j);
      while (j < n && isWs(src.charCodeAt(j))) j++;
      let value = '';
      if (src.charCodeAt(j) === 61 /* = */) {
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
      if (an) node.attrs[an] = decodeXmlEntities(value);
      else j++; // defensive: skip a stray character
    }
    stack[stack.length - 1].children.push(node);
    if (!selfClosing && stack.length < MAX_DEPTH) stack.push(node);
    i = j;
  }
  return root;
}

// ---------------------------------------------------------------------------
// Query helpers

export function child(node: XNode | undefined, name: string): XNode | undefined {
  if (!node) return undefined;
  for (const c of node.children) if (typeof c !== 'string' && c.name === name) return c;
  return undefined;
}

export function childrenNamed(node: XNode | undefined, name: string): XNode[] {
  const out: XNode[] = [];
  if (!node) return out;
  for (const c of node.children) if (typeof c !== 'string' && c.name === name) out.push(c);
  return out;
}

export function elements(node: XNode | undefined): XNode[] {
  const out: XNode[] = [];
  if (!node) return out;
  for (const c of node.children) if (typeof c !== 'string') out.push(c);
  return out;
}

/** Depth-first search for the first descendant with the given name. */
export function find(node: XNode | undefined, name: string): XNode | undefined {
  if (!node) return undefined;
  for (const c of node.children) {
    if (typeof c === 'string') continue;
    if (c.name === name) return c;
    const f = find(c, name);
    if (f) return f;
  }
  return undefined;
}

/** First descendant element matching a predicate (depth-first). */
export function findBy(node: XNode | undefined, pred: (n: XNode) => boolean): XNode | undefined {
  if (!node) return undefined;
  for (const c of node.children) {
    if (typeof c === 'string') continue;
    if (pred(c)) return c;
    const f = findBy(c, pred);
    if (f) return f;
  }
  return undefined;
}

export function findAll(node: XNode | undefined, name: string, out: XNode[] = []): XNode[] {
  if (!node) return out;
  for (const c of node.children) {
    if (typeof c === 'string') continue;
    if (c.name === name) out.push(c);
    else findAll(c, name, out);
  }
  return out;
}

export function textContent(node: XNode | undefined): string {
  if (!node) return '';
  let s = '';
  for (const c of node.children) s += typeof c === 'string' ? c : textContent(c);
  return s;
}

/** Root element of a parsed document (skips the synthetic #root). */
export function rootElement(doc: XNode): XNode | undefined {
  return elements(doc)[0];
}

/** Local name without namespace prefix. */
export const localName = (name: string): string => {
  const i = name.indexOf(':');
  return i < 0 ? name : name.slice(i + 1);
};
