import type { Block, Doc } from '../model';
import { cleanText } from '../util';

/** Plain text → Doc. Blank lines separate paragraphs; single newlines are kept as line breaks. */
export function textToDoc(text: string, title?: string): Doc {
  const normalized = cleanText(text.replace(/\r\n?/g, '\n'));
  const blocks: Block[] = [];
  for (const chunk of normalized.split(/\n[ \t]*\n+/)) {
    const t = chunk.replace(/^\n+|\s+$/g, '');
    if (!t.trim()) continue;
    if (/^\f/.test(t)) blocks.push({ type: 'pagebreak' });
    blocks.push({ type: 'paragraph', runs: [{ text: t.replace(/\f/g, '') }] });
  }
  return { title, blocks };
}
