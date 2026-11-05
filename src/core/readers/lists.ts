/** Builds nested list blocks from a flat sequence of numbered paragraphs (Word-style ilvl). */
import type { Block, ListItem } from '../model';

type ListBlock = Extract<Block, { type: 'list' }>;

export class ListNester {
  private stack: { list: ListBlock; key: string }[] = [];
  constructor(private out: Block[]) {}

  /** A paragraph that is not part of a list closes all open lists. */
  block(...blocks: Block[]): void {
    this.stack = [];
    this.out.push(...blocks);
  }

  item(blocks: Block[], info: { key: string; level: number; ordered: boolean; start?: number }): void {
    const level = Math.max(0, Math.min(info.level, 8));
    const stack = this.stack;
    if (stack.length && level === 0 && stack[0].key !== info.key) stack.length = 0;
    while (stack.length > level + 1) stack.pop();
    if (stack.length === level + 1 && stack[level].list.ordered !== info.ordered) stack.pop();
    while (stack.length < level + 1) {
      const fresh: ListBlock = { type: 'list', ordered: info.ordered, items: [] };
      if (info.ordered && info.start && info.start !== 1 && stack.length === level) fresh.start = info.start;
      if (!stack.length) this.out.push(fresh);
      else {
        const parent = stack[stack.length - 1].list;
        if (!parent.items.length) parent.items.push({ blocks: [] });
        parent.items[parent.items.length - 1].blocks.push(fresh);
      }
      stack.push({ list: fresh, key: info.key });
    }
    const item: ListItem = { blocks: blocks.length ? blocks : [{ type: 'paragraph', runs: [] }] };
    stack[level].list.items.push(item);
  }

  /** Content that continues the current list item (e.g. an image inside it). */
  append(blocks: Block[]): boolean {
    const top = this.stack[this.stack.length - 1];
    const it = top?.list.items[top.list.items.length - 1];
    if (!it) return false;
    it.blocks.push(...blocks);
    return true;
  }

  close(): void {
    this.stack = [];
  }
}
