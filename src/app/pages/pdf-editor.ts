import { h } from '../ui';

export function mount(root: HTMLElement): void {
  root.append(h('p', null, 'Скоро'));
}
