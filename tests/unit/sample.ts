import { encodePng } from '../../src/core/png';
import type { Doc } from '../../src/core/model';

/** Rich document used by several test files. */
const png = encodePng(
  new Uint8Array(120 * 80 * 3).map((_, i) => (i % 3 === 0 ? 220 : i % 3 === 1 ? 80 : (i >> 4) & 255)),
  120,
  80,
  3,
);

export const longPara = 'Съешь же ещё этих мягких французских булок, да выпей чаю. '.repeat(12).trim();

export const sample: Doc = {
  title: 'Тестовый документ',
  blocks: [
    { type: 'heading', level: 1, runs: [{ text: 'Отчёт о работе' }] },
    {
      type: 'paragraph',
      runs: [
        { text: 'Обычный текст, ' },
        { text: 'жирный', bold: true },
        { text: ', ' },
        { text: 'курсив', italic: true },
        { text: ', ' },
        { text: 'подчёркнутый', underline: true },
        { text: ' и ' },
        { text: 'ссылка', link: 'https://example.com/' },
        { text: '. E=mc' },
        { text: '2', sup: true },
      ],
    },
    { type: 'heading', level: 2, runs: [{ text: 'Список задач' }] },
    {
      type: 'list',
      ordered: false,
      items: [
        { blocks: [{ type: 'paragraph', runs: [{ text: 'Первый пункт' }] }] },
        {
          blocks: [
            { type: 'paragraph', runs: [{ text: 'Второй пункт' }] },
            { type: 'list', ordered: true, items: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Вложенный один' }] }] }, { blocks: [{ type: 'paragraph', runs: [{ text: 'Вложенный два' }] }] }] },
          ],
        },
      ],
    },
    { type: 'paragraph', runs: [{ text: longPara }], align: 'justify' },
    {
      type: 'table',
      rows: [
        { cells: [{ header: true, blocks: [{ type: 'paragraph', runs: [{ text: 'Город' }] }] }, { header: true, blocks: [{ type: 'paragraph', runs: [{ text: 'Население' }] }] }] },
        { cells: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Москва' }] }] }, { blocks: [{ type: 'paragraph', runs: [{ text: '13 100 000' }] }] }] },
        { cells: [{ blocks: [{ type: 'paragraph', runs: [{ text: 'Казань' }] }] }, { blocks: [{ type: 'paragraph', runs: [{ text: '1 300 000' }] }] }] },
      ],
    },
    { type: 'image', data: png, mime: 'image/png', width: 120, height: 80, alt: 'Диаграмма' },
    { type: 'code', text: 'function hello() {\n  return "Привет";\n}' },
    { type: 'quote', blocks: [{ type: 'paragraph', runs: [{ text: 'Цитата великого человека.' }] }] },
    { type: 'pagebreak' },
    { type: 'heading', level: 2, runs: [{ text: 'Вторая страница' }], align: 'center' },
    { type: 'paragraph', runs: [{ text: 'Конец документа.' }] },
  ],
};

