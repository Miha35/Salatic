/** Office-style chart parts (c:chartSpace with cached values) for tests. */
const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const strCache = (vals: string[]): string => `<c:strRef><c:f>Sheet1!$A$2:$A$${vals.length + 1}</c:f><c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`;
const numCache = (vals: number[], fmt = 'General'): string =>
  `<c:numRef><c:f>Sheet1!$B$2:$B$${vals.length + 1}</c:f><c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;

export interface SeriesSpec {
  name: string;
  values: number[];
  color?: string;
  xs?: number[];
}

const ser = (s: SeriesSpec, i: number, cats: string[], extra = ''): string =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strCache([s.name])}</c:tx>` +
  (s.color ? `<c:spPr><a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill></a:ln></c:spPr>` : '') +
  extra +
  (s.xs ? `<c:xVal>${numCache(s.xs)}</c:xVal><c:yVal>${numCache(s.values)}</c:yVal>` : `<c:cat>${strCache(cats)}</c:cat><c:val>${numCache(s.values)}</c:val>`) +
  '</c:ser>';

const axes = (catId = 1, valId = 2, opts: { pos?: string; grid?: boolean; fmt?: string; xVal?: boolean } = {}): string =>
  (opts.xVal
    ? `<c:valAx><c:axId val="${catId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:crossAx val="${valId}"/></c:valAx>`
    : `<c:catAx><c:axId val="${catId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${valId}"/></c:catAx>`) +
  `<c:valAx><c:axId val="${valId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${opts.pos ?? 'l'}"/>${opts.grid === false ? '' : '<c:majorGridlines/>'}<c:numFmt formatCode="${opts.fmt ?? 'General'}" sourceLinked="0"/><c:crossAx val="${catId}"/></c:valAx>`;

const wrap = (title: string | null, plot: string, legend = 'b', view3D = ''): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace ${NS}><c:chart>` +
  (title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>') +
  view3D +
  `<c:plotArea><c:layout/>${plot}</c:plotArea>` +
  (legend ? `<c:legend><c:legendPos val="${legend}"/><c:overlay val="0"/></c:legend>` : '') +
  '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

export function barChart(cats: string[], series: SeriesSpec[], opts: { title?: string; horizontal?: boolean; grouping?: 'clustered' | 'stacked' | 'percentStacked'; labels?: boolean; fmt?: string } = {}): string {
  return wrap(
    opts.title ?? null,
    `<c:barChart><c:barDir val="${opts.horizontal ? 'bar' : 'col'}"/><c:grouping val="${opts.grouping ?? 'clustered'}"/><c:varyColors val="0"/>` +
      series.map((s, i) => ser(s, i, cats)).join('') +
      (opts.labels ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>' : '') +
      `<c:gapWidth val="150"/>${opts.grouping && opts.grouping !== 'clustered' ? '<c:overlap val="100"/>' : ''}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      axes(1, 2, { fmt: opts.fmt }),
  );
}

export function comboChart(cats: string[], bars: SeriesSpec[], lines: SeriesSpec[], title: string): string {
  return wrap(
    title,
    `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${bars.map((s, i) => ser(s, i, cats)).join('')}<c:gapWidth val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${lines.map((s, i) => ser(s, bars.length + i, cats, '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>')).join('')}<c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>` +
      axes(1, 2) +
      '<c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="4"/></c:catAx><c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="0%" sourceLinked="0"/><c:crossAx val="3"/><c:crosses val="max"/></c:valAx>',
  );
}

export function pieChart(cats: string[], values: number[], opts: { title?: string; doughnut?: boolean } = {}): string {
  const kind = opts.doughnut ? 'doughnutChart' : 'pieChart';
  return wrap(
    opts.title ?? null,
    `<c:${kind}><c:varyColors val="1"/>${ser({ name: opts.title ?? 'Доли', values }, 0, cats)}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls><c:firstSliceAng val="0"/>${opts.doughnut ? '<c:holeSize val="55"/>' : ''}</c:${kind}>`,
    'r',
  );
}

export function scatterChart(series: SeriesSpec[], title: string): string {
  return wrap(title, `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series.map((s, i) => ser(s, i, [])).join('')}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>` + axes(1, 2, { xVal: true }));
}

export function areaChart(cats: string[], series: SeriesSpec[], title: string): string {
  return wrap(title, `<c:areaChart><c:grouping val="stacked"/><c:varyColors val="0"/>${series.map((s, i) => ser(s, i, cats)).join('')}<c:axId val="1"/><c:axId val="2"/></c:areaChart>` + axes(1, 2));
}

export const SAMPLE_CATS = ['Янв', 'Фев', 'Мар', 'Апр'];

export function sampleCharts(): Record<string, string> {
  const cats = SAMPLE_CATS;
  return {
    bar: barChart(cats, [{ name: 'План', values: [120, 135, 150, 160] }, { name: 'Факт', values: [110, 140, 155, 170] }, { name: 'Прогноз', values: [100, 120, 160, 190] }], { title: 'Продажи, тыс. ₽', labels: true }),
    hbar: barChart(['Москва', 'Санкт-Петербург', 'Казань'], [{ name: 'Q1', values: [30, 20, 10] }, { name: 'Q2', values: [35, 25, 12] }], { title: 'По городам', horizontal: true, grouping: 'stacked' }),
    combo: comboChart(cats, [{ name: 'Выручка', values: [1200000, 1350000, 1500000, 1600000] }], [{ name: 'Маржа', values: [0.21, 0.24, 0.23, 0.27] }], 'Выручка и маржа'),
    pie: pieChart(['Москва', 'СПб', 'Казань', 'Прочие'], [45, 25, 10, 20], { title: 'Доля рынка' }),
    donut: pieChart(['А', 'Б', 'В'], [50, 30, 20], { title: 'Кольцо', doughnut: true }),
    scatter: scatterChart([{ name: 'Замеры', values: [2, 4, 3, 7, 6], xs: [1, 2, 3, 4, 5] }], 'Точки'),
    area: areaChart(cats, [{ name: 'A', values: [3, 4, 5, 4] }, { name: 'B', values: [2, 2, 3, 5] }], 'Области'),
  };
}

// ---------------------------------------------------------------- more chart kinds

const v3 = (rotX: number, rotY: number, rAng = true): string => `<c:view3D><c:rotX val="${rotX}"/><c:rotY val="${rotY}"/><c:rAngAx val="${rAng ? 1 : 0}"/></c:view3D>`;
const serAx = '<c:serAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:serAx>';
const noLine = '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
const markerOf = (sym: string): string => `<c:marker><c:symbol val="${sym}"/><c:size val="5"/></c:marker>`;
const serRaw = (s: SeriesSpec, i: number, cats: string[], extra = ''): string =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strCache([s.name])}</c:tx>${extra}<c:cat>${strCache(cats)}</c:cat><c:val>${numCache(s.values)}</c:val></c:ser>`;

export function radarChart(cats: string[], series: SeriesSpec[], style: 'standard' | 'marker' | 'filled', title: string): string {
  return wrap(title, `<c:radarChart><c:radarStyle val="${style}"/><c:varyColors val="0"/>${series.map((s, i) => ser(s, i, cats, style === 'marker' ? markerOf('circle') : '')).join('')}<c:axId val="1"/><c:axId val="2"/></c:radarChart>` + axes(1, 2));
}

/** open, high, low, close */
export function stockChart(cats: string[], ohlc: [number[], number[], number[], number[]], title: string): string {
  const names = ['Открытие', 'Максимум', 'Минимум', 'Закрытие'];
  const sers = ohlc.map((values, i) => serRaw({ name: names[i], values }, i, cats, noLine + markerOf(i === 3 ? 'dash' : 'none'))).join('');
  return wrap(title, `<c:stockChart>${sers}<c:hiLowLines/><c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars><c:axId val="1"/><c:axId val="2"/></c:stockChart>` + axes(1, 2), 'b');
}

export function bubbleChart(series: { name: string; xs: number[]; ys: number[]; sizes: number[] }[], title: string): string {
  const sers = series
    .map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strCache([s.name])}</c:tx><c:invertIfNegative val="0"/><c:xVal>${numCache(s.xs)}</c:xVal><c:yVal>${numCache(s.ys)}</c:yVal><c:bubbleSize>${numCache(s.sizes)}</c:bubbleSize><c:bubble3D val="0"/></c:ser>`)
    .join('');
  return wrap(title, `<c:bubbleChart><c:varyColors val="0"/>${sers}<c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>` + axes(1, 2, { xVal: true }));
}

export function bar3DChart(cats: string[], series: SeriesSpec[], opts: { title: string; shape?: string; horizontal?: boolean; grouping?: string }): string {
  return wrap(
    opts.title,
    `<c:bar3DChart><c:barDir val="${opts.horizontal ? 'bar' : 'col'}"/><c:grouping val="${opts.grouping ?? 'clustered'}"/><c:varyColors val="0"/>${series.map((s, i) => ser(s, i, cats)).join('')}<c:gapWidth val="150"/><c:shape val="${opts.shape ?? 'box'}"/><c:axId val="1"/><c:axId val="2"/><c:axId val="0"/></c:bar3DChart>` + axes(1, 2),
    'b',
    v3(15, 20),
  );
}

export function pie3DChart(cats: string[], values: number[], title: string): string {
  return wrap(title, `<c:pie3DChart><c:varyColors val="1"/>${ser({ name: title, values }, 0, cats)}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls></c:pie3DChart>`, 'r', v3(30, 0, false));
}

export function line3DChart(cats: string[], series: SeriesSpec[], title: string): string {
  return wrap(title, `<c:line3DChart><c:grouping val="standard"/><c:varyColors val="0"/>${series.map((s, i) => ser(s, i, cats)).join('')}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:line3DChart>` + axes(1, 2) + serAx, 'b', v3(15, 20, false));
}

export function area3DChart(cats: string[], series: SeriesSpec[], title: string): string {
  return wrap(title, `<c:area3DChart><c:grouping val="standard"/><c:varyColors val="0"/>${series.map((s, i) => ser(s, i, cats)).join('')}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:area3DChart>` + axes(1, 2) + serAx, 'b', v3(15, 20, false));
}

export function surfaceChart(cats: string[], rows: SeriesSpec[], title: string, threeD: boolean): string {
  const kind = threeD ? 'surface3DChart' : 'surfaceChart';
  return wrap(title, `<c:${kind}><c:wireframe val="0"/>${rows.map((s, i) => ser(s, i, cats)).join('')}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:${kind}>` + axes(1, 2) + serAx, 'r', threeD ? v3(15, 20, false) : v3(90, 0, false));
}

export function ofPieChart(cats: string[], values: number[], title: string, bar: boolean): string {
  return wrap(
    title,
    `<c:ofPieChart><c:ofPieType val="${bar ? 'bar' : 'pie'}"/><c:varyColors val="1"/>${ser({ name: title, values }, 0, cats)}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls><c:gapWidth val="100"/><c:splitType val="pos"/><c:splitPos val="3"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart>`,
    'b',
  );
}

export function moreCharts(): Record<string, string> {
  const cats = SAMPLE_CATS;
  const six = ['Скорость', 'Цена', 'Надёжность', 'Дизайн', 'Сервис', 'Гарантия'];
  return {
    radar: radarChart(six, [{ name: 'Модель А', values: [8, 6, 7, 9, 5, 7] }, { name: 'Модель Б', values: [5, 9, 6, 6, 8, 4] }], 'marker', 'Сравнение моделей'),
    radarFilled: radarChart(six, [{ name: 'Модель А', values: [8, 6, 7, 9, 5, 7] }, { name: 'Модель Б', values: [5, 9, 6, 6, 8, 4] }], 'filled', 'Заполненная'),
    stock: stockChart(['01.03', '02.03', '03.03', '04.03', '05.03'], [[44, 25, 38, 50, 34], [55, 57, 57, 58, 58], [11, 12, 13, 11, 25], [25, 38, 50, 35, 43]], 'Котировки'),
    bubble: bubbleChart([{ name: 'Регионы', xs: [1, 2.5, 4, 5.5], ys: [12, 30, 18, 25], sizes: [5, 20, 10, 40] }, { name: 'Партнёры', xs: [2, 3.5, 5], ys: [22, 8, 35], sizes: [15, 30, 8] }], 'Пузырьки'),
    column3d: bar3DChart(cats, [{ name: 'План', values: [120, 135, 150, 160] }, { name: 'Факт', values: [110, 140, 155, 170] }], { title: 'Объёмная гистограмма' }),
    cylinder3d: bar3DChart(cats, [{ name: 'План', values: [120, 135, 150, 160] }], { title: 'Цилиндры', shape: 'cylinder' }),
    bar3d: bar3DChart(['Москва', 'Казань', 'Сочи'], [{ name: 'Q1', values: [30, 20, 10] }, { name: 'Q2', values: [35, 25, 12] }], { title: 'Объёмные полосы', horizontal: true, grouping: 'stacked' }),
    pie3d: pie3DChart(['Москва', 'СПб', 'Казань', 'Прочие'], [45, 25, 10, 20], 'Объёмная круговая'),
    line3d: line3DChart(cats, [{ name: 'A', values: [3, 5, 4, 6] }, { name: 'B', values: [2, 3, 5, 4] }], 'Объёмный график'),
    area3d: area3DChart(cats, [{ name: 'A', values: [3, 5, 4, 6] }, { name: 'B', values: [2, 3, 5, 4] }], 'Объёмные области'),
    contour: surfaceChart(['0', '1', '2', '3', '4'], [{ name: 'y0', values: [1, 2, 3, 2, 1] }, { name: 'y1', values: [2, 4, 6, 4, 2] }, { name: 'y2', values: [3, 6, 9, 6, 3] }, { name: 'y3', values: [2, 4, 6, 4, 2] }], 'Контурная', false),
    surface3d: surfaceChart(['0', '1', '2', '3', '4'], [{ name: 'y0', values: [1, 2, 3, 2, 1] }, { name: 'y1', values: [2, 4, 6, 4, 2] }, { name: 'y2', values: [3, 6, 9, 6, 3] }, { name: 'y3', values: [2, 4, 6, 4, 2] }], 'Поверхность', true),
    pieOfPie: ofPieChart(['Москва', 'СПб', 'Казань', 'Сочи', 'Омск', 'Тула'], [40, 25, 15, 8, 7, 5], 'Вторичная круговая', false),
    barOfPie: ofPieChart(['Москва', 'СПб', 'Казань', 'Сочи', 'Омск', 'Тула'], [40, 25, 15, 8, 7, 5], 'Вторичная гистограмма', true),
  };
}
