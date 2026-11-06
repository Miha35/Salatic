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

const wrap = (title: string | null, plot: string, legend = 'b'): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace ${NS}><c:chart>` +
  (title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>') +
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
