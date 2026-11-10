const fields = (names, children = {}) => ({
  ...Object.fromEntries(names.split(' ').filter(Boolean).map(name => [name, null])),
  ...children,
});
const dictionaryValues = Symbol('dictionaryValues');
const dictionary = schema => ({ [dictionaryValues]: schema });
const fontFamily = fields('latin ea mac win');
const border = fields('color width style');
const shadow = fields('color blur offset');
const fill = fields('type color gradientType angle src opacity', {
  stops: fields('position color'),
});
const textStyle = fields('fontSize bold italic underline strike baseline highlight color lineHeight letterSpacing wrap align verticalAlign textDirection fit overflow style literalEscapes', {
  fontFamily, shadow, gradient: fill,
});
const paragraph = fields('fontSize marginLeft marginRight indent align lineHeight lineSpacing spaceBefore spaceAfter', {
  runs: fields('text', { options: textStyle }),
  numbering: fields('type startAt'),
  tabStops: fields('position alignment'),
});
const textContent = fields('text nativeTextBody textBox margins fontScale sourceAutoFit', {
  ...textStyle, style: textStyle, paragraphs: paragraph,
});
const cell = fields('text rowSpan colSpan', { ...textStyle, textStyle, fill, border });
const tableStyle = fields('rowOverColumn', {
  cellStyle: cell, firstRowStyle: cell, lastRowStyle: cell,
  firstColumnStyle: cell, lastColumnStyle: cell, bodyStyles: cell,
});
const axis = fields('min max majorUnit title hidden majorTick minorTick type position inverse', {
  label: fields('fontSize color rotate position', { fontFamily }),
  line: border, gridLine: border,
});
const dataLabels = fields('show content fontSize color position', { fontFamily });
const series = fields('type name width stack nullHandling innerRadius yAxisIndex xAxisIndex smooth', {
  encode: fields('x y category value size open close low high source target weight color label'),
  fill, lineColor: fill, areaColor: fill, border,
  marker: fields('shape size width', { fill, lineColor: fill, border }),
  dataLabels,
  dataFilter: fields('col value'),
  upBars: fields('', { fill, border }), downBars: fields('', { fill, border }),
  totalBars: fields('', { fill, border }), increaseBars: fields('', { fill, border }),
  decreaseBars: fields('', { fill, border }),
});
const base = fields('elementId elementType sourceLayout bounds');
const transformed = fields('rotation opacity flip', base);
export const ELEMENT_SCHEMAS = {
  text: fields('', { ...transformed, content: textContent }),
  shape: fields('shapeName adjustments viewBox path', {
    ...transformed, fill, border, shadow,
    nativeShape: fields('geometry lineEnds fill line modelDigest paintDigest lineDigest'),
  }),
  line: fields('viewBox points curve arrow', { ...transformed, border, shadow }),
  image: fields('src', {
    ...transformed, border, shadow,
    fit: fields('mode'), crop: fields('left top right bottom'),
    cropShape: fields('shapeName adjustments viewBox path'),
  }),
  icon: fields('iconName color', { ...transformed, fill, border, shadow }),
  table: fields('columnWidths rowHeights', { ...base, rows: cell, style: tableStyle, fill, shadow }),
  chart: fields('barWidth barGap categoryGap', {
    ...base, fontFamily, fill, border, shadow,
    data: fields('cols rows'), series,
    seriesDefaults: fields('bar line area scatter bubble pie radar candlestick waterfall heatmap treemap sunburst sankey',
      Object.fromEntries('bar line area scatter bubble pie radar candlestick waterfall heatmap treemap sunburst sankey'.split(' ').map(type => [type, series]))),
    xAxis: axis, yAxis: axis, spokeAxis: axis,
    plotLayout: fields('x y w h'), plotBorder: border,
    title: fields('text fontSize color', { fontFamily }),
    legend: fields('show position fontSize color', { fontFamily }), dataLabels,
  }),
};
export const MANIFEST_SCHEMA = fields('version title size pages', {
  template: fields('id name sourceFile sourceSha256'),
  theme: fields('', {
    colors: dictionary(null), fonts: dictionary(fontFamily),
    textStyles: dictionary(textStyle), tableStyles: dictionary(tableStyle),
  }),
});
export const PAGE_SCHEMA = fields('pageType notes elements', {
  templateReference: fields('sourceSlideNumber rationale contentRelationship'),
  background: fill,
});
export const MANIFEST_FIELDS = new Set(Object.keys(MANIFEST_SCHEMA));
export const PAGE_FIELDS = new Set(Object.keys(PAGE_SCHEMA));
export const ELEMENT_FIELDS = Object.fromEntries(Object.entries(ELEMENT_SCHEMAS).map(([type, schema]) => [type, new Set(Object.keys(schema))]));
export const MISPLACED_TEXT_STYLE_FIELDS = new Set(Object.keys(textStyle));

/** Field names are closed; theme entry names and tabular data remain user-defined. */
export function unknownFieldPaths(value, schema, prefix = '', ancestors = new Set()) {
  if (schema === null || value === null || typeof value !== 'object' || ancestors.has(value)) return [];
  ancestors.add(value);
  const unknown = [];
  if (Array.isArray(value)) {
    for (const [index, item] of value.entries()) unknown.push(...unknownFieldPaths(item, schema, `${prefix}[${index}]`, ancestors));
  } else {
    for (const [key, child] of Object.entries(value)) {
      const field = prefix === '' ? key : `${prefix}.${key}`;
      if (dictionaryValues in schema) unknown.push(...unknownFieldPaths(child, schema[dictionaryValues], field, ancestors));
      else if (Object.hasOwn(schema, key)) unknown.push(...unknownFieldPaths(child, schema[key], field, ancestors));
      else unknown.push(field);
    }
  }
  ancestors.delete(value);
  return unknown;
}
