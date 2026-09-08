/**
 * Chart spec construction, merging and summarising — all pure.
 *
 * The Sheets `ChartSpec` is a nine-way union with a very large surface. These
 * tools expose a curated subset instead: a chart type, A1 source ranges, and
 * the handful of labels people actually ask for. Everything here translates
 * between that subset and the API shape, so the tool handlers stay I/O only.
 */

import { assertBareA1Range, assertSingleCell, parseA1Range, gridRangeToA1 } from './a1.js';

/** Chart types this connector builds. COMBO is excluded: it needs a per-series type. */
export const CURATED_CHART_TYPES = [
  'COLUMN', 'BAR', 'LINE', 'AREA', 'SCATTER', 'STEPPED_AREA', 'PIE',
] as const;
export type CuratedChartType = (typeof CURATED_CHART_TYPES)[number];

export const LEGEND_POSITIONS = [
  'BOTTOM', 'LEFT', 'RIGHT', 'TOP', 'NONE', 'LABELED',
] as const;
export type CuratedLegendPosition = (typeof LEGEND_POSITIONS)[number];

export const STACKED_TYPES = ['NOT_STACKED', 'STACKED', 'PERCENT_STACKED'] as const;
export type CuratedStackedType = (typeof STACKED_TYPES)[number];

/**
 * A chart with more series than this is unreadable, and each one costs a
 * source range in the spec. The cap turns a runaway request into a message
 * rather than a wall of JSON Google would reject anyway.
 */
export const MAX_SERIES = 50;

/** Charts reported by list_charts before the response is marked truncated. */
export const MAX_LISTED_CHARTS = 200;

/**
 * The nine mutually exclusive chart-type fields of a ChartSpec. Only one is
 * ever set; everything else at the top level is shared across types.
 */
const CHART_UNION_MEMBERS = new Set([
  'basicChart', 'pieChart', 'bubbleChart', 'candlestickChart', 'orgChart',
  'histogramChart', 'waterfallChart', 'treemapChart', 'scorecardChart',
]);

/**
 * `__proto__` from parsed JSON reassigns the prototype instead of becoming an
 * own property, which would shadow the `pieChart` / `basicChart` reads that
 * follow the merge. `constructor` and `prototype` become ordinary properties
 * and are swept up with it rather than because they carry the same hazard.
 */
const PROTOTYPE_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

function isPieType(chartType: CuratedChartType): boolean {
  return chartType === 'PIE';
}

/**
 * Which axis a series plots its values against.
 *
 * A bar chart runs horizontally, so its value axis is the bottom one and
 * Google rejects a bar series targeting any other with "Bar charts series may
 * only target the BOTTOM_AXIS". Every other basic type puts values on a
 * vertical axis.
 */
export function valueAxisFor(chartType: CuratedChartType): 'BOTTOM_AXIS' | 'LEFT_AXIS' {
  return chartType === 'BAR' ? 'BOTTOM_AXIS' : 'LEFT_AXIS';
}

/**
 * Only the filled types stack. Google rejects the field outright on the
 * others — "stackedType not supported when chartType is LINE" — rather than
 * ignoring it, so an inherited value has to be dropped on a retype and an
 * explicit one has to be refused before the request is sent. Confirmed
 * against the live API for both LINE and SCATTER.
 */
const STACKABLE_TYPES = new Set<CuratedChartType>([
  'COLUMN', 'BAR', 'AREA', 'STEPPED_AREA',
]);

export function supportsStacking(chartType: CuratedChartType): boolean {
  return STACKABLE_TYPES.has(chartType);
}

const INTERPOLATES_NULLS = new Set<CuratedChartType>(['LINE', 'AREA']);
const POINT_STYLED_TYPES = new Set<CuratedChartType>(['LINE', 'AREA', 'SCATTER']);

/**
 * Google renders 3D for pie and bar only. COLUMN looks like it should qualify
 * and does not: the live API answers 400 with "threeDimensional not supported
 * when chartType is COLUMN". Verified against the live API in both directions.
 */
export function supportsThreeDimensional(chartType: CuratedChartType): boolean {
  return chartType === 'PIE' || chartType === 'BAR';
}

function assertThreeDimensional(chartType: CuratedChartType): void {
  if (!supportsThreeDimensional(chartType)) {
    throw new Error(
      `three_dimensional is not supported on a ${chartType} chart. Google renders 3D for `
      + 'PIE and BAR only.',
    );
  }
}

/**
 * `LABELED_LEGEND` exists only on pie charts, and the two enums are otherwise
 * spelled the same, so one mapper serves both with a single guard.
 */
export function toApiLegendPosition(
  position: CuratedLegendPosition,
  chartType: CuratedChartType,
): string {
  if (position === 'LABELED') {
    if (!isPieType(chartType)) {
      throw new Error(
        'legend_position "LABELED" is only available on PIE charts. Use BOTTOM, LEFT, RIGHT, TOP or NONE.',
      );
    }
    return 'LABELED_LEGEND';
  }
  return position === 'NONE' ? 'NO_LEGEND' : `${position}_LEGEND`;
}

export function fromApiLegendPosition(position: string | undefined): CuratedLegendPosition | undefined {
  if (!position) return undefined;
  if (position === 'NO_LEGEND') return 'NONE';
  if (position === 'LABELED_LEGEND') return 'LABELED';
  const stem = position.replace(/_LEGEND$/, '');
  return (LEGEND_POSITIONS as readonly string[]).includes(stem)
    ? stem as CuratedLegendPosition
    : undefined;
}

type Orientation = 'columns' | 'rows';

interface RangeShape {
  rows: number;
  columns: number;
  orientation: Orientation;
  /** Number of data points the range contributes, along its long side. */
  length: number;
}

/**
 * A chart series is one column or one row. A rectangle is rejected rather
 * than silently reinterpreted: Google accepts the request but what it plots
 * is not what a caller passing "B2:D9" as one series expects.
 */
export function rangeShape(range: string, label: string): RangeShape {
  const grid = parseA1Range(assertBareA1Range(range, label));
  const rows = grid.endRowIndex - grid.startRowIndex;
  const columns = grid.endColumnIndex - grid.startColumnIndex;

  if (columns === 1) return { rows, columns, orientation: 'columns', length: rows };
  if (rows === 1) return { rows, columns, orientation: 'rows', length: columns };

  throw new Error(
    `${label} "${range}" is ${rows}x${columns}. Each range must be a single column `
    + '(e.g. "B1:B20") or a single row (e.g. "B1:T1"). Pass one range per series.',
  );
}

/**
 * Domain and series must line up or the chart plots against the wrong labels.
 * This is exact A1 arithmetic, so it costs nothing and runs before any call
 * to Google — the mismatch is reported in the caller's own terms rather than
 * as a 400, or worse, as a chart that silently drops points.
 */
export function assertAlignedRanges(domainRange: string, seriesRanges: string[]): void {
  if (seriesRanges.length === 0) {
    throw new Error('series_ranges must contain at least one range');
  }
  if (seriesRanges.length > MAX_SERIES) {
    throw new Error(
      `series_ranges has ${seriesRanges.length} entries, over the ${MAX_SERIES}-series limit.`,
    );
  }

  const domain = rangeShape(domainRange, 'domain_range');
  seriesRanges.forEach((range, i) => {
    const shape = rangeShape(range, `series_ranges[${i}]`);
    if (shape.orientation !== domain.orientation) {
      throw new Error(
        `series_ranges[${i}] "${range}" runs across ${shape.orientation} but domain_range `
        + `"${domainRange}" runs across ${domain.orientation}. All ranges must share an orientation.`,
      );
    }
    if (shape.length !== domain.length) {
      throw new Error(
        `series_ranges[${i}] "${range}" covers ${shape.length} cells but domain_range `
        + `"${domainRange}" covers ${domain.length}. Every series must be the same length as the domain.`,
      );
    }
  });
}

function chartData(range: string, sheetId: number): any {
  return { sourceRange: { sources: [{ sheetId, ...parseA1Range(range) }] } };
}

interface ChartDefinition {
  chartType: CuratedChartType;
  domainRange: string;
  seriesRanges: string[];
  title?: string;
  subtitle?: string;
  axisTitles?: { bottom?: string; left?: string; right?: string };
  legendPosition?: CuratedLegendPosition;
  stackedType?: CuratedStackedType;
  headerCount?: number;
  pieHole?: number;
  threeDimensional?: boolean;
}

/**
 * A bar chart has no usable right-hand axis. Verified live: Google ACCEPTS a
 * bar spec carrying a RIGHT_AXIS entry — 200, chart created — and then simply
 * does not persist it; reading the chart back shows BOTTOM and LEFT only.
 *
 * So this is refused rather than passed through, because passing it through
 * would report success for a title that silently never exists. Nothing in
 * BasicChartAxis documents the restriction, so the check rests on that
 * observation rather than on the docs.
 */
function assertAxisTitles(
  titles: ChartDefinition['axisTitles'],
  chartType: CuratedChartType,
): void {
  if (titles?.right !== undefined && chartType === 'BAR') {
    throw new Error(
      'axis_titles.right is not supported on a BAR chart — every bar series plots against '
      + 'the bottom axis, so a right-hand axis would have nothing to label.',
    );
  }
}

function buildAxes(titles: ChartDefinition['axisTitles']): any[] {
  const positions: Array<[keyof NonNullable<ChartDefinition['axisTitles']>, string]> = [
    ['bottom', 'BOTTOM_AXIS'],
    ['left', 'LEFT_AXIS'],
    ['right', 'RIGHT_AXIS'],
  ];
  return positions
    .filter(([key]) => titles?.[key] !== undefined)
    .map(([key, position]) => ({ position, title: titles![key] }));
}

export function buildChartSpec(definition: ChartDefinition, sourceSheetId: number): any {
  const {
    chartType, domainRange, seriesRanges, title, subtitle,
    axisTitles, legendPosition, stackedType, headerCount, pieHole, threeDimensional,
  } = definition;

  assertAlignedRanges(domainRange, seriesRanges);

  const spec: any = {};
  if (title !== undefined) spec.title = title;
  if (subtitle !== undefined) spec.subtitle = subtitle;

  if (isPieType(chartType)) {
    if (seriesRanges.length > 1) {
      throw new Error(
        `A PIE chart plots a single series, but ${seriesRanges.length} were given. `
        + 'Pass one range in series_ranges, or choose COLUMN/BAR/LINE to show several.',
      );
    }
    const pie: any = {
      domain: chartData(domainRange, sourceSheetId),
      series: chartData(seriesRanges[0], sourceSheetId),
    };
    if (legendPosition) pie.legendPosition = toApiLegendPosition(legendPosition, chartType);
    if (pieHole !== undefined) pie.pieHole = pieHole;
    if (threeDimensional !== undefined) pie.threeDimensional = threeDimensional;
    spec.pieChart = pie;
    return spec;
  }

  const basic: any = {
    chartType,
    domains: [{ domain: chartData(domainRange, sourceSheetId) }],
    series: seriesRanges.map((range) => ({
      series: chartData(range, sourceSheetId),
      targetAxis: valueAxisFor(chartType),
    })),
    headerCount: headerCount ?? 1,
  };
  if (legendPosition) basic.legendPosition = toApiLegendPosition(legendPosition, chartType);
  if (threeDimensional !== undefined) {
    assertThreeDimensional(chartType);
    basic.threeDimensional = threeDimensional;
  }
  if (stackedType) {
    if (!supportsStacking(chartType)) {
      throw new Error(
        `stacked_type is not supported on a ${chartType} chart — Google rejects it rather than `
        + 'ignoring it. Stacking applies to COLUMN, BAR, AREA and STEPPED_AREA only.',
      );
    }
    basic.stackedType = stackedType;
  }
  assertAxisTitles(axisTitles, chartType);
  const axis = buildAxes(axisTitles);
  if (axis.length > 0) basic.axis = axis;
  spec.basicChart = basic;
  return spec;
}

export interface ChartPlacement {
  newSheet?: boolean;
  anchorCell?: string;
  anchorSheetId?: number;
  widthPixels?: number;
  heightPixels?: number;
  offsetXPixels?: number;
  offsetYPixels?: number;
}

/**
 * Placement is required rather than defaulted. Anchoring somewhere arbitrary
 * would overlay the data it charts, and picking a "free" cell for the caller
 * is a layout decision the caller is better placed to make.
 */
export function buildPosition(placement: ChartPlacement): any {
  const { newSheet, anchorCell, anchorSheetId } = placement;

  if (newSheet && anchorCell) {
    throw new Error(
      'Pass either new_sheet: true or anchor_cell, not both — a chart lives on its own '
      + 'sheet or overlays an existing one.',
    );
  }
  if (newSheet) return { newSheet: true };
  if (!anchorCell) {
    throw new Error(
      'Placement is required: pass anchor_cell (e.g. "F2") to overlay the chart on a sheet, '
      + 'or new_sheet: true to put it on its own sheet.',
    );
  }
  if (anchorSheetId === undefined) {
    throw new Error('anchor_cell requires a resolved sheet id');
  }

  const cell = parseA1Range(assertSingleCell(anchorCell, 'anchor_cell'));
  const overlay: any = {
    anchorCell: {
      sheetId: anchorSheetId,
      rowIndex: cell.startRowIndex,
      columnIndex: cell.startColumnIndex,
    },
  };
  for (const key of ['widthPixels', 'heightPixels', 'offsetXPixels', 'offsetYPixels'] as const) {
    if (placement[key] !== undefined) overlay[key] = placement[key];
  }
  return { overlayPosition: overlay };
}

/** The placement wording add_chart and move_chart both report back. */
export function describePlacement(
  newSheet: boolean | undefined,
  sheetName: string | undefined,
  anchorCell: string | undefined,
): string {
  return newSheet ? 'own chart sheet' : `overlaying ${sheetName}!${anchorCell}`;
}

/**
 * `updateEmbeddedObjectPosition` needs a field mask naming the OverlayPosition
 * fields to write. Omitted pixel fields are left alone rather than reset, so
 * resizing a chart does not also move it.
 */
export function overlayFieldMask(placement: ChartPlacement): string {
  const fields = ['anchorCell'];
  for (const key of ['widthPixels', 'heightPixels', 'offsetXPixels', 'offsetYPixels'] as const) {
    if (placement[key] !== undefined) fields.push(key);
  }
  return fields.join(',');
}

interface ExtractedData {
  domain?: any;
  series: any[];
  /** The surviving BasicChartSeries entries, index-aligned with `series`. */
  seriesEntries?: any[];
  headerCount?: number;
}

/** Reads both union members into one shape, so either can be converted to the other. */
function extractChartData(spec: any): ExtractedData {
  if (spec?.pieChart) {
    return {
      domain: spec.pieChart.domain,
      series: spec.pieChart.series ? [spec.pieChart.series] : [],
    };
  }
  const basic = spec?.basicChart;
  // Entries are filtered as pairs. Mapping to `.series` and filtering after
  // would leave the surviving data at indices that no longer line up with
  // `basicChart.series`, landing each entry's styling on the wrong series.
  const entries = (basic?.series ?? []).filter((s: any) => s?.series);
  return {
    domain: basic?.domains?.[0]?.domain,
    series: entries.map((s: any) => s.series),
    seriesEntries: entries,
    headerCount: basic?.headerCount,
  };
}

export function chartTypeOf(spec: any): CuratedChartType | undefined {
  if (spec?.pieChart) return 'PIE';
  const type = spec?.basicChart?.chartType;
  return (CURATED_CHART_TYPES as readonly string[]).includes(type)
    ? type as CuratedChartType
    : undefined;
}

/**
 * The name to report for a spec, as opposed to the name these tools can build
 * from. A chart this connector cannot rebuild is still listed, moved, deleted
 * and retitled, so it still has to be named. COMBO is the near case — a
 * basicChart whose own chartType is more informative than the member name — so
 * prefer that and fall back to the member for the genuinely foreign kinds.
 */
export function chartTypeLabel(spec: any): string | undefined {
  const type = chartTypeOf(spec);
  if (type) return type;
  if (!spec) return undefined;
  return spec.basicChart?.chartType ?? Object.keys(spec).find((k) => k.endsWith('Chart'));
}

interface ChartPatch {
  chartType?: CuratedChartType;
  domainRange?: string;
  seriesRanges?: string[];
  title?: string;
  subtitle?: string;
  axisTitles?: { bottom?: string; left?: string; right?: string };
  legendPosition?: CuratedLegendPosition;
  stackedType?: CuratedStackedType;
  headerCount?: number;
  pieHole?: number;
  threeDimensional?: boolean;
}

/**
 * Resolves the axis array for a retype.
 *
 * A bar chart transposes the plot — categories up the left, values along the
 * bottom, the reverse of every other basic type — so titles carried across such
 * a conversion swap positions to follow their data. Patch-supplied titles are
 * already in the target type's terms and are applied after the swap.
 */
function mergeAxes(
  previousBasic: any,
  patchTitles: ChartDefinition['axisTitles'],
  chartType: CuratedChartType,
): any[] {
  const flipAxes = previousBasic !== undefined
    && (previousBasic.chartType === 'BAR') !== (chartType === 'BAR');
  const swapPosition = (position: string) => {
    if (!flipAxes) return position;
    if (position === 'BOTTOM_AXIS') return 'LEFT_AXIS';
    if (position === 'LEFT_AXIS') return 'BOTTOM_AXIS';
    return position;
  };

  const axes = new Map<string, any>();
  for (const axis of previousBasic?.axis ?? []) {
    if (!axis?.position) continue;
    // Google does not persist a RIGHT_AXIS entry on a bar chart — see
    // assertAxisTitles. Carrying one over would put it in the request and
    // then read it back missing.
    if (chartType === 'BAR' && axis.position === 'RIGHT_AXIS') continue;
    const position = swapPosition(axis.position);
    axes.set(position, { ...axis, position });
  }
  for (const axis of buildAxes(patchTitles)) {
    axes.set(axis.position, { ...(axes.get(axis.position) ?? {}), ...axis });
  }
  return [...axes.values()];
}

/**
 * `UpdateChartSpecRequest` carries no field mask: whatever is sent replaces
 * the chart's spec entirely. So an edit is read-modify-write, and this is the
 * modify — every field the caller did not name has to be carried across from
 * the existing spec or it is silently dropped.
 *
 * Two merges are not plain overwrites. Axes merge by `position`, since
 * BOTTOM/LEFT/RIGHT are the identity of an axis and its index is not stable.
 * Source ranges replace wholesale, because a half-updated set of series would
 * be misaligned with its domain.
 */
export function mergeChartSpec(
  existing: any,
  patch: ChartPatch,
  sourceSheetId?: number,
): any {
  // Renaming needs no understanding of the chart. Short-circuiting here lets
  // update_chart retitle a COMBO, waterfall, treemap or scorecard — kinds
  // list_charts reports but the merge below cannot rebuild — instead of
  // refusing, or offering a conversion that would flatten their per-series
  // types. The union member is passed through untouched, so the result has no
  // modeled type and callers must name it with chartTypeLabel.
  const patchesOnlyText = Object.entries(patch)
    .every(([key, value]) => value === undefined || key === 'title' || key === 'subtitle');
  if (patchesOnlyText && chartTypeOf(existing) === undefined) {
    const renamed: any = { ...existing };
    if (patch.title !== undefined) renamed.title = patch.title;
    if (patch.subtitle !== undefined) renamed.subtitle = patch.subtitle;
    return renamed;
  }

  const previousType = chartTypeOf(existing);
  const chartType = patch.chartType ?? previousType;
  if (!chartType) {
    throw new Error(
      'This chart is a type these tools do not model (only column, bar, line, area, '
      + 'scatter, stepped-area and pie are supported), so this edit cannot be applied to it. '
      + 'Title and subtitle can still be changed on their own. For anything else pass '
      + 'chart_type to convert it, or delete and recreate it.',
    );
  }

  if ((patch.domainRange === undefined) !== (patch.seriesRanges === undefined)) {
    throw new Error(
      'domain_range and series_ranges must be changed together — a new domain with the old '
      + 'series (or the reverse) would be misaligned.',
    );
  }

  const previous = extractChartData(existing);
  let domain = previous.domain;
  let series = previous.series;

  if (patch.domainRange !== undefined && patch.seriesRanges !== undefined) {
    if (sourceSheetId === undefined) {
      throw new Error('Changing source ranges requires a resolved sheet id');
    }
    assertAlignedRanges(patch.domainRange, patch.seriesRanges);
    domain = chartData(patch.domainRange, sourceSheetId);
    series = patch.seriesRanges.map((range) => chartData(range, sourceSheetId));
  }

  if (!domain || series.length === 0) {
    throw new Error(
      'The existing chart has no readable source ranges, so it cannot be patched. '
      + 'Pass domain_range and series_ranges to set them.',
    );
  }

  // Everything except the chart-type union is carried across untouched. An
  // allowlist was wrong here: a spec is sent whole, so any field left out is
  // erased, and ChartSpec holds plenty this connector does not model —
  // titleTextFormat, altText, hiddenDimensionStrategy, filterSpecs. Dropping
  // a title style someone set in the Sheets UI is exactly the silent loss
  // read-modify-write exists to prevent.
  const spec: any = {};
  for (const [key, value] of Object.entries(existing ?? {})) {
    if (!CHART_UNION_MEMBERS.has(key) && !PROTOTYPE_KEYS.has(key)) spec[key] = value;
  }
  if (patch.title !== undefined) spec.title = patch.title;
  if (patch.subtitle !== undefined) spec.subtitle = patch.subtitle;

  // Legend placement means the same thing in both unions, so it survives a
  // conversion. LABELED is the one value that does not exist outside pie, and
  // converting away from pie drops it rather than failing the whole edit.
  const previousLegend = fromApiLegendPosition(
    existing?.pieChart?.legendPosition ?? existing?.basicChart?.legendPosition,
  );
  const legendPosition = patch.legendPosition
    ?? (previousLegend === 'LABELED' && !isPieType(chartType) ? undefined : previousLegend);

  if (isPieType(chartType)) {
    if (series.length > 1) {
      throw new Error(
        `A PIE chart plots a single series, but this chart has ${series.length}. `
        + 'Pass series_ranges with exactly one range to convert it.',
      );
    }
    // Seeded, not allowlisted — see the carry-over loop above.
    const pie: any = {
      ...(existing?.pieChart ?? {}),
      domain,
      series: series[0],
    };
    if (legendPosition) pie.legendPosition = toApiLegendPosition(legendPosition, chartType);
    const pieHole = patch.pieHole ?? existing?.pieChart?.pieHole;
    if (pieHole !== undefined) pie.pieHole = pieHole;
    const threeD = patch.threeDimensional
      ?? existing?.pieChart?.threeDimensional
      ?? existing?.basicChart?.threeDimensional;
    if (threeD !== undefined) pie.threeDimensional = threeD;
    spec.pieChart = pie;
    return spec;
  }

  // The previous member is read directly rather than matched against the
  // curated type: COMBO is a basicChart this connector cannot name, so a
  // type comparison would treat COMBO -> COLUMN as a union change and throw
  // the whole basicChart away.
  const previousBasic = existing?.basicChart;

  // Seeded, not allowlisted. BasicChartSpec carries more than this connector
  // models — interpolateNulls, lineSmoothing, compareMode, totalDataLabel — and
  // listing only the known keys would erase the rest on every edit. The prune
  // block below then drops back the few the target type cannot carry.
  const basic: any = {
    ...(previousBasic ?? {}),
    chartType,
    domains: [{ ...(previousBasic?.domains?.[0] ?? {}), domain }],
    // Per-series styling is preserved where a series survives in place and the
    // target type can render it — the prune block below drops the rest; a
    // series added by the patch has none, which is what a new series should be.
    //
    // The target axis cannot simply be carried over: converting to BAR must
    // move every series to BOTTOM_AXIS, and converting away from it must move
    // them back. Outside BAR a deliberate RIGHT_AXIS assignment survives, so a
    // dual-axis chart keeps its second axis; converting TO bar collapses it,
    // because Google allows bar series on the bottom axis only.
    series: series.map((data, i) => {
      const prior = previous.seriesEntries?.[i];
      return {
        ...(prior ?? {}),
        series: data,
        targetAxis: chartType !== 'BAR' && prior?.targetAxis === 'RIGHT_AXIS'
          ? 'RIGHT_AXIS'
          : valueAxisFor(chartType),
      };
    }),
    headerCount: patch.headerCount ?? previous.headerCount ?? 1,
  };
  if (legendPosition) basic.legendPosition = toApiLegendPosition(legendPosition, chartType);
  // Dropping LABELED takes a delete, not just a skipped write: `basic` is
  // spread-seeded, so LABELED_LEGEND arrives on its own and BasicChartSpec has
  // no such value. Only this one value is removed — a legend position this
  // connector cannot map is unmodeled, not invalid, and the seeding keeps it.
  else if (previousLegend === 'LABELED') delete basic.legendPosition;

  // Carried over only where the target type accepts it. A retype from COLUMN to
  // LINE inherits STACKED, which Google refuses outright, so an edit that never
  // mentioned stacking would fail on a field the caller did not set.
  if (patch.stackedType && !supportsStacking(chartType)) {
    throw new Error(
      `stacked_type is not supported on a ${chartType} chart — Google rejects it rather than `
      + 'ignoring it. Stacking applies to COLUMN, BAR, AREA and STEPPED_AREA only.',
    );
  }
  const stackedType = patch.stackedType ?? previousBasic?.stackedType;
  if (stackedType && supportsStacking(chartType)) basic.stackedType = stackedType;
  else delete basic.stackedType;

  // Fields the spread carries that Google only accepts under a condition the
  // new type may no longer meet. `stackedType` on LINE is a confirmed 400, and
  // the docs phrase these the same way ("valid only if" / "applies to"), so all
  // are pruned by the same principle rather than a subset — a retype must not
  // fail on a field the caller never named.
  //
  // `NOT_STACKED` is truthy, so a plain "turn stacking off" edit used to keep a
  // totalDataLabel that Google then rejected. Test the two stacking values.
  const isStacked = basic.stackedType === 'STACKED' || basic.stackedType === 'PERCENT_STACKED';
  if (!isStacked) delete basic.totalDataLabel;
  if (chartType !== 'LINE') delete basic.lineSmoothing;
  if (!INTERPOLATES_NULLS.has(chartType)) delete basic.interpolateNulls;
  if (!POINT_STYLED_TYPES.has(chartType)) {
    for (const entry of basic.series) {
      delete entry.lineStyle;
      delete entry.pointStyle;
      // Same "valid only if AREA, LINE, or SCATTER" wording, one level down.
      // Copied, not mutated: `entry` is a shallow spread of the caller's series,
      // so its styleOverrides array and objects are still the caller's. This
      // function is pure and a deleting write here reaches back into `existing`.
      if (entry.styleOverrides) {
        entry.styleOverrides = entry.styleOverrides.map(
          ({ pointStyle, ...rest }: any) => rest,
        );
      }
    }
  }
  // A custom data label reads from the column beside its series ("this source
  // data must come from the next column in the source data"), so replacing the
  // series ranges leaves it pointing at the old data. The label text is not
  // recoverable from here, so the custom labelling is dropped rather than left
  // silently wrong.
  if (patch.seriesRanges !== undefined) {
    for (const entry of basic.series) {
      if (entry.dataLabel?.customLabelData) delete entry.dataLabel;
    }
  }

  if (patch.threeDimensional !== undefined) assertThreeDimensional(chartType);
  const carriedThreeD = patch.threeDimensional ?? existing?.pieChart?.threeDimensional;
  if (carriedThreeD !== undefined) basic.threeDimensional = carriedThreeD;
  // NOT redundant with the assignment above: `basic` is spread-seeded, so a
  // previous threeDimensional arrives without ever being assigned here. This
  // delete is the only thing that removes it on a retype away from bar/column.
  if (!supportsThreeDimensional(chartType)) delete basic.threeDimensional;
  for (const entry of basic.series) delete entry.type;

  assertAxisTitles(patch.axisTitles, chartType);
  // Settled on both paths. The spread above already put the previous `axis` on
  // `basic`, so writing only when non-empty would leave a stale array behind
  // exactly when it empties — the RIGHT_AXIS-only case the drop exists for.
  const axis = mergeAxes(previousBasic, patch.axisTitles, chartType);
  if (axis.length > 0) basic.axis = axis;
  else delete basic.axis;

  spec.basicChart = basic;
  return spec;
}

interface RangeRef {
  sheetName?: string;
  range: string;
}

function describeChartData(
  data: any,
  sheetNameById: Map<number, string>,
): RangeRef | undefined {
  const source = data?.sourceRange?.sources?.[0];
  if (!source) return undefined;
  try {
    return {
      ...(source.sheetId !== undefined && sheetNameById.has(source.sheetId)
        ? { sheetName: sheetNameById.get(source.sheetId)! }
        : {}),
      range: gridRangeToA1(source),
    };
  } catch {
    // An unbounded source (a whole-column chart range) has no A1 spelling.
    return undefined;
  }
}

export interface ChartSummary {
  chartId: number;
  title?: string;
  chartType?: string;
  onOwnSheet: boolean;
  anchor?: { sheetName?: string; cell: string };
  domain?: RangeRef;
  series: RangeRef[];
  /**
   * Source ranges omitted because they have no bounded A1 spelling — a chart
   * built over whole columns, which is what the Sheets UI produces. Without
   * this an empty `series` is indistinguishable from a chart that genuinely
   * has none.
   */
  unreportableRanges?: number;
  spec?: any;
}

export function summarizeChart(
  chart: any,
  sheetNameById: Map<number, string>,
  objectSheetIds: Set<number>,
  includeSpec: boolean,
): ChartSummary {
  const spec = chart?.spec;
  const { domain, series } = extractChartData(spec);
  const anchorCell = chart?.position?.overlayPosition?.anchorCell;
  const positionSheetId = chart?.position?.sheetId ?? anchorCell?.sheetId;

  const describedSeries = series.map((data: any) => describeChartData(data, sheetNameById));
  const domainRef = domain ? describeChartData(domain, sheetNameById) : undefined;
  const unreportable = describedSeries.filter((ref) => ref === undefined).length
    + (domain !== undefined && domainRef === undefined ? 1 : 0);

  const summary: ChartSummary = {
    chartId: chart?.chartId,
    onOwnSheet: positionSheetId !== undefined && objectSheetIds.has(positionSheetId),
    series: describedSeries.filter((ref): ref is RangeRef => ref !== undefined),
    ...(unreportable > 0 && { unreportableRanges: unreportable }),
  };

  if (spec?.title !== undefined) summary.title = spec.title;
  const named = chartTypeLabel(spec);
  if (named) summary.chartType = named;

  if (anchorCell) {
    summary.anchor = {
      ...(anchorCell.sheetId !== undefined && sheetNameById.has(anchorCell.sheetId)
        ? { sheetName: sheetNameById.get(anchorCell.sheetId)! }
        : {}),
      cell: gridRangeToA1({
        startRowIndex: anchorCell.rowIndex ?? 0,
        endRowIndex: (anchorCell.rowIndex ?? 0) + 1,
        startColumnIndex: anchorCell.columnIndex ?? 0,
        endColumnIndex: (anchorCell.columnIndex ?? 0) + 1,
      }).split(':')[0],
    };
  }

  if (domainRef) summary.domain = domainRef;
  if (includeSpec) summary.spec = spec;

  return summary;
}
