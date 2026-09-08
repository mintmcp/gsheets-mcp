import { describe, it, expect } from 'vitest';
import {
  toApiLegendPosition,
  fromApiLegendPosition,
  rangeShape,
  assertAlignedRanges,
  buildChartSpec,
  buildPosition,
  describePlacement,
  overlayFieldMask,
  mergeChartSpec,
  chartTypeOf,
  chartTypeLabel,
  summarizeChart,
  valueAxisFor,
  supportsStacking,
  supportsThreeDimensional,
  MAX_SERIES,
} from '../lib/chart.js';
import { gridRangeToA1 } from '../lib/a1.js';

describe('gridRangeToA1', () => {
  it('is the inverse of parseA1Range', () => {
    expect(gridRangeToA1({
      startRowIndex: 0, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 1,
    })).toBe('A1:A8');
    expect(gridRangeToA1({
      startRowIndex: 1, endRowIndex: 3, startColumnIndex: 1, endColumnIndex: 4,
    })).toBe('B2:D3');
  });

  it('handles columns past Z', () => {
    expect(gridRangeToA1({
      startRowIndex: 0, endRowIndex: 1, startColumnIndex: 26, endColumnIndex: 27,
    })).toBe('AA1:AA1');
  });

  it('rejects an unbounded range rather than guessing', () => {
    expect(() => gridRangeToA1({ startRowIndex: 0, startColumnIndex: 0 }))
      .toThrow(/fully bounded/);
  });
});

describe('legend position mapping', () => {
  it('suffixes the enum', () => {
    expect(toApiLegendPosition('BOTTOM', 'COLUMN')).toBe('BOTTOM_LEGEND');
    expect(toApiLegendPosition('RIGHT', 'LINE')).toBe('RIGHT_LEGEND');
  });

  it('maps NONE to NO_LEGEND', () => {
    expect(toApiLegendPosition('NONE', 'COLUMN')).toBe('NO_LEGEND');
  });

  it('allows LABELED only on pie', () => {
    expect(toApiLegendPosition('LABELED', 'PIE')).toBe('LABELED_LEGEND');
    expect(() => toApiLegendPosition('LABELED', 'COLUMN')).toThrow(/only available on PIE/);
  });

  it('round-trips', () => {
    for (const p of ['BOTTOM', 'LEFT', 'RIGHT', 'TOP', 'NONE'] as const) {
      expect(fromApiLegendPosition(toApiLegendPosition(p, 'COLUMN'))).toBe(p);
    }
    expect(fromApiLegendPosition(toApiLegendPosition('LABELED', 'PIE'))).toBe('LABELED');
  });

  it('returns undefined for absent or unknown values', () => {
    expect(fromApiLegendPosition(undefined)).toBeUndefined();
    expect(fromApiLegendPosition('SOMETHING_LEGEND')).toBeUndefined();
  });
});

describe('rangeShape', () => {
  it('reads a single column', () => {
    expect(rangeShape('B1:B20', 'r')).toMatchObject({ orientation: 'columns', length: 20 });
  });

  it('reads a single row', () => {
    expect(rangeShape('B1:T1', 'r')).toMatchObject({ orientation: 'rows', length: 19 });
  });

  it('treats a single cell as column-oriented', () => {
    expect(rangeShape('B1', 'r')).toMatchObject({ orientation: 'columns', length: 1 });
  });

  it('rejects a rectangle, naming the parameter', () => {
    expect(() => rangeShape('B2:D9', 'series_ranges[0]'))
      .toThrow(/series_ranges\[0\] "B2:D9" is 8x3/);
  });
});

describe('assertAlignedRanges', () => {
  it('accepts matching column ranges', () => {
    expect(() => assertAlignedRanges('A1:A8', ['B1:B8', 'C1:C8'])).not.toThrow();
  });

  it('accepts matching row ranges', () => {
    expect(() => assertAlignedRanges('B1:H1', ['B2:H2'])).not.toThrow();
  });

  it('rejects a length mismatch', () => {
    expect(() => assertAlignedRanges('A1:A8', ['B1:B9']))
      .toThrow(/covers 9 cells but domain_range "A1:A8" covers 8/);
  });

  it('rejects mixed orientation', () => {
    expect(() => assertAlignedRanges('A1:A8', ['B1:H1']))
      .toThrow(/must share an orientation/);
  });

  it('rejects an empty series list', () => {
    expect(() => assertAlignedRanges('A1:A8', [])).toThrow(/at least one range/);
  });

  it('caps the series count', () => {
    const many = Array.from({ length: MAX_SERIES + 1 }, () => 'B1:B8');
    expect(() => assertAlignedRanges('A1:A8', many)).toThrow(/over the 50-series limit/);
  });
});

describe('buildChartSpec', () => {
  const base = { domainRange: 'A1:A8', seriesRanges: ['B1:B8', 'C1:C8'] };

  it('builds a basic chart with GridRange sources on the given sheet', () => {
    const spec = buildChartSpec({ chartType: 'COLUMN', ...base, title: 'Q1' }, 42);
    expect(spec.title).toBe('Q1');
    expect(spec.basicChart.chartType).toBe('COLUMN');
    expect(spec.basicChart.domains[0].domain.sourceRange.sources[0]).toEqual({
      sheetId: 42, startRowIndex: 0, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 1,
    });
    expect(spec.basicChart.series).toHaveLength(2);
    expect(spec.basicChart.series[1].series.sourceRange.sources[0].startColumnIndex).toBe(2);
    expect(spec.basicChart.series[0].targetAxis).toBe('LEFT_AXIS');
  });

  it('defaults headerCount to 1 and honours an explicit 0', () => {
    expect(buildChartSpec({ chartType: 'BAR', ...base }, 1).basicChart.headerCount).toBe(1);
    expect(buildChartSpec({ chartType: 'BAR', ...base, headerCount: 0 }, 1).basicChart.headerCount).toBe(0);
  });

  it('emits only the axes that were titled', () => {
    const spec = buildChartSpec(
      { chartType: 'LINE', ...base, axisTitles: { bottom: 'Model', left: 'Sales' } }, 1,
    );
    expect(spec.basicChart.axis).toEqual([
      { position: 'BOTTOM_AXIS', title: 'Model' },
      { position: 'LEFT_AXIS', title: 'Sales' },
    ]);
  });

  it('omits the axis key entirely when no titles are given', () => {
    expect(buildChartSpec({ chartType: 'LINE', ...base }, 1).basicChart.axis).toBeUndefined();
  });

  it('builds a pie chart from a single series', () => {
    const spec = buildChartSpec({
      chartType: 'PIE', domainRange: 'A1:A8', seriesRanges: ['B1:B8'], pieHole: 0.4,
    }, 7);
    expect(spec.basicChart).toBeUndefined();
    expect(spec.pieChart.pieHole).toBe(0.4);
    expect(spec.pieChart.series.sourceRange.sources[0].sheetId).toBe(7);
  });

  it('refuses a pie chart with several series', () => {
    expect(() => buildChartSpec({ chartType: 'PIE', ...base }, 1))
      .toThrow(/PIE chart plots a single series, but 2 were given/);
  });

  it('validates ranges before building', () => {
    expect(() => buildChartSpec({ chartType: 'COLUMN', domainRange: 'A1:A8', seriesRanges: ['B1:B9'] }, 1))
      .toThrow(/covers 9 cells/);
  });
});

describe('buildPosition', () => {
  it('builds an own-sheet position', () => {
    expect(buildPosition({ newSheet: true })).toEqual({ newSheet: true });
  });

  it('builds an overlay position from an A1 cell', () => {
    expect(buildPosition({ anchorCell: 'F2', anchorSheetId: 3, widthPixels: 600 }))
      .toEqual({
        overlayPosition: {
          anchorCell: { sheetId: 3, rowIndex: 1, columnIndex: 5 },
          widthPixels: 600,
        },
      });
  });

  it('rejects both placements at once', () => {
    expect(() => buildPosition({ newSheet: true, anchorCell: 'F2', anchorSheetId: 1 }))
      .toThrow(/not both/);
  });

  it('rejects neither, rather than defaulting over the data', () => {
    expect(() => buildPosition({})).toThrow(/Placement is required/);
  });

  it('rejects a range where a cell is required', () => {
    expect(() => buildPosition({ anchorCell: 'F2:G3', anchorSheetId: 1 }))
      .toThrow(/anchor_cell must be a single cell/);
  });
});

describe('overlayFieldMask', () => {
  it('always names the anchor', () => {
    expect(overlayFieldMask({ anchorCell: 'A1' })).toBe('anchorCell');
  });

  it('names only the pixel fields that were supplied', () => {
    expect(overlayFieldMask({ anchorCell: 'A1', widthPixels: 400, offsetYPixels: 10 }))
      .toBe('anchorCell,widthPixels,offsetYPixels');
  });
});

describe('describePlacement', () => {
  it('names the chart sheet when there is no anchor', () => {
    expect(describePlacement(true, 'Data', undefined)).toBe('own chart sheet');
  });

  it('qualifies the anchor with the tab it overlays', () => {
    expect(describePlacement(undefined, 'Data', 'F2')).toBe('overlaying Data!F2');
  });
});

describe('chartTypeOf', () => {
  it('reads both union members', () => {
    expect(chartTypeOf({ basicChart: { chartType: 'COLUMN' } })).toBe('COLUMN');
    expect(chartTypeOf({ pieChart: {} })).toBe('PIE');
  });

  it('returns undefined for a kind these tools do not model', () => {
    expect(chartTypeOf({ waterfallChart: {} })).toBeUndefined();
    expect(chartTypeOf({ basicChart: { chartType: 'COMBO' } })).toBeUndefined();
  });
});

describe('mergeChartSpec', () => {
  const source = (col: number) => ({
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 8,
        startColumnIndex: col, endColumnIndex: col + 1,
      }],
    },
  });

  const existing = {
    title: 'Original',
    fontName: 'Roboto',
    basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'BOTTOM_LEGEND',
      headerCount: 1,
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Model' },
        { position: 'LEFT_AXIS', title: 'Sales' },
      ],
      domains: [{ domain: source(0) }],
      series: [
        { series: source(1), targetAxis: 'LEFT_AXIS', colorStyle: { themeColor: 'ACCENT1' } },
      ],
    },
  };

  it('carries over every field the patch does not name', () => {
    const merged = mergeChartSpec(existing, { title: 'Renamed' });
    expect(merged.title).toBe('Renamed');
    expect(merged.fontName).toBe('Roboto');
    expect(merged.basicChart.chartType).toBe('COLUMN');
    expect(merged.basicChart.legendPosition).toBe('BOTTOM_LEGEND');
    expect(merged.basicChart.headerCount).toBe(1);
    expect(merged.basicChart.domains[0].domain).toEqual(source(0));
  });

  it('preserves per-series styling on a series that survives', () => {
    const merged = mergeChartSpec(existing, { title: 'x' });
    expect(merged.basicChart.series[0].colorStyle).toEqual({ themeColor: 'ACCENT1' });
  });

  it('merges axes by position, not index', () => {
    const merged = mergeChartSpec(existing, { axisTitles: { left: 'Revenue' } });
    expect(merged.basicChart.axis).toEqual([
      { position: 'BOTTOM_AXIS', title: 'Model' },
      { position: 'LEFT_AXIS', title: 'Revenue' },
    ]);
  });

  it('adds an axis the chart did not have', () => {
    const merged = mergeChartSpec(existing, { axisTitles: { right: 'Margin' } });
    expect(merged.basicChart.axis).toHaveLength(3);
    expect(merged.basicChart.axis[2]).toEqual({ position: 'RIGHT_AXIS', title: 'Margin' });
  });

  it('changes chart type within the basic union', () => {
    expect(mergeChartSpec(existing, { chartType: 'LINE' }).basicChart.chartType).toBe('LINE');
  });

  it('converts a basic chart to pie, keeping its data', () => {
    const merged = mergeChartSpec(existing, { chartType: 'PIE' });
    expect(merged.basicChart).toBeUndefined();
    expect(merged.pieChart.domain).toEqual(source(0));
    expect(merged.pieChart.series).toEqual(source(1));
    expect(merged.pieChart.legendPosition).toBe('BOTTOM_LEGEND');
  });

  it('converts pie back to a basic chart', () => {
    const pie = { title: 'P', pieChart: { domain: source(0), series: source(1), pieHole: 0.5 } };
    const merged = mergeChartSpec(pie, { chartType: 'BAR' });
    expect(merged.pieChart).toBeUndefined();
    expect(merged.basicChart.chartType).toBe('BAR');
    expect(merged.basicChart.domains[0].domain).toEqual(source(0));
    expect(merged.basicChart.series[0].series).toEqual(source(1));
  });

  it('refuses to convert a multi-series chart to pie', () => {
    const twoSeries = {
      basicChart: {
        ...existing.basicChart,
        series: [{ series: source(1) }, { series: source(2) }],
      },
    };
    expect(() => mergeChartSpec(twoSeries, { chartType: 'PIE' }))
      .toThrow(/PIE chart plots a single series, but this chart has 2/);
  });

  it('replaces source ranges wholesale', () => {
    const merged = mergeChartSpec(
      existing, { domainRange: 'A1:A20', seriesRanges: ['B1:B20', 'C1:C20'] }, 9,
    );
    expect(merged.basicChart.domains[0].domain.sourceRange.sources[0]).toEqual({
      sheetId: 9, startRowIndex: 0, endRowIndex: 20, startColumnIndex: 0, endColumnIndex: 1,
    });
    expect(merged.basicChart.series).toHaveLength(2);
  });

  it('rejects changing one side of the source ranges', () => {
    expect(() => mergeChartSpec(existing, { domainRange: 'A1:A20' }, 9))
      .toThrow(/must be changed together/);
  });

  it('requires a sheet id when ranges change', () => {
    expect(() => mergeChartSpec(existing, { domainRange: 'A1:A20', seriesRanges: ['B1:B20'] }))
      .toThrow(/requires a resolved sheet id/);
  });

  it('refuses a STRUCTURAL change to a chart kind it cannot model', () => {
    // A title-only patch is now short-circuited and succeeds — see the
    // "renaming a chart kind these tools cannot rebuild" suite. Anything that
    // would require rebuilding the spec still has to refuse.
    expect(() => mergeChartSpec({ waterfallChart: {} }, { chartType: undefined, stackedType: 'STACKED' }))
      .toThrow(/type these tools do not model/);
  });

  it('refuses a chart whose source ranges cannot be read', () => {
    expect(() => mergeChartSpec({ basicChart: { chartType: 'COLUMN' } }, { title: 'x' }))
      .toThrow(/no readable source ranges/);
  });
});

describe('summarizeChart', () => {
  const names = new Map([[1, 'Data'], [2, 'Chart tab']]);
  const objects = new Set([2]);

  const chart = {
    chartId: 55,
    position: {
      overlayPosition: { anchorCell: { sheetId: 1, rowIndex: 1, columnIndex: 5 } },
    },
    spec: {
      title: 'Q1 Sales',
      basicChart: {
        chartType: 'COLUMN',
        domains: [{
          domain: {
            sourceRange: {
              sources: [{
                sheetId: 1, startRowIndex: 0, endRowIndex: 8,
                startColumnIndex: 0, endColumnIndex: 1,
              }],
            },
          },
        }],
        series: [{
          series: {
            sourceRange: {
              sources: [{
                sheetId: 1, startRowIndex: 0, endRowIndex: 8,
                startColumnIndex: 1, endColumnIndex: 2,
              }],
            },
          },
        }],
      },
    },
  };

  it('renders ranges back to A1 with their tab', () => {
    const summary = summarizeChart(chart, names, objects, false);
    expect(summary).toMatchObject({
      chartId: 55,
      title: 'Q1 Sales',
      chartType: 'COLUMN',
      onOwnSheet: false,
      anchor: { sheetName: 'Data', cell: 'F2' },
      domain: { sheetName: 'Data', range: 'A1:A8' },
      series: [{ sheetName: 'Data', range: 'B1:B8' }],
    });
  });

  it('omits the spec unless asked', () => {
    expect(summarizeChart(chart, names, objects, false).spec).toBeUndefined();
    expect(summarizeChart(chart, names, objects, true).spec).toBe(chart.spec);
  });

  it('flags a chart living on its own chart sheet', () => {
    const onOwn = { ...chart, position: { sheetId: 2 } };
    const summary = summarizeChart(onOwn, names, objects, false);
    expect(summary.onOwnSheet).toBe(true);
    expect(summary.anchor).toBeUndefined();
  });

  it('still lists a chart kind these tools cannot build', () => {
    const summary = summarizeChart(
      { chartId: 9, spec: { waterfallChart: {} } }, names, objects, false,
    );
    expect(summary.chartId).toBe(9);
    expect(summary.chartType).toBe('waterfallChart');
    expect(summary.series).toEqual([]);
  });

  it('drops a source range that has no A1 spelling', () => {
    const unbounded = {
      chartId: 3,
      spec: {
        basicChart: {
          chartType: 'LINE',
          domains: [{ domain: { sourceRange: { sources: [{ sheetId: 1, startRowIndex: 0 }] } } }],
          series: [],
        },
      },
    };
    expect(summarizeChart(unbounded, names, objects, false).domain).toBeUndefined();
  });
});

describe('mergeChartSpec legend across a union conversion', () => {
  const data = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };

  it('drops LABELED when converting away from pie, rather than failing', () => {
    const pie = {
      pieChart: { domain: data, series: data, legendPosition: 'LABELED_LEGEND' },
    };
    const merged = mergeChartSpec(pie, { chartType: 'COLUMN' });
    expect(merged.basicChart.legendPosition).toBeUndefined();
  });

  it('keeps LABELED when the chart stays a pie', () => {
    const pie = {
      pieChart: { domain: data, series: data, legendPosition: 'LABELED_LEGEND' },
    };
    expect(mergeChartSpec(pie, { title: 'x' }).pieChart.legendPosition).toBe('LABELED_LEGEND');
  });
});

describe('series target axis', () => {
  const base = { domainRange: 'A1:A9', seriesRanges: ['B1:B9', 'C1:C9'] };

  it('points a BAR chart at the bottom axis', () => {
    // Google rejects anything else: "Bar charts series may only target the
    // BOTTOM_AXIS." Caught against the live API, not in review.
    const spec = buildChartSpec({ chartType: 'BAR', ...base }, 1);
    expect(spec.basicChart.series.map((s: any) => s.targetAxis))
      .toEqual(['BOTTOM_AXIS', 'BOTTOM_AXIS']);
  });

  it('points every other basic type at a vertical axis', () => {
    for (const t of ['COLUMN', 'LINE', 'AREA', 'SCATTER', 'STEPPED_AREA'] as const) {
      const spec = buildChartSpec({ chartType: t, ...base }, 1);
      expect(spec.basicChart.series[0].targetAxis).toBe('LEFT_AXIS');
    }
  });

  it('valueAxisFor names the value axis per type', () => {
    expect(valueAxisFor('BAR')).toBe('BOTTOM_AXIS');
    expect(valueAxisFor('COLUMN')).toBe('LEFT_AXIS');
  });

  const columnChart = (targetAxis: string) => ({
    basicChart: {
      chartType: 'COLUMN',
      domains: [{ domain: { sourceRange: { sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }] } } }],
      series: [{
        targetAxis,
        series: { sourceRange: { sources: [{
          sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 1, endColumnIndex: 2,
        }] } },
      }],
    },
  });

  it('moves series to the bottom axis when converting to BAR', () => {
    const merged = mergeChartSpec(columnChart('LEFT_AXIS'), { chartType: 'BAR' });
    expect(merged.basicChart.series[0].targetAxis).toBe('BOTTOM_AXIS');
  });

  it('moves series off the bottom axis when converting away from BAR', () => {
    const bar = columnChart('BOTTOM_AXIS');
    bar.basicChart.chartType = 'BAR';
    const merged = mergeChartSpec(bar, { chartType: 'COLUMN' });
    expect(merged.basicChart.series[0].targetAxis).toBe('LEFT_AXIS');
  });

  it('preserves a deliberate RIGHT_AXIS on a non-bar chart', () => {
    const merged = mergeChartSpec(columnChart('RIGHT_AXIS'), { title: 'x' });
    expect(merged.basicChart.series[0].targetAxis).toBe('RIGHT_AXIS');
  });
});

describe('mergeChartSpec carries unmodeled spec fields', () => {
  const data = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };

  // A spec is sent whole, so anything left out is erased. These are real
  // fields Google returns that the curated schema has no argument for.
  const styled = {
    title: 'Original',
    titleTextFormat: { bold: true, foregroundColor: { red: 1 } },
    subtitleTextFormat: { italic: true },
    hiddenDimensionStrategy: 'SHOW_ALL',
    altText: 'described for screen readers',
    backgroundColorStyle: { themeColor: 'BACKGROUND' },
    basicChart: {
      chartType: 'COLUMN',
      domains: [{ domain: data }],
      series: [{ series: data, targetAxis: 'LEFT_AXIS' }],
    },
  };

  it('preserves styling the curated schema cannot express', () => {
    const merged = mergeChartSpec(styled, { title: 'Renamed' });
    expect(merged.title).toBe('Renamed');
    expect(merged.titleTextFormat).toEqual({ bold: true, foregroundColor: { red: 1 } });
    expect(merged.subtitleTextFormat).toEqual({ italic: true });
    expect(merged.hiddenDimensionStrategy).toBe('SHOW_ALL');
    expect(merged.altText).toBe('described for screen readers');
    expect(merged.backgroundColorStyle).toEqual({ themeColor: 'BACKGROUND' });
  });

  it('never carries the previous union member across a conversion', () => {
    const merged = mergeChartSpec(styled, { chartType: 'PIE' });
    expect(merged.basicChart).toBeUndefined();
    expect(merged.pieChart).toBeDefined();
    expect(merged.altText).toBe('described for screen readers');
  });
});


describe('mergeChartSpec against a hostile carry-over key', () => {
  const data = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };

  // Only JSON.parse produces an own "__proto__" key, and an upstream spec is
  // exactly that: parsed JSON. A plain object literal cannot express this.
  const hostile = JSON.parse(JSON.stringify({
    title: 'Original',
    basicChart: { chartType: 'COLUMN', domains: [{ domain: data }], series: [{ series: data }] },
  }).replace('{', '{"__proto__":{"pieChart":{"stolen":true}},'));

  it('has an own __proto__ key to begin with', () => {
    expect(Object.prototype.hasOwnProperty.call(hostile, '__proto__')).toBe(true);
  });

  it('does not let it shadow the union member the merge reads back', () => {
    const merged = mergeChartSpec(hostile, { title: 'Renamed' });
    expect(Object.getPrototypeOf(merged)).toBe(Object.prototype);
    expect(merged.pieChart).toBeUndefined();
    expect(merged.basicChart.chartType).toBe('COLUMN');
    expect(merged.title).toBe('Renamed');
  });

  it('leaves Object.prototype alone either way', () => {
    mergeChartSpec(hostile, { title: 'Renamed' });
    expect(({} as any).pieChart).toBeUndefined();
  });
});

describe('basic-to-basic retype preserves the whole basicChart', () => {
  // Regression: `sameUnion` compared the curated chart type, so COLUMN -> LINE
  // read as a union switch and silently discarded axis titles, stacking and
  // per-series styling. Found by review, proven against the real merge.
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };
  const column = {
    title: 'T',
    basicChart: {
      chartType: 'COLUMN',
      stackedType: 'STACKED',
      legendPosition: 'BOTTOM_LEGEND',
      headerCount: 1,
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Model' },
        { position: 'LEFT_AXIS', title: 'Sales' },
      ],
      domains: [{ domain: src }],
      series: [{ series: src, targetAxis: 'RIGHT_AXIS', colorStyle: { themeColor: 'ACCENT1' } }],
    },
  };

  it('keeps axis titles and series styling across COLUMN -> LINE', () => {
    const merged = mergeChartSpec(column, { chartType: 'LINE' });
    expect(merged.basicChart.axis).toEqual([
      { position: 'BOTTOM_AXIS', title: 'Model' },
      { position: 'LEFT_AXIS', title: 'Sales' },
    ]);
    // NOT stackedType: LINE cannot stack and Google rejects the field, so the
    // merge drops it. Covered by the stacking suite below.
    expect(merged.basicChart.stackedType).toBeUndefined();
    expect(merged.basicChart.series[0].colorStyle).toEqual({ themeColor: 'ACCENT1' });
    expect(merged.basicChart.series[0].targetAxis).toBe('RIGHT_AXIS');
    expect(merged.basicChart.legendPosition).toBe('BOTTOM_LEGEND');
  });

  it('still treats basic <-> pie as a real union switch', () => {
    const merged = mergeChartSpec(column, { chartType: 'PIE' });
    expect(merged.basicChart).toBeUndefined();
    expect(merged.pieChart.pieHole).toBeUndefined();
  });

  it('swaps axis titles when converting to BAR, whose axes are transposed', () => {
    const merged = mergeChartSpec(column, { chartType: 'BAR' });
    const byPosition = Object.fromEntries(
      merged.basicChart.axis.map((a: any) => [a.position, a.title]),
    );
    expect(byPosition.LEFT_AXIS).toBe('Model');
    expect(byPosition.BOTTOM_AXIS).toBe('Sales');
    expect(merged.basicChart.series[0].targetAxis).toBe('BOTTOM_AXIS');
  });

  it('swaps them back when converting away from BAR', () => {
    const bar = {
      basicChart: {
        ...column.basicChart,
        chartType: 'BAR',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Sales' },
          { position: 'LEFT_AXIS', title: 'Model' },
        ],
        series: [{ series: src, targetAxis: 'BOTTOM_AXIS' }],
      },
    };
    const merged = mergeChartSpec(bar, { chartType: 'COLUMN' });
    const byPosition = Object.fromEntries(
      merged.basicChart.axis.map((a: any) => [a.position, a.title]),
    );
    expect(byPosition.BOTTOM_AXIS).toBe('Model');
    expect(byPosition.LEFT_AXIS).toBe('Sales');
    expect(merged.basicChart.series[0].targetAxis).toBe('LEFT_AXIS');
  });

  it('keeps stacking when retyping to another stackable type', () => {
    expect(mergeChartSpec(column, { chartType: 'AREA' }).basicChart.stackedType).toBe('STACKED');
  });

  it('leaves axis positions alone when neither side is BAR', () => {
    const merged = mergeChartSpec(column, { chartType: 'AREA' });
    const byPosition = Object.fromEntries(
      merged.basicChart.axis.map((a: any) => [a.position, a.title]),
    );
    expect(byPosition.BOTTOM_AXIS).toBe('Model');
    expect(byPosition.LEFT_AXIS).toBe('Sales');
  });
});

describe('mergeChartSpec carries unmodeled basicChart fields', () => {
  // Second round of the same defect: the top-level allowlist was fixed, but
  // basicChart was still rebuilt key-by-key, so anything BasicChartSpec carries
  // that this connector does not model was erased on every edit — not just on
  // a conversion. Found by review, verified by execution.
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };
  const styled = {
    basicChart: {
      chartType: 'LINE',
      threeDimensional: true,
      interpolateNulls: true,
      lineSmoothing: true,
      compareMode: 'CATEGORY',
      totalDataLabel: { type: 'DATA' },
      domains: [{ domain: src, reversed: true }],
      series: [{ series: src, targetAxis: 'LEFT_AXIS' }],
    },
  };

  it('keeps BasicChartSpec fields the curated schema cannot express', () => {
    const merged = mergeChartSpec(styled, { title: 'Renamed' });
    expect(merged.basicChart.interpolateNulls).toBe(true);
    expect(merged.basicChart.lineSmoothing).toBe(true);
    expect(merged.basicChart.compareMode).toBe('CATEGORY');
    // NOT totalDataLabel or threeDimensional: this fixture is a LINE chart with
    // no stacking, and Google accepts neither there. They are pruned with their
    // prerequisites — see the type-dependent suite below.
    expect(merged.basicChart.totalDataLabel).toBeUndefined();
    expect(merged.basicChart.threeDimensional).toBeUndefined();
  });

  it('drops line-only options when retyping to a type that rejects them', () => {
    // Same "valid only if" phrasing Google uses for the fields already pruned,
    // so they are pruned by the same rule rather than left to chance.
    const merged = mergeChartSpec(styled, { chartType: 'COLUMN' });
    expect(merged.basicChart.lineSmoothing).toBeUndefined();
    expect(merged.basicChart.interpolateNulls).toBeUndefined();
  });

  it('keeps interpolateNulls on AREA, which also accepts it', () => {
    expect(mergeChartSpec(styled, { chartType: 'AREA' }).basicChart.interpolateNulls).toBe(true);
    expect(mergeChartSpec(styled, { chartType: 'AREA' }).basicChart.lineSmoothing).toBeUndefined();
  });

  it('keeps per-domain options such as reversed', () => {
    const merged = mergeChartSpec(styled, { title: 'x' });
    expect(merged.basicChart.domains[0].reversed).toBe(true);
    expect(merged.basicChart.domains[0].domain).toEqual(src);
  });

  it('treats COMBO as a basicChart rather than a union switch', () => {
    // chartTypeOf returns undefined for COMBO, which previously made the merge
    // think the union had changed and throw the whole basicChart away.
    const combo = {
      basicChart: {
        chartType: 'COMBO',
        stackedType: 'STACKED',
        axis: [{ position: 'BOTTOM_AXIS', title: 'Model' }],
        domains: [{ domain: src }],
        series: [{ series: src, targetAxis: 'LEFT_AXIS', colorStyle: { themeColor: 'ACCENT2' } }],
      },
    };
    const merged = mergeChartSpec(combo, { chartType: 'COLUMN' });
    expect(merged.basicChart.chartType).toBe('COLUMN');
    expect(merged.basicChart.stackedType).toBe('STACKED');
    expect(merged.basicChart.axis).toEqual([{ position: 'BOTTOM_AXIS', title: 'Model' }]);
    expect(merged.basicChart.series[0].colorStyle).toEqual({ themeColor: 'ACCENT2' });
  });

  it('does not carry a basicChart into a pie conversion', () => {
    const merged = mergeChartSpec(styled, { chartType: 'PIE' });
    expect(merged.basicChart).toBeUndefined();
    expect(merged.pieChart.domain).toEqual(src);
  });
});

describe('summarizeChart names a COMBO chart by its own type', () => {
  it('prefers basicChart.chartType over the union member name', () => {
    const summary = summarizeChart(
      { chartId: 4, spec: { basicChart: { chartType: 'COMBO' } } },
      new Map(), new Set(), false,
    );
    expect(summary.chartType).toBe('COMBO');
  });

  it('still falls back to the member name for a foreign kind', () => {
    const summary = summarizeChart(
      { chartId: 5, spec: { waterfallChart: {} } }, new Map(), new Set(), false,
    );
    expect(summary.chartType).toBe('waterfallChart');
  });
});

describe('BAR conversion drops the orphaned right axis', () => {
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };
  const dualAxis = {
    basicChart: {
      chartType: 'COLUMN',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Month' },
        { position: 'LEFT_AXIS', title: 'Revenue' },
        { position: 'RIGHT_AXIS', title: 'Margin %' },
      ],
      domains: [{ domain: src }],
      series: [
        { series: src, targetAxis: 'LEFT_AXIS' },
        { series: src, targetAxis: 'RIGHT_AXIS' },
      ],
    },
  };

  it('removes the right axis title, since no bar series can target it', () => {
    const merged = mergeChartSpec(dualAxis, { chartType: 'BAR' });
    const positions = merged.basicChart.axis.map((a: any) => a.position);
    expect(positions).not.toContain('RIGHT_AXIS');
    const byPosition = Object.fromEntries(
      merged.basicChart.axis.map((a: any) => [a.position, a.title]),
    );
    expect(byPosition.LEFT_AXIS).toBe('Month');
    expect(byPosition.BOTTOM_AXIS).toBe('Revenue');
    expect(merged.basicChart.series.map((s: any) => s.targetAxis))
      .toEqual(['BOTTOM_AXIS', 'BOTTOM_AXIS']);
  });

  it('keeps a right axis on a non-bar retype', () => {
    const merged = mergeChartSpec(dualAxis, { chartType: 'LINE' });
    const byPosition = Object.fromEntries(
      merged.basicChart.axis.map((a: any) => [a.position, a.title]),
    );
    expect(byPosition.RIGHT_AXIS).toBe('Margin %');
    expect(merged.basicChart.series[1].targetAxis).toBe('RIGHT_AXIS');
  });
});

describe('stacking is only sent to types that accept it', () => {
  // Google REJECTS stackedType on LINE/SCATTER rather than ignoring it —
  // "stackedType not supported when chartType is LINE" — confirmed against the
  // live API. Carrying it across a retype therefore broke COLUMN -> LINE.
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };
  const base = { domainRange: 'A1:A9', seriesRanges: ['B1:B9'] };

  it('knows which types stack', () => {
    for (const t of ['COLUMN', 'BAR', 'AREA', 'STEPPED_AREA'] as const) {
      expect(supportsStacking(t)).toBe(true);
    }
    for (const t of ['LINE', 'SCATTER', 'PIE'] as const) {
      expect(supportsStacking(t)).toBe(false);
    }
  });

  it('refuses an explicit stacked_type on a type that cannot stack', () => {
    expect(() => buildChartSpec({ chartType: 'LINE', ...base, stackedType: 'STACKED' }, 1))
      .toThrow(/not supported on a LINE chart/);
    expect(() => buildChartSpec({ chartType: 'SCATTER', ...base, stackedType: 'STACKED' }, 1))
      .toThrow(/not supported on a SCATTER chart/);
  });

  it('still accepts it on the stackable types', () => {
    expect(buildChartSpec({ chartType: 'AREA', ...base, stackedType: 'PERCENT_STACKED' }, 1)
      .basicChart.stackedType).toBe('PERCENT_STACKED');
  });

  it('drops an inherited stackedType when retyping to a non-stacking type', () => {
    const stackedColumn = {
      basicChart: {
        chartType: 'COLUMN',
        stackedType: 'STACKED',
        domains: [{ domain: src }],
        series: [{ series: src, targetAxis: 'LEFT_AXIS' }],
      },
    };
    // The caller never mentioned stacking; failing the edit over an inherited
    // field would be the merge breaking an operation it was asked to preserve.
    const merged = mergeChartSpec(stackedColumn, { chartType: 'LINE' });
    expect(merged.basicChart.stackedType).toBeUndefined();
    expect(merged.basicChart.chartType).toBe('LINE');
  });

  it('keeps an inherited stackedType when the target still stacks', () => {
    const stackedColumn = {
      basicChart: {
        chartType: 'COLUMN',
        stackedType: 'STACKED',
        domains: [{ domain: src }],
        series: [{ series: src, targetAxis: 'LEFT_AXIS' }],
      },
    };
    expect(mergeChartSpec(stackedColumn, { chartType: 'BAR' }).basicChart.stackedType)
      .toBe('STACKED');
  });
});

describe('type-dependent fields are pruned with their prerequisite', () => {
  // Round 3: seeding `basic` from the previous spec made several Google
  // conditional fields survive into types that reject them, and turned the
  // guarded `axis` write into a stale-value leak.
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };

  it('drops a stale axis array when the map empties', () => {
    // The RIGHT_AXIS drop emptied the map, but the spread had already put the
    // old array on `basic`, so the guarded write never replaced it.
    const rightOnly = {
      basicChart: {
        chartType: 'COLUMN',
        axis: [{ position: 'RIGHT_AXIS', title: 'Margin %' }],
        domains: [{ domain: src }],
        series: [{ series: src, targetAxis: 'RIGHT_AXIS' }],
      },
    };
    const merged = mergeChartSpec(rightOnly, { chartType: 'BAR' });
    expect(merged.basicChart.axis).toBeUndefined();
    expect(merged.basicChart.series[0].targetAxis).toBe('BOTTOM_AXIS');
  });

  it('normalises out an axis entry that has no position', () => {
    const orphan = {
      basicChart: {
        chartType: 'COLUMN',
        axis: [{ title: 'Orphan' }],
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    };
    expect(mergeChartSpec(orphan, { title: 'x' }).basicChart.axis).toBeUndefined();
  });

  it('refuses an explicit stacked_type on update, matching add_chart', () => {
    const line = {
      basicChart: { chartType: 'LINE', domains: [{ domain: src }], series: [{ series: src }] },
    };
    expect(() => mergeChartSpec(line, { stackedType: 'STACKED' }))
      .toThrow(/not supported on a LINE chart/);
  });

  it('drops totalDataLabel when its stackedType goes', () => {
    const stacked = {
      basicChart: {
        chartType: 'COLUMN',
        stackedType: 'STACKED',
        totalDataLabel: { type: 'DATA' },
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    };
    const merged = mergeChartSpec(stacked, { chartType: 'LINE' });
    expect(merged.basicChart.stackedType).toBeUndefined();
    expect(merged.basicChart.totalDataLabel).toBeUndefined();
  });

  it('keeps totalDataLabel when stacking survives', () => {
    const stacked = {
      basicChart: {
        chartType: 'COLUMN',
        stackedType: 'STACKED',
        totalDataLabel: { type: 'DATA' },
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    };
    expect(mergeChartSpec(stacked, { chartType: 'AREA' }).basicChart.totalDataLabel)
      .toEqual({ type: 'DATA' });
  });

  it('strips per-series COMBO type when the chart is no longer a combo', () => {
    const combo = {
      basicChart: {
        chartType: 'COMBO',
        domains: [{ domain: src }],
        series: [
          { series: src, targetAxis: 'LEFT_AXIS', type: 'LINE', colorStyle: { themeColor: 'ACCENT1' } },
          { series: src, targetAxis: 'LEFT_AXIS', type: 'COLUMN' },
        ],
      },
    };
    const merged = mergeChartSpec(combo, { chartType: 'COLUMN' });
    expect(merged.basicChart.series.map((s: any) => s.type)).toEqual([undefined, undefined]);
    // styling still survives — only the COMBO-only field is stripped
    expect(merged.basicChart.series[0].colorStyle).toEqual({ themeColor: 'ACCENT1' });
  });

  it('drops threeDimensional when retyping away from bar/column', () => {
    const threeD = {
      basicChart: {
        chartType: 'COLUMN',
        threeDimensional: true,
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    };
    expect(mergeChartSpec(threeD, { chartType: 'LINE' }).basicChart.threeDimensional)
      .toBeUndefined();
    expect(mergeChartSpec(threeD, { chartType: 'BAR' }).basicChart.threeDimensional)
      .toBe(true);
  });

  it('seeds the pie branch instead of allowlisting it', () => {
    const pie = {
      pieChart: {
        domain: src, series: src, pieHole: 0.3,
        someFutureField: 'kept',
      },
    };
    expect(mergeChartSpec(pie, { title: 'x' }).pieChart.someFutureField).toBe('kept');
  });
});

describe('three_dimensional is honoured, not swallowed', () => {
  // It was wired to the pie branch only, so asking for a 3D column chart was
  // accepted and silently discarded while stacked_type on a bad type threw.
  const base = { domainRange: 'A1:A9', seriesRanges: ['B1:B9'] };
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };

  it('knows which types render in 3D', () => {
    for (const t of ['PIE', 'BAR'] as const) {
      expect(supportsThreeDimensional(t)).toBe(true);
    }
    // COLUMN belongs with the refusals, not the allowed set: the live API
    // answers "threeDimensional not supported when chartType is COLUMN".
    for (const t of ['COLUMN', 'LINE', 'AREA', 'SCATTER', 'STEPPED_AREA'] as const) {
      expect(supportsThreeDimensional(t)).toBe(false);
    }
  });

  it('applies it to a BAR chart instead of dropping it', () => {
    expect(buildChartSpec({ chartType: 'BAR', ...base, threeDimensional: true }, 1)
      .basicChart.threeDimensional).toBe(true);
  });

  it('refuses it on COLUMN, which Google rejects despite bar accepting it', () => {
    expect(() => buildChartSpec({ chartType: 'COLUMN', ...base, threeDimensional: true }, 1))
      .toThrow(/not supported on a COLUMN chart/);
  });

  it('refuses it on a type Google cannot render in 3D', () => {
    expect(() => buildChartSpec({ chartType: 'LINE', ...base, threeDimensional: true }, 1))
      .toThrow(/not supported on a LINE chart/);
  });

  it('lets update_chart turn 3D off on a bar chart', () => {
    const threeD = {
      basicChart: {
        chartType: 'BAR', threeDimensional: true,
        domains: [{ domain: src }], series: [{ series: src }],
      },
    };
    expect(mergeChartSpec(threeD, { threeDimensional: false }).basicChart.threeDimensional)
      .toBe(false);
  });

  it('drops an inherited 3D when retyping a bar chart to COLUMN', () => {
    const bar3d = {
      basicChart: {
        chartType: 'BAR', threeDimensional: true,
        domains: [{ domain: src }], series: [{ series: src }],
      },
    };
    expect(mergeChartSpec(bar3d, { chartType: 'COLUMN' }).basicChart.threeDimensional)
      .toBeUndefined();
  });

  it('carries 3D across a pie/basic conversion, like legend position', () => {
    const pie3d = { pieChart: { domain: src, series: src, threeDimensional: true } };
    expect(mergeChartSpec(pie3d, { chartType: 'BAR' }).basicChart.threeDimensional)
      .toBe(true);
    const bar3d = {
      basicChart: {
        chartType: 'BAR', threeDimensional: true,
        domains: [{ domain: src }], series: [{ series: src }],
      },
    };
    expect(mergeChartSpec(bar3d, { chartType: 'PIE' }).pieChart.threeDimensional).toBe(true);
  });
});

describe('a right-axis title is refused on a BAR chart', () => {
  // Found adversarially: the carried-over path dropped a RIGHT_AXIS entry on a
  // bar chart, but a patch-supplied one bypassed that guard — so the title
  // appeared, then vanished on the NEXT edit of any kind. Idempotent, but not a
  // fixed point, which is why single-conversion tests missed it.
  const base = { domainRange: 'A1:A9', seriesRanges: ['B1:B9'] };
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };
  const column = {
    basicChart: {
      chartType: 'COLUMN',
      axis: [{ position: 'BOTTOM_AXIS', title: 'Month' }, { position: 'LEFT_AXIS', title: 'USD' }],
      domains: [{ domain: src }],
      series: [{ series: src, targetAxis: 'LEFT_AXIS' }],
    },
  };

  it('refuses it at creation', () => {
    expect(() => buildChartSpec(
      { chartType: 'BAR', ...base, axisTitles: { right: 'Secondary' } }, 1,
    )).toThrow(/axis_titles\.right is not supported on a BAR chart/);
  });

  it('refuses it on a conversion to BAR', () => {
    expect(() => mergeChartSpec(column, {
      chartType: 'BAR', axisTitles: { right: 'Secondary' },
    })).toThrow(/axis_titles\.right is not supported on a BAR chart/);
  });

  it('still allows a right title on types that have a right axis', () => {
    const merged = mergeChartSpec(column, { chartType: 'LINE', axisTitles: { right: 'Margin' } });
    const byPosition = Object.fromEntries(
      merged.basicChart.axis.map((a: any) => [a.position, a.title]),
    );
    expect(byPosition.RIGHT_AXIS).toBe('Margin');
  });

  it('reaches a fixed point — a following no-op edit changes nothing', () => {
    const once = mergeChartSpec(column, { chartType: 'BAR' });
    const twice = mergeChartSpec(once, { title: 'renamed' });
    expect(twice.basicChart.axis).toEqual(once.basicChart.axis);
  });
});

describe('unrenderable source ranges are reported, not hidden', () => {
  // A whole-column source (what the Sheets UI produces) has no bounded A1 form,
  // so it was dropped from `series` with no signal — making a 3-series chart
  // indistinguishable from a genuine 2-series one.
  const bounded = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 1, endColumnIndex: 2,
      }],
    },
  };
  const unbounded = {
    sourceRange: { sources: [{ sheetId: 1, startRowIndex: 0, startColumnIndex: 0 }] },
  };

  it('counts a series whose range cannot be spelled in A1', () => {
    const chart = {
      chartId: 1,
      spec: {
        basicChart: {
          chartType: 'COLUMN',
          domains: [{ domain: bounded }],
          series: [{ series: bounded }, { series: unbounded }],
        },
      },
    };
    const summary = summarizeChart(chart, new Map([[1, 'Data']]), new Set(), false);
    expect(summary.series).toHaveLength(1);
    expect(summary.unreportableRanges).toBe(1);
  });

  it('counts an unrenderable domain too', () => {
    const chart = {
      chartId: 2,
      spec: {
        basicChart: {
          chartType: 'LINE',
          domains: [{ domain: unbounded }],
          series: [{ series: unbounded }],
        },
      },
    };
    const summary = summarizeChart(chart, new Map([[1, 'Data']]), new Set(), false);
    expect(summary.domain).toBeUndefined();
    expect(summary.unreportableRanges).toBe(2);
  });

  it('stays absent when every range is reportable', () => {
    const chart = {
      chartId: 3,
      spec: {
        basicChart: {
          chartType: 'COLUMN',
          domains: [{ domain: bounded }],
          series: [{ series: bounded }],
        },
      },
    };
    expect(summarizeChart(chart, new Map([[1, 'Data']]), new Set(), false).unreportableRanges)
      .toBeUndefined();
  });
});

describe('conditional-field pruning is complete and NOT_STACKED-aware', () => {
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };

  it('drops totalDataLabel when stacking is turned OFF explicitly', () => {
    // 'NOT_STACKED' is truthy, so the old `if (!basic.stackedType)` guard never
    // fired on the ordinary "turn stacking off" edit and Google rejected the write.
    const stacked = {
      basicChart: {
        chartType: 'COLUMN',
        stackedType: 'STACKED',
        totalDataLabel: { type: 'DATA' },
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    };
    const merged = mergeChartSpec(stacked, { stackedType: 'NOT_STACKED' });
    expect(merged.basicChart.stackedType).toBe('NOT_STACKED');
    expect(merged.basicChart.totalDataLabel).toBeUndefined();
  });

  it('keeps totalDataLabel for PERCENT_STACKED as well as STACKED', () => {
    const pct = {
      basicChart: {
        chartType: 'COLUMN',
        stackedType: 'PERCENT_STACKED',
        totalDataLabel: { type: 'DATA' },
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    };
    expect(mergeChartSpec(pct, { title: 'x' }).basicChart.totalDataLabel)
      .toEqual({ type: 'DATA' });
  });

  it('drops per-series line and point styling on a type that rejects them', () => {
    const line = {
      basicChart: {
        chartType: 'LINE',
        domains: [{ domain: src }],
        series: [{
          series: src,
          lineStyle: { width: 3 },
          pointStyle: { shape: 'CIRCLE' },
          colorStyle: { themeColor: 'ACCENT1' },
        }],
      },
    };
    const merged = mergeChartSpec(line, { chartType: 'COLUMN' });
    expect(merged.basicChart.series[0].lineStyle).toBeUndefined();
    expect(merged.basicChart.series[0].pointStyle).toBeUndefined();
    // colorStyle is unconditional — it must survive
    expect(merged.basicChart.series[0].colorStyle).toEqual({ themeColor: 'ACCENT1' });
  });

  it('keeps line and point styling on SCATTER, which accepts them', () => {
    const line = {
      basicChart: {
        chartType: 'LINE',
        domains: [{ domain: src }],
        series: [{ series: src, lineStyle: { width: 3 }, pointStyle: { shape: 'CIRCLE' } }],
      },
    };
    const merged = mergeChartSpec(line, { chartType: 'SCATTER' });
    expect(merged.basicChart.series[0].lineStyle).toEqual({ width: 3 });
    expect(merged.basicChart.series[0].pointStyle).toEqual({ shape: 'CIRCLE' });
  });
});

describe('conditionals one level below the series', () => {
  const src = (col: number) => ({
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9,
        startColumnIndex: col, endColumnIndex: col + 1,
      }],
    },
  });

  it('strips per-point pointStyle overrides on a type that rejects them', () => {
    // Same "Valid only if AREA, LINE, or SCATTER" wording as series.pointStyle,
    // one level down — it was being carried across by the spread.
    const line = {
      basicChart: {
        chartType: 'LINE',
        domains: [{ domain: src(0) }],
        series: [{
          series: src(1),
          styleOverrides: [
            { index: 0, pointStyle: { shape: 'STAR' }, colorStyle: { themeColor: 'ACCENT1' } },
          ],
        }],
      },
    };
    const merged = mergeChartSpec(line, { chartType: 'COLUMN' });
    const override = merged.basicChart.series[0].styleOverrides[0];
    expect(override.pointStyle).toBeUndefined();
    // the unconditional part of the override survives
    expect(override.colorStyle).toEqual({ themeColor: 'ACCENT1' });
  });

  it('keeps per-point overrides when the target still accepts them', () => {
    const line = {
      basicChart: {
        chartType: 'LINE',
        domains: [{ domain: src(0) }],
        series: [{ series: src(1), styleOverrides: [{ index: 0, pointStyle: { shape: 'STAR' } }] }],
      },
    };
    expect(mergeChartSpec(line, { chartType: 'SCATTER' }).basicChart.series[0].styleOverrides[0].pointStyle)
      .toEqual({ shape: 'STAR' });
  });

  it('drops a custom data label when the series ranges are replaced', () => {
    const withCustom = {
      basicChart: {
        chartType: 'COLUMN',
        domains: [{ domain: src(0) }],
        series: [{
          series: src(1),
          dataLabel: { type: 'CUSTOM', customLabelData: { sourceRange: { sources: [src(2)] } } },
        }],
      },
    };
    const merged = mergeChartSpec(
      withCustom, { domainRange: 'A1:A20', seriesRanges: ['B1:B20'] }, 1,
    );
    expect(merged.basicChart.series[0].dataLabel).toBeUndefined();
  });

  it('keeps a plain data label across a range change', () => {
    const plain = {
      basicChart: {
        chartType: 'COLUMN',
        domains: [{ domain: src(0) }],
        series: [{ series: src(1), dataLabel: { type: 'DATA' } }],
      },
    };
    const merged = mergeChartSpec(plain, { domainRange: 'A1:A20', seriesRanges: ['B1:B20'] }, 1);
    expect(merged.basicChart.series[0].dataLabel).toEqual({ type: 'DATA' });
  });
});

describe('mergeChartSpec does not mutate its input', () => {
  it('leaves the caller spec untouched when pruning nested overrides', () => {
    // `entry` is a shallow spread, so its styleOverrides array and objects were
    // still the caller's — a deleting prune wrote back through into `existing`.
    const src = {
      sourceRange: {
        sources: [{
          sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
        }],
      },
    };
    const existing = {
      basicChart: {
        chartType: 'LINE',
        domains: [{ domain: src }],
        series: [{ series: src, styleOverrides: [{ index: 0, pointStyle: { shape: 'STAR' } }] }],
      },
    };
    mergeChartSpec(existing, { chartType: 'COLUMN' });
    expect(existing.basicChart.series[0].styleOverrides[0].pointStyle)
      .toEqual({ shape: 'STAR' });
    // and a second merge to a type that accepts it still sees the value
    expect(mergeChartSpec(existing, { chartType: 'SCATTER' })
      .basicChart.series[0].styleOverrides[0].pointStyle).toEqual({ shape: 'STAR' });
  });

  it('pairs series with their styling before filtering, not after', () => {
    const at = (col: number) => ({
      sourceRange: {
        sources: [{
          sheetId: 1, startRowIndex: 0, endRowIndex: 9,
          startColumnIndex: col, endColumnIndex: col + 1,
        }],
      },
    });
    const skewed = {
      basicChart: {
        chartType: 'COLUMN',
        domains: [{ domain: at(0) }],
        // A leading entry with no `series` used to shift every later entry's
        // styling onto the wrong series.
        series: [{}, { series: at(1), colorStyle: { themeColor: 'ACCENT1' } }],
      },
    };
    const merged = mergeChartSpec(skewed, { title: 't' });
    expect(merged.basicChart.series).toHaveLength(1);
    expect(merged.basicChart.series[0].colorStyle).toEqual({ themeColor: 'ACCENT1' });
  });
});

describe('renaming a chart kind these tools cannot rebuild', () => {
  const src = {
    sourceRange: {
      sources: [{
        sheetId: 1, startRowIndex: 0, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1,
      }],
    },
  };

  it('retitles a COMBO chart instead of refusing', () => {
    const combo = {
      title: 'Old',
      basicChart: {
        chartType: 'COMBO',
        domains: [{ domain: src }],
        series: [{ series: src, type: 'LINE' }, { series: src, type: 'COLUMN' }],
      },
    };
    const merged = mergeChartSpec(combo, { title: 'New' });
    expect(merged.title).toBe('New');
    // the combo is passed through whole — per-series types survive
    expect(merged.basicChart.series.map((s: any) => s.type)).toEqual(['LINE', 'COLUMN']);
  });

  it('retitles a waterfall chart', () => {
    const merged = mergeChartSpec({ title: 'A', waterfallChart: { x: 1 } }, { subtitle: 'B' });
    expect(merged.title).toBe('A');
    expect(merged.subtitle).toBe('B');
    expect(merged.waterfallChart).toEqual({ x: 1 });
  });

  it('still refuses a structural change it cannot model', () => {
    expect(() => mergeChartSpec({ waterfallChart: {} }, { legendPosition: 'TOP' }))
      .toThrow(/type these tools do not model/);
  });

  // update_chart names the type it wrote from the merge result. The retitle
  // path returns a spec with no basicChart, so reading one directly threw
  // *after* the write had already succeeded and reported a failed rename.
  it.each([
    'waterfallChart', 'treemapChart', 'scorecardChart', 'orgChart',
    'histogramChart', 'bubbleChart', 'candlestickChart',
  ])('names the retitled %s so the caller is not told the write failed', (member) => {
    const merged = mergeChartSpec({ [member]: {} }, { title: 'New' });
    expect(chartTypeOf(merged)).toBeUndefined();
    expect(chartTypeLabel(merged)).toBe(member);
  });

  it('names a retitled COMBO by its own chartType, not its union member', () => {
    const merged = mergeChartSpec(
      { basicChart: { chartType: 'COMBO', series: [] } },
      { title: 'New' },
    );
    expect(chartTypeLabel(merged)).toBe('COMBO');
  });

  it('drops a LABELED legend the spread carried onto a basic chart', () => {
    const merged = mergeChartSpec({
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'LABELED_LEGEND',
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    }, { chartType: 'LINE' });
    expect(merged.basicChart.legendPosition).toBeUndefined();
  });

  it('keeps a legend position it cannot map, which is unmodeled rather than invalid', () => {
    const merged = mergeChartSpec({
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'SOME_FUTURE_LEGEND',
        domains: [{ domain: src }],
        series: [{ series: src }],
      },
    }, { chartType: 'LINE' });
    expect(merged.basicChart.legendPosition).toBe('SOME_FUTURE_LEGEND');
  });

  it('does not short-circuit a chart it can model', () => {
    const column = {
      title: 'Old',
      basicChart: { chartType: 'COLUMN', domains: [{ domain: src }], series: [{ series: src }] },
    };
    const merged = mergeChartSpec(column, { title: 'New' });
    expect(merged.title).toBe('New');
    expect(merged.basicChart.headerCount).toBe(1);
  });
});
