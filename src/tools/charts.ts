/**
 * Chart tools. Every chart operation is a `batchUpdate` request, so all of
 * these are native-only — an uploaded .xlsx has no chart surface here.
 *
 * The model cannot see a chart it creates: the Sheets API has no image
 * export, and Drive only renders Sheets to PDF. So these tools verify what
 * they can before and after the write instead — source ranges are checked for
 * alignment and read for their labels up front, and what was created is
 * echoed back in A1 terms.
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from '../auth.js';
import { wrapHandler, toolResponse } from '../lib/errors.js';
import {
  quoteSheetName,
  assertBareA1Range,
  parseA1Range,
  columnIndexToLetter,
} from '../lib/a1.js';
import { makeSheetsRequest, getSheetId } from '../lib/google.js';
import { applyBatchUpdate } from '../lib/batch.js';
import { nativeOnly } from '../lib/office.js';
import { truncationFields } from '../lib/sheetBudget.js';
import {
  CURATED_CHART_TYPES,
  LEGEND_POSITIONS,
  STACKED_TYPES,
  MAX_LISTED_CHARTS,
  buildChartSpec,
  buildPosition,
  describePlacement,
  overlayFieldMask,
  mergeChartSpec,
  summarizeChart,
  chartTypeOf,
  chartTypeLabel,
  type ChartPlacement,
  type ChartSummary,
} from '../lib/chart.js';

/**
 * Named so the mask keeps their keys, which is what lets summarizeChart report
 * a waterfall or treemap by its member name instead of omitting the type.
 */
const UNMODELED_CHART_FIELDS = [
  'bubbleChart', 'candlestickChart', 'orgChart', 'histogramChart',
  'waterfallChart', 'treemapChart', 'scorecardChart',
].join(',');

/**
 * Drops the parts of a spec that scale with the data, chiefly per-point
 * `styleOverrides`, so listing charts stays bounded.
 */
const SUMMARY_FIELDS =
  'sheets(properties(sheetId,title,sheetType),charts(chartId,position,'
  + 'spec(title,basicChart(chartType,domains,series(series)),'
  + `pieChart(domain,series),${UNMODELED_CHART_FIELDS})))`;

const FULL_FIELDS = 'sheets(properties(sheetId,title,sheetType),charts(chartId,position,spec))';

const SHEET_ONLY_FIELDS = 'sheets(properties(sheetId,title,sheetType))';

/** Enough to find a chart and tell whether it sits on its own sheet. */
const LOCATE_FIELDS = 'sheets(properties(sheetId,title,sheetType),charts(chartId,position))';

interface ChartContext {
  sheetNameById: Map<number, string>;
  objectSheetIds: Set<number>;
  charts: Array<{ chart: any; sheetId?: number }>;
}

async function fetchChartContext(
  spreadsheetId: string,
  accessToken: string,
  fields: string,
): Promise<ChartContext> {
  const metadata = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}?fields=${encodeURIComponent(fields)}`,
    accessToken,
    { method: 'GET' },
  ) as {
    sheets?: Array<{
      properties: { sheetId: number; title: string; sheetType?: string };
      charts?: any[];
    }>;
  };

  const sheetNameById = new Map<number, string>();
  const objectSheetIds = new Set<number>();
  const charts: ChartContext['charts'] = [];

  for (const sheet of metadata.sheets ?? []) {
    const { sheetId, title, sheetType } = sheet.properties;
    sheetNameById.set(sheetId, title);
    if (sheetType === 'OBJECT') objectSheetIds.add(sheetId);
    for (const chart of sheet.charts ?? []) {
      charts.push({ chart, sheetId });
    }
  }

  return { sheetNameById, objectSheetIds, charts };
}

/**
 * The hint differs per tool: delete_chart's mentions embedded images, because
 * refusing an id that is not a chart's is the whole point of looking it up.
 */
function findChartOrThrow(
  ctx: ChartContext,
  chartId: number,
  hint: string,
): ChartContext['charts'][number] {
  const found = ctx.charts.find(({ chart }) => chart?.chartId === chartId);
  if (!found) {
    throw new Error(`No chart with id ${chartId} in this spreadsheet. ${hint}`);
  }
  return found;
}

function topLeftCell(range: string): string {
  const grid = parseA1Range(range);
  return `${columnIndexToLetter(grid.startColumnIndex)}${grid.startRowIndex + 1}`;
}

/**
 * Catches the mistakes that actually happen — a range pointing at the wrong
 * column, or at a tab that holds nothing — while the caller can still act on
 * it. With the default headerCount of 1 these cells are also the series names,
 * so reporting them says what the chart will be labelled.
 */
async function readRangeLabels(
  spreadsheetId: string,
  sheetName: string,
  ranges: string[],
  accessToken: string,
): Promise<Array<string | undefined>> {
  const params = new URLSearchParams({ fields: 'valueRanges(values)' });
  for (const range of ranges) {
    params.append('ranges', `${quoteSheetName(sheetName)}!${topLeftCell(range)}`);
  }

  const result = await makeSheetsRequest(
    `/${encodeURIComponent(spreadsheetId)}/values:batchGet?${params}`,
    accessToken,
    { method: 'GET' },
  ) as { valueRanges?: Array<{ values?: string[][] }> };

  return ranges.map((_, i) => result.valueRanges?.[i]?.values?.[0]?.[0]);
}

const chartTypeSchema = z.enum(CURATED_CHART_TYPES);
const legendSchema = z.enum(LEGEND_POSITIONS);
const stackedSchema = z.enum(STACKED_TYPES);

const spreadsheetIdSchema = z.string().describe('Google Sheets spreadsheet ID');
const anchorCellSchema = z.string().optional().describe('Single A1 cell the chart\'s top-left corner overlays, e.g. "F2". Mutually exclusive with new_sheet.');
const threeDimensionalSchema = z.boolean().optional().describe('Render in 3D. PIE and BAR only — other types, COLUMN included, are refused.');

/**
 * Shared only up to the type: each tool describes its own `chart_id`, and a
 * tool's field descriptions are what steer a model towards using it correctly.
 */
const chartIdSchema = z.coerce.number().int();

const axisTitlesSchema = z.object({
  bottom: z.string().optional().describe('Title for the horizontal axis'),
  left: z.string().optional().describe('Title for the left vertical axis'),
  right: z.string().optional().describe('Title for the right vertical axis'),
});

const rangeRefSchema = z.object({
  sheetName: z.string().optional(),
  range: z.string(),
});

const chartSummarySchema = z.object({
  chartId: z.number(),
  title: z.string().optional(),
  chartType: z.string().optional().describe('Named for every chart, including kinds these tools cannot build (COMBO, waterfallChart and the rest). Omitted only when the chart carries no specification.'),
  onOwnSheet: z.boolean().describe('True when the chart occupies a dedicated chart sheet'),
  anchor: z.object({
    sheetName: z.string().optional(),
    cell: z.string(),
  }).optional().describe('Top-left cell the chart overlays; absent for own-sheet charts'),
  domain: rangeRefSchema.optional(),
  series: z.array(rangeRefSchema),
  unreportableRanges: z.number().optional()
    .describe('Source ranges omitted because they have no bounded A1 form — a whole-column range, or a data-source column reference. When present, `series`/`domain` are incomplete.'),
  spec: z.any().optional().describe('Raw Google ChartSpec, when include_spec is set'),
});

export const chartTools = {
      add_chart: {
        description:
          'Add a chart to a spreadsheet. Supports COLUMN, BAR, LINE, AREA, SCATTER, STEPPED_AREA and PIE. '
          + 'Source data is given as bounded bare A1 ranges on one tab: `domain_range` holds the category '
          + 'labels (the x axis) and each entry of `series_ranges` holds one plotted series. Every range '
          + 'must be a SINGLE column ("B1:B20") or a SINGLE row ("B1:T1"), all sharing one orientation and '
          + 'the same length as the domain — a rectangle like "B2:D9" is rejected, pass one range per series '
          + 'instead. Do NOT include a sheet prefix; use `sheet_name`. By default `header_count` is 1, so the '
          + 'first cell of each range is used as its series name rather than as data; pass 0 if your ranges '
          + 'hold no header. Placement is required: pass `anchor_cell` to overlay the chart on a tab, or '
          + '`new_sheet: true` to give it a dedicated chart sheet. PIE plots exactly one series and ignores '
          + '`axis_titles`, `stacked_type` and `header_count` — a pie has no axes and no header row, so its '
          + 'first cell is plotted as a slice. Returns the new `chart_id`, needed by update_chart, move_chart and '
          + 'delete_chart. NOTE: charts cannot be rendered back — nothing can show you the finished chart — '
          + 'so check the returned `resolved` block, which reports the ranges as stored and the label found '
          + 'in the first cell of each.',
        outputSchema: {
          id: z.string(),
          chartId: z.number(),
          resolved: z.object({
            chartType: z.string(),
            sheetName: z.string(),
            domain: z.object({ range: z.string(), label: z.string().optional() }),
            series: z.array(z.object({ range: z.string(), label: z.string().optional() })),
            placement: z.string(),
          }).describe('What was actually created, for verification'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          sheet_name: z.string().describe('Name of the tab holding the source data'),
          chart_type: chartTypeSchema.describe('Chart type. PIE plots a single series.'),
          domain_range: z.string().describe('Bounded bare A1 range holding the category labels, e.g. "A1:A20". Single column or single row.'),
          series_ranges: z.array(z.string()).min(1).describe('One bounded bare A1 range per plotted series, e.g. ["B1:B20", "C1:C20"]. Each must match domain_range in orientation and length.'),
          title: z.string().optional().describe('Chart title'),
          subtitle: z.string().optional().describe('Chart subtitle'),
          axis_titles: axisTitlesSchema.optional().describe('Axis titles. Ignored for PIE. A bar chart has no right-hand axis, so `right` is refused on BAR.'),
          legend_position: legendSchema.optional().describe('Legend placement. LABELED is PIE-only.'),
          stacked_type: stackedSchema.optional().describe('Stacking mode. COLUMN, BAR, AREA and STEPPED_AREA only — refused on LINE and SCATTER, which Google rejects; ignored on PIE, which has no stacking.'),
          header_count: z.coerce.number().int().min(0).optional().describe('Leading cells of each range treated as labels rather than data. Defaults to 1.'),
          pie_hole: z.coerce.number().min(0).max(1).optional().describe('PIE only: hole size 0..1, making a donut'),
          three_dimensional: threeDimensionalSchema,
          anchor_cell: anchorCellSchema,
          anchor_sheet_name: z.string().optional().describe('Tab to place the chart on when using anchor_cell. Defaults to sheet_name.'),
          new_sheet: z.boolean().optional().describe('Put the chart on its own dedicated chart sheet instead of overlaying a tab'),
          width_pixels: z.coerce.number().int().positive().optional().describe('Chart width when using anchor_cell'),
          height_pixels: z.coerce.number().int().positive().optional().describe('Chart height when using anchor_cell'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async (args: any, context: any) => {
          const { accessToken } = context;
          const {
            spreadsheet_id, sheet_name, chart_type, domain_range, series_ranges,
            title, subtitle, axis_titles, legend_position, stacked_type, header_count,
            pie_hole, three_dimensional,
            anchor_cell, anchor_sheet_name, new_sheet, width_pixels, height_pixels,
          } = args;

          const domainRange = assertBareA1Range(domain_range, 'domain_range');
          const seriesRanges = series_ranges.map(
            (r: unknown, i: number) => assertBareA1Range(r, `series_ranges[${i}]`),
          );

          const sourceSheetId = await getSheetId(spreadsheet_id, sheet_name, accessToken);
          const anchorSheetName = anchor_sheet_name ?? sheet_name;
          const anchorSheetId = anchor_cell
            ? (anchorSheetName === sheet_name
                ? sourceSheetId
                : await getSheetId(spreadsheet_id, anchorSheetName, accessToken))
            : undefined;

          const placement: ChartPlacement = {
            newSheet: new_sheet,
            anchorCell: anchor_cell,
            anchorSheetId,
            widthPixels: width_pixels,
            heightPixels: height_pixels,
          };
          // Built before the labels are read so a bad placement fails without
          // spending a call on data that is about to be discarded.
          const position = buildPosition(placement);

          const spec = buildChartSpec({
            chartType: chart_type,
            domainRange,
            seriesRanges,
            title,
            subtitle,
            axisTitles: axis_titles,
            legendPosition: legend_position,
            stackedType: stacked_type,
            headerCount: header_count,
            pieHole: pie_hole,
            threeDimensional: three_dimensional,
          }, sourceSheetId);

          const labels = await readRangeLabels(
            spreadsheet_id, sheet_name, [domainRange, ...seriesRanges], accessToken,
          );
          if (labels.every((label) => label === undefined)) {
            throw new Error(
              `No source range on tab "${sheet_name}" has a value where the chart starts reading — domain_range `
              + `"${domainRange}" and series_ranges ${JSON.stringify(seriesRanges)} start at cells `
              + 'with no value. Check the tab and ranges with get_metadata or get_sheet_data '
              + 'before charting them.',
            );
          }

          const replies = await applyBatchUpdate<{ addChart?: { chart?: { chartId?: number } } }>(
            spreadsheet_id,
            [{ addChart: { chart: { spec, position } } }],
            accessToken,
            { fields: 'replies.addChart.chart.chartId' },
          );

          const chartId = replies[0]?.addChart?.chart?.chartId;
          if (chartId === undefined) {
            throw new Error(
              'Google accepted the chart but returned no chart id. The chart WAS created — '
              + 'call list_charts to find its id rather than retrying, which would add a second chart.',
            );
          }

          return toolResponse({
            id: spreadsheet_id,
            chartId,
            resolved: {
              chartType: chart_type,
              sheetName: sheet_name,
              domain: { range: domainRange, ...(labels[0] !== undefined && { label: labels[0] }) },
              series: seriesRanges.map((range: string, i: number) => ({
                range,
                ...(labels[i + 1] !== undefined && { label: labels[i + 1] }),
              })),
              placement: describePlacement(new_sheet, anchorSheetName, anchor_cell),
            },
            message: `Chart ${chartId} created. It cannot be rendered back — verify from the resolved ranges above.`,
          });
        }))),
      },

      list_charts: {
        description:
          'List the charts in a spreadsheet with their `chart_id`, type, title, source ranges and placement. '
          + 'Call this to find the `chart_id` that update_chart, move_chart and delete_chart need — a chart id '
          + 'is otherwise only returned at creation time. Source ranges are reported in A1 with the tab they '
          + 'live on, which may differ from the tab the chart sits on. By default the raw Google ChartSpec is '
          + 'omitted, since a styled chart\'s spec is large; pass include_spec to get it.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          charts: z.array(chartSummarySchema),
          truncated: z.boolean().optional(),
          message: z.string().optional(),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          include_spec: z.boolean().optional().describe('Include each chart\'s raw Google ChartSpec. Large for styled charts.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, include_spec }: any, context: any) => {
          const { accessToken } = context;

          const ctx = await fetchChartContext(
            spreadsheet_id, accessToken, include_spec ? FULL_FIELDS : SUMMARY_FIELDS,
          );

          const listed = ctx.charts.slice(0, MAX_LISTED_CHARTS);
          const omitted = ctx.charts.length - listed.length;
          const charts: ChartSummary[] = listed.map(({ chart }) =>
            summarizeChart(chart, ctx.sheetNameById, ctx.objectSheetIds, Boolean(include_spec)));

          const notes: string[] = [];
          if (omitted > 0) {
            notes.push(`${omitted} further chart(s) are not listed; this spreadsheet has more than the ${MAX_LISTED_CHARTS}-chart limit.`);
          }
          // An omitted range would otherwise be invisible: `series` is a required
          // field, so a short list reads as fact rather than as missing data.
          const withUnreportable = charts.filter((c) => c.unreportableRanges).length;
          if (withUnreportable > 0) {
            notes.push(
              `${withUnreportable} chart(s) have source ranges with no bounded A1 form (whole-column ranges, `
              + 'or data-source column references); their `series`/`domain` are incomplete and carry '
              + '`unreportableRanges`. '
              + 'Pass include_spec for the raw sources.',
            );
          }

          return toolResponse({
            id: spreadsheet_id,
            charts,
            ...truncationFields(notes),
          });
        }))),
      },

      update_chart: {
        description:
          'Change an existing chart\'s type, titles, axis labels, legend, stacking or source ranges. '
          + 'Find `chart_id` with list_charts. IMPORTANT: the Google API replaces a chart\'s whole '
          + 'specification on every update — there is no partial write — so this tool reads the current '
          + 'spec, applies your changes and writes the result back. Fields you do not pass are carried '
          + 'over, but a concurrent edit made between the read and the write is overwritten (last write '
          + 'wins). `domain_range` and `series_ranges` must be changed together, and require `sheet_name`; '
          + 'the same shape rules as add_chart apply. `header_count`, `axis_titles` and `stacked_type` have no '
          + 'effect on a PIE chart, and a pie stores no header count at all — so converting TO pie and back '
          + 'RESETS `header_count` to 1. If yours was 0, pass it again when converting back or row 1 stops '
          + 'being plotted. '
          + 'A chart of a kind these tools do not model (waterfall, treemap, scorecard, COMBO and the '
          + 'rest) can still be retitled: pass title and/or subtitle alone and the rest of its '
          + 'specification is written back untouched. Any other field requires chart_type to convert it '
          + 'first. IMPORTANT: converting to or from BAR transposes the chart — a bar '
          + 'chart runs categories up the left and values along the bottom — so existing axis titles are '
          + 'swapped to follow their data, and any `axis_titles` you pass name the axes of the NEW type, '
          + 'not the old one. Do not echo back the positions list_charts reported before the conversion. '
          + 'A bar chart also has no right-hand axis, so converting to BAR drops a RIGHT axis title and '
          + 'moves every series onto the bottom axis. Changing `chart_type` between PIE and the others '
          + 'converts the chart, keeping its source data. To move or resize a chart use move_chart, which '
          + 'does not rewrite the spec.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          chartId: z.number(),
          chartType: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          chart_id: chartIdSchema.describe('ID of the chart to update (from list_charts or add_chart)'),
          chart_type: chartTypeSchema.optional().describe('New chart type. Converts between PIE and the others.'),
          title: z.string().optional().describe('New chart title'),
          subtitle: z.string().optional().describe('New chart subtitle'),
          axis_titles: axisTitlesSchema.optional().describe('Axis titles to set. Axes you do not name keep their titles. `right` is refused on BAR; a right-hand title inherited from the previous type is dropped instead.'),
          legend_position: legendSchema.optional().describe('New legend placement'),
          stacked_type: stackedSchema.optional().describe('New stacking mode. Refused on LINE and SCATTER; ignored on PIE. Stacking inherited from the previous type is dropped instead of refused.'),
          header_count: z.coerce.number().int().min(0).optional().describe('Leading cells of each range treated as labels'),
          pie_hole: z.coerce.number().min(0).max(1).optional().describe('PIE only: hole size 0..1'),
          three_dimensional: threeDimensionalSchema,
          sheet_name: z.string().optional().describe('Tab holding the source data. Required when changing ranges.'),
          domain_range: z.string().optional().describe('New bounded bare A1 domain range. Must be passed with series_ranges.'),
          series_ranges: z.array(z.string()).optional().describe('New bounded bare A1 series ranges. Must be passed with domain_range.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async (args: any, context: any) => {
          const { accessToken } = context;
          const {
            spreadsheet_id, chart_id, sheet_name, domain_range, series_ranges,
            chart_type, title, subtitle, axis_titles, legend_position, stacked_type,
            header_count, pie_hole, three_dimensional,
          } = args;

          // Both checks run before any request: mergeChartSpec would catch the
          // unpaired case too, but only after a metadata read and a sheet
          // lookup have already been spent on an edit that cannot proceed.
          const changesRanges = domain_range !== undefined || series_ranges !== undefined;
          if ((domain_range === undefined) !== (series_ranges === undefined)) {
            throw new Error(
              'domain_range and series_ranges must be changed together — a new domain with the old '
              + 'series (or the reverse) would be misaligned.',
            );
          }
          if (changesRanges && !sheet_name) {
            throw new Error('sheet_name is required when changing domain_range or series_ranges');
          }

          const ctx = await fetchChartContext(spreadsheet_id, accessToken, FULL_FIELDS);
          const found = findChartOrThrow(
            ctx,
            chart_id,
            'Call list_charts to see the charts that exist and their ids.',
          );

          const sourceSheetId = changesRanges
            ? await getSheetId(spreadsheet_id, sheet_name, accessToken)
            : undefined;

          const spec = mergeChartSpec(found.chart.spec, {
            chartType: chart_type,
            title,
            subtitle,
            axisTitles: axis_titles,
            legendPosition: legend_position,
            stackedType: stacked_type,
            headerCount: header_count,
            pieHole: pie_hole,
            threeDimensional: three_dimensional,
            ...(domain_range !== undefined && {
              domainRange: assertBareA1Range(domain_range, 'domain_range'),
            }),
            ...(series_ranges !== undefined && {
              seriesRanges: series_ranges.map(
                (r: unknown, i: number) => assertBareA1Range(r, `series_ranges[${i}]`),
              ),
            }),
          }, sourceSheetId);

          await applyBatchUpdate(
            spreadsheet_id,
            [{ updateChartSpec: { chartId: chart_id, spec } }],
            accessToken,
            { fields: 'spreadsheetId' },
          );

          // A spec the merge only retitled has no modeled type and was passed
          // through whole, so it neither can be named from `basicChart` nor
          // suffered any of the drops the rewrite message describes.
          const rebuilt = chartTypeOf(spec);
          const chartType = chartTypeLabel(spec) ?? 'unknown';
          return toolResponse({
            id: spreadsheet_id,
            chartId: chart_id,
            chartType,
            message: rebuilt
              ? `Chart ${chart_id} updated to ${chartType}. Its full specification was rewritten, `
                + 'so any change made to it since this call began has been overwritten, along with any '
                + 'setting the result cannot carry: stacking and 3D the new type does not support, a '
                + 'right-hand axis title on a bar chart, a total data label once stacking is off, '
                + 'per-series line and point styling outside LINE, AREA and SCATTER, and custom point '
                + 'labels when the source ranges change.'
              : `Chart ${chart_id} retitled. It is a ${chartType} chart, a type these tools do not `
                + 'model, so only its title and subtitle changed and the rest of its specification was '
                + 'written back unaltered.',
          });
        }))),
      },

      move_chart: {
        description:
          'Move or resize an existing chart without touching its data or styling. Find `chart_id` with '
          + 'list_charts. Pass `anchor_cell` together with `anchor_sheet_name` (and optional '
          + '`width_pixels`, `height_pixels` and pixel offsets) to place the chart over a tab, or '
          + '`new_sheet: true` to move '
          + 'it onto its own dedicated chart sheet. Pixel fields you omit are left as they are, so resizing '
          + 'does not also reposition. Unlike update_chart this does not rewrite the chart specification.',
        outputSchema: {
          id: z.string(),
          chartId: z.number(),
          placement: z.string(),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          chart_id: chartIdSchema.describe('ID of the chart to move (from list_charts or add_chart)'),
          anchor_cell: anchorCellSchema,
          anchor_sheet_name: z.string().optional().describe('Tab to move the chart onto. Required with anchor_cell.'),
          new_sheet: z.boolean().optional().describe('Move the chart onto its own dedicated chart sheet'),
          width_pixels: z.coerce.number().int().positive().optional().describe('New chart width'),
          height_pixels: z.coerce.number().int().positive().optional().describe('New chart height'),
          offset_x_pixels: z.coerce.number().int().optional().describe('Horizontal offset from the anchor cell'),
          offset_y_pixels: z.coerce.number().int().optional().describe('Vertical offset from the anchor cell'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async (args: any, context: any) => {
          const { accessToken } = context;
          const {
            spreadsheet_id, chart_id, anchor_cell, anchor_sheet_name, new_sheet,
            width_pixels, height_pixels, offset_x_pixels, offset_y_pixels,
          } = args;

          if (anchor_cell && !anchor_sheet_name) {
            throw new Error('anchor_sheet_name is required with anchor_cell — name the tab to move the chart onto.');
          }

          const anchorSheetId = anchor_cell
            ? await getSheetId(spreadsheet_id, anchor_sheet_name, accessToken)
            : undefined;

          const placement: ChartPlacement = {
            newSheet: new_sheet,
            anchorCell: anchor_cell,
            anchorSheetId,
            widthPixels: width_pixels,
            heightPixels: height_pixels,
            offsetXPixels: offset_x_pixels,
            offsetYPixels: offset_y_pixels,
          };
          const newPosition = buildPosition(placement);

          // `updateEmbeddedObjectPosition` reaches images and other embedded
          // objects too, exactly like the delete request. Without this lookup
          // the tool would happily relocate an image and then report "Chart N
          // moved", which is a lie the caller has no way to catch.
          findChartOrThrow(
            await fetchChartContext(spreadsheet_id, accessToken, LOCATE_FIELDS),
            chart_id,
            'The id may belong to an image or another embedded object, which this tool will not move. '
            + 'Call list_charts to see the charts that exist.',
          );

          await applyBatchUpdate(
            spreadsheet_id,
            [{
              updateEmbeddedObjectPosition: {
                objectId: chart_id,
                newPosition,
                // Only meaningful for an overlay; Google ignores it otherwise.
                ...(newPosition.overlayPosition && { fields: overlayFieldMask(placement) }),
              },
            }],
            accessToken,
            { fields: 'spreadsheetId' },
          );

          const placementText = describePlacement(new_sheet, anchor_sheet_name, anchor_cell);

          return toolResponse({
            id: spreadsheet_id,
            chartId: chart_id,
            placement: placementText,
            message: `Chart ${chart_id} moved to ${placementText}`,
          });
        }))),
      },

      delete_chart: {
        description:
          'Delete a chart by `chart_id`. Find the id with list_charts. The underlying Google request also '
          + 'deletes images and other embedded objects, so this tool first confirms the id belongs to a '
          + 'chart and refuses otherwise. If the chart occupies its own chart sheet, the response reports '
          + 'whether that sheet was removed along with it. Deleting a chart does not touch the data it '
          + 'was built from.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          chartId: z.number(),
          wasOnOwnSheet: z.boolean(),
          ownSheetRemoved: z.boolean().optional().describe('Reported only when the chart had its own chart sheet'),
          message: z.string(),
        },
        schema: {
          spreadsheet_id: spreadsheetIdSchema,
          chart_id: chartIdSchema.describe('ID of the chart to delete (from list_charts)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/spreadsheets", wrapHandler(nativeOnly(async ({ spreadsheet_id, chart_id }: any, context: any) => {
          const { accessToken } = context;

          const ctx = await fetchChartContext(spreadsheet_id, accessToken, LOCATE_FIELDS);
          const found = findChartOrThrow(
            ctx,
            chart_id,
            'The id may belong to an image or another embedded object, which this tool will not '
            + 'delete. Call list_charts to see the charts that exist.',
          );

          const hostSheetId = found.sheetId;
          const wasOnOwnSheet = hostSheetId !== undefined && ctx.objectSheetIds.has(hostSheetId);

          await applyBatchUpdate(
            spreadsheet_id,
            [{ deleteEmbeddedObject: { objectId: chart_id } }],
            accessToken,
            { fields: 'spreadsheetId' },
          );

          // Google does not document what becomes of a chart sheet whose only
          // object is removed, so the outcome is observed rather than assumed.
          //
          // The delete has already succeeded by this point, so a failure here
          // must not fail the tool: reporting a completed destructive action
          // as an error invites a retry that then says the chart is missing.
          // The observation is best-effort and degrades to "not reported".
          let ownSheetRemoved: boolean | undefined;
          if (wasOnOwnSheet) {
            try {
              const after = await fetchChartContext(
                spreadsheet_id, accessToken, SHEET_ONLY_FIELDS,
              );
              ownSheetRemoved = !after.sheetNameById.has(hostSheetId!);
            } catch {
              ownSheetRemoved = undefined;
            }
          }

          const sheetNote = !wasOnOwnSheet
            ? ''
            : ownSheetRemoved === undefined
              ? ' It was on its own chart sheet; whether that sheet survived could not be confirmed.'
              : ownSheetRemoved
                ? ' Its chart sheet was removed with it.'
                : ' Its chart sheet remains, now empty.';

          return toolResponse({
            id: spreadsheet_id,
            chartId: chart_id,
            wasOnOwnSheet,
            ...(ownSheetRemoved !== undefined && { ownSheetRemoved }),
            message: `Chart ${chart_id} deleted.${sheetNote}`,
          });
        }))),
      },
};
