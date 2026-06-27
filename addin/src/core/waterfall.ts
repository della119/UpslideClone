/*
 * Waterfall builder
 * -----------------
 * Reproduces Upslide's waterfall WITHOUT a native waterfall chart type:
 * a standard STACKED COLUMN chart where an invisible "Base" series floats each
 * delta bar to the correct height. Non-applicable series cells are set to #N/A
 * (=NA()) so Excel draws no bar and no zero-label for them — the same trick the
 * UPSLIDE_Waterfall helper tab uses.
 *
 * Anchor (total) rows are auto-detected: a row whose value equals the running
 * cumulative of the preceding deltas is treated as a total (full-height bar).
 * The first row is always an anchor.
 *
 * This logic is intentionally framework-light so it ports cleanly to C#/VSTO.
 */

const COLORS = {
  increase: "#86BC25", // brand green
  decrease: "#E0301E", // brand red
  total: "#53565A", // brand cool gray
};
const ANCHOR_TOLERANCE = 1e-6;

interface Point {
  label: string;
  value: number;
}
interface WfRow {
  label: string;
  base: number;
  decrease: number | "=NA()";
  increase: number | "=NA()";
  total: number | "=NA()";
}

/** Pure geometry — unit-testable, no Office dependency. */
export function computeWaterfall(points: Point[]): WfRow[] {
  const rows: WfRow[] = [];
  let running = 0;
  points.forEach((p, i) => {
    const isAnchor =
      i === 0 || Math.abs(p.value - running) < ANCHOR_TOLERANCE;
    if (isAnchor) {
      running = p.value;
      rows.push({ label: p.label, base: 0, decrease: "=NA()", increase: "=NA()", total: p.value });
    } else if (p.value >= 0) {
      rows.push({ label: p.label, base: running, decrease: "=NA()", increase: p.value, total: "=NA()" });
      running += p.value;
    } else {
      running += p.value; // running now at the new (lower) level
      rows.push({ label: p.label, base: running, decrease: -p.value, increase: "=NA()", total: "=NA()" });
    }
  });
  return rows;
}

export async function buildWaterfallFromSelection(): Promise<string> {
  return Excel.run(async (context) => {
    const sel = context.workbook.getSelectedRange();
    sel.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    const sheet = sel.worksheet;
    sheet.load("name");
    await context.sync();

    // Parse label/value pairs; skip non-numeric (title/header) rows.
    const points: Point[] = [];
    for (const r of sel.values) {
      const label = r[0] == null ? "" : String(r[0]);
      const raw = r[1];
      if (typeof raw === "number" && !isNaN(raw)) points.push({ label, value: raw });
    }
    if (points.length < 2) {
      throw new Error("Select a two-column table (labels + values) with at least 2 numeric rows.");
    }

    const rows = computeWaterfall(points);

    // Write the helper block below the selection.
    const startRow = sel.rowIndex + sel.rowCount + 2;
    const startCol = sel.columnIndex;
    const header = ["Item", "Base", "Decrease", "Increase", "Total"];
    const block: (string | number)[][] = [header];
    rows.forEach((w) => block.push([w.label, w.base, w.decrease, w.increase, w.total]));

    const helper = sheet.getRangeByIndexes(startRow, startCol, block.length, 5);
    helper.formulas = block; // numbers stay numbers, "=NA()" becomes #N/A
    helper.format.font.name = "Calibri";
    helper.format.font.size = 9;

    // Build the stacked column chart from the helper block.
    const chart = sheet.charts.add(Excel.ChartType.columnStacked, helper, Excel.ChartSeriesBy.columns);
    chart.title.text = "Waterfall";
    chart.legend.position = Excel.ChartLegendPosition.bottom;
    chart.load("series");
    await context.sync();

    chart.series.load("items/name");
    await context.sync();

    // Series order matches helper columns 2..5: Base, Decrease, Increase, Total.
    const byName: Record<string, Excel.ChartSeries> = {};
    chart.series.items.forEach((s) => (byName[s.name] = s));

    if (byName["Base"]) byName["Base"].format.fill.clear(); // invisible floor
    if (byName["Increase"]) byName["Increase"].format.fill.setSolidColor(COLORS.increase);
    if (byName["Decrease"]) byName["Decrease"].format.fill.setSolidColor(COLORS.decrease);
    if (byName["Total"]) byName["Total"].format.fill.setSolidColor(COLORS.total);

    // Data labels on the visible series only.
    for (const n of ["Increase", "Decrease", "Total"]) {
      try {
        if (byName[n]) {
          byName[n].hasDataLabels = true;
          byName[n].dataLabels.numberFormat = "#,##0;(#,##0)";
        }
      } catch (e) {
        /* older API: skip per-series labels */
      }
    }

    // Hide the "Base" entry from the legend.
    try {
      chart.legend.load("legendEntries");
      await context.sync();
      const baseIdx = chart.series.items.findIndex((s) => s.name === "Base");
      if (baseIdx >= 0) chart.legend.legendEntries.getItemAt(baseIdx).visible = false;
    } catch (e) {
      /* legend entry API not available — leave as is */
    }

    chart.top = 10;
    chart.left = 10;
    await context.sync();
    return `Waterfall built from ${points.length} points.`;
  });
}
