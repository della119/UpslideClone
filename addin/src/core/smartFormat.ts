/*
 * Smart Format (tables)
 * ---------------------
 * One-click branded formatting for a selected financial table, reproducing the
 * "Expected Result" in the Upslide training workbook:
 *   - header row: dark fill, white bold font, bottom border
 *   - result/total rows (EBITDA, Gross Margin, ...): bold + top border + light fill
 *   - number formats: thousands with parens for negatives; % for ratio columns
 *   - font Calibri; labels left-aligned, numbers right-aligned
 *   - outer border
 *
 * Heuristics are deliberate and conservative; the task pane will later expose
 * manual overrides (mark header / mark result row) like Upslide's toggles.
 * Framework-light for a clean C#/VSTO port.
 */

const BRAND = {
  headerFill: "#53565A", // brand cool gray
  headerFont: "#FFFFFF",
  resultFill: "#EFEFEF",
  border: "#BFBFBF",
};

const NUM_FMT = "#,##0;(#,##0)";
const PCT_FMT = "0.0%";

const RESULT_RE = /\b(ebitda|ebit|gross margin|net income|net profit|operating (income|profit)|total|net sales|grand total)\b/i;

/** Decide if a column holds ratios (→ percent format). Pure / testable. */
export function isPercentColumn(headerText: string, values: number[]): boolean {
  if (/%|cagr|margin|growth|rate|ratio/i.test(headerText)) return true;
  const nums = values.filter((v) => typeof v === "number" && !isNaN(v));
  if (nums.length === 0) return false;
  return nums.every((v) => Math.abs(v) <= 1 && v !== Math.trunc(v));
}

export async function smartFormatSelection(): Promise<string> {
  return Excel.run(async (context) => {
    const sel = context.workbook.getSelectedRange();
    sel.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    const sheet = sel.worksheet;
    await context.sync();

    const nRows = sel.rowCount;
    const nCols = sel.columnCount;
    if (nRows < 2 || nCols < 2) throw new Error("Select a table with at least 2 rows and 2 columns.");

    const vals = sel.values;

    // Column 0 = labels; columns 1.. = data. Per data column, decide % vs number.
    const headerRow = vals[0].map((c) => (c == null ? "" : String(c)));
    const pctCol: boolean[] = [];
    for (let c = 1; c < nCols; c++) {
      const colVals = vals.slice(1).map((r) => r[c]).filter((x) => typeof x === "number") as number[];
      pctCol[c] = isPercentColumn(headerRow[c], colVals);
    }

    // Base styling on the whole range.
    sel.format.font.name = "Calibri";
    sel.format.font.size = 10;

    // Number formats (build a full 2D array).
    const fmt: string[][] = [];
    for (let r = 0; r < nRows; r++) {
      const row: string[] = [];
      for (let c = 0; c < nCols; c++) {
        if (r === 0 || c === 0) row.push("General");
        else row.push(pctCol[c] ? PCT_FMT : NUM_FMT);
      }
      fmt.push(row);
    }
    sel.numberFormat = fmt;

    // Alignment: labels left, numbers right.
    const labelCol = sheet.getRangeByIndexes(sel.rowIndex + 1, sel.columnIndex, nRows - 1, 1);
    labelCol.format.horizontalAlignment = Excel.HorizontalAlignment.left;
    const dataCols = sheet.getRangeByIndexes(sel.rowIndex + 1, sel.columnIndex + 1, nRows - 1, nCols - 1);
    dataCols.format.horizontalAlignment = Excel.HorizontalAlignment.right;

    // Header row.
    const header = sheet.getRangeByIndexes(sel.rowIndex, sel.columnIndex, 1, nCols);
    header.format.fill.color = BRAND.headerFill;
    header.format.font.color = BRAND.headerFont;
    header.format.font.bold = true;
    header.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;

    // Result / total rows.
    for (let r = 1; r < nRows; r++) {
      const label = vals[r][0] == null ? "" : String(vals[r][0]);
      if (RESULT_RE.test(label)) {
        const rng = sheet.getRangeByIndexes(sel.rowIndex + r, sel.columnIndex, 1, nCols);
        rng.format.font.bold = true;
        rng.format.fill.color = BRAND.resultFill;
        rng.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
      }
    }

    // Outer border.
    for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]) {
      const b = sel.format.borders.getItem(edge as Excel.BorderIndex);
      b.style = Excel.BorderLineStyle.continuous;
      b.color = BRAND.border;
    }

    await context.sync();
    return `Formatted ${nRows}×${nCols} table.`;
  });
}
