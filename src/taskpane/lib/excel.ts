/* global Excel, Office */

export const HELPER_SHEET = "ROLLSI Data";
export const TAG = "ROLLSI|";

export function colLetter(index0: number): string {
  let s = "";
  let n = index0 + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function quoteSheet(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

export function apiSupported(version: string): boolean {
  return Office.context.requirements.isSetSupported("ExcelApi", version);
}

export function requireApi(version: string, feature: string) {
  if (!apiSupported(version)) {
    throw new Error(
      `${feature} needs a newer Excel (ExcelApi ${version}). Update Excel or use Excel on the web.`
    );
  }
}

export interface DataSelection {
  sheetName: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  values: any[][];
  numberFormat: any[][];
  hasHeader: boolean;
  headers: string[];
  dataRows: number;
  left: number;
  top: number;
  width: number;
  height: number;
  /** Sheet-qualified absolute reference to the data rows of column c (0-based within the selection). */
  colRef(c: number): string;
  /** Data values of column c (header excluded). */
  column(c: number): any[];
}

/**
 * Reads the selected block. A single selected cell expands to its surrounding region.
 * The first row is treated as a header when it has text and the next row has a number.
 */
export async function readSelection(
  context: Excel.RequestContext,
  minCols = 1,
  what = "your data"
): Promise<DataSelection> {
  let range = context.workbook.getSelectedRange();
  range.load("rowCount,columnCount");
  await context.sync();
  if (range.rowCount === 1 && range.columnCount === 1 && apiSupported("1.13")) {
    range = range.getSurroundingRegion();
  }
  range.load(
    "address,values,numberFormat,rowIndex,columnIndex,rowCount,columnCount,left,top,width,height"
  );
  const worksheet = context.workbook.worksheets.getActiveWorksheet();
  worksheet.load("name");
  await context.sync();

  const values = range.values;
  const hasHeader =
    range.rowCount > 1 &&
    values[0].some((v) => typeof v === "string" && v.trim() !== "") &&
    values[1].some((v) => typeof v === "number");
  const dataRows = range.rowCount - (hasHeader ? 1 : 0);
  if (range.columnCount < minCols || dataRows < 1) {
    throw new Error(
      `Select ${what} first (at least ${minCols} column${minCols > 1 ? "s" : ""} with a header row).`
    );
  }
  const sheetName = worksheet.name;
  const firstRow = range.rowIndex + 1 + (hasHeader ? 1 : 0);
  const lastRow = range.rowIndex + range.rowCount;
  const headers = values[0].map((v, c) =>
    hasHeader && String(v).trim() ? String(v).trim() : `Series ${c + 1}`
  );
  return {
    sheetName,
    address: range.address,
    rowIndex: range.rowIndex,
    columnIndex: range.columnIndex,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    values,
    numberFormat: range.numberFormat,
    hasHeader,
    headers,
    dataRows,
    left: range.left,
    top: range.top,
    width: range.width,
    height: range.height,
    colRef(c: number) {
      const col = colLetter(range.columnIndex + c);
      return `${quoteSheet(sheetName)}!$${col}$${firstRow}:$${col}$${lastRow}`;
    },
    column(c: number) {
      return values.slice(hasHeader ? 1 : 0).map((r) => r[c]);
    },
  };
}

export interface HelperBlock {
  sheet: Excel.Worksheet;
  col0: number;
  /** Data rows of block column c (below the header). */
  data(c: number, rows?: number): Excel.Range;
}

/**
 * Writes a block of live formulas on the "ROLLSI Data" sheet, to the right of existing blocks:
 * row 1 = title, row 2 = headers, row 3 = formulas (which spill down).
 */
export async function writeHelperBlock(
  context: Excel.RequestContext,
  title: string,
  headers: string[],
  formulas: [number, string][],
  rows: number,
  formats: Record<number, string> = {}
): Promise<HelperBlock> {
  let sheet = context.workbook.worksheets.getItemOrNullObject(HELPER_SHEET);
  await context.sync();
  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(HELPER_SHEET);
    sheet.tabColor = "#107C41";
  }
  const used = sheet.getUsedRangeOrNullObject();
  used.load("columnIndex,columnCount");
  await context.sync();
  const col0 = used.isNullObject ? 0 : used.columnIndex + used.columnCount + 1;

  const titleCell = sheet.getRangeByIndexes(0, col0, 1, 1);
  titleCell.values = [[title]];
  titleCell.format.font.bold = true;
  titleCell.format.font.color = "#107C41";
  const head = sheet.getRangeByIndexes(1, col0, 1, headers.length);
  head.values = [headers];
  head.format.font.bold = true;
  head.format.fill.color = "#E8F3EC";
  formulas.forEach(([c, f]) => {
    sheet.getRangeByIndexes(2, col0 + c, 1, 1).formulas = [[f]];
  });
  Object.entries(formats).forEach(([c, fmt]) => {
    sheet.getRangeByIndexes(2, col0 + Number(c), rows, 1).numberFormat = Array.from(
      { length: rows },
      () => [fmt]
    );
  });
  await context.sync();
  return {
    sheet,
    col0,
    data: (c: number, n = rows) => sheet.getRangeByIndexes(2, col0 + c, n, 1),
  };
}

export interface SeriesSpec {
  name: string;
  values: Excel.Range;
  type?: Excel.ChartType | string;
  secondary?: boolean;
}

/**
 * Creates a chart on the source sheet next to the selection, with explicitly defined series
 * (so Excel's guessing of headers/categories never gets in the way).
 */
export async function createChart(
  context: Excel.RequestContext,
  sel: DataSelection,
  type: Excel.ChartType | string,
  specs: SeriesSpec[],
  categories: Excel.Range,
  title: string,
  place: { left?: number; top?: number; width?: number; height?: number } = {}
): Promise<{ chart: Excel.Chart; series: Excel.ChartSeries[] }> {
  const ws = context.workbook.worksheets.getItem(sel.sheetName);
  const chart = ws.charts.add(
    type as Excel.ChartType,
    specs[0].values,
    Excel.ChartSeriesBy.columns
  );
  chart.series.load("items");
  await context.sync();
  chart.series.items.forEach((s) => s.delete());
  const series = specs.map((sp) => {
    const s = chart.series.add(sp.name);
    s.setValues(sp.values);
    s.setXAxisValues(categories);
    if (sp.type) s.chartType = sp.type as Excel.ChartType;
    if (sp.secondary) s.axisGroup = Excel.ChartAxisGroup.secondary;
    return s;
  });
  chart.title.text = title;
  chart.title.visible = true;
  chart.left = place.left ?? sel.left + sel.width + 24;
  chart.top = place.top ?? sel.top;
  chart.width = place.width ?? 520;
  chart.height = place.height ?? 300;
  chart.name = `ROLLSI ${title}`.slice(0, 60) + ` ${Date.now() % 100000}`;
  await context.sync();
  return { chart, series };
}

export function formatOf(sel: DataSelection, c: number): string {
  const row = sel.hasHeader ? 1 : 0;
  const fmt = sel.numberFormat[row]?.[c];
  return typeof fmt === "string" && fmt !== "General" ? fmt : "General";
}

/** Groups freshly created shapes (ids are synced first, which every Excel build accepts) and names the group. */
export async function groupShapes(
  context: Excel.RequestContext,
  shapes: Excel.ShapeCollection,
  parts: Excel.Shape[],
  name: string
): Promise<void> {
  parts.forEach((p) => p.load("id"));
  await context.sync();
  const grp = shapes.addGroup(parts.map((p) => p.id));
  grp.name = name;
  await context.sync();
}
