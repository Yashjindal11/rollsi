/* global Excel */
import { pareto, shortNumber, splitBar, toNum, waterfall } from "../../core";
import type { Num } from "../../core";
import { createChart, formatOf, readSelection, writeHelperBlock } from "./excel";
import type { DataSelection } from "./excel";
import { currentPalette } from "./palette";
import { applyStyle } from "./style";

const bool = (b: boolean) => (b ? "TRUE" : "FALSE");

/** Label column (first column when there are 2+ columns) or a 1..n sequence. */
function labelFormula(sel: DataSelection): { formula: string; header: string; format: string } {
  if (sel.columnCount >= 2) {
    return { formula: `=${sel.colRef(0)}`, header: sel.headers[0], format: formatOf(sel, 0) };
  }
  return { formula: `=SEQUENCE(${sel.dataRows})`, header: "#", format: "General" };
}

function valueCol(sel: DataSelection): number {
  return sel.columnCount >= 2 ? 1 : 0;
}

export async function buildTrendMA(
  context: Excel.RequestContext,
  window: number,
  sel?: DataSelection
) {
  sel =
    sel ??
    (await readSelection(
      context,
      1,
      "a column of numbers (optionally with dates/labels on the left)"
    ));
  const v = valueCol(sel);
  const label = labelFormula(sel);
  const ref = sel.colRef(v);
  const maName = `MA ${window}`;
  const block = await writeHelperBlock(
    context,
    `Trend + MA · ${sel.address}`,
    [label.header, sel.headers[v], maName],
    [
      [0, label.formula],
      [1, `=ROLLSI.NA(${ref})`],
      [2, `=ROLLSI.NA(ROLLSI.MA(${ref},${window}))`],
    ],
    sel.dataRows,
    { 0: label.format, 1: formatOf(sel, v), 2: formatOf(sel, v) }
  );
  const { chart } = await createChart(
    context,
    sel,
    Excel.ChartType.line,
    [
      { name: sel.headers[v], values: block.data(1) },
      { name: maName, values: block.data(2) },
    ],
    block.data(0),
    `${sel.headers[v]}: ${window}-period moving average`
  );
  await applyStyle(context, chart, currentPalette(), { labels: false });
  return chart;
}

export async function buildControl(
  context: Excel.RequestContext,
  window: number | null,
  k: number
) {
  const sel = await readSelection(
    context,
    1,
    "a column of numbers (optionally with dates/labels on the left)"
  );
  const v = valueCol(sel);
  const label = labelFormula(sel);
  const ref = sel.colRef(v);
  const block = await writeHelperBlock(
    context,
    `Control chart · ${sel.address}`,
    [label.header, sel.headers[v], "Center", "Upper", "Lower", "Outlier"],
    [
      [0, label.formula],
      [1, `=ROLLSI.NA(${ref})`],
      [2, `=ROLLSI.CONTROL(${ref},${window ?? ""},${k})`],
    ],
    sel.dataRows,
    { 0: label.format }
  );
  const names = [sel.headers[v], "Center", "Upper", "Lower", "Outlier"];
  const { chart } = await createChart(
    context,
    sel,
    Excel.ChartType.line,
    names.map((name, i) => ({ name, values: block.data(i + 1) })),
    block.data(0),
    `${sel.headers[v]}: control chart (±${k}σ${window ? `, rolling ${window}` : ""})`
  );
  await applyStyle(context, chart, currentPalette(), { labels: false });
  return chart;
}

export async function buildWaterfall(
  context: Excel.RequestContext,
  firstIsStart: boolean,
  lastIsTotal: boolean
) {
  const sel = await readSelection(context, 2, "labels and values (2 columns)");
  const labelsRef = sel.colRef(0);
  const valuesRef = sel.colRef(1);
  const names = [
    "_base+",
    "_base-",
    "Increase",
    "Increase (-)",
    "Decrease",
    "Decrease (-)",
    "Total",
    "Total (-)",
  ];
  const block = await writeHelperBlock(
    context,
    `Waterfall · ${sel.address}`,
    [sel.headers[0], ...names],
    [
      [0, `=${labelsRef}`],
      [
        1,
        `=ROLLSI.WATERFALL(${labelsRef},${valuesRef},${bool(firstIsStart)},${bool(lastIsTotal)})`,
      ],
    ],
    sel.dataRows
  );
  const { chart, series } = await createChart(
    context,
    sel,
    Excel.ChartType.columnStacked,
    names.map((name, i) => ({ name, values: block.data(i + 1) })),
    block.data(0),
    `${sel.headers[1]}: bridge`
  );
  await applyStyle(context, chart, currentPalette(), { labels: false });
  series[0].gapWidth = 40;

  // Static labels with the step change (re-run to refresh after data changes).
  const steps = waterfall(sel.column(0), sel.column(1).map(toNum), firstIsStart, lastIsTotal);
  steps.forEach((s, i) => {
    const p = splitBar(s.from, s.to);
    const slot = s.kind === "up" ? 2 : s.kind === "down" ? 4 : 6;
    const idx = p.pos !== 0 || p.neg === 0 ? slot : slot + 1;
    const point = series[idx].points.getItemAt(i);
    point.hasDataLabel = true;
    const signed = s.kind === "up" ? "+" + shortNumber(s.label) : shortNumber(s.label);
    point.dataLabel.text = signed;
    point.dataLabel.format.font.size = 9;
    point.dataLabel.format.font.bold = s.kind === "total" || s.kind === "start";
    point.dataLabel.format.font.color = "#FFFFFF";
  });
  chart.legend.visible = false;
  await context.sync();
  return chart;
}

export async function buildActualVsTarget(context: Excel.RequestContext) {
  const sel = await readSelection(context, 3, "labels, actual and target (3 columns)");
  const targetName = /^target/i.test(sel.headers[2])
    ? sel.headers[2]
    : `Target (${sel.headers[2]})`;
  const block = await writeHelperBlock(
    context,
    `Actual vs target · ${sel.address}`,
    [sel.headers[0], sel.headers[1], targetName, "Over", "Under"],
    [
      [0, `=${sel.colRef(0)}`],
      [1, `=ROLLSI.VSTARGET(${sel.colRef(1)},${sel.colRef(2)})`],
    ],
    sel.dataRows,
    { 0: formatOf(sel, 0) }
  );
  const palette = currentPalette();
  const main = await createChart(
    context,
    sel,
    Excel.ChartType.columnClustered,
    [
      { name: sel.headers[1], values: block.data(1) },
      { name: targetName, values: block.data(2), type: Excel.ChartType.lineMarkers },
    ],
    block.data(0),
    `${sel.headers[1]} vs ${sel.headers[2]}`
  );
  await applyStyle(context, main.chart, palette, { labels: true });
  main.chart.load("top,height,left");
  await context.sync();
  const variance = await createChart(
    context,
    sel,
    Excel.ChartType.columnStacked,
    [
      { name: "Over", values: block.data(3) },
      { name: "Under", values: block.data(4) },
    ],
    block.data(0),
    "Variance vs target",
    { left: main.chart.left, top: main.chart.top + main.chart.height + 8, height: 170 }
  );
  await applyStyle(context, variance.chart, palette, { labels: true });
  variance.chart.legend.visible = false;
  await context.sync();
  return main.chart;
}

export async function buildPareto(context: Excel.RequestContext) {
  const sel = await readSelection(context, 2, "labels and values (2 columns)");
  const rows = pareto(sel.column(0), sel.column(1).map(toNum)).length;
  if (rows === 0) throw new Error("No positive values to chart.");
  const block = await writeHelperBlock(
    context,
    `Pareto · ${sel.address}`,
    [sel.headers[0], sel.headers[1], "Cumulative %", "80% line"],
    [[0, `=ROLLSI.PARETO(${sel.colRef(0)},${sel.colRef(1)})`]],
    rows,
    { 2: "0%", 3: "0%" }
  );
  const { chart } = await createChart(
    context,
    sel,
    Excel.ChartType.columnClustered,
    [
      { name: sel.headers[1], values: block.data(1) },
      { name: "Cumulative %", values: block.data(2), type: Excel.ChartType.line, secondary: true },
      { name: "80% line", values: block.data(3), type: Excel.ChartType.line, secondary: true },
    ],
    block.data(0),
    `Pareto: ${sel.headers[1]}`
  );
  await applyStyle(context, chart, currentPalette(), { labels: false });
  const secondary = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
  secondary.minimum = 0;
  secondary.maximum = 1;
  secondary.numberFormat = "0%";
  secondary.majorGridlines.visible = false;
  await context.sync();
  return chart;
}

export async function buildGantt(context: Excel.RequestContext) {
  const sel = await readSelection(context, 3, "tasks, start dates and end dates (3 columns)");
  const starts = sel
    .column(1)
    .map(toNum)
    .filter((x: Num): x is number => x !== null);
  if (!starts.length) throw new Error("The second column needs start dates.");
  const block = await writeHelperBlock(
    context,
    `Gantt · ${sel.address}`,
    [sel.headers[0], "_start", "Duration (days)"],
    [
      [0, `=${sel.colRef(0)}`],
      [1, `=${sel.colRef(1)}`],
      [2, `=${sel.colRef(2)}-${sel.colRef(1)}`],
    ],
    sel.dataRows,
    { 1: "yyyy-mm-dd" }
  );
  const { chart } = await createChart(
    context,
    sel,
    Excel.ChartType.barStacked,
    [
      { name: "_start", values: block.data(1) },
      { name: "Duration (days)", values: block.data(2) },
    ],
    block.data(0),
    "Timeline",
    { height: Math.max(200, 24 * sel.dataRows + 70) }
  );
  await applyStyle(context, chart, currentPalette(), { labels: false });
  chart.axes.categoryAxis.reversePlotOrder = true;
  const va = chart.axes.valueAxis;
  va.minimum = Math.min(...starts);
  va.numberFormat = "d mmm";
  va.visible = true;
  va.majorGridlines.visible = true;
  va.majorGridlines.format.line.color = "#EDEDED";
  chart.legend.visible = false;
  await context.sync();
  return chart;
}
