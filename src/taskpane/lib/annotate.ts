/* global Excel */
import { cagrLabel, categoryX, diffLabel, shortNumber, valueY } from "../../core/viz";
import type { DiffMode, PlotBox } from "../../core/viz";
import { groupShapes, requireApi, TAG } from "./excel";
import { SEMANTIC } from "./palette";

interface ChartGeometry {
  chart: Excel.Chart;
  sheet: Excel.Worksheet;
  name: string;
  box: PlotBox;
  n: number;
  /** Top of each category: stack top for stacked charts, first visible series otherwise. */
  tops: number[];
  /** Net total per category across visible series. */
  totals: number[];
  catWidth: number;
}

async function activeChart(context: Excel.RequestContext): Promise<Excel.Chart> {
  requireApi("1.12", "Annotations");
  const chart = context.workbook.getActiveChartOrNullObject();
  await context.sync();
  if (chart.isNullObject) throw new Error("Click a chart first, then use the annotation buttons.");
  return chart;
}

async function geometry(context: Excel.RequestContext, chart: Excel.Chart): Promise<ChartGeometry> {
  chart.load("name,left,top,chartType");
  chart.worksheet.load("name");
  const pa = chart.plotArea;
  pa.load("insideLeft,insideTop,insideWidth,insideHeight");
  const va = chart.axes.valueAxis;
  va.load("minimum,maximum");
  chart.series.load("items/name");
  await context.sync();

  const type = String(chart.chartType);
  if (/^bar|pie|doughnut|scatter|radar|area|bubble/i.test(type)) {
    throw new Error("Annotations work on column and line charts.");
  }
  const results = chart.series.items.map((s) => ({
    name: s.name,
    values: s.getDimensionValues(Excel.ChartSeriesDimension.values),
  }));
  await context.sync();

  const series = results.map((r) => ({
    name: r.name,
    values: r.values.value.map((x) => Number(x) || 0),
  }));
  const n = series.length ? series[0].values.length : 0;
  if (!n) throw new Error("The chart has no data.");
  const stacked = /stacked/i.test(type);
  const visible = series.filter((s) => !s.name.startsWith("_"));
  const tops: number[] = [];
  const totals: number[] = [];
  for (let i = 0; i < n; i++) {
    const all = series.map((s) => s.values[i] ?? 0);
    tops.push(
      stacked
        ? all.filter((v) => v > 0).reduce((a, b) => a + b, 0)
        : (visible[0] ?? series[0]).values[i]
    );
    totals.push(visible.reduce((a, s) => a + (s.values[i] ?? 0), 0));
  }
  const box: PlotBox = {
    left: chart.left + pa.insideLeft,
    top: chart.top + pa.insideTop,
    width: pa.insideWidth,
    height: pa.insideHeight,
    min: va.minimum,
    max: va.maximum,
  };
  return {
    chart,
    sheet: chart.worksheet,
    name: chart.name,
    box,
    n,
    tops,
    totals,
    catWidth: box.width / n,
  };
}

function line(
  shapes: Excel.ShapeCollection,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  opts: { color?: string; weight?: number; dash?: boolean; arrow?: boolean } = {}
): Excel.Shape {
  const s = shapes.addLine(x1, y1, x2, y2, Excel.ConnectorType.straight);
  s.lineFormat.color = opts.color ?? SEMANTIC.ink;
  s.lineFormat.weight = opts.weight ?? 1;
  if (opts.dash) s.lineFormat.dashStyle = Excel.ShapeLineDashStyle.dash;
  if (opts.arrow) s.line.endArrowheadStyle = Excel.ArrowheadStyle.triangle;
  return s;
}

function label(
  shapes: Excel.ShapeCollection,
  text: string,
  cx: number,
  cy: number,
  opts: { bubble?: boolean; plain?: boolean; color?: string } = {}
): Excel.Shape {
  const w = 6 * text.length + (opts.bubble ? 22 : 12);
  const h = opts.bubble ? 22 : 16;
  const s = opts.plain
    ? shapes.addTextBox(text)
    : shapes.addGeometricShape(
        opts.bubble ? Excel.GeometricShapeType.ellipse : Excel.GeometricShapeType.roundRectangle
      );
  s.left = cx - w / 2;
  s.top = cy - h / 2;
  s.width = w;
  s.height = h;
  if (opts.plain) {
    s.fill.clear();
    s.lineFormat.visible = false;
  } else {
    s.fill.setSolidColor("#FFFFFF");
    s.lineFormat.color = opts.color ?? SEMANTIC.ink;
    s.lineFormat.weight = 0.75;
    s.textFrame.textRange.text = text;
  }
  const tf = s.textFrame;
  tf.leftMargin = 1;
  tf.rightMargin = 1;
  tf.topMargin = 0;
  tf.bottomMargin = 0;
  tf.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
  tf.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
  tf.textRange.font.size = 9;
  tf.textRange.font.bold = true;
  tf.textRange.font.color = opts.color ?? SEMANTIC.ink;
  return s;
}

function group(
  context: Excel.RequestContext,
  g: ChartGeometry,
  shapes: Excel.Shape[],
  tag: string
) {
  return groupShapes(context, g.sheet.shapes, shapes, `${TAG}${g.name}|${tag}`);
}

function drawDiff(
  context: Excel.RequestContext,
  g: ChartGeometry,
  a: number,
  b: number,
  mode: DiffMode
) {
  if (a < 0 || b < 0 || a >= g.n || b >= g.n || a === b) {
    throw new Error(`Pick two different categories between 1 and ${g.n}.`);
  }
  const shapes = g.sheet.shapes;
  const xa = categoryX(g.box, a, g.n);
  const xb = categoryX(g.box, b, g.n);
  const ya = valueY(g.box, g.tops[a]);
  const yb = valueY(g.box, g.tops[b]);
  const half = g.catWidth * 0.32;
  const xr = Math.max(xa, xb) + g.catWidth * 0.5;
  const text = diffLabel(g.tops[a], g.tops[b], mode);
  const color = g.tops[b] >= g.tops[a] ? SEMANTIC.pos : SEMANTIC.neg;
  const parts = [
    line(shapes, xa + half, ya, xr, ya, { color: SEMANTIC.band, dash: true, weight: 0.75 }),
    line(shapes, xb + half, yb, xr, yb, { color: SEMANTIC.band, dash: true, weight: 0.75 }),
  ];
  if (Math.abs(ya - yb) > 2)
    parts.push(line(shapes, xr, ya, xr, yb, { color, weight: 1.25, arrow: true }));
  parts.push(label(shapes, text, xr + 6 * text.length * 0.5 + 10, (ya + yb) / 2, { color }));
  return group(context, g, parts, `diff|${a}|${b}|${mode}`);
}

function drawCagr(context: Excel.RequestContext, g: ChartGeometry, a: number, b: number) {
  if (a < 0 || b >= g.n || b <= a)
    throw new Error(`Pick a start before the end, between 1 and ${g.n}.`);
  const shapes = g.sheet.shapes;
  const lift = 16;
  const x1 = categoryX(g.box, a, g.n);
  const x2 = categoryX(g.box, b, g.n);
  const y1 = valueY(g.box, g.tops[a]) - lift;
  const y2 = valueY(g.box, g.tops[b]) - lift;
  const text = cagrLabel(g.tops[a], g.tops[b], b - a);
  const parts = [
    line(shapes, x1, y1, x2, y2, { weight: 1.25, arrow: true }),
    label(shapes, text, (x1 + x2) / 2, (y1 + y2) / 2 - 14, { bubble: true }),
  ];
  return group(context, g, parts, `cagr|${a}|${b}`);
}

function drawTotals(context: Excel.RequestContext, g: ChartGeometry) {
  const shapes = g.sheet.shapes;
  const parts = g.totals.map((t, i) =>
    label(shapes, shortNumber(t), categoryX(g.box, i, g.n), valueY(g.box, g.tops[i]) - 10, {
      plain: true,
    })
  );
  return group(context, g, parts, "totals");
}

export async function addDifferenceArrow(
  context: Excel.RequestContext,
  from1: number,
  to1: number,
  mode: DiffMode
) {
  const g = await geometry(context, await activeChart(context));
  await drawDiff(context, g, from1 - 1, to1 - 1, mode);
  await context.sync();
}

export async function addCagrArrow(
  context: Excel.RequestContext,
  from1?: number | null,
  to1?: number | null
) {
  const g = await geometry(context, await activeChart(context));
  await drawCagr(context, g, (from1 ?? 1) - 1, (to1 ?? g.n) - 1);
  await context.sync();
}

export async function addTotals(context: Excel.RequestContext) {
  const g = await geometry(context, await activeChart(context));
  await drawTotals(context, g);
  await context.sync();
}

async function taggedShapes(context: Excel.RequestContext): Promise<Excel.Shape[]> {
  const shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
  shapes.load("items/name,items/left,items/top");
  await context.sync();
  return shapes.items.filter((s) => s.name.startsWith(TAG));
}

/** Deletes and redraws every ROLLSI annotation on the active sheet (after data or size changes). */
export async function refreshAnnotations(
  context: Excel.RequestContext,
  redrawMekko: (source: string, left: number, top: number) => Promise<void>
): Promise<number> {
  const tagged = await taggedShapes(context);
  const jobs = tagged.map((s) => ({
    parts: s.name.slice(TAG.length).split("|"),
    left: s.left,
    top: s.top,
  }));
  tagged.forEach((s) => s.delete());
  await context.sync();

  const sheet = context.workbook.worksheets.getActiveWorksheet();
  for (const job of jobs) {
    const [owner, kind, ...args] = job.parts;
    if (owner === "MEKKO") {
      await redrawMekko([kind, ...args].join("|"), job.left, job.top);
      continue;
    }
    const chart = sheet.charts.getItemOrNullObject(owner);
    await context.sync();
    if (chart.isNullObject) continue;
    const g = await geometry(context, chart);
    if (kind === "diff")
      await drawDiff(context, g, Number(args[0]), Number(args[1]), args[2] as DiffMode);
    if (kind === "cagr") await drawCagr(context, g, Number(args[0]), Number(args[1]));
    if (kind === "totals") await drawTotals(context, g);
    await context.sync();
  }
  return jobs.length;
}

/** Removes annotations of the selected chart, or all ROLLSI annotations on the sheet if no chart is selected. */
export async function removeAnnotations(context: Excel.RequestContext): Promise<number> {
  const chart = context.workbook.getActiveChartOrNullObject();
  chart.load("name");
  await context.sync();
  const prefix = chart.isNullObject ? TAG : `${TAG}${chart.name}|`;
  const tagged = (await taggedShapes(context)).filter(
    (s) => s.name.startsWith(prefix) && !s.name.startsWith(`${TAG}MEKKO|`)
  );
  tagged.forEach((s) => s.delete());
  await context.sync();
  return tagged.length;
}
