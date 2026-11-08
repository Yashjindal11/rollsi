/* global Excel */
import { buildTrendMA, buildWaterfall } from "./charts";

export const DEMO_SHEET = "ROLLSI Demo";

/** Deterministic pseudo-random numbers so the demo looks the same every time. */
function rng(seed: number) {
  return () => {
    seed = (seed * 1664525 + 1013904223) % 4294967296;
    return seed / 4294967296;
  };
}

function demoData(): [number, number][] {
  const rand = rng(42);
  const start = 46023; // 2026-01-01
  const skipped = new Set([20, 21, 50, 77]); // missing days, to show ROLLDAYS vs MA
  const rows: [number, number][] = [];
  for (let d = 0; d < 120; d++) {
    if (skipped.has(d)) continue;
    const weekly = [0.85, 1.0, 1.05, 1.1, 1.2, 1.35, 0.9][d % 7];
    let sales = (1000 + d * 6) * weekly * (0.9 + rand() * 0.2);
    if (d === 60) sales *= 2.6;
    if (d === 95) sales *= 0.35;
    rows.push([start + d, Math.round(sales)]);
  }
  return rows;
}

const BRIDGE: [string, number | string][] = [
  ["Q1 revenue", 1200],
  ["Price", 180],
  ["Volume", 240],
  ["Mix", -90],
  ["Subtotal", ""],
  ["Cost inflation", -310],
  ["FX", 45],
  ["Q2 revenue", ""],
];

export async function createDemo(context: Excel.RequestContext) {
  const sheets = context.workbook.worksheets;
  const existing = sheets.getItemOrNullObject(DEMO_SHEET);
  await context.sync();
  if (!existing.isNullObject) existing.delete();
  const sheet = sheets.add(DEMO_SHEET);
  sheet.activate();

  const data = demoData();
  const last = data.length + 1;
  const B = `B2:B${last}`;
  const A = `A2:A${last}`;

  const columns: [string, string, string][] = [
    ["7-row avg (last point)", `=ROLLSI.MA(${B},7)`, "#,##0"],
    ["7-row avg (mid point)", `=ROLLSI.MA(${B},7,1,1)`, "#,##0"],
    ["7-row SUM (code 9)", `=ROLLSI.MA(${B},7,9)`, "#,##0"],
    ["MAX of 3 above + 3 below", `=ROLLSI.MA(${B},,4,3,3,3)`, "#,##0"],
    ["7-day avg/day (calendar)", `=ROLLSI.ROLLDAYS(${A},${B},7,"perday")`, "#,##0"],
    ["WoW %", `=ROLLSI.PERIODCHANGE(${A},${B},"W")`, "0.0%"],
    ["Anomaly score (14)", `=ROLLSI.ROLLZSCORE(${B},14)`, "0.00"],
    ["Outlier?", `=ROLLSI.OUTLIERS(${B},"mad")`, "General"],
  ];

  sheet.getRange("A1:J1").values = [["Date", "Sales", ...columns.map((c) => c[0])]];
  sheet.getRange(`A2:B${last}`).values = data;
  sheet.getRange(A).numberFormat = data.map(() => ["yyyy-mm-dd"]);
  sheet.getRange(B).numberFormat = data.map(() => ["#,##0"]);
  columns.forEach(([, formula, format], i) => {
    const col = String.fromCharCode(67 + i); // C, D, ...
    sheet.getRange(`${col}2`).formulas = [[formula]];
    sheet.getRange(`${col}2:${col}${last}`).numberFormat = data.map(() => [format]);
  });

  const recent = `B${last - 29}:B${last}`;
  sheet.getRange("L1:M1").values = [["KPI", "Value"]];
  sheet.getRange("L2:M9").formulas = [
    ["Last 30 days", `=ROLLSI.SPARK(${recent})`],
    ["Avg of last 7 points", `=ROLLSI.LAST(${B},7)`],
    ["Week over week", `=ROLLSI.ARROW(ROLLSI.LAST(${B},7)/ROLLSI.LAST(B2:B${last - 7},7)-1)`],
    ["Sum of last 7 points", `=ROLLSI.SHORT(ROLLSI.LAST(${B},7,"sum"))`],
    ["Median of last 30 points", `=ROLLSI.LAST(${B},30,"median")`],
    ["Volatility (stdev, last 30)", `=ROLLSI.LAST(${B},30,"stdev")`],
    ["Daily growth trend (CAGR)", `=ROLLSI.CAGR(B2,B${last},${data.length - 1})`],
    ["Best day vs max", `=ROLLSI.BAR(MAX(${recent}),MAX(${B}))`],
  ];
  sheet.getRange("M3").numberFormat = [["#,##0"]];
  sheet.getRange("M6:M7").numberFormat = [["#,##0"], ["#,##0"]];
  sheet.getRange("M8").numberFormat = [["0.00%"]];
  sheet.getRange("L11").values = [["Next 14 days (Holt forecast)"]];
  sheet.getRange("L12").formulas = [[`=ROLLSI.FORECAST(${B},14)`]];
  sheet.getRange("L12:L25").numberFormat = Array.from({ length: 14 }, () => ["#,##0"]);

  sheet.getRange("L28:M28").values = [["Revenue bridge", "Change"]];
  sheet.getRange(`L29:M${28 + BRIDGE.length}`).values = BRIDGE;

  [sheet.getRange("A1:J1"), sheet.getRange("L1:M1"), sheet.getRange("L28:M28")].forEach((h) => {
    h.format.font.bold = true;
    h.format.fill.color = "#107C41";
    h.format.font.color = "#FFFFFF";
    h.format.wrapText = true;
  });
  sheet.getRange("L11").format.font.bold = true;
  sheet.freezePanes.freezeRows(1);
  sheet.getRange("A:M").format.autofitColumns();
  sheet.getRange("M2").format.font.color = "#107C41";
  await context.sync();

  const anchor = sheet.getRange("O2");
  anchor.load("left,top");
  await context.sync();

  sheet.getRange(`A1:B${last}`).select();
  const trend = await buildTrendMA(context, 7);
  trend.left = anchor.left;
  trend.top = anchor.top;

  sheet.getRange(`L28:M${28 + BRIDGE.length}`).select();
  const bridge = await buildWaterfall(context, true, true);
  bridge.left = anchor.left;
  bridge.top = anchor.top + 320;

  sheet.getRange("A1").select();
  await context.sync();
}
