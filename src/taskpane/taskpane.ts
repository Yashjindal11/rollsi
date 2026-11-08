/* global document, Excel, HTMLButtonElement, HTMLInputElement, Office */
import { CATALOG } from "./catalog";
import type { FunctionDoc } from "./catalog";

const DEMO_SHEET = "ROLLSI Demo";

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!
  );
}

function render(filter: string) {
  const list = document.getElementById("list")!;
  const q = filter.trim().toLowerCase();
  const matches = (f: FunctionDoc) =>
    !q || [f.name, f.group, f.summary, f.syntax].some((s) => s.toLowerCase().includes(q));
  const groups = new Map<string, FunctionDoc[]>();
  CATALOG.filter(matches).forEach((f) => groups.set(f.group, [...(groups.get(f.group) ?? []), f]));
  list.innerHTML = groups.size
    ? [...groups]
        .map(
          ([group, fns]) =>
            `<h2>${escapeHtml(group)}</h2>` +
            fns
              .map(
                (f) =>
                  `<div class="card"><div class="syntax">${escapeHtml(f.syntax)}</div>` +
                  `<p>${escapeHtml(f.summary)}</p><code>${escapeHtml(f.example)}</code></div>`
              )
              .join("")
        )
        .join("")
    : "<p>No matching functions.</p>";
}

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
  const skipped = new Set([20, 21, 50, 77]); // missing days, to show ROLLDAYS vs SMA
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

async function createDemo() {
  const button = document.getElementById("demo") as HTMLButtonElement;
  const status = document.getElementById("status")!;
  button.disabled = true;
  status.className = "";
  status.textContent = "Building demo sheet...";
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const existing = sheets.getItemOrNullObject(DEMO_SHEET);
      await context.sync();
      if (!existing.isNullObject) existing.delete();
      const sheet = sheets.add(DEMO_SHEET);

      const data = demoData();
      const last = data.length + 1;
      const B = `B2:B${last}`;
      const A = `A2:A${last}`;

      const columns: [string, string, string][] = [
        ["7-row SMA", `=ROLLSI.SMA(${B},7)`, "#,##0"],
        ["14-row SMA", `=ROLLSI.SMA(${B},14)`, "#,##0"],
        ["7-row EMA", `=ROLLSI.EMA(${B},7)`, "#,##0"],
        ["7-day sum (calendar)", `=ROLLSI.ROLLDAYS(${A},${B},7,"sum")`, "#,##0"],
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

      sheet.getRange("L1:M1").values = [["KPI", "Value"]];
      sheet.getRange("L2:M6").formulas = [
        ["Avg of last 7 points", `=ROLLSI.LAST(${B},7)`],
        ["Sum of last 7 points", `=ROLLSI.LAST(${B},7,"sum")`],
        ["Median of last 30 points", `=ROLLSI.LAST(${B},30,"median")`],
        ["Volatility (stdev, last 30)", `=ROLLSI.LAST(${B},30,"stdev")`],
        ["Daily growth trend (CAGR)", `=ROLLSI.CAGR(B2,B${last},${data.length - 1})`],
      ];
      sheet.getRange("M2:M5").numberFormat = [["#,##0"], ["#,##0"], ["#,##0"], ["#,##0"]];
      sheet.getRange("M6").numberFormat = [["0.00%"]];
      sheet.getRange("L8").values = [["Next 14 days (Holt forecast)"]];
      sheet.getRange("L9").formulas = [[`=ROLLSI.FORECAST(${B},14)`]];
      sheet.getRange("L9:L22").numberFormat = Array.from({ length: 14 }, () => ["#,##0"]);

      const header = sheet.getRange("A1:J1");
      header.format.font.bold = true;
      header.format.fill.color = "#107C41";
      header.format.font.color = "#FFFFFF";
      header.format.wrapText = true;
      const kpiHeader = sheet.getRange("L1:M1");
      kpiHeader.format.font.bold = true;
      kpiHeader.format.fill.color = "#107C41";
      kpiHeader.format.font.color = "#FFFFFF";
      sheet.getRange("L8").format.font.bold = true;
      sheet.freezePanes.freezeRows(1);

      const chart = sheet.charts.add(
        Excel.ChartType.line,
        sheet.getRange(`A1:D${last}`),
        Excel.ChartSeriesBy.columns
      );
      chart.title.text = "Daily sales with 7- and 14-row moving averages";
      chart.legend.position = Excel.ChartLegendPosition.bottom;
      chart.setPosition("O2", "X22");

      sheet.getRange("A:M").format.autofitColumns();
      sheet.activate();
      await context.sync();
    });
    status.textContent = `Done. See the "${DEMO_SHEET}" sheet.`;
  } catch (e) {
    status.className = "error";
    status.textContent = `Could not create the demo: ${e instanceof Error ? e.message : String(e)}`;
  } finally {
    button.disabled = false;
  }
}

Office.onReady(() => {
  render("");
  document
    .getElementById("search")!
    .addEventListener("input", (e) => render((e.target as HTMLInputElement).value));
  document.getElementById("demo")!.addEventListener("click", createDemo);
});
