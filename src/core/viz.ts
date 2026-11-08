import type { Num } from "./series";
import { RollsiError, valid } from "./series";
import { mean, stdev } from "./stats";

// ---------------------------------------------------------------- Number formatting

/** Excel number format that shows 1.2K / 3.4M (Excel allows two conditions per format). */
export const SHORT_NUMBER_FORMAT = '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0';

/** 1234 -> "1.2K", 5600000 -> "5.6M", -2500000000 -> "-2.5B". */
export function shortNumber(value: number, decimals = 1): string {
  if (!Number.isFinite(value)) return "";
  const abs = Math.abs(value);
  const units: [number, string][] = [
    [1e12, "T"],
    [1e9, "B"],
    [1e6, "M"],
    [1e3, "K"],
  ];
  const [div, suffix] = units.find(([d]) => abs >= d) ?? [1, ""];
  const scaled = (abs / div).toFixed(Math.max(0, decimals));
  const trimmed = scaled.includes(".") ? scaled.replace(/\.?0+$/, "") : scaled;
  return (value < 0 ? "-" : "") + trimmed + suffix;
}

export function formatPct(value: number, decimals = 1, signed = true): string {
  const s = (value * 100).toFixed(decimals) + "%";
  return signed && value > 0 ? "+" + s : s;
}

// ---------------------------------------------------------------- In-cell charts

const BLOCKS = "▁▂▃▄▅▆▇█";
const EIGHTHS = "▏▎▍▌▋▊▉";

export type SparkType = "bars" | "winloss";

export function parseSparkType(input: unknown): SparkType {
  if (input === null || input === undefined || input === "") return "bars";
  const key = String(input).trim().toLowerCase();
  if (key === "bars" || key === "bar" || key === "line" || key === "column") return "bars";
  if (key === "winloss" || key === "win-loss" || key === "wl") return "winloss";
  throw new RollsiError(`Unknown sparkline type "${input}". Use bars or winloss.`);
}

/** Text sparkline such as ▁▂▄▆█. Blanks become spaces. */
export function sparkline(
  values: Num[],
  type: SparkType = "bars",
  min?: number | null,
  max?: number | null
): string {
  if (type === "winloss") {
    return values.map((v) => (v === null ? " " : v > 0 ? "▀" : v < 0 ? "▄" : "─")).join("");
  }
  const v = valid(values);
  if (v.length === 0) return "";
  const lo = min ?? v.reduce((a, b) => (b < a ? b : a));
  const hi = max ?? v.reduce((a, b) => (b > a ? b : a));
  return values
    .map((x) => {
      if (x === null) return " ";
      if (hi === lo) return BLOCKS[3];
      const t = Math.min(1, Math.max(0, (x - lo) / (hi - lo)));
      return BLOCKS[Math.round(t * (BLOCKS.length - 1))];
    })
    .join("");
}

/** In-cell horizontal bar with 1/8-character precision. Negative values get a leading minus. */
export function textBar(value: number, max: number, width = 10): string {
  if (!(max > 0)) throw new RollsiError("max must be > 0.");
  if (!Number.isInteger(width) || width < 1 || width > 255) {
    throw new RollsiError("width must be a whole number between 1 and 255.");
  }
  const cells = Math.min(1, Math.abs(value) / max) * width;
  let full = Math.floor(cells);
  let rem = Math.round((cells - full) * 8);
  if (rem === 8) {
    full += 1;
    rem = 0;
  }
  const bar = "█".repeat(full) + (rem > 0 ? EIGHTHS[rem - 1] : "");
  return value < 0 && bar ? "-" + bar : bar;
}

export type ArrowFormat = "pct" | "num" | "arrow";

export function parseArrowFormat(input: unknown): ArrowFormat {
  if (input === null || input === undefined || input === "") return "pct";
  const key = String(input).trim().toLowerCase();
  if (key === "pct" || key === "%" || key === "percent") return "pct";
  if (key === "num" || key === "number" || key === "abs") return "num";
  if (key === "arrow" || key === "none") return "arrow";
  throw new RollsiError(`Unknown format "${input}". Use pct, num or arrow.`);
}

/** ▲ +12.3% / ▼ -4.0% / ► 0.0%. Changes within ±threshold count as flat. */
export function arrow(
  change: number,
  threshold = 0,
  format: ArrowFormat = "pct",
  decimals = 1
): string {
  const symbol = change > threshold ? "▲" : change < -threshold ? "▼" : "►";
  if (format === "arrow") return symbol;
  if (format === "num") return `${symbol} ${change > 0 ? "+" : ""}${shortNumber(change, decimals)}`;
  return `${symbol} ${formatPct(change, decimals)}`;
}

// ---------------------------------------------------------------- Chart data

export interface ControlLimits {
  center: Num[];
  upper: Num[];
  lower: Num[];
  outlier: Num[];
}

/**
 * Control-chart lines: center ± k standard deviations. With a window, limits are computed from the
 * previous `window` rows (rolling SPC); without, from the whole series. Outliers keep their value,
 * everything else is null (so charts show only the flagged points).
 */
export function controlLimits(values: Num[], window?: number | null, k = 3): ControlLimits {
  if (!(k > 0)) throw new RollsiError("k must be > 0.");
  if (window !== null && window !== undefined && (!Number.isInteger(window) || window < 2)) {
    throw new RollsiError("window must be a whole number >= 2.");
  }
  const n = values.length;
  const out: ControlLimits = { center: [], upper: [], lower: [], outlier: [] };
  const all = valid(values);
  const gm = mean(all);
  const gs = stdev(all);
  for (let i = 0; i < n; i++) {
    let m: Num = gm;
    let s: Num = gs;
    if (window) {
      if (i < window) {
        m = null;
        s = null;
      } else {
        const prev = valid(values.slice(i - window, i));
        m = mean(prev);
        s = stdev(prev);
      }
    }
    const x = values[i];
    const ok = m !== null && s !== null;
    out.center.push(ok ? m : null);
    out.upper.push(ok ? (m as number) + k * (s as number) : null);
    out.lower.push(ok ? (m as number) - k * (s as number) : null);
    out.outlier.push(
      ok && x !== null && Math.abs(x - (m as number)) > k * (s as number) ? x : null
    );
  }
  return out;
}

export interface ParetoRow {
  label: unknown;
  value: number;
  cumulative: number;
}

/** Sorts descending and adds the cumulative share. Blank or negative values are dropped. */
export function pareto(labels: unknown[], values: Num[]): ParetoRow[] {
  if (labels.length !== values.length)
    throw new RollsiError("labels and values must be the same size.");
  const rows = values
    .map((v, i) => ({ label: labels[i], value: v }))
    .filter((r): r is { label: unknown; value: number } => r.value !== null && r.value >= 0)
    .sort((a, b) => b.value - a.value);
  const total = rows.reduce((a, r) => a + r.value, 0);
  let run = 0;
  return rows.map((r) => {
    run += r.value;
    return { ...r, cumulative: total === 0 ? 0 : run / total };
  });
}

/** Splits a floating bar [from, to] into stacked-column parts (Excel stacks + and - separately). */
export function splitBar(
  from: number,
  to: number
): { basePos: number; baseNeg: number; pos: number; neg: number } {
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  if (lo >= 0) return { basePos: lo, baseNeg: 0, pos: hi - lo, neg: 0 };
  if (hi <= 0) return { basePos: 0, baseNeg: hi, pos: 0, neg: lo - hi };
  return { basePos: 0, baseNeg: 0, pos: hi, neg: lo };
}

export type StepKind = "start" | "up" | "down" | "total";

export interface WaterfallStep {
  kind: StepKind;
  from: number;
  to: number;
  /** Number to print on the bar: the change for up/down, the level for start/total. */
  label: number;
}

const TOTAL_LABEL = /\b(sub-?)?totals?\b/i;

/**
 * Bridge / waterfall steps. The first row is the starting level (if firstIsStart), rows whose label
 * contains "total" or "subtotal" show the running total (their value is ignored), and the last row
 * is a computed total if lastIsTotal.
 */
export function waterfall(
  labels: unknown[],
  values: Num[],
  firstIsStart = true,
  lastIsTotal = true
): WaterfallStep[] {
  if (labels.length !== values.length)
    throw new RollsiError("labels and values must be the same size.");
  let level = 0;
  return values.map((v, i) => {
    const isTotal =
      (lastIsTotal && i === values.length - 1 && i > 0) ||
      TOTAL_LABEL.test(String(labels[i] ?? ""));
    if (isTotal) return { kind: "total", from: 0, to: level, label: level };
    if (i === 0 && firstIsStart) {
      level = v ?? 0;
      return { kind: "start", from: 0, to: level, label: level };
    }
    const d = v ?? 0;
    const step: WaterfallStep = {
      kind: d >= 0 ? "up" : "down",
      from: level,
      to: level + d,
      label: d,
    };
    level += d;
    return step;
  });
}

export const WATERFALL_SERIES = [
  "_base+",
  "_base-",
  "Increase",
  "Increase (-)",
  "Decrease",
  "Decrease (-)",
  "Total",
  "Total (-)",
] as const;

/** One row per step with the 8 stacked-series values named in WATERFALL_SERIES. */
export function waterfallMatrix(steps: WaterfallStep[]): number[][] {
  return steps.map((s) => {
    const p = splitBar(s.from, s.to);
    const row = [p.basePos, p.baseNeg, 0, 0, 0, 0, 0, 0];
    const slot = s.kind === "up" ? 2 : s.kind === "down" ? 4 : 6;
    row[slot] = p.pos;
    row[slot + 1] = p.neg;
    return row;
  });
}

export interface MekkoCell {
  seg: number;
  cat: number;
  x: number;
  w: number;
  y: number;
  h: number;
  value: number;
  share: number;
}

export interface MekkoLayout {
  columns: { x: number; w: number; total: number; share: number }[];
  cells: MekkoCell[];
  grandTotal: number;
}

/**
 * Marimekko layout in a unit square: column widths are each category's share of the grand total,
 * cell heights are each segment's share within its category. Segment 0 sits at the bottom (y = 0).
 * values[seg][cat] must be >= 0; blanks count as 0.
 */
export function mekkoLayout(values: Num[][]): MekkoLayout {
  const segs = values.length;
  const cats = segs ? values[0].length : 0;
  if (!segs || !cats)
    throw new RollsiError("Mekko needs at least one segment row and one category column.");
  const v = values.map((row) =>
    row.map((x) => {
      if (x !== null && x < 0) throw new RollsiError("Mekko values must be >= 0.");
      return x ?? 0;
    })
  );
  const totals = Array.from({ length: cats }, (_, c) => v.reduce((a, row) => a + row[c], 0));
  const grandTotal = totals.reduce((a, b) => a + b, 0);
  if (grandTotal === 0) throw new RollsiError("Mekko values add up to 0.");
  const columns: MekkoLayout["columns"] = [];
  const cells: MekkoCell[] = [];
  let x = 0;
  totals.forEach((total, c) => {
    const w = total / grandTotal;
    columns.push({ x, w, total, share: w });
    let y = 0;
    for (let s = 0; s < segs; s++) {
      const share = total === 0 ? 0 : v[s][c] / total;
      cells.push({ seg: s, cat: c, x, w, y, h: share, value: v[s][c], share });
      y += share;
    }
    x += w;
  });
  return { columns, cells, grandTotal };
}

// ---------------------------------------------------------------- Annotation geometry

export interface PlotBox {
  left: number;
  top: number;
  width: number;
  height: number;
  min: number;
  max: number;
}

/** Horizontal centre (points) of category i of n in a column/line chart. */
export function categoryX(box: PlotBox, i: number, n: number): number {
  return box.left + ((i + 0.5) * box.width) / n;
}

/** Vertical position (points) of a value on the value axis. */
export function valueY(box: PlotBox, v: number): number {
  const range = box.max - box.min || 1;
  const t = Math.min(1, Math.max(0, (v - box.min) / range));
  return box.top + box.height * (1 - t);
}

export type DiffMode = "pct" | "abs" | "both";

export function diffLabel(a: number, b: number, mode: DiffMode = "pct"): string {
  const abs = `${b - a > 0 ? "+" : ""}${shortNumber(b - a)}`;
  if (mode === "abs" || a === 0) return abs;
  const pct = formatPct((b - a) / Math.abs(a), 0);
  return mode === "both" ? `${abs} (${pct})` : pct;
}

export function cagrLabel(a: number, b: number, periods: number): string {
  if (!(a > 0) || !(b >= 0) || !(periods > 0)) return "CAGR n/a";
  return `CAGR ${formatPct(Math.pow(b / a, 1 / periods) - 1, 1)}`;
}

/** Actual vs target variance split into an "over" and an "under" series (null where not applicable). */
export function varianceSplit(actual: Num[], target: Num[]): { over: Num[]; under: Num[] } {
  if (actual.length !== target.length)
    throw new RollsiError("actual and target must be the same size.");
  const over: Num[] = [];
  const under: Num[] = [];
  actual.forEach((a, i) => {
    const t = target[i];
    const d = a === null || t === null ? null : a - t;
    over.push(d !== null && d >= 0 ? d : null);
    under.push(d !== null && d < 0 ? d : null);
  });
  return { over, under };
}
