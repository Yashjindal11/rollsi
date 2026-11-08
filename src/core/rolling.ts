import type { Num } from "./series";
import { RollsiError, requirePositiveInt, valid } from "./series";
import type { Stat, StatSpec } from "./stats";
import { aggregate } from "./stats";

export type Align = "last" | "mid" | "first" | "custom";

/**
 * Where the current row sits in its window. Codes: 0 = last (trailing), 1 = mid (centered),
 * 2 = first (forward-looking), 3 = custom (explicit rows above / below).
 */
export function parseAlign(input: unknown): Align | null {
  if (input === null || input === undefined || input === "") return null;
  const key = String(input).trim().toLowerCase();
  switch (key) {
    case "0":
    case "last":
    case "trailing":
    case "end":
    case "right":
      return "last";
    case "1":
    case "mid":
    case "middle":
    case "center":
    case "centre":
    case "centered":
    case "centred":
      return "mid";
    case "2":
    case "first":
    case "leading":
    case "forward":
    case "start":
    case "left":
      return "first";
    case "3":
    case "custom":
      return "custom";
  }
  throw new RollsiError(
    `Unknown alignment "${input}". Use 0/last (trailing), 1/mid (centered), 2/first (forward) or 3/custom.`
  );
}

function requireNonNegativeInt(value: unknown, name: string): number {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
    throw new RollsiError(`${name} must be a whole number >= 0.`);
  }
  return value;
}

/**
 * Converts a window size + alignment into rows above / below the current row.
 * - last:   window-1 rows above (classic trailing moving average)
 * - mid:    centered; for even windows the extra row goes above (like pandas center=True)
 * - first:  window-1 rows below
 * - custom: exactly `above` rows above and `below` rows below; the window argument is ignored.
 * If no alignment is given but `above` or `below` is, custom is assumed.
 */
export function resolveWindow(
  window: number | null | undefined,
  align: Align | null,
  above?: number | null,
  below?: number | null
): { above: number; below: number } {
  const hasCustom =
    (above !== null && above !== undefined) || (below !== null && below !== undefined);
  const mode = align ?? (hasCustom ? "custom" : "last");
  if (mode === "custom") {
    if (!hasCustom) throw new RollsiError("Custom alignment needs rows_above and/or rows_below.");
    return {
      above: requireNonNegativeInt(above ?? 0, "rows_above"),
      below: requireNonNegativeInt(below ?? 0, "rows_below"),
    };
  }
  const w = requirePositiveInt(window, "window");
  if (mode === "mid") {
    const up = Math.floor(w / 2);
    return { above: up, below: w - 1 - up };
  }
  return mode === "first" ? { above: 0, below: w - 1 } : { above: w - 1, below: 0 };
}

export interface MovingOptions {
  above: number;
  below: number;
  stat: Stat | StatSpec;
  /** Compute rows whose window runs past the top/bottom of the data from what is available. */
  partial?: boolean;
  /** Which cells are non-empty (text included), for COUNTA. Defaults to "is a number". */
  present?: boolean[];
}

/**
 * General moving window: for each row, aggregates rows [i - above, i + below].
 * Blank cells inside a window are skipped. Rows whose window runs past the edge of the data are
 * blank unless `partial` is true.
 */
export function moving(values: Num[], opts: MovingOptions): Num[] {
  const { above, below, partial = false } = opts;
  const spec: StatSpec = typeof opts.stat === "string" ? { stat: opts.stat } : opts.stat;
  const n = values.length;
  const out: Num[] = new Array(n).fill(null);
  for (let i = 0; i < n; i++) {
    const lo = i - above;
    const hi = i + below;
    if (!partial && (lo < 0 || hi >= n)) continue;
    const from = Math.max(0, lo);
    const to = Math.min(n - 1, hi);
    if (spec.stat === "counta" && opts.present) {
      let c = 0;
      for (let j = from; j <= to; j++) if (opts.present[j]) c++;
      out[i] = c;
    } else {
      out[i] = aggregate(values.slice(from, to + 1), spec);
    }
  }
  return out;
}

/**
 * Trailing rolling window over positions (rows).
 * - Blank cells inside a window are skipped; the stat is computed on the numbers that are present.
 * - Until `window` rows are available the result is blank, unless `partial` is true.
 */
export function rolling(
  values: Num[],
  window: number,
  stat: Stat | StatSpec,
  partial = false
): Num[] {
  requirePositiveInt(window, "window");
  return moving(values, { above: window - 1, below: 0, stat, partial });
}

export function sma(values: Num[], window: number, partial = false): Num[] {
  return rolling(values, window, "mean", partial);
}

/**
 * Exponential moving average: ema = alpha * x + (1 - alpha) * prev, seeded with the first number
 * (same as pandas ewm(span, adjust=False)). alpha = 2 / (span + 1) unless given explicitly.
 * Blank cells return blank and do not update the average. The first span-1 rows are a warm-up
 * period and are blank unless `partial` is true.
 */
export function ema(values: Num[], span: number, partial = false, alpha?: number): Num[] {
  requirePositiveInt(span, "span");
  const a = alpha ?? 2 / (span + 1);
  if (!(a > 0 && a <= 1)) throw new RollsiError("alpha must be between 0 (exclusive) and 1.");
  const out: Num[] = new Array(values.length).fill(null);
  let prev: number | null = null;
  for (let i = 0; i < values.length; i++) {
    const x = values[i];
    if (x === null) continue;
    prev = prev === null ? x : a * x + (1 - a) * prev;
    if (partial || i >= span - 1) out[i] = prev;
  }
  return out;
}

/** Linearly weighted moving average: the newest row gets weight N, the oldest weight 1. */
export function wma(values: Num[], window: number, partial = false): Num[] {
  requirePositiveInt(window, "window");
  const out: Num[] = new Array(values.length).fill(null);
  for (let i = 0; i < values.length; i++) {
    if (!partial && i < window - 1) continue;
    let num = 0;
    let den = 0;
    for (let k = 0; k < window && i - k >= 0; k++) {
      const x = values[i - k];
      if (x === null) continue;
      const w = window - k;
      num += w * x;
      den += w;
    }
    out[i] = den === 0 ? null : num / den;
  }
  return out;
}

export type CumStat = "sum" | "mean" | "min" | "max" | "count";

/** Running (expanding) aggregate from the first row to the current row. Blanks stay blank. */
export function cumulative(values: Num[], stat: CumStat): Num[] {
  let sum = 0;
  let n = 0;
  let min = Infinity;
  let max = -Infinity;
  return values.map((x) => {
    if (x === null) return null;
    sum += x;
    n += 1;
    min = Math.min(min, x);
    max = Math.max(max, x);
    switch (stat) {
      case "sum":
        return sum;
      case "mean":
        return sum / n;
      case "min":
        return min;
      case "max":
        return max;
      case "count":
        return n;
    }
  });
}

/**
 * Aggregate of the last `n` numbers in the series, ignoring blanks (including trailing blank rows),
 * e.g. "average of the last 7 data points".
 */
export function last(values: Num[], n: number, stat: Stat | StatSpec): Num {
  requirePositiveInt(n, "n");
  const v = valid(values);
  return aggregate(v.slice(Math.max(0, v.length - n)), stat);
}
