import type { Num } from "./series";
import { RollsiError, requirePositiveInt, valid } from "./series";
import type { Stat } from "./stats";
import { aggregate } from "./stats";

/**
 * Trailing rolling window over positions (rows).
 * - Blank cells inside a window are skipped; the stat is computed on the numbers that are present.
 * - Until `window` rows are available the result is blank, unless `partial` is true.
 */
export function rolling(values: Num[], window: number, stat: Stat, partial = false): Num[] {
  requirePositiveInt(window, "window");
  const out: Num[] = new Array(values.length).fill(null);
  for (let i = 0; i < values.length; i++) {
    if (!partial && i < window - 1) continue;
    const start = Math.max(0, i - window + 1);
    const slice = values.slice(start, i + 1);
    out[i] = valid(slice).length === 0 && stat !== "count" ? null : aggregate(slice, stat);
  }
  return out;
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
export function last(values: Num[], n: number, stat: Stat): Num {
  requirePositiveInt(n, "n");
  const v = valid(values);
  return aggregate(v.slice(Math.max(0, v.length - n)), stat);
}
