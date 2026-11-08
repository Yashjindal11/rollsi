import type { Num } from "./series";
import { RollsiError, requirePositiveInt, valid } from "./series";
import { mean, quantile, median, stdev } from "./stats";

/** Standard score of each value against the whole series (sample stdev). */
export function zscore(values: Num[]): Num[] {
  const v = valid(values);
  const m = mean(v);
  const s = stdev(v);
  return values.map((x) => (x === null || m === null || !s ? null : (x - m) / s));
}

/**
 * Rolling anomaly score: how many standard deviations each value sits from the mean of the
 * previous `window` rows (the current row is excluded so a spike cannot mask itself).
 */
export function rollingZscore(values: Num[], window: number): Num[] {
  requirePositiveInt(window, "window");
  return values.map((x, i) => {
    if (x === null || i < window) return null;
    const prev = valid(values.slice(i - window, i));
    const m = mean(prev);
    const s = stdev(prev);
    return m === null || !s ? null : (x - m) / s;
  });
}

export type OutlierMethod = "iqr" | "z" | "mad";

export function parseOutlierMethod(input: unknown): OutlierMethod {
  if (input === null || input === undefined || input === "") return "iqr";
  const key = String(input).trim().toLowerCase();
  if (key === "iqr" || key === "z" || key === "mad") return key;
  if (key === "zscore") return "z";
  throw new RollsiError(`Unknown method "${input}". Use iqr, z or mad.`);
}

const DEFAULT_THRESHOLD: Record<OutlierMethod, number> = { iqr: 1.5, z: 3, mad: 3.5 };

/**
 * TRUE/FALSE outlier flags.
 * - iqr: outside [Q1 - k*IQR, Q3 + k*IQR] (k defaults to 1.5, the Tukey fence)
 * - z:   |z-score| > k (default 3)
 * - mad: |0.6745 * (x - median) / MAD| > k (default 3.5), robust to the outliers themselves
 */
export function outliers(
  values: Num[],
  method: OutlierMethod = "iqr",
  threshold?: number
): (boolean | null)[] {
  const k = threshold ?? DEFAULT_THRESHOLD[method];
  if (!(k > 0)) throw new RollsiError("threshold must be > 0.");
  const v = valid(values);
  if (v.length === 0) return values.map(() => null);

  let isOut: (x: number) => boolean;
  if (method === "iqr") {
    const q1 = quantile(v, 0.25) as number;
    const q3 = quantile(v, 0.75) as number;
    const iqr = q3 - q1;
    isOut = (x) => x < q1 - k * iqr || x > q3 + k * iqr;
  } else if (method === "z") {
    const m = mean(v) as number;
    const s = stdev(v);
    isOut = (x) => !!s && Math.abs((x - m) / s) > k;
  } else {
    const med = median(v) as number;
    const mad = median(v.map((x) => Math.abs(x - med))) as number;
    isOut = (x) => mad > 0 && Math.abs((0.6745 * (x - med)) / mad) > k;
  }
  return values.map((x) => (x === null ? null : isOut(x)));
}

/** Percentile rank in [0, 1], matching Excel PERCENTRANK.INC. */
export function pctRank(values: Num[]): Num[] {
  const sorted = valid(values).sort((a, b) => a - b);
  const n = sorted.length;
  return values.map((x) => {
    if (x === null) return null;
    if (n === 1) return 1;
    let lo = 0;
    let hi = n;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (sorted[mid] < x) lo = mid + 1;
      else hi = mid;
    }
    return lo / (n - 1);
  });
}

export type NormalizeMethod = "minmax" | "z" | "share";

export function parseNormalizeMethod(input: unknown): NormalizeMethod {
  if (input === null || input === undefined || input === "") return "minmax";
  const key = String(input).trim().toLowerCase();
  if (key === "minmax" || key === "z" || key === "share") return key;
  if (key === "zscore") return "z";
  if (key === "pct" || key === "percent" || key === "total") return "share";
  throw new RollsiError(`Unknown method "${input}". Use minmax, z or share.`);
}

/** Scales values: minmax -> 0..1, z -> z-score, share -> fraction of the total. */
export function normalize(values: Num[], method: NormalizeMethod = "minmax"): Num[] {
  if (method === "z") return zscore(values);
  const v = valid(values);
  if (v.length === 0) return values.map(() => null);
  if (method === "share") {
    const total = v.reduce((a, b) => a + b, 0);
    return values.map((x) => (x === null || total === 0 ? null : x / total));
  }
  const min = v.reduce((a, b) => (b < a ? b : a));
  const range = v.reduce((a, b) => (b > a ? b : a)) - min;
  return values.map((x) => (x === null ? null : range === 0 ? 0 : (x - min) / range));
}
