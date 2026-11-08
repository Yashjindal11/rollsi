import type { Num } from "./series";
import { RollsiError, valid } from "./series";

export type Stat =
  "mean" | "sum" | "min" | "max" | "median" | "stdev" | "var" | "count" | "product";

const STAT_ALIASES: Record<string, Stat> = {
  mean: "mean",
  avg: "mean",
  average: "mean",
  sum: "sum",
  total: "sum",
  min: "min",
  max: "max",
  median: "median",
  stdev: "stdev",
  std: "stdev",
  sd: "stdev",
  var: "var",
  variance: "var",
  count: "count",
  n: "count",
  product: "product",
  prod: "product",
};

export function parseStat(input: unknown, fallback: Stat = "mean"): Stat {
  if (input === null || input === undefined || input === "") return fallback;
  const key = String(input).trim().toLowerCase();
  const stat = STAT_ALIASES[key];
  if (!stat) {
    throw new RollsiError(
      `Unknown stat "${input}". Use avg, sum, min, max, median, stdev, var, count or product.`
    );
  }
  return stat;
}

export function mean(xs: number[]): Num {
  if (xs.length === 0) return null;
  let s = 0;
  for (const x of xs) s += x;
  return s / xs.length;
}

/** Sample variance (n-1), matching Excel VAR.S. */
export function variance(xs: number[]): Num {
  if (xs.length < 2) return null;
  const m = mean(xs) as number;
  let ss = 0;
  for (const x of xs) ss += (x - m) ** 2;
  return ss / (xs.length - 1);
}

/** Sample standard deviation (n-1), matching Excel STDEV.S. */
export function stdev(xs: number[]): Num {
  const v = variance(xs);
  return v === null ? null : Math.sqrt(v);
}

/** Linear-interpolated quantile, matching Excel PERCENTILE.INC / QUARTILE.INC. */
export function quantile(xs: number[], q: number): Num {
  if (xs.length === 0) return null;
  const sorted = [...xs].sort((a, b) => a - b);
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

export function median(xs: number[]): Num {
  return quantile(xs, 0.5);
}

export function aggregate(xs: Num[] | number[], stat: Stat): Num {
  const v = valid(xs as Num[]);
  switch (stat) {
    case "count":
      return v.length;
    case "sum":
      return v.length ? v.reduce((a, b) => a + b, 0) : null;
    case "mean":
      return mean(v);
    case "min":
      return v.length ? v.reduce((a, b) => (b < a ? b : a)) : null;
    case "max":
      return v.length ? v.reduce((a, b) => (b > a ? b : a)) : null;
    case "median":
      return median(v);
    case "stdev":
      return stdev(v);
    case "var":
      return variance(v);
    case "product":
      return v.length ? v.reduce((a, b) => a * b, 1) : null;
  }
}
