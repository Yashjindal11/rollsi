import type { Num } from "./series";
import { RollsiError, valid } from "./series";

export type Stat =
  | "mean"
  | "count"
  | "counta"
  | "max"
  | "min"
  | "product"
  | "stdev"
  | "stdevp"
  | "sum"
  | "var"
  | "varp"
  | "median"
  | "mode"
  | "large"
  | "small"
  | "percentile"
  | "quartile"
  | "percentile_exc"
  | "quartile_exc"
  | "first"
  | "last"
  | "range"
  | "geomean"
  | "harmean"
  | "avedev"
  | "skew"
  | "kurt";

export interface StatSpec {
  stat: Stat;
  /** Parameter for large, small, percentile and quartile stats. */
  k?: number;
}

/** Excel AGGREGATE function_num codes 1-19, in order. */
export const AGGREGATE_CODES: Stat[] = [
  "mean",
  "count",
  "counta",
  "max",
  "min",
  "product",
  "stdev",
  "stdevp",
  "sum",
  "var",
  "varp",
  "median",
  "mode",
  "large",
  "small",
  "percentile",
  "quartile",
  "percentile_exc",
  "quartile_exc",
];

const NEEDS_K = new Set<Stat>([
  "large",
  "small",
  "percentile",
  "quartile",
  "percentile_exc",
  "quartile_exc",
]);

const STAT_ALIASES: Record<string, Stat> = {
  mean: "mean",
  avg: "mean",
  average: "mean",
  count: "count",
  n: "count",
  counta: "counta",
  max: "max",
  min: "min",
  product: "product",
  prod: "product",
  stdev: "stdev",
  "stdev.s": "stdev",
  std: "stdev",
  sd: "stdev",
  stdevp: "stdevp",
  "stdev.p": "stdevp",
  sum: "sum",
  total: "sum",
  var: "var",
  "var.s": "var",
  variance: "var",
  varp: "varp",
  "var.p": "varp",
  median: "median",
  mode: "mode",
  "mode.sngl": "mode",
  large: "large",
  small: "small",
  percentile: "percentile",
  "percentile.inc": "percentile",
  pct: "percentile",
  quartile: "quartile",
  "quartile.inc": "quartile",
  "percentile.exc": "percentile_exc",
  percentile_exc: "percentile_exc",
  "quartile.exc": "quartile_exc",
  quartile_exc: "quartile_exc",
  first: "first",
  last: "last",
  range: "range",
  geomean: "geomean",
  harmean: "harmean",
  avedev: "avedev",
  skew: "skew",
  kurt: "kurt",
};

const STAT_HELP =
  "Use a code 1-19 (Excel AGGREGATE: 1 AVERAGE, 2 COUNT, 3 COUNTA, 4 MAX, 5 MIN, 6 PRODUCT, 7 STDEV.S, 8 STDEV.P, " +
  "9 SUM, 10 VAR.S, 11 VAR.P, 12 MEDIAN, 13 MODE, 14 LARGE, 15 SMALL, 16 PERCENTILE.INC, 17 QUARTILE.INC, " +
  "18 PERCENTILE.EXC, 19 QUARTILE.EXC) or a name such as avg, sum, first, last, range, geomean, harmean, avedev, skew, kurt.";

export function statNeedsK(stat: Stat): boolean {
  return NEEDS_K.has(stat);
}

function checkK(stat: Stat, k: number | undefined): number | undefined {
  if (!NEEDS_K.has(stat)) return undefined;
  if (typeof k !== "number" || !Number.isFinite(k)) {
    throw new RollsiError(
      `"${stat}" needs k, e.g. "${stat}:2" as the stat, or pass the k argument.`
    );
  }
  const bad = (rule: string) => new RollsiError(`k for "${stat}" must be ${rule}.`);
  switch (stat) {
    case "large":
    case "small":
      if (!Number.isInteger(k) || k < 1) throw bad("a whole number >= 1");
      break;
    case "percentile":
      if (k < 0 || k > 1) throw bad("between 0 and 1");
      break;
    case "percentile_exc":
      if (k <= 0 || k >= 1) throw bad("between 0 and 1 (exclusive)");
      break;
    case "quartile":
      if (!Number.isInteger(k) || k < 0 || k > 4) throw bad("0, 1, 2, 3 or 4");
      break;
    case "quartile_exc":
      if (!Number.isInteger(k) || k < 1 || k > 3) throw bad("1, 2 or 3");
      break;
  }
  return k;
}

/**
 * Parses a stat given as an Excel AGGREGATE code (1-19), a name ("avg", "sum", "stdev.p"...),
 * a name with a parameter ("large:2", "percentile:0.9") or a percentile shorthand ("p90").
 * Blank or 0 means the fallback (average).
 */
export function parseStat(input: unknown, k?: number | null, fallback: Stat = "mean"): StatSpec {
  const kArg = k === null ? undefined : k;
  if (input === null || input === undefined || input === "" || input === 0)
    return { stat: fallback };

  let stat: Stat | undefined;
  let embeddedK: number | undefined;
  const text = String(input).trim().toLowerCase();

  if (typeof input === "number" || /^\d+$/.test(text)) {
    const code = Number(text);
    stat = Number.isInteger(code) ? AGGREGATE_CODES[code - 1] : undefined;
    if (!stat) throw new RollsiError(`Unknown stat code ${input}. ${STAT_HELP}`);
  } else {
    const pct = /^p(\d+(?:\.\d+)?)$/.exec(text);
    const withParam = /^([a-z._]+)\s*[:(]\s*(-?\d+(?:\.\d+)?)\s*\)?$/.exec(text);
    if (pct) {
      stat = "percentile";
      embeddedK = Number(pct[1]) / 100;
    } else if (withParam) {
      stat = STAT_ALIASES[withParam[1]];
      embeddedK = Number(withParam[2]);
    } else {
      stat = STAT_ALIASES[text];
    }
    if (!stat) throw new RollsiError(`Unknown stat "${input}". ${STAT_HELP}`);
  }
  return { stat, k: checkK(stat, embeddedK ?? kArg) };
}

export function mean(xs: number[]): Num {
  if (xs.length === 0) return null;
  let s = 0;
  for (const x of xs) s += x;
  return s / xs.length;
}

function sumSquaredDev(xs: number[]): number {
  const m = mean(xs) as number;
  let ss = 0;
  for (const x of xs) ss += (x - m) ** 2;
  return ss;
}

/** Sample variance (n-1), matching Excel VAR.S. */
export function variance(xs: number[]): Num {
  return xs.length < 2 ? null : sumSquaredDev(xs) / (xs.length - 1);
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

/** Matches Excel PERCENTILE.EXC: rank = q * (n + 1); blank when outside 1..n. */
export function quantileExc(xs: number[], q: number): Num {
  const n = xs.length;
  const rank = q * (n + 1);
  if (n === 0 || rank < 1 || rank > n) return null;
  const sorted = [...xs].sort((a, b) => a - b);
  const lo = Math.floor(rank);
  const hi = Math.min(lo + 1, n);
  return sorted[lo - 1] + (sorted[hi - 1] - sorted[lo - 1]) * (rank - lo);
}

export function median(xs: number[]): Num {
  return quantile(xs, 0.5);
}

/** Matches Excel MODE.SNGL: most frequent value, earliest on ties, blank if nothing repeats. */
function mode(xs: number[]): Num {
  const counts = new Map<number, number>();
  let best: Num = null;
  let bestCount = 1;
  for (const x of xs) {
    const c = (counts.get(x) ?? 0) + 1;
    counts.set(x, c);
    if (c > bestCount) {
      best = x;
      bestCount = c;
    }
  }
  if (best === null) return null;
  // Earliest value among those tied for the highest count.
  return xs.find((x) => counts.get(x) === bestCount) ?? best;
}

function kth(xs: number[], k: number, largest: boolean): Num {
  if (k > xs.length) return null;
  const sorted = [...xs].sort((a, b) => (largest ? b - a : a - b));
  return sorted[k - 1];
}

function skew(xs: number[]): Num {
  const n = xs.length;
  const s = stdev(xs);
  if (n < 3 || !s) return null;
  const m = mean(xs) as number;
  let acc = 0;
  for (const x of xs) acc += ((x - m) / s) ** 3;
  return (n / ((n - 1) * (n - 2))) * acc;
}

function kurt(xs: number[]): Num {
  const n = xs.length;
  const s = stdev(xs);
  if (n < 4 || !s) return null;
  const m = mean(xs) as number;
  let acc = 0;
  for (const x of xs) acc += ((x - m) / s) ** 4;
  return (
    ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * acc - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
  );
}

export function aggregate(xs: Num[] | number[], spec: Stat | StatSpec): Num {
  const { stat, k } = typeof spec === "string" ? { stat: spec, k: undefined } : spec;
  const v = valid(xs as Num[]);
  if (stat === "count" || stat === "counta") return v.length;
  if (v.length === 0) return null;
  switch (stat) {
    case "sum":
      return v.reduce((a, b) => a + b, 0);
    case "mean":
      return mean(v);
    case "min":
      return v.reduce((a, b) => (b < a ? b : a));
    case "max":
      return v.reduce((a, b) => (b > a ? b : a));
    case "range":
      return v.reduce((a, b) => (b > a ? b : a)) - v.reduce((a, b) => (b < a ? b : a));
    case "median":
      return median(v);
    case "mode":
      return mode(v);
    case "stdev":
      return stdev(v);
    case "stdevp":
      return Math.sqrt(sumSquaredDev(v) / v.length);
    case "var":
      return variance(v);
    case "varp":
      return sumSquaredDev(v) / v.length;
    case "product":
      return v.reduce((a, b) => a * b, 1);
    case "first":
      return v[0];
    case "last":
      return v[v.length - 1];
    case "geomean":
      return v.every((x) => x > 0)
        ? Math.exp(v.reduce((a, x) => a + Math.log(x), 0) / v.length)
        : null;
    case "harmean":
      return v.every((x) => x > 0) ? v.length / v.reduce((a, x) => a + 1 / x, 0) : null;
    case "avedev": {
      const m = mean(v) as number;
      return v.reduce((a, x) => a + Math.abs(x - m), 0) / v.length;
    }
    case "skew":
      return skew(v);
    case "kurt":
      return kurt(v);
    case "large":
      return kth(v, k as number, true);
    case "small":
      return kth(v, k as number, false);
    case "percentile":
      return quantile(v, k as number);
    case "quartile":
      return quantile(v, (k as number) / 4);
    case "percentile_exc":
      return quantileExc(v, k as number);
    case "quartile_exc":
      return quantileExc(v, (k as number) / 4);
  }
}
