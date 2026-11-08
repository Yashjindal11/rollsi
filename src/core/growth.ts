import type { Num } from "./series";
import { RollsiError, requirePositiveInt } from "./series";
import type { ChangeMode } from "./dates";
import { change } from "./dates";

/** Row-over-row change versus the row `periods` rows earlier (pct, diff or prior value). */
export function shiftChange(values: Num[], periods = 1, mode: ChangeMode = "pct"): Num[] {
  requirePositiveInt(periods, "periods");
  return values.map((x, i) => {
    const prior = i >= periods ? values[i - periods] : null;
    return x === null || prior === null ? null : change(x, prior, mode);
  });
}

export function pctChange(values: Num[], periods = 1): Num[] {
  return shiftChange(values, periods, "pct");
}

export function diff(values: Num[], periods = 1): Num[] {
  return shiftChange(values, periods, "diff");
}

/** Rebases the series so the first number equals `base` (e.g. 100), for comparing growth paths. */
export function rebase(values: Num[], base = 100): Num[] {
  const first = values.find((v): v is number => v !== null);
  if (first === undefined) return values.map(() => null);
  if (first === 0) throw new RollsiError("Cannot rebase: the first value is 0.");
  return values.map((v) => (v === null ? null : (v / first) * base));
}

/** Compound annual growth rate: (end / start) ^ (1 / periods) - 1. */
export function cagr(start: number, end: number, periods: number): number {
  if (!(start > 0) || !(end >= 0)) throw new RollsiError("CAGR needs start > 0 and end >= 0.");
  if (!(periods > 0)) throw new RollsiError("periods must be > 0.");
  return Math.pow(end / start, 1 / periods) - 1;
}
