import type { Num } from "./series";
import { RollsiError } from "./series";
import type { Stat } from "./stats";
import { aggregate, parseStat } from "./stats";

const MS_PER_DAY = 86400000;
const EXCEL_EPOCH_OFFSET = 25569; // serial of 1970-01-01

export function serialToDate(serial: number): Date {
  return new Date(Math.round((Math.floor(serial) - EXCEL_EPOCH_OFFSET) * MS_PER_DAY));
}

export function dateToSerial(d: Date): number {
  return Math.round(d.getTime() / MS_PER_DAY) + EXCEL_EPOCH_OFFSET;
}

function daysInMonth(year: number, month0: number): number {
  return new Date(Date.UTC(year, month0 + 1, 0)).getUTCDate();
}

/** Like Excel EDATE, but a month-end date maps to the target month's end (Feb 28 -> Jan 31). */
export function addMonths(serial: number, months: number): number {
  const d = serialToDate(serial);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth();
  const day = d.getUTCDate();
  const isMonthEnd = day === daysInMonth(y, m);
  const target = new Date(Date.UTC(y, m + months, 1));
  const tMax = daysInMonth(target.getUTCFullYear(), target.getUTCMonth());
  target.setUTCDate(isMonthEnd ? tMax : Math.min(day, tMax));
  return dateToSerial(target);
}

type Period = { days: number } | { months: number };

const PERIODS: Record<string, Period> = {
  d: { days: 1 },
  day: { days: 1 },
  dod: { days: 1 },
  w: { days: 7 },
  week: { days: 7 },
  wow: { days: 7 },
  m: { months: 1 },
  month: { months: 1 },
  mom: { months: 1 },
  q: { months: 3 },
  quarter: { months: 3 },
  qoq: { months: 3 },
  y: { months: 12 },
  year: { months: 12 },
  yoy: { months: 12 },
};

export function parsePeriod(input: unknown): Period {
  if (typeof input === "number" && Number.isInteger(input) && input > 0) return { days: input };
  const p =
    PERIODS[
      String(input ?? "")
        .trim()
        .toLowerCase()
    ];
  if (!p)
    throw new RollsiError(
      `Unknown period "${input}". Use D, W, M, Q, Y (or DoD, WoW, MoM, QoQ, YoY) or a number of days.`
    );
  return p;
}

function shift(day: number, period: Period, sign: 1 | -1): number {
  return "days" in period ? day + sign * period.days : addMonths(day, sign * period.months);
}

export type ChangeMode = "pct" | "diff" | "prior";

export function parseChangeMode(input: unknown): ChangeMode {
  if (input === null || input === undefined || input === "") return "pct";
  const key = String(input).trim().toLowerCase();
  if (key === "pct" || key === "%" || key === "percent") return "pct";
  if (key === "diff" || key === "abs" || key === "delta") return "diff";
  if (key === "prior" || key === "prev" || key === "value") return "prior";
  throw new RollsiError(`Unknown mode "${input}". Use pct, diff or prior.`);
}

export function change(current: number, prior: number, mode: ChangeMode): Num {
  if (mode === "diff") return current - prior;
  if (mode === "prior") return prior;
  return prior === 0 ? null : (current - prior) / prior;
}

/**
 * Calendar-aware period-over-period change (DoD, WoW, MoM, QoQ, YoY). For each row it looks up the
 * row whose date is exactly one period earlier, so gaps and unsorted data are handled correctly.
 */
export function periodChange(
  dates: Num[],
  values: Num[],
  period: unknown,
  mode: ChangeMode = "pct"
): Num[] {
  if (dates.length !== values.length)
    throw new RollsiError("dates and values must be the same size.");
  const p = parsePeriod(period);
  const byDay = new Map<number, number>();
  dates.forEach((d, i) => {
    if (d === null) return;
    const day = Math.floor(d);
    if (byDay.has(day)) {
      throw new RollsiError("Each date may appear only once. Aggregate to one row per date first.");
    }
    const v = values[i];
    if (v !== null) byDay.set(day, v);
  });
  return dates.map((d, i) => {
    const v = values[i];
    if (d === null || v === null) return null;
    const prior = byDay.get(shift(Math.floor(d), p, -1));
    return prior === undefined ? null : change(v, prior, mode);
  });
}

export type DayStat = Stat | "perday";

export function parseDayStat(input: unknown): DayStat {
  const key = String(input ?? "")
    .trim()
    .toLowerCase();
  if (key === "perday" || key === "daily" || key === "per_day") return "perday";
  return parseStat(input);
}

/**
 * Date-aware trailing window: aggregates every row whose date falls in the last `days` calendar
 * days up to and including the current row's date. Missing days are simply absent, so a 7-day
 * window always spans 7 calendar days. Multiple rows per date are all included.
 * Use stat "perday" for sum / days, i.e. a daily average that treats missing days as zero.
 */
export function rollDays(
  dates: Num[],
  values: Num[],
  days: number,
  stat: DayStat,
  partial = false
): Num[] {
  if (dates.length !== values.length)
    throw new RollsiError("dates and values must be the same size.");
  if (typeof days !== "number" || !Number.isInteger(days) || days < 1) {
    throw new RollsiError("days must be a whole number >= 1.");
  }
  const idx = dates
    .map((d, i) => ({ day: d === null ? null : Math.floor(d), i }))
    .filter((r): r is { day: number; i: number } => r.day !== null)
    .sort((a, b) => a.day - b.day);
  if (idx.length === 0) return dates.map(() => null);
  const sortedDays = idx.map((r) => r.day);
  const firstDay = sortedDays[0];

  const lowerBound = (x: number) => {
    let lo = 0;
    let hi = sortedDays.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (sortedDays[mid] < x) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };

  return dates.map((d) => {
    if (d === null) return null;
    const day = Math.floor(d);
    const start = day - days + 1;
    if (!partial && start < firstDay) return null;
    const from = lowerBound(start);
    const to = lowerBound(day + 1);
    const window: Num[] = [];
    for (let k = from; k < to; k++) window.push(values[idx[k].i]);
    if (stat === "perday") {
      const sum = aggregate(window, "sum");
      return sum === null ? null : sum / days;
    }
    return aggregate(window, stat);
  });
}
