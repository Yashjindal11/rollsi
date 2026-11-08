/* global CustomFunctions */
import * as core from "../core";

function run<T>(fn: () => T): T {
  try {
    return fn();
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
  }
}

function spill(
  range: any[][],
  transform: (values: core.Num[]) => (core.OutCell | null)[]
): any[][] {
  return run(() => {
    const { values, orientation } = core.toSeries(range);
    return core.toRange(transform(values), orientation);
  });
}

function spillPair(
  dates: any[][],
  values: any[][],
  transform: (d: core.Num[], v: core.Num[]) => core.Num[]
): any[][] {
  return run(() => {
    const d = core.toSeries(dates);
    const v = core.toSeries(values);
    return core.toRange(transform(d.values, v.values), v.orientation);
  });
}

function scalar(value: core.Num): any {
  return value === null ? "" : value;
}

// ---------------------------------------------------------------- Moving averages

/**
 * Simple moving average over the last N rows, e.g. a 7-day or 14-day moving average. Spills a full column.
 * @customfunction SMA
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window, e.g. 7.
 * @param [partial] TRUE to also compute the first N-1 rows from the data available. Default FALSE (blank).
 * @returns Moving average for every row.
 */
export function sma(values: any[][], window: number, partial?: boolean): any[][] {
  return spill(values, (v) => core.sma(v, window, partial ?? false));
}

/**
 * Exponential moving average that reacts faster to recent changes. alpha = 2 / (span + 1).
 * @customfunction EMA
 * @param values Column (or row) of numbers.
 * @param span Span in rows, e.g. 7.
 * @param [partial] TRUE to show the first span-1 warm-up rows. Default FALSE.
 * @param [alpha] Optional smoothing factor between 0 and 1. Overrides span.
 * @returns Exponential moving average for every row.
 */
export function ema(values: any[][], span: number, partial?: boolean, alpha?: number): any[][] {
  return spill(values, (v) => core.ema(v, span, partial ?? false, alpha ?? undefined));
}

/**
 * Weighted moving average: the newest row has weight N, the oldest has weight 1.
 * @customfunction WMA
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window.
 * @param [partial] TRUE to also compute the first N-1 rows. Default FALSE.
 * @returns Weighted moving average for every row.
 */
export function wma(values: any[][], window: number, partial?: boolean): any[][] {
  return spill(values, (v) => core.wma(v, window, partial ?? false));
}

/**
 * Single value from the last N data points (blanks ignored), e.g. the latest 7-day average for a KPI card.
 * @customfunction LAST
 * @param values Column (or row) of numbers. Can include empty rows at the bottom.
 * @param n How many of the most recent numbers to use.
 * @param [stat] avg (default), sum, min, max, median, stdev, var, count, product.
 * @returns The statistic over the last N numbers.
 */
export function last(values: any[][], n: number, stat?: string): any {
  return run(() => scalar(core.last(core.toSeries(values).values, n, core.parseStat(stat))));
}

// ---------------------------------------------------------------- Rolling statistics

/**
 * Any rolling statistic over the last N rows.
 * @customfunction ROLLING
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window.
 * @param [stat] avg (default), sum, min, max, median, stdev, var, count, product.
 * @param [partial] TRUE to also compute the first N-1 rows. Default FALSE.
 * @returns Rolling statistic for every row.
 */
export function rollingStat(
  values: any[][],
  window: number,
  stat?: string,
  partial?: boolean
): any[][] {
  return spill(values, (v) => core.rolling(v, window, core.parseStat(stat), partial ?? false));
}

/**
 * Rolling sum over the last N rows, e.g. a trailing 7-day total.
 * @customfunction ROLLSUM
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window.
 * @param [partial] TRUE to also compute the first N-1 rows. Default FALSE.
 * @returns Rolling sum for every row.
 */
export function rollSum(values: any[][], window: number, partial?: boolean): any[][] {
  return spill(values, (v) => core.rolling(v, window, "sum", partial ?? false));
}

/**
 * Rolling minimum over the last N rows.
 * @customfunction ROLLMIN
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window.
 * @param [partial] TRUE to also compute the first N-1 rows. Default FALSE.
 * @returns Rolling minimum for every row.
 */
export function rollMin(values: any[][], window: number, partial?: boolean): any[][] {
  return spill(values, (v) => core.rolling(v, window, "min", partial ?? false));
}

/**
 * Rolling maximum over the last N rows.
 * @customfunction ROLLMAX
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window.
 * @param [partial] TRUE to also compute the first N-1 rows. Default FALSE.
 * @returns Rolling maximum for every row.
 */
export function rollMax(values: any[][], window: number, partial?: boolean): any[][] {
  return spill(values, (v) => core.rolling(v, window, "max", partial ?? false));
}

/**
 * Rolling median over the last N rows (robust to spikes).
 * @customfunction ROLLMEDIAN
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window.
 * @param [partial] TRUE to also compute the first N-1 rows. Default FALSE.
 * @returns Rolling median for every row.
 */
export function rollMedian(values: any[][], window: number, partial?: boolean): any[][] {
  return spill(values, (v) => core.rolling(v, window, "median", partial ?? false));
}

/**
 * Rolling sample standard deviation over the last N rows (volatility).
 * @customfunction ROLLSTDEV
 * @param values Column (or row) of numbers.
 * @param window Number of rows in the window.
 * @param [partial] TRUE to also compute the first N-1 rows. Default FALSE.
 * @returns Rolling standard deviation for every row.
 */
export function rollStdev(values: any[][], window: number, partial?: boolean): any[][] {
  return spill(values, (v) => core.rolling(v, window, "stdev", partial ?? false));
}

/**
 * Date-aware rolling window: the last N calendar days up to each row's date. Handles missing days and multiple rows per day.
 * @customfunction ROLLDAYS
 * @param dates Column of dates (same size as values). Does not need to be sorted.
 * @param values Column of numbers.
 * @param days Window length in calendar days, e.g. 7.
 * @param [stat] avg (default), sum, min, max, median, stdev, var, count, product, or perday (sum / days, so missing days count as 0).
 * @param [partial] TRUE to also compute rows with less than N days of history. Default FALSE.
 * @returns Rolling statistic for every row.
 */
export function rollDays(
  dates: any[][],
  values: any[][],
  days: number,
  stat?: string,
  partial?: boolean
): any[][] {
  return spillPair(dates, values, (d, v) =>
    core.rollDays(d, v, days, core.parseDayStat(stat), partial ?? false)
  );
}

// ---------------------------------------------------------------- Cumulative

/**
 * Running total (cumulative sum).
 * @customfunction CUMSUM
 * @param values Column (or row) of numbers.
 * @returns Running total for every row.
 */
export function cumSum(values: any[][]): any[][] {
  return spill(values, (v) => core.cumulative(v, "sum"));
}

/**
 * Running average from the first row to each row.
 * @customfunction CUMAVG
 * @param values Column (or row) of numbers.
 * @returns Running average for every row.
 */
export function cumAvg(values: any[][]): any[][] {
  return spill(values, (v) => core.cumulative(v, "mean"));
}

/**
 * Running maximum (all-time high so far).
 * @customfunction CUMMAX
 * @param values Column (or row) of numbers.
 * @returns Running maximum for every row.
 */
export function cumMax(values: any[][]): any[][] {
  return spill(values, (v) => core.cumulative(v, "max"));
}

/**
 * Running minimum (all-time low so far).
 * @customfunction CUMMIN
 * @param values Column (or row) of numbers.
 * @returns Running minimum for every row.
 */
export function cumMin(values: any[][]): any[][] {
  return spill(values, (v) => core.cumulative(v, "min"));
}

// ---------------------------------------------------------------- Growth

/**
 * Percent change versus N rows earlier (format the result as %).
 * @customfunction PCTCHANGE
 * @param values Column (or row) of numbers.
 * @param [periods] Rows to look back. Default 1.
 * @returns Percent change for every row.
 */
export function pctChange(values: any[][], periods?: number): any[][] {
  return spill(values, (v) => core.pctChange(v, periods ?? 1));
}

/**
 * Difference versus N rows earlier.
 * @customfunction DIFF
 * @param values Column (or row) of numbers.
 * @param [periods] Rows to look back. Default 1.
 * @returns Difference for every row.
 */
export function diff(values: any[][], periods?: number): any[][] {
  return spill(values, (v) => core.diff(v, periods ?? 1));
}

/**
 * Calendar-aware period-over-period change: DoD, WoW, MoM, QoQ or YoY. Matches each date with the date exactly one period earlier.
 * @customfunction PERIODCHANGE
 * @param dates Column of dates (one row per date).
 * @param values Column of numbers.
 * @param period D, W, M, Q, Y (or DoD, WoW, MoM, QoQ, YoY), or a number of days.
 * @param [mode] pct (default), diff, or prior (the earlier value itself).
 * @returns Period-over-period change for every row.
 */
export function periodChange(dates: any[][], values: any[][], period: any, mode?: string): any[][] {
  return spillPair(dates, values, (d, v) =>
    core.periodChange(d, v, period, core.parseChangeMode(mode))
  );
}

/**
 * Rebases a series so the first value equals 100 (or another base), to compare growth paths.
 * @customfunction REBASE
 * @param values Column (or row) of numbers.
 * @param [base] Starting value. Default 100.
 * @returns Rebased series.
 */
export function rebase(values: any[][], base?: number): any[][] {
  return spill(values, (v) => core.rebase(v, base ?? 100));
}

/**
 * Compound annual growth rate.
 * @customfunction CAGR
 * @param start Starting value (> 0).
 * @param end Ending value.
 * @param periods Number of periods (years) between start and end.
 * @returns CAGR as a decimal (format as %).
 */
export function cagr(start: number, end: number, periods: number): number {
  return run(() => core.cagr(start, end, periods));
}

// ---------------------------------------------------------------- Analysis

/**
 * Z-score of each value against the whole series.
 * @customfunction ZSCORE
 * @param values Column (or row) of numbers.
 * @returns Z-score for every row.
 */
export function zscore(values: any[][]): any[][] {
  return spill(values, (v) => core.zscore(v));
}

/**
 * Rolling anomaly score: standard deviations from the mean of the previous N rows.
 * @customfunction ROLLZSCORE
 * @param values Column (or row) of numbers.
 * @param window Number of previous rows to compare against.
 * @returns Rolling z-score for every row.
 */
export function rollZscore(values: any[][], window: number): any[][] {
  return spill(values, (v) => core.rollingZscore(v, window));
}

/**
 * Flags outliers with TRUE/FALSE.
 * @customfunction OUTLIERS
 * @param values Column (or row) of numbers.
 * @param [method] iqr (default, Tukey fences), z (z-score), or mad (robust median-based).
 * @param [threshold] Sensitivity. Defaults: iqr 1.5, z 3, mad 3.5.
 * @returns TRUE where a value is an outlier.
 */
export function outliers(values: any[][], method?: string, threshold?: number): any[][] {
  return spill(values, (v) =>
    core.outliers(v, core.parseOutlierMethod(method), threshold ?? undefined)
  );
}

/**
 * Percentile rank between 0 and 1 (same as PERCENTRANK.INC).
 * @customfunction PCTRANK
 * @param values Column (or row) of numbers.
 * @returns Percentile rank for every row.
 */
export function pctRank(values: any[][]): any[][] {
  return spill(values, (v) => core.pctRank(v));
}

/**
 * Rescales values: minmax (0 to 1), z (z-score), or share (fraction of total).
 * @customfunction NORMALIZE
 * @param values Column (or row) of numbers.
 * @param [method] minmax (default), z, or share.
 * @returns Normalized value for every row.
 */
export function normalize(values: any[][], method?: string): any[][] {
  return spill(values, (v) => core.normalize(v, core.parseNormalizeMethod(method)));
}

// ---------------------------------------------------------------- Trend & forecast

/**
 * Straight-line (least squares) trend value for every row, to overlay on a chart.
 * @customfunction TRENDLINE
 * @param values Column (or row) of numbers.
 * @returns Trend value for every row.
 */
export function trendline(values: any[][]): any[][] {
  return spill(values, (v) => core.trendline(v));
}

/**
 * Forecasts the next N periods after the last number.
 * @customfunction FORECAST
 * @param values Column (or row) of historical numbers.
 * @param horizon How many periods ahead to forecast.
 * @param [method] holt (default, level + trend), linear, or ses (flat).
 * @param [alpha] Level smoothing between 0 and 1. Default 0.3.
 * @param [beta] Trend smoothing between 0 and 1 (holt only). Default 0.1.
 * @returns Forecast values, spilled in the same direction as the input.
 */
export function forecast(
  values: any[][],
  horizon: number,
  method?: string,
  alpha?: number,
  beta?: number
): any[][] {
  return spill(values, (v) =>
    core.forecast(v, horizon, core.parseForecastMethod(method), alpha ?? 0.3, beta ?? 0.1)
  );
}
