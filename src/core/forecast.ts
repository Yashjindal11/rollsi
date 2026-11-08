import type { Num } from "./series";
import { RollsiError, requirePositiveInt } from "./series";

function points(values: Num[]): { x: number[]; y: number[] } {
  const x: number[] = [];
  const y: number[] = [];
  values.forEach((v, i) => {
    if (v !== null) {
      x.push(i);
      y.push(v);
    }
  });
  return { x, y };
}

/** Ordinary least squares fit of value against row position (blanks skipped). */
export function linearFit(values: Num[]): { slope: number; intercept: number } {
  const { x, y } = points(values);
  const n = x.length;
  if (n < 2) throw new RollsiError("Need at least 2 numbers to fit a trend.");
  const mx = x.reduce((a, b) => a + b, 0) / n;
  const my = y.reduce((a, b) => a + b, 0) / n;
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < n; i++) {
    sxy += (x[i] - mx) * (y[i] - my);
    sxx += (x[i] - mx) ** 2;
  }
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx };
}

/** Straight-line trend value for every row (useful to overlay on a chart). */
export function trendline(values: Num[]): Num[] {
  const { slope, intercept } = linearFit(values);
  return values.map((_, i) => intercept + slope * i);
}

export type ForecastMethod = "linear" | "holt" | "ses";

export function parseForecastMethod(input: unknown): ForecastMethod {
  if (input === null || input === undefined || input === "") return "holt";
  const key = String(input).trim().toLowerCase();
  if (key === "linear" || key === "holt" || key === "ses") return key;
  throw new RollsiError(`Unknown method "${input}". Use holt, linear or ses.`);
}

function checkSmoothing(value: number, name: string): number {
  if (!(value > 0 && value <= 1))
    throw new RollsiError(`${name} must be between 0 (exclusive) and 1.`);
  return value;
}

/**
 * Forecasts the next `horizon` periods after the last number in the series.
 * - holt:   Holt's linear exponential smoothing (level + trend), alpha=0.3, beta=0.1 by default
 * - ses:    simple exponential smoothing (flat forecast), alpha=0.3 by default
 * - linear: extends the least-squares trend line
 */
export function forecast(
  values: Num[],
  horizon: number,
  method: ForecastMethod = "holt",
  alpha = 0.3,
  beta = 0.1
): number[] {
  requirePositiveInt(horizon, "horizon");
  const { x, y } = points(values);
  if (y.length === 0) throw new RollsiError("No numbers to forecast from.");
  const steps = Array.from({ length: horizon }, (_, k) => k + 1);

  if (method === "linear") {
    const { slope, intercept } = linearFit(values);
    const lastX = x[x.length - 1];
    return steps.map((k) => intercept + slope * (lastX + k));
  }

  checkSmoothing(alpha, "alpha");
  if (method === "ses") {
    let level = y[0];
    for (let i = 1; i < y.length; i++) level = alpha * y[i] + (1 - alpha) * level;
    return steps.map(() => level);
  }

  checkSmoothing(beta, "beta");
  if (y.length < 2) throw new RollsiError("Holt forecasting needs at least 2 numbers.");
  let level = y[0];
  let trend = y[1] - y[0];
  for (let i = 1; i < y.length; i++) {
    const prevLevel = level;
    level = alpha * y[i] + (1 - alpha) * (level + trend);
    trend = beta * (level - prevLevel) + (1 - beta) * trend;
  }
  return steps.map((k) => level + k * trend);
}
