export type Num = number | null;
export type Orientation = "column" | "row";
export type OutCell = number | string | boolean;

export class RollsiError extends Error {}

/** Converts any cell value to a finite number, or null for blanks/text/booleans/errors. */
export function toNum(v: unknown): Num {
  return typeof v === "number" && Number.isFinite(v) ? v : null;
}

/** Flattens a single-column or single-row Excel range into a 1-D series and remembers its shape. */
export function toSeries(range: unknown[][]): { values: Num[]; orientation: Orientation } {
  if (!Array.isArray(range) || range.length === 0) return { values: [], orientation: "column" };
  const rows = range.length;
  const cols = range[0].length;
  if (rows > 1 && cols > 1) {
    throw new RollsiError("Pass a single column or a single row, not a 2-D block.");
  }
  if (rows === 1 && cols > 1) return { values: range[0].map(toNum), orientation: "row" };
  return { values: range.map((r) => toNum(r[0])), orientation: "column" };
}

/** Converts a 1-D result back into a spill array with the original orientation. Nulls become blanks. */
export function toRange(values: (OutCell | null)[], orientation: Orientation): OutCell[][] {
  const cells = values.map((v) =>
    v === null || (typeof v === "number" && !Number.isFinite(v)) ? "" : v
  );
  if (cells.length === 0) return [[""]];
  return orientation === "row" ? [cells] : cells.map((v) => [v]);
}

export function requirePositiveInt(value: unknown, name: string): number {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
    throw new RollsiError(`${name} must be a whole number >= 1.`);
  }
  return value;
}

export function valid(values: Num[]): number[] {
  return values.filter((v): v is number => v !== null);
}
