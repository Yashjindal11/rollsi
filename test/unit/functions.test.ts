import * as assert from "assert";
import "mocha";

/* global describe, it, global, require */

class MockCfError extends Error {
  constructor(
    public code: string,
    message?: string
  ) {
    super(message);
  }
}
(global as any).CustomFunctions = { Error: MockCfError, ErrorCode: { invalidValue: "#VALUE!" } };

// eslint-disable-next-line @typescript-eslint/no-require-imports
const fx = require("../../src/functions/functions");

const D = 46023;

describe("Excel wrappers", () => {
  it("spill a column the same shape as the input, with blanks for warm-up rows", () => {
    // Excel passes empty cells as "" and omitted optional args as null.
    assert.deepStrictEqual(fx.sma([[1], [2], [""], [4]], 2, null), [[""], [1.5], [2], [4]]);
  });

  it("spill a row when given a row", () => {
    assert.deepStrictEqual(fx.cumSum([[1, 2, 3]]), [[1, 3, 6]]);
  });

  it("LAST returns a single value, blank when there is no data", () => {
    assert.strictEqual(fx.last([[1], [2], [3], [""]], 2, null), 2.5);
    assert.strictEqual(fx.last([["Sales"], [5], [7]], 5, "sum"), 12);
    assert.strictEqual(fx.last([[""]], 3, null), "");
  });

  it("ROLLDAYS and PERIODCHANGE take dates + values", () => {
    const dates = [[D], [D + 1], [D + 7], [D + 8]];
    const vals = [[10], [20], [30], [40]];
    assert.deepStrictEqual(fx.periodChange(dates, vals, "W", null), [[""], [""], [2], [1]]);
    assert.deepStrictEqual(fx.rollDays(dates, vals, 2, "sum", true), [[10], [30], [30], [70]]);
  });

  it("OUTLIERS spills booleans", () => {
    assert.deepStrictEqual(fx.outliers([[1], [2], [3], [4], [100]], null, null), [
      [false],
      [false],
      [false],
      [false],
      [true],
    ]);
  });

  it("FORECAST spills the horizon", () => {
    assert.deepStrictEqual(fx.forecast([[1], [3], [5]], 2, "linear", null, null), [[7], [9]]);
  });

  it("turns bad input into #VALUE! with a helpful message", () => {
    assert.throws(
      () => fx.sma([[1], [2]], 0, null),
      (e: any) => e instanceof MockCfError && e.code === "#VALUE!" && /window/.test(e.message)
    );
    assert.throws(
      () =>
        fx.rollingStat(
          [
            [1, 2],
            [3, 4],
          ],
          2,
          null,
          null
        ),
      MockCfError
    );
    assert.throws(() => fx.rollingStat([[1], [2]], 2, "mode", null), /Unknown stat/);
  });
});
