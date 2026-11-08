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
(global as any).CustomFunctions = {
  Error: MockCfError,
  ErrorCode: { invalidValue: "#VALUE!", notAvailable: "#N/A" },
};

// eslint-disable-next-line @typescript-eslint/no-require-imports
const fx = require("../../src/functions/functions");

const D = 46023;

const isNA = (v: any) => v instanceof MockCfError && v.code === "#N/A";

describe("Excel chart helpers & in-cell charts", () => {
  it("SPARK, BAR, ARROW, SHORT", () => {
    assert.strictEqual(fx.spark([[1], [2]], null, null, null), "▁█");
    assert.strictEqual(fx.bar(5, 10, null), "█████");
    assert.strictEqual(fx.arrow(0.1, null, null, null), "▲ +10.0%");
    assert.strictEqual(fx.short(1234, null), "1.2K");
  });
  it("NA turns blanks into #N/A", () => {
    const out = fx.chartNa([[1], [""], ["x"]]);
    assert.strictEqual(out[0][0], 1);
    assert.ok(isNA(out[1][0]));
    assert.strictEqual(out[2][0], "x");
  });
  it("CONTROL spills 4 columns with #N/A gaps", () => {
    const out = fx.control([[1], [2], [3], [4], [100]], null, 1);
    assert.strictEqual(out.length, 5);
    assert.strictEqual(out[0].length, 4);
    assert.ok(isNA(out[0][3]));
    assert.strictEqual(out[4][3], 100);
  });
  it("PARETO, WATERFALL, VSTARGET shapes", () => {
    assert.deepStrictEqual(fx.paretoTable([["a"], ["b"]], [[1], [3]]), [
      ["b", 3, 0.75, 0.8],
      ["a", 1, 1, 0.8],
    ]);
    assert.strictEqual(
      fx.waterfallTable([["S"], ["x"], ["E"]], [[10], [5], [""]], null, null)[2][6],
      15
    );
    const v = fx.vsTarget([[10], [5]], [[8], [7]]);
    assert.deepStrictEqual(v[0].slice(0, 3), [10, 8, 2]);
    assert.ok(isNA(v[0][3]));
    assert.strictEqual(v[1][3], -2);
  });
});

describe("Excel wrappers", () => {
  const five = [[1], [2], [3], [4], [5]];

  it("MA: trailing average by default", () => {
    assert.deepStrictEqual(fx.ma(five, 3, null, null, null, null, null, null), [
      [""],
      [""],
      [2],
      [3],
      [4],
    ]);
  });

  it("MA: stat code + alignment code", () => {
    // 9 = SUM, 1 = mid point
    assert.deepStrictEqual(fx.ma(five, 3, 9, 1, null, null, null, null), [
      [""],
      [6],
      [9],
      [12],
      [""],
    ]);
    // custom: 1 row above, 1 row below, window left blank
    assert.deepStrictEqual(fx.ma(five, null, "sum", 3, 1, 1, null, null), [
      [""],
      [6],
      [9],
      [12],
      [""],
    ]);
    // 14 = LARGE with k = 2 over a trailing 3-row window
    assert.deepStrictEqual(fx.ma(five, 3, 14, 0, null, null, 2, null), [[""], [""], [2], [3], [4]]);
  });

  it("MA: COUNTA counts text cells, partial fills the edges", () => {
    assert.deepStrictEqual(fx.ma([[1], ["x"], [""], [4]], 2, 3, null, null, null, null, null), [
      [""],
      [2],
      [1],
      [1],
    ]);
    assert.deepStrictEqual(fx.ma(five, 3, "avg", "mid", null, null, null, true), [
      [1.5],
      [2],
      [3],
      [4],
      [4.5],
    ]);
  });

  it("MA: helpful errors", () => {
    assert.throws(() => fx.ma(five, null, null, null, null, null, null, null), /window/);
    assert.throws(() => fx.ma(five, 3, 99, null, null, null, null, null), /Unknown stat code/);
    assert.throws(() => fx.ma(five, 3, 14, null, null, null, null, null), /needs k/);
  });

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
    assert.throws(() => fx.rollingStat([[1], [2]], 2, "bogus", null), /Unknown stat/);
  });
});
