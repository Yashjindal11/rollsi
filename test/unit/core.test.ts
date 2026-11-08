import * as assert from "assert";
import "mocha";
import * as r from "../../src/core";

/* global describe, it */

const close = (actual: (number | null)[], expected: (number | null)[], eps = 1e-9) => {
  assert.strictEqual(
    actual.length,
    expected.length,
    `length ${actual.length} != ${expected.length}`
  );
  actual.forEach((a, i) => {
    const e = expected[i];
    if (e === null || a === null) assert.strictEqual(a, e, `index ${i}: ${a} != ${e}`);
    else assert.ok(Math.abs(a - e) < eps, `index ${i}: ${a} != ${e}`);
  });
};

const D = 46023; // 2026-01-01 as an Excel serial date

describe("series", () => {
  it("reads a column and keeps orientation", () => {
    const s = r.toSeries([[1], [""], ["text"], [true], [2]]);
    assert.deepStrictEqual(s, { values: [1, null, null, null, 2], orientation: "column" });
  });
  it("reads a row", () => {
    assert.deepStrictEqual(r.toSeries([[1, 2, 3]]).orientation, "row");
  });
  it("rejects 2-D blocks", () => {
    assert.throws(
      () =>
        r.toSeries([
          [1, 2],
          [3, 4],
        ]),
      r.RollsiError
    );
  });
  it("writes blanks for nulls and NaN", () => {
    assert.deepStrictEqual(r.toRange([1, null, NaN], "column"), [[1], [""], [""]]);
    assert.deepStrictEqual(r.toRange([1, 2], "row"), [[1, 2]]);
  });
});

describe("moving averages", () => {
  it("SMA leaves the warm-up rows blank", () => {
    close(r.sma([1, 2, 3, 4, 5], 3), [null, null, 2, 3, 4]);
  });
  it("SMA partial computes from available rows", () => {
    close(r.sma([1, 2, 3, 4, 5], 3, true), [1, 1.5, 2, 3, 4]);
  });
  it("SMA skips blanks inside a window", () => {
    close(r.sma([1, null, 3, 4], 2), [null, 1, 3, 3.5]);
  });
  it("EMA matches pandas ewm(adjust=False)", () => {
    close(r.ema([1, 2, 3], 3, true), [1, 1.5, 2.25]);
    close(r.ema([1, 2, 3], 3), [null, null, 2.25]);
  });
  it("EMA rejects bad alpha", () => {
    assert.throws(() => r.ema([1], 3, false, 1.5), r.RollsiError);
  });
  it("WMA weights recent rows more", () => {
    close(r.wma([1, 2, 3], 3), [null, null, 14 / 6]);
  });
  it("rejects a non-integer window", () => {
    assert.throws(() => r.sma([1, 2], 0), r.RollsiError);
    assert.throws(() => r.sma([1, 2], 2.5), r.RollsiError);
  });
  it("LAST ignores trailing blanks", () => {
    const v = [1, 2, 3, null, null];
    assert.strictEqual(r.last(v, 2, "mean"), 2.5);
    assert.strictEqual(r.last(v, 2, "sum"), 5);
    assert.strictEqual(r.last(v, 10, "count"), 3);
  });
});

describe("rolling stats", () => {
  const v = [2, 4, 4, 4, 5, 5, 7, 9];
  it("sum, min, max, median, count", () => {
    close(r.rolling(v, 3, "sum"), [null, null, 10, 12, 13, 14, 17, 21]);
    close(r.rolling(v, 3, "min"), [null, null, 2, 4, 4, 4, 5, 5]);
    close(r.rolling(v, 3, "max"), [null, null, 4, 4, 5, 5, 7, 9]);
    close(r.rolling(v, 4, "median"), [null, null, null, 4, 4, 4.5, 5, 6]);
    close(r.rolling([1, null, 3], 2, "count"), [null, 1, 1]);
  });
  it("sample stdev matches STDEV.S", () => {
    const s = r.rolling(v, 8, "stdev");
    assert.ok(Math.abs((s[7] as number) - 2.138089935299395) < 1e-12);
  });
  it("parses stat aliases", () => {
    assert.strictEqual(r.parseStat("AVG"), "mean");
    assert.strictEqual(r.parseStat(undefined), "mean");
    assert.strictEqual(r.parseStat("std"), "stdev");
    assert.throws(() => r.parseStat("mode"), r.RollsiError);
  });
  it("cumulative", () => {
    close(r.cumulative([1, null, 3], "sum"), [1, null, 4]);
    close(r.cumulative([1, null, 3], "mean"), [1, null, 2]);
    close(r.cumulative([3, 1, 2], "max"), [3, 3, 3]);
    close(r.cumulative([3, 1, 2], "min"), [3, 1, 1]);
  });
});

describe("dates", () => {
  it("converts Excel serials", () => {
    assert.strictEqual(r.serialToDate(D).toISOString().slice(0, 10), "2026-01-01");
    assert.strictEqual(r.dateToSerial(new Date(Date.UTC(2026, 0, 1))), D);
  });
  it("adds months with month-end awareness", () => {
    assert.strictEqual(r.addMonths(D + 30, 1), D + 58); // Jan 31 -> Feb 28
    assert.strictEqual(r.addMonths(D + 58, -1), D + 30); // Feb 28 (month end) -> Jan 31
    assert.strictEqual(r.addMonths(D + 14, 12), D + 14 + 365); // Jan 15 -> next Jan 15
  });
  it("PERIODCHANGE week over week", () => {
    close(r.periodChange([D, D + 7, D + 14], [100, 110, 99], "W"), [null, 0.1, -0.1]);
  });
  it("PERIODCHANGE matches by date, not by row", () => {
    close(r.periodChange([D, D + 1, D + 8], [100, 200, 300], "WoW"), [null, null, 0.5]);
    close(r.periodChange([D + 8, D, D + 1], [300, 100, 200], "w", "diff"), [100, null, null]);
  });
  it("PERIODCHANGE month over month on month-end dates", () => {
    close(r.periodChange([D + 30, D + 58, D + 89], [10, 20, 30], "M"), [null, 1, 0.5]);
  });
  it("PERIODCHANGE rejects duplicate dates and bad periods", () => {
    assert.throws(() => r.periodChange([D, D], [1, 2], "D"), r.RollsiError);
    assert.throws(() => r.periodChange([D], [1], "fortnight"), r.RollsiError);
  });
  it("ROLLDAYS uses calendar days, not rows", () => {
    const dates = [D, D + 1, D + 3, D + 4];
    const vals = [1, 2, 3, 4];
    close(r.rollDays(dates, vals, 3, "sum"), [null, null, 5, 7]);
    close(r.rollDays(dates, vals, 3, "sum", true), [1, 3, 5, 7]);
    close(r.rollDays(dates, vals, 3, "perday"), [null, null, 5 / 3, 7 / 3]);
  });
  it("ROLLDAYS handles unsorted input and many rows per day", () => {
    close(r.rollDays([D + 4, D + 3, D + 1, D], [4, 3, 2, 1], 3, "sum"), [7, 5, null, null]);
    close(r.rollDays([D, D, D + 1], [1, 2, 3], 2, "sum", true), [3, 3, 6]);
  });
});

describe("growth", () => {
  it("PCTCHANGE and DIFF", () => {
    close(r.pctChange([100, 110, null, 121]), [null, 0.1, null, null]);
    close(r.diff([1, 2, 4, 7], 2), [null, null, 3, 5]);
    close(r.pctChange([0, 5]), [null, null]);
  });
  it("REBASE", () => {
    close(r.rebase([null, 50, 75, null]), [null, 100, 150, null]);
  });
  it("CAGR", () => {
    assert.ok(Math.abs(r.cagr(100, 121, 2) - 0.1) < 1e-12);
    assert.throws(() => r.cagr(0, 10, 1), r.RollsiError);
  });
});

describe("analysis", () => {
  it("ZSCORE", () => {
    close(r.zscore([1, 2, 3, null]), [-1, 0, 1, null]);
  });
  it("ROLLZSCORE compares to the previous window", () => {
    close(r.rollingZscore([1, 2, 3, 10], 3), [null, null, null, 8]);
  });
  it("OUTLIERS iqr, z and mad", () => {
    assert.deepStrictEqual(r.outliers([1, 2, 3, 4, 100]), [false, false, false, false, true]);
    assert.deepStrictEqual(r.outliers([1, 2, 3, 4, 100, null], "mad"), [
      false,
      false,
      false,
      false,
      true,
      null,
    ]);
    assert.deepStrictEqual(r.outliers([1, 2, 3], "z"), [false, false, false]);
  });
  it("PCTRANK matches PERCENTRANK.INC", () => {
    close(r.pctRank([10, 20, 30, 40, 50]), [0, 0.25, 0.5, 0.75, 1]);
    close(r.pctRank([1, 2, 2, 3]), [0, 1 / 3, 1 / 3, 1]);
  });
  it("NORMALIZE", () => {
    close(r.normalize([0, 5, 10]), [0, 0.5, 1]);
    close(r.normalize([1, 3], "share"), [0.25, 0.75]);
    close(r.normalize([4, 4]), [0, 0]);
  });
});

describe("forecast", () => {
  it("TRENDLINE fits a straight line, skipping blanks", () => {
    close(r.trendline([1, 3, 5]), [1, 3, 5]);
    close(r.trendline([1, null, 5]), [1, 3, 5]);
  });
  it("linear forecast continues after the last number", () => {
    close(r.forecast([1, 3, 5], 2, "linear"), [7, 9]);
    close(r.forecast([1, null, 5, null], 1, "linear"), [7]);
  });
  it("holt follows a perfect trend", () => {
    close(r.forecast([2, 4, 6, 8], 2, "holt"), [10, 12]);
  });
  it("ses is flat", () => {
    close(r.forecast([5, 5, 5], 3, "ses"), [5, 5, 5]);
  });
  it("validates input", () => {
    assert.throws(() => r.forecast([1], 2, "holt"), r.RollsiError);
    assert.throws(() => r.forecast([1, 2], 2, "holt", 0), r.RollsiError);
    assert.throws(() => r.parseForecastMethod("arima"), r.RollsiError);
  });
});
