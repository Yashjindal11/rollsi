import * as assert from "assert";
import "mocha";
import * as r from "../../src/core";

/* global describe, it */

describe("viz: formatting & in-cell charts", () => {
  it("shortNumber", () => {
    assert.strictEqual(r.shortNumber(1234), "1.2K");
    assert.strictEqual(r.shortNumber(5600000), "5.6M");
    assert.strictEqual(r.shortNumber(-2500000000), "-2.5B");
    assert.strictEqual(r.shortNumber(999), "999");
    assert.strictEqual(r.shortNumber(1000), "1K");
    assert.strictEqual(r.shortNumber(1500, 0), "2K");
  });
  it("sparkline", () => {
    assert.strictEqual(r.sparkline([1, 2, 3, 4, 5, 6, 7, 8]), "▁▂▃▄▅▆▇█");
    assert.strictEqual(r.sparkline([1, null, 1]), "▄ ▄");
    assert.strictEqual(r.sparkline([1, -1, 0, null], "winloss"), "▀▄─ ");
    assert.strictEqual(r.sparkline([5], "bars", 0, 10), "▅");
  });
  it("textBar", () => {
    assert.strictEqual(r.textBar(5, 10), "█████");
    assert.strictEqual(r.textBar(1.25, 10), "█▎");
    assert.strictEqual(r.textBar(-5, 10), "-█████");
    assert.strictEqual(r.textBar(20, 10, 4), "████");
    assert.strictEqual(r.textBar(0, 10), "");
    assert.throws(() => r.textBar(1, 0), r.RollsiError);
  });
  it("arrow", () => {
    assert.strictEqual(r.arrow(0.123), "▲ +12.3%");
    assert.strictEqual(r.arrow(-0.04), "▼ -4.0%");
    assert.strictEqual(r.arrow(0.001, 0.01), "► +0.1%");
    assert.strictEqual(r.arrow(1500, 0, "num"), "▲ +1.5K");
    assert.strictEqual(r.arrow(-2, 0, "arrow"), "▼");
  });
});

describe("viz: chart data", () => {
  it("control limits (whole series)", () => {
    const c = r.controlLimits([1, 2, 3, 4, 100], null, 1);
    assert.deepStrictEqual(c.outlier, [null, null, null, null, 100]);
    assert.ok(c.center.every((v) => v === 22));
  });
  it("control limits (rolling, previous window)", () => {
    const c = r.controlLimits([10, 10, 12, 10, 50], 3, 3);
    assert.deepStrictEqual(c.center.slice(0, 3), [null, null, null]);
    assert.deepStrictEqual(c.outlier, [null, null, null, null, 50]);
  });
  it("pareto", () => {
    assert.deepStrictEqual(r.pareto(["a", "b", "c"], [1, 3, null]), [
      { label: "b", value: 3, cumulative: 0.75 },
      { label: "a", value: 1, cumulative: 1 },
    ]);
  });
  it("splitBar handles zero crossings", () => {
    assert.deepStrictEqual(r.splitBar(2, 5), { basePos: 2, baseNeg: 0, pos: 3, neg: 0 });
    assert.deepStrictEqual(r.splitBar(-1, -4), { basePos: 0, baseNeg: -1, pos: 0, neg: -3 });
    assert.deepStrictEqual(r.splitBar(3, -2), { basePos: 0, baseNeg: 0, pos: 3, neg: -2 });
  });
  it("waterfall steps and stacked matrix", () => {
    const steps = r.waterfall(["Start", "Price", "Cost", "End"], [100, 20, -50, null]);
    assert.deepStrictEqual(
      steps.map((s) => [s.kind, s.from, s.to, s.label]),
      [
        ["start", 0, 100, 100],
        ["up", 100, 120, 20],
        ["down", 120, 70, -50],
        ["total", 0, 70, 70],
      ]
    );
    assert.deepStrictEqual(r.waterfallMatrix(steps), [
      [0, 0, 0, 0, 0, 0, 100, 0],
      [100, 0, 20, 0, 0, 0, 0, 0],
      [70, 0, 0, 0, 50, 0, 0, 0],
      [0, 0, 0, 0, 0, 0, 70, 0],
    ]);
  });
  it("waterfall subtotals and crossing zero", () => {
    const steps = r.waterfall(["Start", "A", "Subtotal", "B", "End"], [10, 5, null, -20, null]);
    assert.deepStrictEqual(r.waterfallMatrix(steps).slice(2), [
      [0, 0, 0, 0, 0, 0, 15, 0],
      [0, 0, 0, 0, 15, -5, 0, 0],
      [0, 0, 0, 0, 0, 0, 0, -5],
    ]);
  });
  it("mekko layout", () => {
    const m = r.mekkoLayout([
      [1, 3],
      [1, 1],
    ]);
    assert.strictEqual(m.grandTotal, 6);
    assert.ok(Math.abs(m.columns[0].w - 1 / 3) < 1e-12);
    assert.ok(Math.abs(m.columns[1].x - 1 / 3) < 1e-12);
    const c = m.cells.find((x) => x.seg === 1 && x.cat === 1)!;
    assert.ok(Math.abs(c.y - 0.75) < 1e-12 && Math.abs(c.h - 0.25) < 1e-12);
    assert.throws(() => r.mekkoLayout([[1, -1]]), r.RollsiError);
  });
  it("annotation geometry and labels", () => {
    const box = { left: 10, top: 0, width: 100, height: 100, min: 0, max: 50 };
    assert.strictEqual(r.categoryX(box, 0, 4), 22.5);
    assert.strictEqual(r.valueY(box, 25), 50);
    assert.strictEqual(r.diffLabel(100, 120), "+20%");
    assert.strictEqual(r.diffLabel(100, 120, "abs"), "+20");
    assert.strictEqual(r.diffLabel(100, 120, "both"), "+20 (+20%)");
    assert.strictEqual(r.cagrLabel(100, 121, 2), "CAGR +10.0%");
  });
  it("variance split", () => {
    assert.deepStrictEqual(r.varianceSplit([10, 5, null], [8, 7, 1]), {
      over: [2, null, null],
      under: [null, -2, null],
    });
  });
});
