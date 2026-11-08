/* global Excel */

export interface ExportSize {
  label: string;
  width?: number;
  height?: number;
}

export const EXPORT_SIZES: ExportSize[] = [
  { label: "As shown in Excel" },
  { label: "Full slide 16:9 (1920 × 1080)", width: 1920, height: 1080 },
  { label: "Half slide (960 × 1080)", width: 960, height: 1080 },
  { label: "Wide strip (1920 × 540)", width: 1920, height: 540 },
];

export interface ChartImage {
  name: string;
  base64: string;
}

/** PNG images of the selected chart, or every chart on the active sheet. */
export async function exportCharts(
  context: Excel.RequestContext,
  scope: "selected" | "sheet",
  size: ExportSize
): Promise<ChartImage[]> {
  let charts: Excel.Chart[];
  if (scope === "selected") {
    const chart = context.workbook.getActiveChartOrNullObject();
    await context.sync();
    if (chart.isNullObject)
      throw new Error("Click a chart first, or choose 'All charts on this sheet'.");
    charts = [chart];
  } else {
    const coll = context.workbook.worksheets.getActiveWorksheet().charts;
    coll.load("items");
    await context.sync();
    charts = coll.items;
    if (!charts.length) throw new Error("There are no charts on this sheet.");
  }
  const jobs = charts.map((c) => {
    c.load("name");
    return {
      chart: c,
      image: size.width
        ? c.getImage(size.width, size.height, Excel.ImageFittingMode.fit)
        : c.getImage(),
    };
  });
  await context.sync();
  return jobs.map((j) => ({ name: j.chart.name, base64: j.image.value }));
}
