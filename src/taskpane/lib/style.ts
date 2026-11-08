/* global Excel */
import { SHORT_NUMBER_FORMAT } from "../../core/viz";
import { SEMANTIC } from "./palette";

type Role =
  | "hidden"
  | "pos"
  | "neg"
  | "total"
  | "band"
  | "center"
  | "outlier"
  | "ref"
  | "target"
  | "smooth"
  | "data";

/** Series roles are inferred from names, so builders and the Style button agree on colours. */
export function roleOf(name: string): Role {
  if (name.startsWith("_")) return "hidden";
  if (/^(increase|over)\b/i.test(name)) return "pos";
  if (/^(decrease|under)\b/i.test(name)) return "neg";
  if (/^(total|start)\b/i.test(name)) return "total";
  if (/^(upper|lower)\b/i.test(name)) return "band";
  if (/^center\b/i.test(name)) return "center";
  if (/^outlier/i.test(name)) return "outlier";
  if (/^80% line/i.test(name)) return "ref";
  if (/^target\b/i.test(name)) return "target";
  if (/^(ma|ema|wma|trend)\b/i.test(name)) return "smooth";
  return "data";
}

const isLine = (t: string) => /line|scatter/i.test(t);
const isStacked = (t: string) => /stacked/i.test(t);
const isBarLike = (t: string) => /column|bar/i.test(t);

export interface StyleOptions {
  /** Direct data labels on column/bar series (think-cell style). Default: on for column/bar charts. */
  labels?: boolean;
}

/**
 * Consultant-style formatting: no gridlines, no border, muted axes, short number formats, palette
 * colours by series role, legend only when it is needed, optional direct labels.
 */
export async function applyStyle(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  palette: string[],
  opts: StyleOptions = {}
): Promise<void> {
  chart.load("chartType");
  chart.series.load("items/name,items/chartType");
  await context.sync();
  const series = chart.series.items;
  const hasHidden = series.some((s) => s.name.startsWith("_"));
  const chartType = String(chart.chartType);
  const wantLabels = (opts.labels ?? isBarLike(chartType)) && !hasHidden;

  chart.format.fill.setSolidColor("#FFFFFF");
  chart.format.border.clear();
  chart.format.font.color = "#404040";
  chart.format.font.size = 9;
  chart.title.format.font.size = 12;
  chart.title.format.font.bold = true;
  chart.title.format.font.color = SEMANTIC.ink;

  const va = chart.axes.valueAxis;
  va.majorGridlines.visible = false;
  va.numberFormat = SHORT_NUMBER_FORMAT;
  va.format.font.color = "#8C8C8C";
  va.format.line.clear();
  va.visible = !(wantLabels && isBarLike(chartType));
  const ca = chart.axes.categoryAxis;
  ca.format.line.color = "#BFBFBF";
  ca.format.font.color = "#595959";

  const hasSmooth = series.some((s) => roleOf(s.name) === "smooth");
  let next = hasSmooth ? 1 : 0;
  let visibleCount = 0;

  series.forEach((s) => {
    const role = roleOf(s.name);
    const type = String(s.chartType || chartType);
    const line = isLine(type);
    let color = palette[0];
    switch (role) {
      case "pos":
        color = SEMANTIC.pos;
        break;
      case "neg":
        color = SEMANTIC.neg;
        break;
      case "total":
        color = SEMANTIC.total;
        break;
      case "band":
      case "ref":
        color = SEMANTIC.band;
        break;
      case "center":
        color = SEMANTIC.center;
        break;
      case "outlier":
        color = SEMANTIC.neg;
        break;
      case "target":
        color = SEMANTIC.ink;
        break;
      case "smooth":
        color = palette[0];
        break;
      default:
        color = palette[next++ % palette.length];
    }

    if (role === "hidden") {
      s.format.fill.clear();
      s.format.line.clear();
      s.hasDataLabels = false;
      return;
    }
    if (!/^(increase|decrease|total) \(-\)/i.test(s.name)) visibleCount++;

    if (line) {
      if (role === "outlier") {
        s.format.line.lineStyle = Excel.ChartLineStyle.none;
        s.markerStyle = Excel.ChartMarkerStyle.circle;
        s.markerSize = 7;
        s.markerBackgroundColor = color;
        s.markerForegroundColor = color;
      } else if (role === "target") {
        s.format.line.lineStyle = Excel.ChartLineStyle.none;
        s.markerStyle = Excel.ChartMarkerStyle.dash;
        s.markerSize = 16;
        s.markerBackgroundColor = color;
        s.markerForegroundColor = color;
      } else {
        s.format.line.color = color;
        s.format.line.weight = role === "smooth" ? 2.75 : role === "data" && hasSmooth ? 1.25 : 2;
        if (role === "band" || role === "ref") s.format.line.lineStyle = Excel.ChartLineStyle.dash;
        s.markerStyle = Excel.ChartMarkerStyle.none;
      }
    } else {
      s.format.fill.setSolidColor(color);
      s.format.line.clear();
      if (isBarLike(type)) s.gapWidth = 60;
    }

    if (wantLabels && !line && role !== "band" && role !== "ref") {
      s.hasDataLabels = true;
      s.dataLabels.numberFormat = SHORT_NUMBER_FORMAT;
      s.dataLabels.format.font.size = 9;
      s.dataLabels.format.font.color = isStacked(type) ? "#FFFFFF" : "#404040";
    }
  });

  chart.legend.visible = visibleCount > 1;
  if (visibleCount > 1) chart.legend.position = Excel.ChartLegendPosition.top;
  await context.sync();
}
