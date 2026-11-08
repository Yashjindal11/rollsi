/* global Excel */
import { mekkoLayout, shortNumber, toNum } from "../../core";
import { apiSupported, groupShapes, requireApi, TAG } from "./excel";
import { currentPalette, textColorOn } from "./palette";

const W = 480;
const H = 300;

function textBox(
  shapes: Excel.ShapeCollection,
  text: string,
  left: number,
  top: number,
  width: number,
  opts: { size?: number; bold?: boolean; color?: string; align?: "Left" | "Center" } = {}
): Excel.Shape {
  const s = shapes.addTextBox(text);
  s.left = left;
  s.top = top;
  s.width = width;
  s.height = (opts.size ?? 8) * 2.6 * Math.max(1, text.split("\n").length) * 0.6 + 4;
  s.fill.clear();
  s.lineFormat.visible = false;
  const tf = s.textFrame;
  tf.leftMargin = 0;
  tf.rightMargin = 0;
  tf.topMargin = 0;
  tf.bottomMargin = 0;
  tf.horizontalAlignment =
    opts.align === "Left"
      ? Excel.ShapeTextHorizontalAlignment.left
      : Excel.ShapeTextHorizontalAlignment.center;
  tf.textRange.font.size = opts.size ?? 8;
  tf.textRange.font.bold = opts.bold ?? false;
  tf.textRange.font.color = opts.color ?? "#404040";
  return s;
}

/**
 * Draws a Marimekko chart from a block where the header row holds categories (columns, widths) and
 * the first column holds segments (stacked shares). Everything is grouped and tagged so
 * "Refresh annotations" can redraw it from the source range.
 */
export async function drawMekko(
  context: Excel.RequestContext,
  source?: { sheet: string; address: string },
  at?: { left: number; top: number }
): Promise<void> {
  requireApi("1.9", "Mekko charts");
  const sheet = source
    ? context.workbook.worksheets.getItem(source.sheet)
    : context.workbook.worksheets.getActiveWorksheet();
  let range: Excel.Range;
  if (source) {
    range = sheet.getRange(source.address);
  } else {
    range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    await context.sync();
    if (range.rowCount === 1 && range.columnCount === 1 && apiSupported("1.13")) {
      range = range.getSurroundingRegion();
    }
  }
  range.load("values,address,left,top,width,rowCount,columnCount");
  sheet.load("name");
  await context.sync();
  if (range.rowCount < 2 || range.columnCount < 2) {
    throw new Error(
      "Select a table: categories across the top row, segments down the first column, values inside."
    );
  }

  const values = range.values;
  const title = String(values[0][0] || "Market map");
  const categories = values[0].slice(1).map(String);
  const segments = values.slice(1).map((r) => String(r[0]));
  const layout = mekkoLayout(values.slice(1).map((r) => r.slice(1).map(toNum)));
  const palette = currentPalette();
  const shapes = sheet.shapes;
  const ox = at?.left ?? range.left + range.width + 24;
  const oy = at?.top ?? range.top;
  const plotTop = oy + 44;
  const parts: Excel.Shape[] = [];

  parts.push(
    textBox(shapes, title, ox, oy, W, { size: 12, bold: true, color: "#262626", align: "Left" })
  );

  layout.cells.forEach((c) => {
    if (c.w <= 0 || c.h <= 0) return;
    const color = palette[c.seg % palette.length];
    const rect = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rect.left = ox + c.x * W;
    rect.top = plotTop + (1 - c.y - c.h) * H;
    rect.width = c.w * W;
    rect.height = c.h * H;
    rect.fill.setSolidColor(color);
    rect.lineFormat.color = "#FFFFFF";
    rect.lineFormat.weight = 1;
    if (c.w * W > 30 && c.h * H > 13) {
      rect.textFrame.textRange.text = `${Math.round(c.share * 100)}%`;
      rect.textFrame.textRange.font.size = 8;
      rect.textFrame.textRange.font.color = textColorOn(color);
      rect.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
      rect.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
    }
    parts.push(rect);
  });

  layout.columns.forEach((col, i) => {
    if (col.w <= 0) return;
    const left = ox + col.x * W;
    const width = Math.max(col.w * W, 30);
    parts.push(
      textBox(shapes, `${Math.round(col.share * 100)}%`, left, plotTop - 14, width, {
        color: "#8C8C8C",
      })
    );
    parts.push(
      textBox(shapes, `${categories[i]}\n${shortNumber(col.total)}`, left, plotTop + H + 4, width)
    );
  });

  const last = layout.columns.length - 1;
  layout.cells
    .filter((c) => c.cat === last && c.h > 0)
    .forEach((c) => {
      const y = plotTop + (1 - c.y - c.h / 2) * H - 6;
      parts.push(
        textBox(shapes, segments[c.seg], ox + W + 6, y, 110, {
          align: "Left",
          color: palette[c.seg % palette.length],
          bold: true,
        })
      );
    });

  const local = range.address.slice(range.address.lastIndexOf("!") + 1);
  await groupShapes(context, shapes, parts, `${TAG}MEKKO|${sheet.name}!${local}`);
}

/** Used by refresh: source is "Sheet!A1:D5". */
export async function redrawMekko(
  context: Excel.RequestContext,
  source: string,
  left: number,
  top: number
) {
  const i = source.lastIndexOf("!");
  await drawMekko(
    context,
    { sheet: source.slice(0, i), address: source.slice(i + 1) },
    { left, top }
  );
}
