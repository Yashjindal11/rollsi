/* global Excel, Office, URL, window */
import type { DiffMode } from "../core/viz";
import {
  addCagrArrow,
  addDifferences,
  addTotals,
  refreshAnnotations,
  removeAnnotations,
} from "../taskpane/lib/annotate";
import type { DiffPick } from "../taskpane/lib/annotate";
import {
  buildActualVsTarget,
  buildControl,
  buildGantt,
  buildPareto,
  buildTrendMA,
  buildWaterfall,
} from "../taskpane/lib/charts";
import { createDemo } from "../taskpane/lib/demo";
import { drawMekko, redrawMekko } from "../taskpane/lib/mekko";
import {
  currentPalette,
  getPref,
  loadPaletteChoice,
  refreshPrefs,
  savePaletteChoice,
} from "../taskpane/lib/palette";
import { applyStyle, styleSelectedOrSheet } from "../taskpane/lib/style";

/** Shows a small message box inside Excel and waits until it is closed. */
function showMessage(text: string): Promise<void> {
  return new Promise((resolve) => {
    const url = new URL("dialog.html", window.location.href);
    url.searchParams.set("msg", text);
    Office.context.ui.displayDialogAsync(
      url.toString(),
      { height: 26, width: 30, displayInIframe: true },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return resolve();
        const dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
          dialog.close();
          resolve();
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve());
      }
    );
  });
}

function register(name: string, fn: (context: Excel.RequestContext) => Promise<unknown>) {
  Office.actions.associate(name, async (event: Office.AddinCommands.Event) => {
    try {
      await refreshPrefs();
      await Excel.run(async (context) => {
        await fn(context);
      });
    } catch (e) {
      await showMessage(e instanceof Error ? e.message : String(e));
    } finally {
      event.completed();
    }
  });
}

const diffMode = () => getPref("diffMode", "pct") as DiffMode;
const labels = () => getPref("styleLabels", "true") === "true";

async function setPalette(context: Excel.RequestContext, name: string) {
  await savePaletteChoice(name, loadPaletteChoice().custom.join(", "));
  const chart = context.workbook.getActiveChartOrNullObject();
  await context.sync();
  if (!chart.isNullObject) await applyStyle(context, chart, currentPalette(), { labels: labels() });
}

Office.onReady(() => {
  [7, 14, 30, 90].forEach((n) => register(`trend${n}`, (c) => buildTrendMA(c, n)));
  register("control2", (c) => buildControl(c, null, 2));
  register("control3", (c) => buildControl(c, null, 3));
  register("control3r30", (c) => buildControl(c, 30, 3));
  register("waterfall", (c) => buildWaterfall(c, true, true));
  register("actualVsTarget", buildActualVsTarget);
  register("pareto", buildPareto);
  register("gantt", buildGantt);
  register("mekko", (c) => drawMekko(c));

  const diff = (pick: DiffPick) => (c: Excel.RequestContext) => addDifferences(c, pick, diffMode());
  register("diffLast", diff("last"));
  register("diffFirstLast", diff("firstLast"));
  register("diffAll", diff("all"));
  register("cagr", (c) => addCagrArrow(c));
  register("totals", addTotals);
  register("refresh", (c) =>
    refreshAnnotations(c, (src, left, top) => redrawMekko(c, src, left, top))
  );
  register("remove", removeAnnotations);

  register("style", (c) => styleSelectedOrSheet(c, currentPalette(), labels()));
  register("paletteRollsi", (c) => setPalette(c, "ROLLSI"));
  register("paletteNavy", (c) => setPalette(c, "Consulting navy"));
  register("paletteMono", (c) => setPalette(c, "Mono + accent"));
  register("paletteColorblind", (c) => setPalette(c, "Colorblind safe"));

  register("demo", createDemo);
});
