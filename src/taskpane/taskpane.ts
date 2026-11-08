/* global atob, Blob, ClipboardItem, document, Excel, HTMLButtonElement, HTMLElement, HTMLInputElement, HTMLSelectElement, navigator, Office */
import { CATALOG } from "./catalog";
import type { FunctionDoc } from "./catalog";
import type { DiffMode } from "../core/viz";
import {
  addCagrArrow,
  addDifferenceArrow,
  addTotals,
  refreshAnnotations,
  removeAnnotations,
} from "./lib/annotate";
import {
  buildActualVsTarget,
  buildControl,
  buildGantt,
  buildPareto,
  buildTrendMA,
  buildWaterfall,
} from "./lib/charts";
import { createDemo, DEMO_SHEET } from "./lib/demo";
import { EXPORT_SIZES, exportCharts } from "./lib/exporter";
import { drawMekko, redrawMekko } from "./lib/mekko";
import { currentPalette, loadPaletteChoice, PALETTES, savePaletteChoice } from "./lib/palette";
import { applyStyle } from "./lib/style";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!
  );
}

function numberInput(id: string): number | null {
  const v = $<HTMLInputElement>(id).value.trim();
  return v === "" ? null : Number(v);
}

function setStatus(text: string, kind: "" | "ok" | "error" = "") {
  const status = $("status");
  status.className = kind;
  status.textContent = text;
}

/** Runs an Excel action with a busy button and a friendly status line. */
async function act(
  button: HTMLButtonElement | null,
  busy: string,
  fn: (ctx: Excel.RequestContext) => Promise<string | void>
) {
  if (button) button.disabled = true;
  setStatus(busy);
  try {
    const result: { message?: string } = {};
    await Excel.run(async (ctx) => {
      result.message = (await fn(ctx)) || undefined;
    });
    setStatus(result.message || "Done.", "ok");
  } catch (e) {
    setStatus(e instanceof Error ? e.message : String(e), "error");
  } finally {
    if (button) button.disabled = false;
  }
}

// ---------------------------------------------------------------- Functions tab

function renderFunctions(filter: string) {
  const list = $("list");
  const q = filter.trim().toLowerCase();
  const matches = (f: FunctionDoc) =>
    !q || [f.name, f.group, f.summary, f.syntax].some((s) => s.toLowerCase().includes(q));
  const groups = new Map<string, FunctionDoc[]>();
  CATALOG.filter(matches).forEach((f) => groups.set(f.group, [...(groups.get(f.group) ?? []), f]));
  list.innerHTML = groups.size
    ? [...groups]
        .map(
          ([group, fns]) =>
            `<h2>${escapeHtml(group)}</h2>` +
            fns
              .map(
                (f) =>
                  `<div class="card"><div class="syntax">${escapeHtml(f.syntax)}</div>` +
                  `<p>${escapeHtml(f.summary)}</p><code>${escapeHtml(f.example)}</code></div>`
              )
              .join("")
        )
        .join("")
    : "<p>No matching functions.</p>";
}

// ---------------------------------------------------------------- Style tab

function renderSwatches() {
  $("swatches").innerHTML = currentPalette()
    .map((c) => `<span style="background:${escapeHtml(c)}" title="${escapeHtml(c)}"></span>`)
    .join("");
}

function initPalette() {
  const select = $<HTMLSelectElement>("palette");
  const custom = $<HTMLInputElement>("custom-palette");
  select.innerHTML = [...Object.keys(PALETTES), "Custom"]
    .map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`)
    .join("");
  const choice = loadPaletteChoice();
  select.value = choice.name;
  custom.value = choice.custom.join(", ");
  custom.style.display = choice.name === "Custom" ? "" : "none";
  const save = async () => {
    custom.style.display = select.value === "Custom" ? "" : "none";
    await savePaletteChoice(select.value, custom.value);
    renderSwatches();
  };
  select.addEventListener("change", save);
  custom.addEventListener("change", save);
  renderSwatches();
}

async function styleCharts(context: Excel.RequestContext, all: boolean): Promise<string> {
  const labels = $<HTMLInputElement>("style-labels").checked;
  let charts: Excel.Chart[];
  if (all) {
    const coll = context.workbook.worksheets.getActiveWorksheet().charts;
    coll.load("items");
    await context.sync();
    charts = coll.items;
  } else {
    const chart = context.workbook.getActiveChartOrNullObject();
    await context.sync();
    if (chart.isNullObject) throw new Error("Click a chart first (or use 'All charts on sheet').");
    charts = [chart];
  }
  for (const c of charts) await applyStyle(context, c, currentPalette(), { labels });
  return `Styled ${charts.length} chart${charts.length === 1 ? "" : "s"}.`;
}

// ---------------------------------------------------------------- Export tab

function base64ToBlob(b64: string): Blob {
  return new Blob([Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))], { type: "image/png" });
}

function renderGallery(images: { name: string; base64: string }[]) {
  const gallery = $("gallery");
  gallery.innerHTML = images
    .map((img, i) => {
      const file = `${img.name.replace(/[^\w-]+/g, "_")}.png`;
      return (
        `<figure><img src="data:image/png;base64,${img.base64}" alt="${escapeHtml(img.name)}" draggable="true" />` +
        `<figcaption><span>${escapeHtml(img.name)}</span>` +
        `<a href="data:image/png;base64,${img.base64}" download="${escapeHtml(file)}">Download</a>` +
        `<button class="secondary" data-copy="${i}">Copy</button></figcaption></figure>`
      );
    })
    .join("");
  gallery.querySelectorAll<HTMLButtonElement>("button[data-copy]").forEach((btn) =>
    btn.addEventListener("click", async () => {
      try {
        const blob = base64ToBlob(images[Number(btn.dataset.copy)].base64);
        await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
        setStatus("Copied. Paste into PowerPoint with Ctrl/Cmd+V.", "ok");
      } catch {
        setStatus(
          "Copy is blocked here. Right-click the image and copy, or drag it into PowerPoint.",
          "error"
        );
      }
    })
  );
}

// ---------------------------------------------------------------- Wiring

const ACTIONS: Record<string, (ctx: Excel.RequestContext) => Promise<string | void>> = {
  trend: async (ctx) => {
    await buildTrendMA(ctx, numberInput("ma-window") ?? 7);
    return "Trend chart added. Its formulas live on the ROLLSI Data sheet.";
  },
  control: async (ctx) => {
    await buildControl(ctx, numberInput("ctl-window"), numberInput("ctl-k") ?? 3);
    return "Control chart added.";
  },
  waterfall: async (ctx) => {
    await buildWaterfall(
      ctx,
      $<HTMLInputElement>("wf-start").checked,
      $<HTMLInputElement>("wf-total").checked
    );
    return "Waterfall added. Re-insert to refresh the step labels after data changes.";
  },
  avt: async (ctx) => {
    await buildActualVsTarget(ctx);
    return "Actual vs target and variance charts added.";
  },
  pareto: async (ctx) => {
    await buildPareto(ctx);
    return "Pareto chart added.";
  },
  gantt: async (ctx) => {
    await buildGantt(ctx);
    return "Timeline added.";
  },
  mekko: async (ctx) => {
    await drawMekko(ctx);
    return "Mekko added. Click Annotate → Refresh all after changing the data.";
  },
  diff: async (ctx) => {
    await addDifferenceArrow(
      ctx,
      numberInput("diff-from") ?? 1,
      numberInput("diff-to") ?? 2,
      $<HTMLSelectElement>("diff-mode").value as DiffMode
    );
    return "Difference arrow added.";
  },
  cagr: async (ctx) => {
    await addCagrArrow(ctx, numberInput("cagr-from"), numberInput("cagr-to"));
    return "CAGR arrow added.";
  },
  totals: async (ctx) => {
    await addTotals(ctx);
    return "Totals added.";
  },
  refresh: async (ctx) => {
    const n = await refreshAnnotations(ctx, (source, left, top) =>
      redrawMekko(ctx, source, left, top)
    );
    return `Refreshed ${n} annotation${n === 1 ? "" : "s"}.`;
  },
  remove: async (ctx) => {
    const n = await removeAnnotations(ctx);
    return `Removed ${n} annotation${n === 1 ? "" : "s"}.`;
  },
  "style-one": (ctx) => styleCharts(ctx, false),
  "style-all": (ctx) => styleCharts(ctx, true),
  export: async (ctx) => {
    const scope = $<HTMLSelectElement>("export-scope").value as "selected" | "sheet";
    const size = EXPORT_SIZES[Number($<HTMLSelectElement>("export-size").value)];
    const images = await exportCharts(ctx, scope, size);
    renderGallery(images);
    return `${images.length} image${images.length === 1 ? "" : "s"} ready below.`;
  },
};

Office.onReady(() => {
  document.querySelectorAll<HTMLButtonElement>(".tabs button").forEach((tab) =>
    tab.addEventListener("click", () => {
      document
        .querySelectorAll(".tabs button, .tab")
        .forEach((el) => el.classList.remove("active"));
      tab.classList.add("active");
      $(`tab-${tab.dataset.tab}`).classList.add("active");
      setStatus("");
    })
  );
  document
    .querySelectorAll<HTMLButtonElement>("button[data-action]")
    .forEach((btn) =>
      btn.addEventListener("click", () => act(btn, "Working…", ACTIONS[btn.dataset.action!]))
    );
  $<HTMLSelectElement>("export-size").innerHTML = EXPORT_SIZES.map(
    (s, i) => `<option value="${i}">${escapeHtml(s.label)}</option>`
  ).join("");
  $("demo").addEventListener("click", (e) =>
    act(e.currentTarget as HTMLButtonElement, "Building demo sheet…", async (ctx) => {
      await createDemo(ctx);
      return `Done. See the "${DEMO_SHEET}" sheet.`;
    })
  );
  $("search").addEventListener("input", (e) =>
    renderFunctions((e.target as HTMLInputElement).value)
  );
  renderFunctions("");
  initPalette();
});
