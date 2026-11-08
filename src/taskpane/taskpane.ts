/* global atob, Blob, ClipboardItem, document, Excel, HTMLButtonElement, HTMLElement, HTMLInputElement, HTMLSelectElement, navigator, Office */
import { CATALOG } from "./catalog";
import type { FunctionDoc } from "./catalog";
import { EXPORT_SIZES, exportCharts } from "./lib/exporter";
import {
  currentPalette,
  getPref,
  loadPaletteChoice,
  PALETTES,
  savePaletteChoice,
  setPref,
} from "./lib/palette";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!
  );
}

function setStatus(text: string, kind: "" | "ok" | "error" = "") {
  const status = $("status");
  status.className = kind;
  status.textContent = text;
}

// ---------------------------------------------------------------- Export

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

async function runExport(button: HTMLButtonElement) {
  button.disabled = true;
  setStatus("Rendering…");
  try {
    const scope = $<HTMLSelectElement>("export-scope").value as "selected" | "sheet";
    const size = EXPORT_SIZES[Number($<HTMLSelectElement>("export-size").value)];
    let images: { name: string; base64: string }[] = [];
    await Excel.run(async (context) => {
      images = await exportCharts(context, scope, size);
    });
    renderGallery(images);
    setStatus(`${images.length} image${images.length === 1 ? "" : "s"} ready.`, "ok");
  } catch (e) {
    setStatus(e instanceof Error ? e.message : String(e), "error");
  } finally {
    button.disabled = false;
  }
}

// ---------------------------------------------------------------- Settings

function renderSwatches() {
  $("swatches").innerHTML = currentPalette()
    .map((c) => `<span style="background:${escapeHtml(c)}" title="${escapeHtml(c)}"></span>`)
    .join("");
}

function initSettings() {
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
    setStatus("Palette saved. Click ROLLSI → Style to apply it.", "ok");
  };
  select.addEventListener("change", save);
  custom.addEventListener("change", save);
  renderSwatches();

  const labels = $<HTMLInputElement>("style-labels");
  labels.checked = getPref("styleLabels", "true") === "true";
  labels.addEventListener("change", () => setPref("styleLabels", String(labels.checked)));

  const diff = $<HTMLSelectElement>("diff-mode");
  diff.value = getPref("diffMode", "pct");
  diff.addEventListener("change", () => setPref("diffMode", diff.value));
}

// ---------------------------------------------------------------- Function reference

function renderFunctions(filter: string) {
  const q = filter.trim().toLowerCase();
  const matches = (f: FunctionDoc) =>
    !q || [f.name, f.group, f.summary, f.syntax].some((s) => s.toLowerCase().includes(q));
  const groups = new Map<string, FunctionDoc[]>();
  CATALOG.filter(matches).forEach((f) => groups.set(f.group, [...(groups.get(f.group) ?? []), f]));
  $("list").innerHTML = groups.size
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

Office.onReady(() => {
  $<HTMLSelectElement>("export-size").innerHTML = EXPORT_SIZES.map(
    (s, i) => `<option value="${i}">${escapeHtml(s.label)}</option>`
  ).join("");
  $("export").addEventListener("click", (e) => runExport(e.currentTarget as HTMLButtonElement));
  $("search").addEventListener("input", (e) =>
    renderFunctions((e.target as HTMLInputElement).value)
  );
  renderFunctions("");
  initSettings();
});
