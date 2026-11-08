/* global Office, localStorage */

export const PALETTES: Record<string, string[]> = {
  ROLLSI: ["#107C41", "#7FBF9B", "#F2A900", "#2F5597", "#A5A5A5", "#C55A11"],
  "Consulting navy": ["#0B2D4F", "#2E6DA4", "#7FA7CC", "#E5A823", "#9E9E9E", "#C7D6E6"],
  "Mono + accent": ["#E4572E", "#404040", "#7F7F7F", "#A6A6A6", "#C9C9C9", "#E3E3E3"],
  "Colorblind safe": ["#0072B2", "#E69F00", "#009E73", "#CC79A7", "#56B4E9", "#D55E00", "#F0E442"],
};

/** Fixed meaning colours used by waterfalls, variance and control charts. */
export const SEMANTIC = {
  pos: "#2E9E5B",
  neg: "#D64541",
  total: "#5B6770",
  band: "#A6A6A6",
  center: "#7F7F7F",
  ink: "#262626",
};

const KEY = "rollsi.palette";
const CUSTOM_KEY = "rollsi.customPalette";

function settings(): Office.Settings | null {
  try {
    return Office.context.document.settings;
  } catch {
    return null;
  }
}

export function parseHexList(text: string): string[] {
  return text
    .split(/[\s,;]+/)
    .map((s) => s.trim())
    .filter((s) => /^#?[0-9a-f]{6}$/i.test(s))
    .map((s) => (s.startsWith("#") ? s : "#" + s).toUpperCase());
}

/** Palette choice is saved in the workbook (travels with the file) and as this user's default. */
export function loadPaletteChoice(): { name: string; custom: string[] } {
  const s = settings();
  const name = (s?.get(KEY) as string) || localStorage.getItem(KEY) || "ROLLSI";
  const customRaw = (s?.get(CUSTOM_KEY) as string) || localStorage.getItem(CUSTOM_KEY) || "";
  return { name, custom: parseHexList(customRaw) };
}

export function savePaletteChoice(name: string, customText: string): Promise<void> {
  localStorage.setItem(KEY, name);
  localStorage.setItem(CUSTOM_KEY, customText);
  const s = settings();
  if (!s) return Promise.resolve();
  s.set(KEY, name);
  s.set(CUSTOM_KEY, customText);
  return new Promise((resolve) => s.saveAsync(() => resolve()));
}

export function currentPalette(): string[] {
  const { name, custom } = loadPaletteChoice();
  if (name === "Custom" && custom.length) return custom;
  return PALETTES[name] ?? PALETTES.ROLLSI;
}

/** Black or white text, whichever reads better on the given fill. */
export function textColorOn(hex: string): string {
  const n = parseInt(hex.replace("#", ""), 16);
  const r = (n >> 16) & 255;
  const g = (n >> 8) & 255;
  const b = n & 255;
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? "#262626" : "#FFFFFF";
}

/** Small preferences shared by the ribbon commands and the panel (workbook settings + local default). */
export function getPref(key: string, fallback: string): string {
  const s = settings();
  return (s?.get(`rollsi.${key}`) as string) ?? localStorage.getItem(`rollsi.${key}`) ?? fallback;
}

export function setPref(key: string, value: string): Promise<void> {
  localStorage.setItem(`rollsi.${key}`, value);
  const s = settings();
  if (!s) return Promise.resolve();
  s.set(`rollsi.${key}`, value);
  return new Promise((resolve) => s.saveAsync(() => resolve()));
}

/** Ribbon commands run in their own runtime, so re-read settings saved by the panel first. */
export function refreshPrefs(): Promise<void> {
  return new Promise((resolve) => {
    const s = settings();
    if (!s || typeof s.refreshAsync !== "function") return resolve();
    try {
      s.refreshAsync(() => resolve());
    } catch {
      resolve();
    }
  });
}
