<p align="center">
  <img src="assets/icon-128.png" width="96" alt="ROLLSI logo" />
</p>

<h1 align="center">ROLLSI</h1>

<p align="center">
  <b>Moving averages, rolling stats, analyst charts and think-cell-style annotations for Excel.</b><br/>
  Use the <b>ROLLSI</b> ribbon tab, or type <code>=ROLLSI.</code> in any cell. Works in Excel for Windows, Mac and the web.
</p>

```excel
=ROLLSI.MA(B2:B366, 7)                        → 7-day moving average (spills the whole column)
=ROLLSI.MA(B2:B366, 15, 9, 1)                 → 15-row SUM (code 9), centered on each row (align 1)
=ROLLSI.LAST(B2:B1000, 7)                     → average of the last 7 data points (one KPI number)
=ROLLSI.ROLLDAYS(A2:A366, B2:B366, 7, "sum")  → trailing 7 *calendar* days, even with missing days
=ROLLSI.PERIODCHANGE(A2:A366, B2:B366, "YoY") → year-over-year %, matched by date
=ROLLSI.OUTLIERS(B2:B366)                     → TRUE/FALSE anomaly flags
=ROLLSI.FORECAST(B2:B366, 14)                 → next 14 periods (Holt)
```

No VBA, no macros, no copying formulas down 1,000 rows. One formula spills the whole result.

---

## Why

Excel can do a 7-day average with `=AVERAGE(B2:B8)` dragged down, but analysts constantly hit the same problems:

| Problem | Plain Excel | ROLLSI |
| --- | --- | --- |
| Moving average for 5,000 rows | Drag a formula, fix the first 6 rows by hand | `=ROLLSI.SMA(B2:B5001, 7)` |
| Missing days in the data | 7 rows ≠ 7 days, so averages are silently wrong | `=ROLLSI.ROLLDAYS(dates, values, 7)` |
| Week-over-week / YoY | `XLOOKUP(date-7, ...)` gymnastics | `=ROLLSI.PERIODCHANGE(dates, values, "W")` |
| "Latest 7-day average" for a dashboard | `AVERAGE(OFFSET(...COUNTA...))` | `=ROLLSI.LAST(B2:B10000, 7)` |
| EMA, rolling median, rolling stdev, z-scores | Helper columns everywhere | One function each |

## Function reference

All `values` arguments accept a single column or a single row; results spill in the same direction.
Blank cells and text are ignored. Optional arguments are in `[brackets]`.

### ⭐ `MA`: one function for any moving window

```excel
=ROLLSI.MA(values, [window], [stat], [align], [rowsAbove], [rowsBelow], [k], [partial])
```

| Example | Meaning |
| --- | --- |
| `=ROLLSI.MA(B2:B100, 7)` | 7-row moving **average**, trailing (current row = last point) |
| `=ROLLSI.MA(B2:B100, 15, 9)` | 15-row moving **SUM** (code 9) |
| `=ROLLSI.MA(B2:B100, 7, 1, 1)` | 7-row average, **centered** (current row = mid point) |
| `=ROLLSI.MA(B2:B100, 7, 4, 2)` | 7-row **MAX**, forward-looking (current row = first point) |
| `=ROLLSI.MA(B2:B100, , "avg", 3, 3, 1)` | **Custom**: average of 3 rows above + current + 1 row below |
| `=ROLLSI.MA(B2:B100, 30, 16, 0, , , 0.9)` | Rolling 90th percentile (code 16, k = 0.9) |
| `=ROLLSI.MA(B2:B100, 30, "p90")` | Same thing, by name |

**`stat`**: the same codes as Excel's `AGGREGATE`, or a name:

| Code | Stat | Code | Stat |
| --- | --- | --- | --- |
| 1 (or blank/0) | AVERAGE | 11 | VAR.P |
| 2 | COUNT (numbers) | 12 | MEDIAN |
| 3 | COUNTA (non-empty) | 13 | MODE.SNGL |
| 4 | MAX | 14 | LARGE (k) |
| 5 | MIN | 15 | SMALL (k) |
| 6 | PRODUCT | 16 | PERCENTILE.INC (k = 0..1) |
| 7 | STDEV.S | 17 | QUARTILE.INC (k = 0..4) |
| 8 | STDEV.P | 18 | PERCENTILE.EXC (k) |
| 9 | SUM | 19 | QUARTILE.EXC (k = 1..3) |
| 10 | VAR.S | | |

Extra stats by name: `"first"`, `"last"`, `"range"` (max − min), `"geomean"`, `"harmean"`, `"avedev"`, `"skew"`, `"kurt"`.
Stats that need `k` can embed it: `"large:2"`, `"percentile:0.9"`, `"p90"`.

**`align`**: where the current row sits in its window:

| Code | Name | Window for `window = 7` |
| --- | --- | --- |
| 0 (default) | `"last"`: last point (trailing) | 6 rows above + current |
| 1 | `"mid"`: mid point (centered) | 3 above + current + 3 below (even windows put the extra row above) |
| 2 | `"first"`: first point (forward-looking) | current + 6 rows below |
| 3 | `"custom"` | `rowsAbove` above + current + `rowsBelow` below (`window` is ignored) |

If you pass `rowsAbove`/`rowsBelow` without an `align`, custom is assumed.

### Moving averages
| Function | What it does |
| --- | --- |
| `SMA(values, window, [partial])` | Simple moving average over the last *window* rows |
| `EMA(values, span, [partial], [alpha])` | Exponential moving average, `alpha = 2/(span+1)`, same as pandas `ewm(adjust=False)` |
| `WMA(values, window, [partial])` | Linearly weighted moving average (newest row weighs most) |
| `LAST(values, n, [stat])` | One value from the last *n* numbers: the latest 7-point average, sum, max… (`stat` as in `MA`) |

### Rolling statistics
| Function | What it does |
| --- | --- |
| `ROLLING(values, window, [stat], [partial])` | Trailing window with any `stat` (same codes/names as `MA`) |
| `ROLLSUM` / `ROLLMIN` / `ROLLMAX` / `ROLLMEDIAN` / `ROLLSTDEV` `(values, window, [partial])` | Shortcuts for the common ones |
| `ROLLDAYS(dates, values, days, [stat], [partial])` | **Date-aware** window: the last *days* calendar days. Handles gaps, unsorted data and many rows per day. `stat` as in `MA`, plus `"perday"` gives sum ÷ days (missing days count as 0) |

### Cumulative
| Function | What it does |
| --- | --- |
| `CUMSUM(values)` | Running total |
| `CUMAVG(values)` | Running average |
| `CUMMAX(values)` / `CUMMIN(values)` | Running high / low (high-water mark) |

### Growth
| Function | What it does |
| --- | --- |
| `PCTCHANGE(values, [periods])` | % change vs *periods* rows earlier |
| `DIFF(values, [periods])` | Difference vs *periods* rows earlier |
| `PERIODCHANGE(dates, values, period, [mode])` | Calendar-matched DoD / WoW / MoM / QoQ / YoY. `period`: `D`, `W`, `M`, `Q`, `Y` (or `DoD`…`YoY`, or a number of days). `mode`: `pct` (default), `diff`, `prior` |
| `REBASE(values, [base])` | Index the series to start at 100 |
| `CAGR(start, end, periods)` | Compound annual growth rate |

### Analysis
| Function | What it does |
| --- | --- |
| `ZSCORE(values)` | Z-score vs the whole series |
| `ROLLZSCORE(values, window)` | Anomaly score vs the **previous** *window* rows |
| `OUTLIERS(values, [method], [threshold])` | TRUE/FALSE flags: `iqr` (default, k=1.5), `z` (3), `mad` (3.5, robust) |
| `PCTRANK(values)` | Percentile rank 0–1 (same as `PERCENTRANK.INC`) |
| `NORMALIZE(values, [method])` | `minmax` (0–1), `z`, or `share` of total |

### Trend & forecast
| Function | What it does |
| --- | --- |
| `TRENDLINE(values)` | Least-squares trend value for every row |
| `FORECAST(values, horizon, [method], [alpha], [beta])` | Next *horizon* periods: `holt` (default, level + trend), `linear`, `ses` |

### Behaviour notes
- **Edge rows**: rows whose window runs past the top (or, for mid/first/custom alignment, the bottom) of the data are left blank. Pass `partial = TRUE` to compute them from the rows available.
- **Blanks inside a window** are skipped; the stat uses the numbers that are present.
- **Standard deviation / variance** are sample statistics (same as `STDEV.S` / `VAR.S`).
- **Month arithmetic** is month-end aware: for MoM, Feb 28 is compared with Jan 31.
- Bad input returns `#VALUE!` with an explanation you can see by hovering over the cell.

---

## The ROLLSI ribbon tab (think-cell style)

Everything happens in Excel's ribbon. There's no separate window and no tabs to switch: **select data → click a button**.

| Group | Buttons |
| --- | --- |
| **Charts** | **Trend ▾** (7 / 14 / 30 / 90-period moving average) · **Control ▾** (±2σ, ±3σ, rolling) · **Waterfall** · **Actual vs Target** · **Pareto** · **Gantt** · **Mekko** |
| **Annotate** | **Difference ▾** (last vs previous, first vs last, every step) · **CAGR** · **Totals** · **Refresh** · **Remove** |
| **Format** | **Style** (selected chart, or every chart on the sheet) · **Palette ▾** (ROLLSI, Consulting navy, Mono + accent, Colorblind safe, Custom…) |
| **Share** | **Export** (PNG at slide sizes) · **Demo** · **Settings & Help** |

If something is missing (for example no data or no chart selected), a small ROLLSI message box appears in Excel. The side panel opens only for **Export** and **Settings & Help**.

### Charts
| Button | Select | What you get |
| --- | --- | --- |
| **Trend** | `[dates] values` | Line + N-period MA, warm-up rows left as gaps (not zeros) |
| **Control** | `[dates] values` | Value, center, ±kσ bands (whole series or rolling), outliers as red dots |
| **Waterfall** | `labels, changes` | Start → steps → total; rows named "Subtotal"/"Total" show running totals; handles zero crossings |
| **Actual vs Target** | `labels, actual, target` | Columns with target markers, plus a green/red variance chart underneath |
| **Pareto** | `labels, values` | Sorted bars, cumulative % on a secondary axis, 80% reference line |
| **Gantt** | `task, start, end` | Horizontal bars from start to end date |
| **Mekko** | matrix: categories across, segments down | Column width = category size, height = segment share (drawn with shapes) |

Charts are **live**: their data comes from formulas (`=ROLLSI.MA`, `=ROLLSI.CONTROL`, `=ROLLSI.WATERFALL`, …) written to a **ROLLSI Data** sheet, so they update when your data changes. Waterfall step labels and Mekko shapes are drawn once; click the button again, or **Refresh**, after big changes.

### Annotations
Click a column or line chart, then **Difference**, **CAGR** or **Totals**. Annotations are grouped shapes tagged with the chart name. **Refresh** redraws every annotation (and Mekko) on the sheet after you resize a chart or change data. **Remove** clears them. Choose percent, absolute or both for difference labels under **Settings & Help**.

### Style & palette
**Style** removes gridlines and borders, mutes the axes, uses K/M number formats, colours series from your palette, shows a legend only when it's needed, and adds direct labels on bars (you can turn labels off in Settings). Paste your company hex codes as a **Custom** palette. The choice is saved in the workbook and as your default.

Series names drive the colours, so you get the same result on any chart: `Increase`/`Over` green, `Decrease`/`Under` red, `Total` grey, `Upper`/`Lower` dashed bands, `Target` marker, names starting with `_` hidden.

### Export
**Export** turns the selected chart, or all charts on the sheet, into PNGs sized for slides (16:9 full slide, half slide, wide strip, or as shown). Then **Copy** and paste into PowerPoint, **Download**, or drag the image into a slide.

### In-cell charts (formulas)
| Function | Example | Result |
| --- | --- | --- |
| `SPARK(values, [type], [min], [max])` | `=ROLLSI.SPARK(B2:B31)` | `▁▂▃▅▇█▆▅` (or `winloss`: `▀▄▀▀`) |
| `BAR(value, max, [width])` | `=ROLLSI.BAR(B2, MAX(B:B))` | `██████▍` |
| `ARROW(change, [threshold], [format], [decimals])` | `=ROLLSI.ARROW(0.123)` | `▲ +12.3%` |
| `SHORT(value, [decimals])` | `=ROLLSI.SHORT(5600000)` | `5.6M` |

### Chart helper formulas
Use these to build your own charts:

| Function | Spills |
| --- | --- |
| `NA(values)` | Same values with blanks as `#N/A`, so charts show gaps instead of zeros |
| `CONTROL(values, [window], [k])` | Center, Upper, Lower, Outlier |
| `PARETO(labels, values)` | Label, Value, Cumulative %, 0.8 |
| `WATERFALL(labels, values, [firstIsStart], [lastIsTotal])` | 8 stacked-column series (hidden bases, Increase, Decrease, Total) |
| `VSTARGET(actual, target)` | Actual, Target, Over, Under |

---

## Install for your whole team (one time, then it just works)

ROLLSI is hosted as static files on GitHub Pages:
**Manifest URL:** `https://yashjindal11.github.io/rollsi/manifest.xml`

### Option A: Microsoft 365 admin center (recommended)
A Microsoft 365 admin does this **once**. ROLLSI then loads automatically for every assigned user, every time they open Excel on Windows, Mac or the web. Nobody needs to install anything.

1. Go to <https://admin.microsoft.com> → **Settings → Integrated apps → Upload custom apps**.
2. App type **Office Add-in** → **Provide link to manifest file** → paste the manifest URL above → **Validate**.
3. Assign it to **Entire organization**, or to a specific group such as your analytics team → **Deploy**.
4. It can take up to 24 hours to appear. Users will see a **ROLLSI** tab in the ribbon, and `=ROLLSI.` autocompletes in every workbook.

### Option B: Each person installs it themselves (no admin needed)
- **Excel on the web:** Home → Add-ins → More Add-ins → My Add-ins → **Upload My Add-in** → choose `manifest.xml` (download it from the URL above).
- **Excel for Mac:** download `manifest.xml` into
  `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create the `wef` folder if needed), restart Excel, then Home → Add-ins → **ROLLSI** (the ROLLSI ribbon tab appears).
- **Excel for Windows:** put `manifest.xml` in a shared network folder, add that folder in File → Options → Trust Center → Trust Center Settings → **Trusted Add-in Catalogs** (tick *Show in Menu*), restart Excel, then Home → Add-ins → **Shared Folder** → ROLLSI.

### Updates are automatic
Every push to `main` is tested and redeployed to GitHub Pages by [the workflow](.github/workflows/deploy.yml). Everyone gets the new functions the next time Excel loads them. You only need to re-upload the manifest if `manifest.xml` itself changes.

> **Sharing workbooks:** anyone who **recalculates** a workbook that uses ROLLSI needs the add-in installed, otherwise they see `#NAME?`. To send results outside the team, copy and use **Paste Values**.

---

## Try it in 10 seconds
Click **ROLLSI → Demo**. You get 120 days of sample sales data (with missing days, a spike and a dip), every major function applied next to it, a KPI block with sparklines and arrows, a 14-day forecast, a live trend + moving-average chart and a revenue bridge (waterfall).

## Use the math from code (npm)
The calculation engine is a dependency-free TypeScript library:

```ts
import { sma, ema, rollDays, periodChange, outliers, forecast } from "rollsi";

sma([1, 2, 3, 4, 5], 3);              // [null, null, 2, 3, 4]
forecast([2, 4, 6, 8], 2, "holt");    // [10, 12]
```

## Development
```bash
npm install
npm test             # unit tests for the math and the Excel wrappers
npm start            # builds, starts https://localhost:3000 and sideloads into Excel desktop
npm run start:web    # same, for Excel on the web
npm run build        # production build into dist/ (set ROLLSI_URL to host somewhere else)
npm run validate:prod
```

Project layout:
```
src/core/        pure TypeScript math (published to npm), no Excel dependency
src/functions/   Excel custom-function wrappers; JSDoc tags generate functions.json
src/commands/    ribbon button handlers + in-Excel message box
src/taskpane/    export/settings panel; lib/ has chart builders, annotations, styling, export, Mekko
scripts/         icon and manifest generators (edit make-manifest.py to change the ribbon)
manifest.xml     generated add-in manifest with the ribbon tab (localhost URLs are rewritten to ROLLSI_URL on build)
```

### Adding a function
1. Implement it in `src/core/` and add tests in `test/unit/core.test.ts`.
2. Add a wrapper with a `@customfunction NAME` JSDoc block in `src/functions/functions.ts`.
3. Add it to `src/taskpane/catalog.ts` and to this README.

PRs are welcome. Ideas: seasonal decomposition, rolling correlation, Holt-Winters, drawdown, business-day windows.

## License
[MIT](LICENSE)
