<p align="center">
  <img src="assets/icon-128.png" width="96" alt="ROLLSI logo" />
</p>

<h1 align="center">ROLLSI</h1>

<p align="center">
  <b>Moving averages, rolling stats, growth, anomaly detection and forecasting, as native Excel functions.</b><br/>
  Type <code>=ROLLSI.</code> in any cell. Works in Excel for Windows, Mac and the web.
</p>

```excel
=ROLLSI.SMA(B2:B366, 7)                       → 7-day moving average (spills the whole column)
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

### Moving averages
| Function | What it does |
| --- | --- |
| `SMA(values, window, [partial])` | Simple moving average over the last *window* rows |
| `EMA(values, span, [partial], [alpha])` | Exponential moving average, `alpha = 2/(span+1)`, same as pandas `ewm(adjust=False)` |
| `WMA(values, window, [partial])` | Linearly weighted moving average (newest row weighs most) |
| `LAST(values, n, [stat])` | One value from the last *n* numbers: the latest 7-point average, sum, max… |

### Rolling statistics
| Function | What it does |
| --- | --- |
| `ROLLING(values, window, [stat], [partial])` | Any rolling stat: `avg` (default), `sum`, `min`, `max`, `median`, `stdev`, `var`, `count`, `product` |
| `ROLLSUM` / `ROLLMIN` / `ROLLMAX` / `ROLLMEDIAN` / `ROLLSTDEV` `(values, window, [partial])` | Shortcuts for the common ones |
| `ROLLDAYS(dates, values, days, [stat], [partial])` | **Date-aware** window: the last *days* calendar days. Handles gaps, unsorted data and many rows per day. `stat = "perday"` gives sum ÷ days (missing days count as 0) |

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
- **Warm-up rows**: rolling functions leave the first `window-1` rows blank. Pass `partial = TRUE` to compute them from the rows available.
- **Blanks inside a window** are skipped; the stat uses the numbers that are present.
- **Standard deviation / variance** are sample statistics (same as `STDEV.S` / `VAR.S`).
- **Month arithmetic** is month-end aware: for MoM, Feb 28 is compared with Jan 31.
- Bad input returns `#VALUE!` with an explanation you can see by hovering over the cell.

---

## Install for your whole team (one time, then it just works)

ROLLSI is hosted as static files on GitHub Pages:
**Manifest URL:** `https://yashjindal11.github.io/rollsi/manifest.xml`

### Option A: Microsoft 365 admin center (recommended)
A Microsoft 365 admin does this **once**. ROLLSI then loads automatically for every assigned user, every time they open Excel on Windows, Mac or the web. Nobody needs to install anything.

1. Go to <https://admin.microsoft.com> → **Settings → Integrated apps → Upload custom apps**.
2. App type **Office Add-in** → **Provide link to manifest file** → paste the manifest URL above → **Validate**.
3. Assign it to **Entire organization**, or to a specific group such as your analytics team → **Deploy**.
4. It can take up to 24 hours to appear. Users will see a **ROLLSI** group on the **Home** tab, and `=ROLLSI.` autocompletes in every workbook.

### Option B: Each person installs it themselves (no admin needed)
- **Excel on the web:** Home → Add-ins → More Add-ins → My Add-ins → **Upload My Add-in** → choose `manifest.xml` (download it from the URL above).
- **Excel for Mac:** download `manifest.xml` into
  `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create the `wef` folder if needed), restart Excel, then Home → Add-ins → **ROLLSI**.
- **Excel for Windows:** put `manifest.xml` in a shared network folder, add that folder in File → Options → Trust Center → Trust Center Settings → **Trusted Add-in Catalogs** (tick *Show in Menu*), restart Excel, then Home → Add-ins → **Shared Folder** → ROLLSI.

### Updates are automatic
Every push to `main` is tested and redeployed to GitHub Pages by [the workflow](.github/workflows/deploy.yml). Everyone gets the new functions the next time Excel loads them. You only need to re-upload the manifest if `manifest.xml` itself changes.

> **Sharing workbooks:** anyone who **recalculates** a workbook that uses ROLLSI needs the add-in installed, otherwise they see `#NAME?`. To send results outside the team, copy and use **Paste Values**.

---

## Try it in 10 seconds
Click **Home → ROLLSI Functions → Create demo sheet**. You get 120 days of sample sales data (with missing days, a spike and a dip) and every major function applied next to it, plus a KPI block, a 14-day forecast and a chart.

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
src/taskpane/    function browser + "Create demo sheet"
manifest.xml     add-in manifest (localhost URLs are rewritten to ROLLSI_URL on build)
```

### Adding a function
1. Implement it in `src/core/` and add tests in `test/unit/core.test.ts`.
2. Add a wrapper with a `@customfunction NAME` JSDoc block in `src/functions/functions.ts`.
3. Add it to `src/taskpane/catalog.ts` and to this README.

PRs are welcome. Ideas: seasonal decomposition, rolling correlation, Holt-Winters, drawdown, business-day windows.

## License
[MIT](LICENSE)
