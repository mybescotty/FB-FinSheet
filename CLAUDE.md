# FB-FinSheet — Financial Dashboard Builder

How the Financial Dashboard Builder works. It is written for whoever maintains it next, person or Claude. Claude Code loads this file automatically in future sessions.

## What it is

`Financial_Dashboard_Builder.html` is one self-contained HTML file (~370 KB, ~2,070 lines). The Franchise Finance Team runs it on their own machine. It has **no build step, no server, no libraries and no network calls**. Everything is vanilla HTML, CSS and JavaScript using built-in browser features.

It does two jobs:

1. **The Builder** (internal, for Finance only). You load the monthly *Depot Financial Reporting* `.xlsx` workbook, plus an optional *Driver Roadmaps* workbook. It checks the workbooks, then generates a dashboard file.
2. **The generated dashboard** (for General Managers). This is another standalone HTML file with one depot's data baked in. GMs double-click it to open it, with no upload and no login. A second button makes a *Finance* version that holds every depot and has a depot dropdown.

The dashboard's code lives **inside** the Builder as one big JavaScript string, `GM_DASHBOARD_TEMPLATE`. Generating a dashboard means: parse the Excel file → build a JSON payload → paste it into that string → download the result.

Banner text in the app: *"Internal tool — do not distribute this file to General Managers. Give them only the dashboard file it generates."*

## File map (`Financial_Dashboard_Builder.html`)

| Lines | What |
|---|---|
| 1–71 | `<head>` and the Builder's CSS. Purple brand colours are CSS variables in `:root` (`--purple-900` etc.). |
| 75–82 | Header. **Line 76 is a ~33 KB base64 PNG logo** on one line. Don't open it in an editor that struggles with long lines. |
| 84–145 | The 5 step cards: `step-1` upload, `step-2` validation, `step-3` depot picker, `step-3b` waterfall workbook (optional), `step-4` page checkboxes + the two Generate buttons. |
| 147–424 | **`LocalXLSX`**, a small home-made `.xlsx` reader (details below). |
| 426–2064 | **Builder logic**: parsing, validation and payload building. |
| **2065** | **`GM_DASHBOARD_TEMPLATE`**: the whole GM dashboard (~225 KB HTML/CSS/JS) as **one JSON-escaped string on one line**. |
| 2066–2068 | Closing tags. |

### How a step "unlocks"
Each `.step` is faded out (`opacity:.45; pointer-events:none`) until JavaScript adds the class `active` (or `done`, which turns the number badge green). `runValidation()` unlocks steps 3, 3b and 4 once the main workbook passes.

## Part 1 — `LocalXLSX` (lines 147–424)

The author chose not to bundle SheetJS (it would be a supply-chain/audit risk). Instead this reads the `.xlsx` ZIP directly:

- `unzip()` finds the ZIP central directory by hand and inflates entries with the browser's `DecompressionStream('deflate-raw')`. This needs a current Edge or Chrome.
- `read()` reads `xl/workbook.xml`, its rels, `sharedStrings.xml`, then each sheet. It also follows each sheet's **drawing** relationship, so text-box commentary (`__drawingXml`) can be read.
- `parseSheetXml()` uses regex over the sheet XML. It keeps cell values plus each row's **`outlineLevel` and `hidden`** flags. The P&L tree comes from Excel's own grouping (the [+]/[-] buttons).
- `sheet_to_json()` mimics SheetJS. With `{header:1}`, `raw[r]` is the **1-indexed Excel row** (`raw[0]` is unused) and columns are 0-indexed (A=0).

Limits: it reads cached values only (no formula evaluation), ignores number formats and dates, and has no ZIP64 support.

## Part 2 — Builder logic (lines 426–2064)

### State
`builderState` (line 429) holds the loaded workbook, depot sheet names, selected depot, the Control-tab values (period, forecast label, week counts), the parsed Risk/AOB/Guide/Agenda tabs, the waterfall data and the enabled pages.

`DASHBOARD_PAGES` (line 437) **must be kept in sync by hand** with the template's own `PAGES` array.

### Workbook tabs it expects

| Tab | Required? | Read by | Used for |
|---|---|---|---|
| `Dashboard requirement` | yes (only checked for existence) | `runValidation` | — |
| `Control` | yes | `getControlPeriod`, `getControlPeriodWeeks`, `getControlForecastLabel` | Current period (e.g. `2027M4`): the row where col B = `Period`, then looked up in a `Period`/`Ref`/`Weeks in period` table that is **found by scanning**, not at fixed columns. Forecast label: the `Column Header` row marked `<<`. |
| `Depot-*` (one per depot) | at least one | `parseDepotSheet` | P&L, trends, KPIs, commentary |
| Risk & Focus (fuzzy match on name containing "risk" + "focus", fallback `Risk&Opps`) | no | `sheetToObjectsWithHeaderSearch(…,'Depot')` | Columns: `Depot`, `Period`, `Type`, `Risk/Focus`, `Commentary` |
| `AOB` | no | same | Each column header after Depot/Period is a category. Text after " eg " in a header is cut off when shown. |
| `Guide` | no | `parseGuideSheet` | `Category`, `KPI`, then the 3rd column (read by position) = "how it's calculated" |
| `Agenda` | no | header `Agenda Item` | `Agenda Item`, `Who` |

Rows with `Depot = "All"` show on every depot (`matchesDepotGroup`).

### `parseDepotSheet` (line 945): the core parser
Layout of a `Depot-*` sheet (0-indexed columns):

- **D1** = depot name (`get(1,3)`). Falls back to the sheet name minus `Depot-`.
- **P&L block**: from row 6 down to the first row with col C = `P&L Trend`.
  - `PERIOD_COLS` F–O (5–14): actual, forecast, vs fcst £/%, budget, vs budget £/%, last year, vs LY £/%.
  - `YTD_COLS` Q–Z (16–25): the same ten fields.
  - `FOCUS_AREA_COLS` AB–AE (27–30): current, YTD average, variance £/%. These feed the Overview "Focus Areas" card.
  - `HISTORY_COLS` AC–AH (28–33): 6 months of prior actuals. **Note: these overlap `FOCUS_AREA_COLS`** (28–30). This looks like it was left behind by a layout change, so check it against the real workbook. History units are auto-detected (×1000 or ÷1000) from the median ratio to the period actual.
  - The header labels (e.g. "Actuals", "Vs Forecast") are found by scanning rows 1–20 for a cell starting "actual" in col F. The row above that holds "Period 4"/"YTD" and the history month names.
  - Tree: built from Excel `outlineLevel`. Outer groups (Revenue/Direct/Semi-Direct/Overheads…) are split on blank rows and named after their last bold row (`isBoldPnlLabel`: starts with "total ", contribution, ebit…).
  - Rows labelled "forecast scenario"/"forecast"/"graph period" are skipped (`PNL_JUNK_LABELS`).
- **P&L Trend section** (col C = `P&L Trend`) and **KPIs section** (col E = `KPIs`), both read by `parseGroups` (line 736). Each block contains:
  - `Current Year` header row (col B), with the block title in E and month labels like `2027M1` in F–Q.
  - `CY Actual`/`Actual` rows, then `CY Forecast`/`Forecast` rows, paired by the code in col D.
  - `Prior Year actual` header, then `PY Actual` rows.
  - 12 values in F–Q = P1..P12.
- Actuals **after the current period are blanked** (`truncateFutureActuals`). "Now" is found by matching the Control period against the CY month-label row.
- **Percent fix-up**: KPIs whose label contains `lost miles`, `sick hrs`, `sick hours` or `efficiency` *and* whose values are all below 1 are multiplied by 100 (`FRACTION_SCALED_KPI_KEYWORDS`). Add new problem KPIs to this list.
- **Commentary**: text boxes on the sheet are split into sections (paragraphs ending "…Commentary") and items ("Label: text" or "Label - text"). They are attached to P&L rows by fuzzy word overlap (`bestCommentaryMatch`): the first word counts double, and a match needs ≥50% coverage. Each group's total row gets the full section's commentary. Commentary only attaches to top-level rows, never to children.

### Driver Cost Waterfall workbook (lines 1159–1298, step 3b)
- `Depot-*` sheets. Two "staging ranges" (Financials, then Hours) are found by scanning columns T–AO for the text `Forecast`. Each is read down to the row labelled `Actual`.
- `E7` = period basis (e.g. "YTD").
- "Scheduling" is renamed to "Scheduling Efficiency".
- `WATERFALL_CALC_GUIDE` holds hover text per bar. **This is hard-coded wording**, so edit it here.
- The workbook is rejected if no depot has a staging range (protects against someone uploading the main workbook here by mistake).

### Validation (lines 1383–1549)
`runValidation` checks the required tabs, reads all shared data, collects warnings, and runs `runStructuralChecks`. That check confirms each depot has P&L rows, a `P&L Trend` section, and `Revenue KPIs` / `Driver KPIs` / `Engineering KPIs` tags in col C. It then fills the depot picker.

### Generate (lines 1551–2064)
- **Single depot**: `buildEmbeddedData()` → JSON → replaces `/*__EMBEDDED_DATA__*/null` in the template → downloads `Depot_Financial_Dashboard_<Depot>_<Period>.html`.
- **Finance, all depots**: `buildFinanceEmbeddedData()` → `{depots:{name:slice}, riskOpp, aob, guide, agenda, …}` → replaces `/*__EMBEDDED_MULTI_DATA__*/null` → `Finance_Dashboard_<Period>.html`.
- `</script` in the JSON is escaped so it can't break out of the script tag.
- `splitNonFinancialGroup` removes the last P&L group as "Non-Financials" when the group before it contains EBIT.
- **Combined depots**: `COMBINED_DEPOT_PAIRS` (line 1818) currently holds Swansea + Cymru West → "Swansea & Cymru West". The two depots are hidden from the picker and replaced by a group option. That group produces a multi-depot file containing both depots plus a virtual combined depot. The combined figures come from the `Depot-SWN&CYW` tab if it exists. Otherwise the code sums the two depots row by row (`sumPnlNodes`, `sumTrendGroups`, `sumWaterfall`). In that case there are no KPI pages, and percentages are recalculated from the summed £ figures.

### Payload shape (what the template receives)
```
depotName, period, forecastLabel, currentPeriodWeeks, ytdPriorWeeks,
pnl[], pnlTree[], pnlNonFinancial[],           // node: {row,code,label,bold,outlineLevel,hidden,period{},ytd{},history[],focusArea{},commentary?,children[]}
periodHeaders, ytdHeaders, periodLabel, ytdLabel, historyMonths,
trend[], kpis[],                                // group: {name, items:[{code,label,actual[12],forecast[12],pyActual[12]}]}
months, cyMonths, pyMonths,
riskOpp[], aob[], aobCategories[], guide[], agenda[],
waterfall {periodBasis, financials[{category,value}], hours[]}, waterfallGuide,
enabledPages[], riskAobGroup[]
```

## Part 3 — The GM dashboard (`GM_DASHBOARD_TEMPLATE`, line 2065)

When decoded, it is ~2,625 lines: CSS (lines 7–386), a small body shell (header, `#app-nav`, `#pages-container`), and one script. **Line numbers below refer to the decoded template**, not the Builder file.

- **Boot** (end of file): if `EMBEDDED_MULTI_DEPOT_DATA` is set, it calls `ingestMultiDepotData` (adds a depot `<select>` to the header). Otherwise, if `EMBEDDED_DEPOT_DATA` is set, it calls `ingestDepotData`. If neither is set, it shows a "no data embedded" message.
- `PAGES` (line 421) → `buildNav` / `buildPagesShell` → `renderActivePage` switches on the page id. Pages are re-rendered each time you click their tab. The Waterfall page is hidden if there is no waterfall data. Pages unticked in the Builder are hidden too.
- **RAG**: `ragStatus` (≈line 764) uses `RAG_TOLERANCE = 0.02` (±2% = amber). Cost lines are judged on magnitude ("lower is better").

| Page id | Renderer | Shows |
|---|---|---|
| `agenda` | `renderAgenda` | Agenda table |
| `overview` | `renderOverview` | Narrative "Depot Summary" for the 5 `PNL_GRAND_TOTAL_LINES`. **Focus Areas** card ranks 6 hard-coded cost lines (`PNL_TREND_LINES`) red/amber/green against their YTD average. Also the P&L Definitions card (hard-coded text). |
| `pnlsummary` | `renderSummaryPnl` | The 8 `SUMMARY_PNL_LINES` |
| `pnlperiod` / `pnlytd` | `renderDetailedPnl` | Expandable P&L tree. Forecast/Budget/Last Year column groups can be toggled. Commentary appears on hover. Revenue/Direct/Semi-Direct rows are wrapped under their totals for display only (`groupPnlLinesUnderTotals`). |
| `pnltrend` | `renderPnlTrend` | P1–P12 line charts per trend group (dropdown) |
| `revenue` / `driver` / `engineering` | `renderKpiPage` | KPI picker, a CY vs LY vs Forecast chart, run-rate text, and the Guide text (`KPI_PAGE_META`) |
| `waterfall` | `renderWaterfall` | Two canvas waterfall charts (£'000 and hours) with hover text |
| `aob` | `renderAob` | One card per AOB category |
| `risks` | `renderRisks` | Risk/Focus cards |

All charts are drawn by hand on `<canvas>` (`drawLineChart`, `buildWaterfallChart`) with custom tooltips. There is no charting library. Chart colours are constants near line 2515.

## Working on it

### Editing the dashboard template
The template is a JSON string on one line, so editing it directly is error-prone. To extract it for reading:

```bash
sed -n '2065p' Financial_Dashboard_Builder.html | python3 -c "
import sys,json; l=sys.stdin.read().strip(); s=l[l.index('=')+1:].strip().rstrip(';')
open('template.html','w',errors='replace').write(json.loads(s))"
```

**Recommended before any big change:** split the template into its own `src/gm_dashboard_template.html`, and add a small build script that JSON-encodes it back into the Builder. Then template edits are normal HTML/JS edits. (Not done yet; this needs agreeing with the owner first.)

### Things to remember
- Keep it **dependency-free and offline**. That is a deliberate requirement: no CDNs, fonts or `fetch`.
- `DASHBOARD_PAGES` (Builder) and `PAGES` (template) must match.
- Most workbook positions are **found by scanning** for marker text rather than hard-coded, because Finance changes the layout often. Keep that pattern. The fixed-column exceptions are `PERIOD_COLS`, `YTD_COLS`, `HISTORY_COLS`, `FOCUS_AREA_COLS` and `SERIES_COLS`.
- Many business rules are hard-coded label lists matched in lower case: `PNL_GRAND_TOTAL_LINES`, `PNL_GOLD_HIGHLIGHT_LINES`, `SUMMARY_PNL_LINES`, `PNL_TREND_LINES`, `FRACTION_SCALED_KPI_KEYWORDS`, `COMMENTARY_SYNONYMS`, `WATERFALL_CALC_GUIDE`, `COMBINED_DEPOT_PAIRS`. Renamed rows in the workbook usually mean editing one of these.
- The code comments are long and record *why* each rule exists (often "confirmed directly with the user"). Read them before changing behaviour.
- Testing: there are no automated tests. Open the Builder in Edge or Chrome, load a real workbook, generate a dashboard, and click through every page. Check the browser console for errors.

### Known issues found during handover
1. **Mixed text encoding.** The file is UTF-8 but contains ~30 stray Windows-1252 bytes (`—` 0x97, `£` 0xA3, `–`, `±`). Most sit in comments, but two reach the screen and show as `�`:
   - P&L Definitions: "Total Operating Costs **–** Continuing Operations"
   - KPI run-rate text: "P1**–**P4"

   Fix by re-saving those characters as UTF-8 or as `–`-style escapes.
2. `getSheetHeaders` is defined twice (Builder lines 520 and 560). They are identical in effect, so one can be deleted.
3. `HISTORY_COLS` (28–33) overlaps `FOCUS_AREA_COLS` (27–30). Check which one matches the current workbook. The P&L 6-month history and the Focus Areas figures may be reading the same cells.
4. `findDepotSheetByName` and the picker call `parseDepotSheet` again for every sheet. This is slow with many depots but works correctly.
