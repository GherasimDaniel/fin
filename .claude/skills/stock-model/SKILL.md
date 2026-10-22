---
name: stock-model
description: Build a Martin Shkreli-style Excel stock model for a public company (5 years of history, quarterly detail, projections, a "Discounted Cash Flow" row, Bear/Base/Bull scenarios, WACC sensitivity, biggest contracts, sources). Use when the user asks to "make a model", "build a DCF", or "value" a company/ticker, e.g. "/stock-model AMAT".
---

# Shkreli-style stock model

Input: a ticker (args). Output: `C:\Users\gherasid\work\workshop\fin\<Company>\<TICKER>_Model.xlsx`.
Pipeline: **fetch (tool) -> research & fill JSON (you) -> build (tool) -> verify (tool) -> review**.
Worked example of a finished input file: `examples/glw.json` (Corning). Mirror its level of detail.

## 1. Fetch SEC data (tool)
```
node scripts/fetch_sec.js <TICKER> <scratchpad>/<ticker>.json
```
This writes a JSON skeleton containing:
- 5 fiscal years and 10 quarters of GAAP figures from SEC XBRL.
- Balance sheet and cash flow, latest share count, cash and debt, and the 10-yr Treasury yield.
- Placeholder assumptions.
It also prints a **TODO list**. Read it: missing tags (some companies use custom XBRL tags, e.g. Corning's capex) and line items that don't tie must be fixed from the 10-K.

## 2. Research and complete the JSON (you)
Work through `_todo`:
- **Basis.** If the company guides on non-GAAP ("core", "adjusted"), set `basis` and replace `gp/sga/rd/pretax/taxRate/ni/epsRep` with the adjusted figures from the earnings-release reconciliations. Keep `gaapOpInc/gaapNi/gaapEps`. Use `scripts/fetchtext.js <url>` for 8-K exhibits and `scripts/rpage.js` for SEC R-pages.
- **Segments.** Replace "Total company" with the reported segments (`q`, `y`, `note`, `qEstGrowth`, `scen` bear/base/bull, `noYoY` for years that aren't comparable after a reorganization). The segments must sum to the sales line.
- **Market.** `price` (previous close), `dps`, `consensusEps`, `avgTarget`, `beta`, `kd`; check `rf`.
- **Assumptions:**
  - `qEst`: anchor the remaining quarters of this year to company guidance.
  - `proj`: margins, capex, D&A, NWC, non-core cash costs.
  - `gmScen`: gross margin by scenario.
  - Make the Bull scenario match management's long-term targets.
- **Contracts.** Use web search for the biggest customer, supply and partnership deals: counterparty, date, value (`null` if undisclosed, never invented), type, term, scope, why it matters, source.
- **`main`** text (segment descriptions, management, guidance), `rowNotes`, `sources` (with URLs), `methodology`.
- Set `outFile`. Delete `_todo` once it is resolved.

## 3. Build (tool)
```
powershell -File scripts/build_model.ps1 -JsonPath <ticker>.json
```
This creates the Main / Model / Scenarios / WACC / Contracts / Sources tabs. Every cell is a live formula with blue/black/green colour coding. The Model tab includes the "Discounted Cash Flow (PV of UFCF)" row, Gordon-growth and exit-multiple valuations, the Shkreli NPV-of-net-income check and a Bear/Base/Bull data table.

## 4. Verify (tool) - must PASS
```
powershell -File scripts/verify_model.ps1 -XlsxPath <file.xlsx> -JsonPath <ticker>.json
```
It checks for zero formula errors, computed EPS against reported EPS, a non-empty sales row, a computed valuation, and that the sensitivity tables tie to the model. Fix the JSON and rebuild until it passes. Then sanity-check the printed rows: history ties to the 10-K, and the current-year estimate is close to consensus.

## 5. Report back
Give the file path, a valuation table (methods × scenarios vs price), the history and projection headlines, the contract list, the key assumptions and caveats, and source links.

## Environment notes
- **No Python here.** Use Node for data and Excel COM from PowerShell for workbooks.
- **COM writes in custom scripts:** always write cell values through `InvokeMember`, as `build_model.ps1` does. Mixed int/double writes break PowerShell's COM binder.
- **PowerShell variable names are case-insensitive** (`$J` and `$j` are the same variable).
- **One-item arrays:** PowerShell unwraps a one-item array returned from a function, so use `return ,@(...)`.
