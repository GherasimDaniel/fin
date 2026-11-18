# The Month Delta

River-delta view of S&P 500 market value over the latest month of trading days, with a day-by-day chart,
a week-by-week table (sectors and largest companies) and a full company table.

## Use it
Double-click `open-month-delta.cmd` (or run `node serve.js`) and the page opens at http://localhost:8787.
Every time the page opens it asks the local helper for new trading days. The helper downloads them,
drops days older than one month, and the page reloads with the new data. The status line under the title says
what happened ("Up to date", "Added 1 new trading day", or why an update failed). Close the console window to stop the helper.

Today's close is added after 4:20 pm New York time; until then the latest day is the previous close.

Published snapshot (does not update by itself): https://claude.ai/artifact/9m5B9KHMriV53uYDtRsdBE
Ask Claude to republish `build/month-delta.body.html` to refresh it.

## Docker
`docker build -t month-delta . && docker run --rm -p 8080:80 month-delta`, then open http://localhost:8080.
The image is nginx serving a static copy built from `data/month.js` at build time. It does not update itself:
run `node scripts/update.js` (or `open-month-delta.cmd`) to fetch new trading days, then rebuild the image.

## Files
- `serve.js` - the local helper (http://localhost:8787, local machine only).
- `open-month-delta.cmd` - starts the helper and opens the page.
- `month-delta.html` - static copy of the page with the current data (no live updates). The helper rewrites it
  (and `build/month-delta.body.html`) when it starts and after every update check, so it always matches `data/month.js`.
- `data/daily_market_values.csv` - one row per company per day in the current window.
- `data/month.js` - the data the page uses. `data/meta.json` - window dates and last update time.
- `data/raw/` - downloaded inputs: Wikipedia member list, SEC share counts, Yahoo closes for the window only.
- `scripts/lib/market.js` - the pipeline: member list, share counts, prices, one-month window.
- `scripts/lib/page.js` - puts the data into `scripts/template.html`.
- `scripts/update.js` - command-line update (`node scripts/update.js`, add `--force` to rebuild without new days).
- `scripts/fetch_cover.js` - reads 10-Q cover pages for multi-class companies (run by hand when those share counts need refreshing; copy the numbers into `COVER` in `market.js`).

## Method and caveats
- Market value = SEC shares outstanding x Yahoo daily close (split-adjusted; clean splits after the share-count date are applied).
- The member list and SEC share counts refresh weekly; prices refresh whenever a new close is available.
- Full market value, not S&P's float-adjusted index value. Share counts are held fixed within the window.
- Multi-class companies are counted once with all classes; Visa's B and C classes use approximate conversion rates.
- Members with no share count at SEC (currently Vylor, a new listing) are left out and listed on the page.
- State Street and iShares holdings downloads are blocked for scripts, so index weights are computed from SEC share counts instead.
