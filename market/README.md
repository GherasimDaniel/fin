# The Market Delta

Animated river-delta view of US stock market value by industry, July 1926 to September 2026.
Published artifact: https://claude.ai/artifact/4eBGq4ZC2f1wsvjf1rF4Sa
Open `market-delta.html` in a browser to view it locally (fonts load from Google Fonts).

## Files
- `market-delta.html` - the page, with the data embedded.
- `data/industry_market_values_monthly.csv` - month, industry, market value ($M), firms, share of total (%), CPI-U, value in Sep 2026 dollars ($M).
- `data/data.js` - the same data in the compact form the page embeds.
- `data/raw/30_Industry_Portfolios.csv` - Kenneth R. French Data Library, 30 Industry Portfolios (CRSP, through 2026-09).
- `data/raw/cpi_u_monthly.csv` - BLS CPI-U all items, via github.com/datasets/cpi-us (Oct 2025 not published by BLS; interpolated in the build).
- `scripts/build_data.js` - builds `data/data.js` from the raw files: market value = number of firms x average firm size.
- `scripts/template.html` - the page source; `/*DATA*/` is replaced with the contents of `data/data.js`.

## Rebuild
```
node scripts/build_data.js data/raw/30_Industry_Portfolios.csv data/raw/cpi_u_monthly.csv data/data.js
node -e "const fs=require('fs');const t=fs.readFileSync('scripts/template.html','utf8');fs.writeFileSync('market-delta.body.html',t.replace('/*DATA*/',()=>fs.readFileSync('data/data.js','utf8').trim()))"
```
The published artifact uses the body-only file; `market-delta.html` wraps it in a full HTML document.

## Caveats
NYSE only until June 1963; AMEX joins July 1963, NASDAQ July 1973. Banks were mostly OTC before 1973, so early finance is understated. Month-end values. The nine sector groups are a grouping made for this page.
