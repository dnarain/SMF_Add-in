# SMF for Excel on Mac — 0.7-dev

This experimental bundle contains version `3.0.2024.07.30-mac.0.7-dev` and three example workbooks.
The full native suite saved 538 results: `{"CHECK_FAILED": 36, "PASS": 502}`.
All 246 focused repair cases passed.
`BUILD-INFO.json` records the exact source, binary and test hashes.
Remaining provider failures stay in the full-suite counts.
Not Supported records expected guards for confirmed retired modes; those passes establish no financial feed.
`NotSupported.csv` lists the source reasons and shipped obsolete default fields.
Supported CSV-URL and Yahoo option modes remain separate.

## Install

1. Keep these files together with their original names.
2. Quit Excel, then run `install_helper.command`.
3. In Excel, select Tools > Excel Add-ins > Browse, then select this folder's `SMF-Mac.xlam`.
4. Enable its checkbox and allow macros when Excel asks.
5. Keep Trust access to the VBA project object model OFF.
6. Set calculation to Automatic, then quit and reopen Excel.
7. Open an example from this folder.

For FRED requests, obtain your own FRED API key.
Save it in a private plain-text file named `smf-FRED-API-Key.txt` beside the loaded add-in.
The file contains only your 32-character key; keep its file permissions private.
The bundle excludes this file. Do not add it to Git or a shared example.
Other sources do not need this FRED key.

## Examples

The basic example checks the version, HTML parsing and two Finviz values.
The data/picture example has 11 passing comparisons, a StockCharts picture and a numeric chart.
Three live-value checks require a positive price, a yield fraction from 0 to 1, and a positive target.
Those checks establish numeric values/ranges for the saved inputs. They do not independently verify accuracy, freshness or provider agreement.
Its AAPL fiscal-date comparison uses the independently verified 2026-09-26 date.
The 0.6 rollback retains its original 2025-09-27 comparison.
The history example has eight passing comparisons and a five-row, two-column result.
Native observations finished at `2026-11-09T22:31:29.488582+00:00`.
Live values can change after recalculation.

`SHA256SUMS` lists the other 23 file hashes.
Packaging preserves the native VBA bytes, formulas, cached values, charts and image bytes.
Packaging removes saved home-folder paths from workbook XML and uses a relative add-in link.
The unchanged VBA project can retain reference and compiler-cache paths.
The 0.6 rollback remains separate.
A fresh Mac and a separate macOS user account remain untested.
