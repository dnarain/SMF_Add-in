# SMF for Excel on Mac — 0.6-dev

This experimental bundle contains the add-in and three example workbooks.
Native Excel checks used version `3.0.2024.07.30-mac.0.6-dev`.
The full suite saved 292 results: `{"CHECK_FAILED": 54, "PASS": 238}`.
The separate nine-case numeric-chart suite passed.
The live data/picture example completed the recorded cold-reopen check.
`BUILD-INFO.json` records the failed-case counts.
The Mac documentation lists provider failures and untested inputs.

## Install

1. Keep all files in this folder together, with their original names.
2. Quit Excel.
3. Double-click `install_helper.command`.
4. Open Excel.
5. Select **Tools > Excel Add-ins > Browse**.
6. Select this folder's `SMF-Mac.xlam` and enable its checkbox.
7. Allow the add-in's macros when Excel asks.
8. Keep **Trust access to the VBA project object model** OFF.
9. Set calculation to **Automatic**.
10. Quit and reopen Excel.
11. Open an example workbook from this folder.

The helper installer preserves an existing helper as a backup.
The installer does not change Excel security settings.

## Examples

| Workbook | Saved native checks |
| --- | --- |
| `SMF-Mac-Checks-20261108.xlsx` | Version, HTML parsing, Finviz earnings and price. |
| `SMF-Mac-Data-Picture-Checks-20261108.xlsx` | 10 of 11 worksheet checks passed; ADVFN date differs from the unchanged expected date, a StockCharts picture, and a native numeric chart. |
| `SMF-Mac-History-Checks-20261108.xlsx` | Eight worksheet checks and a five-row, two-column history result. |

The example receipt records native checks at `2026-11-09T08:17Z`.
Recalculation can change live values.
Fixed expected prices and dates can require updates.
The ADVFN check returned `2026-09-26`; its fixed comparison remains `2025-09-27`.
The receipt records an independent source check.
The workbook retains D9 and its total as FAIL.


## Build record

`BUILD-INFO.json` records source and binary hashes, native checks, and package changes.
`SHA256SUMS` lists the other 22 file hashes.
The package retains the tested add-in's VBA bytes.
Excel retains Forms-reference and compiler-cache paths inside the VBA project.
The package removes saved home-folder paths from workbook XML.
A fresh Mac or separate user account remains untested.

[Mac documentation](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/mac/README.md)
includes the current limits.
