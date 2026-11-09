# 0.6 final-candidate chart checks

The experimental 0.6 candidate passed nine focused native cases and the later cold-reopen checks.
The final 292-case rerun completed 238 PASS and 54 CHECK_FAILED. Cold reopens completed: 4 basic values, 8 history PASS and 10/11 data/picture checks. The one mismatch is a verified newer ADVFN date.
The current source commit is `1cbbd104dd5e6696df9030d2aa261cfc630a8995`.

## Preserved earlier results

| Record | Saved result | Limit |
|---|---|---|
| First history/option focus | 49 PASS, four date-assertion failures | The four worksheet assertions later changed to expect `Value2` date serials. |
| First 0.6 full run | 238 PASS, 54 CHECK_FAILED across 292 cases | No native phase errors or unfinished cases. All 53 added cases passed. |
| Later data-and-picture cold reopens | Two Excel restarts | Diagnostic logs ended with failed chart insertion and chart-error records. |
| B16-disabled control | Automatic calculation and Ready after reopen | A16 remained live. The control isolated the numeric-chart formula for further work. |
| First numeric-reuse focus | Five PASS, four CHECK_FAILED across nine cases | Unchanged formulas could reuse cached UDF results. These rows remain preserved. |

These records remain separate. A passing batch does not establish a later cold reopen.
The first full run and its 292 raw outcomes remain unchanged.
The four original obsolete option functions and the failed positive-bid assertions remain visible in that record.

## Current candidate and focused checks

The Mac numeric-chart path now reuses an owned chart for the same cell.
It updates that chart's values, dimensions, note and visibility. It keeps one series and retains the chart's object ID.
The code still removes owned output for `NONE` and invalid arguments. Image requests retain their existing replacement path.
The numeric chart moves to the front after its visibility changes. Unsupported ordering remains a cosmetic error.

The revised nine cases force fresh UDF entry with changed return markers, mutated chart state and calculation controls.
The saved second focus records nine PASS, zero failed assertions and no unfinished cases.
Its case SHA256 is `cbe068e0aff8170a548a013956ec4f6fce1be36ac77df87f29d3ccd3b94f2b4e`.
Its source fingerprint is `8c7028f8eddc6017698a12d54a44b4886669e835a9554c8f33d35ef1ebad2ec3`.
The earlier focus remains preserved with its original results and assertions.

A separate worksheet probe created a chart, changed its return marker to `Reused`, then selected another cell.
The plot remained visible above its note. The probe then cleared the test chart with `NONE` and emptied its cell.
That probe established the visibility result. The later cold-add-in and installed receipts now record stable example reopens.

## Saved source and final checks

The final clean workbook contains 51 core components and preserves Dictionary's default Item attribute.
Its source audit found 47 modules byte-exact to 0.5 and four changed modules matching the reviewed source after identifier-case changes.
The clean-workbook SHA256 is `606f9ddc73e4b005c600e67afa3a7875ec5cbdb0e3270b09ce67e4fe8d329089`.
The clean-workbook VBA SHA256 is `b92f82ecf861a4da41a45fb7be05f93bd47043b093fbdc7a4fe78ab55c2df284`.
The current `smfCreateComment_.bas` SHA256 is `18b331c4639490490bff1e5c72fc04d699c52a34c3ca993681f727ad89cc8e42`.

| Remaining gate | Status |
|---|---|
| Final 292 native run, ID `smf-0.6-full-final-20261108` | DONE: 238 PASS, 54 CHECK_FAILED, zero phase errors |
| Final saved owner/core audit | All 51 core sources confirmed |
| Reviewed package build and checksums | 23 files, 22 listed checksums passed. Native/bundle VBA bytes match |
| Cold Excel restart and example reopens | Stable. Basic4 expected values, history8 PASS, data/picture10 PASS and1 preserved date FAIL |
| New-Mac or fresh-account installation | Untested |
| Windows runtime | Untested |

The normal 0.5 add-in remains unchanged and is unloaded during the isolated checks.
The 0.5 package and reports remain unchanged.
The documentation review and the history/option repairs are complete for their selected checks.
The prepared 0.6 build is experimental. The reviewed package is saved under `mac/built/0.6-dev` for continued testing in this fork.

The final cold check kept the StockCharts image and unselected numeric plot visible, with Ready and Automatic calculation.
The data/picture example retained 10/11: ADVFN returned 2026-09-26 against the unchanged 2025-09-27 expectation.
A direct public-page check verified the2026 year-end label. The expected example cell was not changed.

Both the production numeric-only ZOrder block and the fresh-entry case controls changed before the nine focused passes.
