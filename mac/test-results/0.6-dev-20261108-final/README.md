# SMF Mac 0.6 final-candidate full run

The final candidate uses source commit `1cbbd104dd5e6696df9030d2aa261cfc630a8995`.
Its 292-case run ID is `smf-0.6-full-final-20261108`.
The final run completed 292 cases: 238 PASS and 54 CHECK_FAILED. It recorded zero native phase errors and zero unfinished cases.
The experimental 0.6 candidate completed its native full run and cold-reopen checks. The reviewed package is saved under `mac/built/0.6-dev` for continued testing in this fork.

## Earlier records

The [first 0.6 full run](../0.6-dev-20261108-first-run/README.md) recorded 238 PASS and 54 CHECK_FAILED across 292 cases.
It recorded zero native phase errors or unfinished cases. All 53 added history/option cases passed.
Two later cold-reopen checks restarted Excel in the chart-insertion lifecycle.
The revised candidate has a different source fingerprint. The earlier full run remains unchanged.

The first numeric-reuse focus recorded five PASS and four CHECK_FAILED.
Before the second focus, both the production numeric-only ZOrder block and the fresh-entry case controls changed.
The revised cases changed return markers and forced fresh entry after state mutations. The nine passes correspond to both changes.
The revised nine-case focus recorded nine PASS. A separate unselected worksheet probe showed the reused plot above its note.
The [chart recheck record](../../CHART-RECHECK-STATUS-20261108.md) preserves the sequence and remaining gates.

## Final 292 results

| Outcome | Cases |
|---|---:|
| PASS | 238 |
| CHECK_FAILED | 54 |
| Native phase errors | 0 |
| Unfinished | 0 |

The 121-function reference uses every selected case. It lists 17 passed data calls, two limited functions, one field dispatcher and 19 failed live calls. It also lists four obsolete functions, 77 local/controlled functions and one URL-dependent downloader.
All 32 worksheet-entry cases passed. All 14 picture-group cases passed. The nine focused fresh-call numeric-reuse cases passed separately.
The 48 provider-data/image assertions include cached values, four numeric-zero comparisons and two StockCharts image checks. The recorded failures remain 29 unavailable requests, 20 unresolved results, one malformed URL and four obsolete functions.
The zero-accepting option comparisons and positive-bid assertions remain separate checks.
Matching a numeric zero in Yahoo JSON does not establish a fresh usable positive quote.

## Final artifact checks

The clean-source audit passed for 51 core components, with 47 unchanged sources and four reviewed modules. The final saved-owner audit confirmed all 51 core sources against the reviewed candidate.
The native add-in source audit passed for 51 components. The 23-file reviewed package passed all 22 listed checksum checks. Its VBA bytes match the native add-in.
A fresh Excel session completed the basic checks (4 expected values), history checks (8 PASS) and data/picture checks (10 PASS, 1 FAIL).
The StockCharts picture and unselected numeric plot remained visible. Excel remained Ready with Automatic calculation after explicit recalculation and reopen.
The experimental 0.6 build is saved under `mac/built/0.6-dev`. The preserved 0.5 bundle remains unchanged.

The generated runner expects source fingerprint `8c7028f8eddc6017698a12d54a44b4886669e835a9554c8f33d35ef1ebad2ec3`.
Its case specification SHA256 is `75196b800984e07692881ed64dff7d3bd498e7bb9d7dbcc6fab6e6dd1c53086f`.
The completed saved native report confirms both run bindings.
The observed workbook SHA256 is `14ed9c3af4e9733deefbde7fac3129dcea888006cd8e47a189b763289132e5a6`.
The observed workbook VBA SHA256 is `31066426d87e287f55c29c3df295f4061b47c7cd295bc48da229e3b86e67bec9`.

The run uses the same Mac, Excel installation and helper. Another Mac and a fresh macOS account remain untested.
Windows runtime and most argument/provider combinations remain untested.
Raw workbooks and diagnostic captures stay private. The preserved 0.5 package and reports remain unchanged.

All 292 raw statuses match the preserved first 0.6 full run. The final candidate has a different source/VBA fingerprint.
The [provenance sidecar](provenance.json) records source commit 1cbb separately from the reporter commit and the unchanged null raw source_commit.

## Preserved live-date mismatch

The data/picture example still compares ADVFN field 5196 with its saved 2025-09-27 expectation.
The cold check returned 2026-09-26, so D9 remains FAIL. The expected cell C9 was not changed.
A current unauthenticated request to the configured URL returned HTTP200 and the label "Year End September 26 2026".
The annual columns were 2022–2026. The [minimal provider evidence](advfn-year-end-evidence.json) records the request time and body hash.
The final 292 case also returned 2026-09-26 and passed its valid-date assertion. Its assertion does not compare against the example's old fixed date.
The [cold-add-in receipt](cold-addin-receipt.json) preserves all three example outcomes, including 10/11 for data/pictures.


## Portable folder and preparation checks

The reviewed bundle passed a separate same-Mac folder move and fresh Excel session.
Data/picture remained 10/11 with both plots visible. History returned 8 PASS and a genuine 5x2 array. Basic returned four expected values.
The operator closed without saving, reopened the examples and quit Excel normally. Package bytes remained unchanged.
The unchanged helper remained installed. The [cold receipt](cold-addin-receipt.json) records these observations.

The fresh-checkout preparation and loopback checks passed 63 checks. The default 292 case IDs, 121 public functions and 32 worksheet entries matched the native manifest.
The derived workbook preserved the bundled VBA bytes. These preparation checks did not execute the full native suite in that new workbook.
The [preparation summary](fresh-checkout-validation-summary.json) records the separate scope.

## Installed add-in verification

The final installed add-in loaded automatically after a normal Excel quit and restart. Its version and directory matched the repository mac/out installation.
The installed data/picture example remained 10/11 with both plots visible. History returned 8 PASS and a genuine 5-by-2 array.
Excel showed Ready and Automatic calculation. VBA project access was OFF, with no security setting changed.
The [installed receipt](installed-addin-receipt.json) records the final check. The 0.5 rollback copies remain preserved and the helper/14 definitions remain unchanged.
