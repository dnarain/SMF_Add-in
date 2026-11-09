# SMF Mac 0.5-dev: full-suite results, November 8, 2026

The native Excel suite completed **239 cases: 187 PASS and 52 CHECK_FAILED**. No selected case remains unfinished.
Excel stayed running through both complete runs and the final version-gate check. The port remains experimental.

The runner now prepares its workbook and fixtures from tracked files. It needs no old `mac/diagnostics/out` folder.
See the [fresh-checkout instructions](../../tests/full_function_suite/README.md) to repeat the run.

## Final results

| Result | Cases | Meaning |
|---|---:|---|
| PASS | 187 | The assertion passed for its recorded input. |
| Failed request | 31 | The native case recorded an unavailable endpoint, access response or transport failure. |
| Failed, cause unresolved | 16 | The saved evidence does not distinguish a parser defect from a provider response problem. |
| Invalid URL | 1 | Google History constructed a malformed URL; curl rejected it before HTTP. |
| Obsolete | 4 | MSN/OptionsXpress compatibility functions explicitly returned obsolete-status text. |
| Unfinished | 0 | All 239 selected cases completed. |
| Runner setup/call/assertion/cleanup errors | 0 | No native phase exception occurred in the final run. |

The four obsolete results are included in the 52 raw CHECK_FAILED rows. They do not provide option data.
The [generated report](test-results.md), [all case results](all-case-results.csv), and [function results](function-results.csv) retain the exact assertions and outcomes.

All 121 inventoried public `.bas` functions have a selected primary call: 108 passed, nine data calls failed and four returned obsolete status.
A passing primary call does not establish every other call to that function. The report includes all cases for each function.

All **25 worksheet-entry cases passed**, including all nine history row/column regressions.
The `worksheet_entry` field includes formula cases tagged as VBA calls or fixtures as well as those tagged WORKSHEET.
All **14 picture/comment cases passed**. The repaired Yahoo history and Yahoo JSON option bid, strikes and expirations passed their selected checks.
The report marks 41 passing provider-data/image checks, including cached results and two StockCharts image checks.
Many check only a positive or nonempty result; financial accuracy and freshness need separate comparisons.

| Case group | PASS | CHECK_FAILED |
|---|---:|---:|
| `derived_history_snapshot` | 7 | 0 |
| `json_and_utils` | 40 | 0 |
| `json_core` | 15 | 0 |
| `live_datasources` | 8 | 22 |
| `live_extraction` | 10 | 9 |
| `live_options` | 5 | 8 |
| `local` | 21 | 0 |
| `market` | 20 | 9 |
| `options` | 13 | 4 |
| `repair_live_history` | 6 | 0 |
| `repair_local` | 17 | 0 |
| `worksheet` | 2 | 0 |
| `worksheet_mac_images` | 14 | 0 |
| `worksheet_regression` | 9 | 0 |

## Failures that need work

- `smfGetGoogleHistoryCSV` inserts unencoded date spaces into its URL. Fixing the encoding still requires a live check of the old CSV route.
- `smfRSITargetPrices` depends on a DNS-failing legacy quote host. Its history call also inherits F/G defaults that the history validator rejects.
- `smfGetADVFNElement` and `smfGetGuruFocusItem` require markers absent from the separately captured pages. Repaired numbered fields on those providers passed separately.
- Other failures include old MSN/YQL/GuruFocus CSV routes, timeouts, access responses and legacy Google/Yahoo HTML option routes.
- Four MSN/OptionsXpress option entry points return obsolete text before requesting data.

See [the failure triage](failure-triage.md) for functions, source lines and the limits of each finding.
The raw statuses remain unchanged. The report classifies each case once, even when source review finds another defect in the same function.

## Data-source coverage

The [source table](datasource-results.csv) lists 35 configured domains. Separate HTTP probes cover those domains; their results are diagnostic evidence.
Of those domains, 32 have associated native cases after the report includes the observed `ca.advfn.com` request.
The remaining three groups have no dedicated native case. `c.stockcharts.com` and `ogres-crypt.com` are retained Windows comment-chart routes; the Mac branch uses StockCharts and local charts.
`fc.yahoo.com` is a helper-internal cookie-bootstrap hop without an isolated native case/hop record.
The native StockCharts image cases passed. Image downloads do not update the text downloader's last-HTTP fields, so the report cannot bind every chart/CDN request to those fields.

Only 11 domain groups have a passing positive native extraction/image case. This does not establish all fields on those sites.
The [58 HTTP probes](http-probe-results.csv) recorded 23 HTTP 200 responses, 17 transport failures, nine 404s, six 401s, two 403s and one 202.
Some 200 responses contain unrelated or parked pages. An HTTP 200 does not establish usable financial data.
The probes were collected separately during this test session. They never override a native case's recorded outcome or supply fixture answers for live URLs.

## What changed between the two full runs

Run 1 recorded 184 PASS, 52 CHECK_FAILED and three CLEANUP_ERROR results.
The tests tried to delete a picture that the text-only, clear-picture and failed-image cases correctly left absent.
The shared cleanup now searches for an owned picture before deleting it. The expectations and production code stayed unchanged.

Run 2 repeated all 239 cases. Only `MacComment_text_only`, `MacComment_clear_none` and `MacComment_failed_update_clears_picture` changed to PASS.
The [run 1 summary](run1-summary.json) preserves the initial failures. All other 236 statuses stayed unchanged.

## What remains untested

- First installation on another Mac or a fresh macOS account. This run used the current Mac's Excel and installed helper.
- Worksheet entry for most public functions; the 25 worksheet cases cover selected inputs only.
- Most of the 1,816 numbered recipes, every field/value meaning, every argument combination and arbitrary caller URLs.
- Full parser behavior where the request failed, plus helper bootstrap hops and retained Windows chart routes without dedicated native cases.
- All menu actions, public Subs, private functions and Dictionary members as separate callable tests.
- Native Stop/Resume recovery after an actual crash, sustained large-workbook use, concurrent requests and repeated daily sessions.
- Windows 32-bit/64-bit runtime behavior, older Excel/Mac builds and other timezones.

The [uncovered call patterns](uncovered-call-patterns.csv) list the per-function worksheet/live-result gaps and exact endpoint evidence gaps.
A missing case-bound endpoint record can also reflect redirects, caching, multi-request calls or image downloads; it does not prove the endpoint was never contacted.

## Native code and additional checks

The disposable test workbook came from `mac/built/0.5-dev/SMF-Mac.xlam`, with only its workbook content type changed before importing the runner.
An independent audit found all 51 core VBA components byte-exact to the package, including the Dictionary class/default Item attribute.
The imported runner is the sole additional component. The production source and packaged binary stayed unchanged.

Excel rewrote numeric literals on 105 runner source lines and changed five descriptive metadata strings during import.
The largest numeric difference was `5E-13`; the audit found no other runner changes. The metadata encoding issue remains visible in some native descriptions.
The [provenance record](provenance.json) retains the workbook hashes, audit results, helper comparison and run delta.

The final runner also requires an exact version token. It rejects suffixes and version matches found only in a directory path.
That guard changed after the full repeat and passed a separate native case covering seven version-token conditions, including the actual version.
All 239 default case specifications still match the full run after run-path normalization.

| Additional check | Result |
|---|---|
| Fresh preparation and fixture checks | 61 passed |
| Actual-generator checks, including worksheet defaults and version guard | 14 passed |
| Saved-result/report fixtures | 27 passed |
| HTTP-runner fixtures | 3 passed |
| Core static compatibility checks | 42 fixtures, 49 modules, 464 signatures across four configurations passed |
| HTML checks / no-prompt build checks | Passed |
| Table recipe recognizers | 31 fixtures passed; 35 unsupported table recipes remain guarded |
| Provider recipe fragments | Eight fragments and four field recipes passed |
| Standalone helper transport / image transport | 24 / 18 passed |

The full suite ran on Apple Silicon, macOS 27.0.1 and Excel 16.114 with timezone `America/Los_Angeles`.
The observed version was `3.0.2024.07.30-mac.0.5-dev`. Both completed runs saved matching case hashes and core-source fingerprints.
Raw workbooks, response bodies, headers and private JSON remain local. The checked-in reports redact anonymous Yahoo cookie and crumb values.

After the checks, the normal add-in and history workbook were restored. Excel showed Automatic calculation and Ready; VBA project access remained off.
The temporary fixture server and keep-awake process stopped.
