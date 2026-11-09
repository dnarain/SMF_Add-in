# SMF Mac 0.7 native results

Excel recorded DONE in both saved workbooks. The tables contain every recorded status and failed assertion.

| Run | Cases | Raw results | Interpreted results | Worksheet entries |
|---|---:|---|---|---:|
| Focused | 246 | 246 PASS | 31 NOT_SUPPORTED, 215 PASS | 19 |
| Full | 538 | 36 CHECK_FAILED, 502 PASS | 44 NOT_SUPPORTED, 458 PASS, 22 PROVIDER_UNAVAILABLE, 14 UNCLASSIFIED_FAIL | 51 |

Excel saved PASS when the selected assertion returned True. The reporter marks passing retired-mode guard checks NOT_SUPPORTED.
Guarded financial modes return #N/A. The transport returns an error string for blocked URLs.

The full run can contain legitimate provider, parser, or compatibility failures. Those failures remain in the exported files.

## Source and transport evidence

Source commit: [`e7d788b4f0871e94dbe6e13131ca8d9b7ff0eb8a`](https://github.com/dnarain/SMF_Add-in/commit/e7d788b4f0871e94dbe6e13131ca8d9b7ff0eb8a).

Both final manifests record the same source fingerprint. Source hashes and saved workbook hashes are in `provenance.json`.
Each run folder includes a native-owner-source-audit.json file. It records the separate saved-source audit against the clean workbook and generated runner.

Columns R:S contain measured fetch-entry and fetchURL-bridge deltas. When the bridge delta is zero, SMF did not call that text-fetch handler.
Image downloads use a separate handler. The reporter also requires matching case and run markers for positive bridge counts.
A later blocked call or smfMacReadResponse can overwrite the last diagnostic within the same case.

## Earlier checkpoints

These operator-recorded checkpoints precede the final runs. Their workbooks are not included.

| Checkpoint | Planned cases | Recorded results | Note |
|---|---:|---|---|
| First focused trial | 179 | 7 CHECK_FAILED, 118 PASS, 53 PENDING, 1 RUNNING | The operator stopped the run at a test compile error. Later work corrected precise-number parsing, worksheet date assertions, and the unavailable Error_Type test call. |
| Focused trial after those corrections | 183 | 2 CHECK_FAILED, 181 PASS | Two FRED metadata checks found extra carriage returns. A local FRED text correction and two newline checks followed. |
| Focused FRED newline checkpoint | 185 | 185 PASS | This completed checkpoint preceded the retired-source guard checks. |
| Earlier retired-source checkpoint | 234 | 234 PASS | This passed before a diagnostic self-test expectation changed from 3 to 19. The current export requires both later runs to use the final source. |
| First full-suite attempt | 526 | 1 CALL_ERROR, 5 CHECK_FAILED, 144 PASS, 376 PENDING | The operator stopped this attempt after 150 completed cases. Case 34 found the old diagnostic self-test expectation. This attempt was not a completed full run. |
| Earlier full run with 13 revised source modules | 526 | 36 CHECK_FAILED, 490 PASS | Runtime finished. The native runner audit found a TAB literal replaced by three spaces during import. The operator rejected that source evidence and regenerated the runner with explicit vbTab. |
| Isolated Google checks, run3 | 13 | 2 CHECK_FAILED, 11 PASS | The financial-data assertion received HTTP 200 HTML. A worksheet test also triggered two requests after formula assignment and Calculate. The test changed to conditional Calculate with strict 1/1 counters. |
| Isolated Google checks, run4 | 13 | 1 CHECK_FAILED, 12 PASS | The strict worksheet 1/1 counter check passed. The financial-data assertion still failed on HTTP 200 HTML. These isolated checks do not establish full-suite completion. |

## Limits

These are selected inputs on one Mac. They do not establish all fields, arguments, accounts, Mac versions, or Windows behavior.
HTTP 200 alone does not prove correct financial data. Separate HTTP probes do not change native case classifications.
Provider values can change after these snapshots. These checks do not establish quote freshness or trading suitability.
The function-results.csv file lists selected case results. A separate function reference can explain broader usage limits.

## Files

- `focused/` and `full/` contain redacted reports, manifests, cases, summaries, saved-source audits, and CSV files
- `registry/` contains the support table and its recorded hashes
- `history.json` records earlier checkpoints
- `provenance.json` records source, registry, counter, and saved-workbook hashes
- `SHA256SUMS` identifies each exported file

Native workbooks, key files, cookies, raw HTTP bodies, and raw response headers remain private.
Cases and manifests are redacted copies. Their published hashes can differ from the original hashes recorded in the native run.

## Function reference and source audits

[The function reference](full/FUNCTION-STATUS.md) describes all 121 public Functions.
[The function CSV](full/function-status.csv) records their selected cases and limits.
The report lists six auxiliary Sub cases separately. The full case CSV records all 538 case IDs.

The full run records 36 failed assertions. All 51 worksheet cases passed.
Three Google URL diagnostics passed. The Google historical-data assertion received HTML and failed.
These diagnostics establish no Google financial-data success.

The full folder also contains the reference checks and reviewed work notes.
The audits folder contains the clean-source and native-add-in source audits.
These audits compare saved VBA sources. They do not replace native calculation or cold-reopen receipts.

The ADVFN evidence records the public request, response hash and three year-end headers.
The raw response body remains private.
The publication manifest records the file allowlist, original hashes and privacy checks.
