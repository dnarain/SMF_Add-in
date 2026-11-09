# Run the SMF 0.7 native suite from a fresh checkout

These tools prepare and read a disposable Excel test workbook. They do not change the packaged add-in or run Excel themselves.
The reviewed specifications contain 526 total cases across 121 public Functions.
That total includes six explicitly tagged auxiliary Sub cases, which remain outside the Function inventory.
Fifty cases enter a worksheet formula. Ten exercise RSI, ADVFN and FRED paths; eight check retired-mode errors or IFNA display.
These counts describe the inputs. The saved native results establish their outcomes.
The existing history row/column and default-item worksheet cases remain included.
The inventory contains 36 provider and diagnostic groups, including retained legacy endpoint diagnostics.
It does not enumerate arbitrary caller-supplied URLs. Registered retired routes are skipped before HTTP.

Run the commands below from the repository root. Use Python 3 and native Excel on Mac.
The three UTC conversion cases assume the Mac timezone is `America/Los_Angeles`.
The short `fnWait` case also requires local time before `23:59:58`.
A fresh checkout still uses the current Mac's Excel installation and installed helper. First-time setup on another Mac remains unverified.

## 1. Prepare a private run folder

```sh
SMF_SUITE_DIR="$(mktemp -d "${TMPDIR:-/tmp}/smf-full-suite.XXXXXX")"
python3 mac/tests/full_function_suite/prepare_suite.py --suite "$SMF_SUITE_DIR"
```

The preparer uses tracked inputs, including `mac/built/0.7-dev/SMF-Mac.xlam`.
It creates `SMF-Full-Suite.xlsm` by changing only the workbook content type. Every VBA-project byte remains unchanged.
It copies 14 definition files, the reviewed case specifications, provider fragments and the required local fixtures.
It records source hashes and a run ID in `preparation.json`. It refuses a nonempty run folder.

No old `mac/diagnostics/out` state or downloaded provider captures are required.
Preparation does not run Excel, install the helper or contact providers.

## 2. Start the local fixture server

Open `PREPARED.txt` in the new run folder. In a second terminal, run its exact `start_server` command.
The command supplies both paths explicitly:

```sh
python3 mac/tests/full_function_suite/fixture_server.py \
  --root "/absolute/run-folder/fixtures" \
  --output "/absolute/run-folder/server-config.json"
```

Use the paths printed by the preparer. Wait for **Loopback fixture server ready**, then leave that terminal open.
The server binds only to `127.0.0.1` and writes its selected port and PID to `server-config.json`.

The local text, HTML table, CSV and PNG are controlled parser/image fixtures.
JSON, option and historical-data fixtures remain the reviewed inline case setup. The preparer also copies their literals for inspection.
The server never supplies answers for live provider URLs.

## 3. Generate the VBA runner

Return to the first terminal, where `SMF_SUITE_DIR` is set:

```sh
SMF_RUN_ID="$(python3 -c 'import json,sys; print(json.load(open(sys.argv[1]))["run_id"])' "$SMF_SUITE_DIR/preparation.json")"
python3 mac/tests/full_function_suite/build_suite.py \
  --suite "$SMF_SUITE_DIR" \
  --server-config "$SMF_SUITE_DIR/server-config.json" \
  --run-id "$SMF_RUN_ID" \
  --expected-version 3.0.2024.07.30-mac.0.7-dev
```

The generator writes `modSMFFunctionSuite.bas`, `test-cases.json`, `function-inventory.json` and `run-manifest.json`.
It also saves `not-supported-registry.json` and `NotSupported.csv` beside those files.
The manifest binds both retirement files by SHA256.
It verifies public-function coverage and records the fixture URL, case hashes and expected version.
Generation does not execute any case. A case is not a pass until native Excel saves its completed result.

For a focused version-gate check, prepare another empty run folder.
Add `--case-file mac/tests/full_function_suite/version_gate_cases.json` to that folder's build command.
This optional one-case run checks the actual version and rejects version suffixes, path-only matches and malformed version text.
It is separate from the default 526 cases.

To rerun the nine numeric-chart reuse cases, prepare another empty run folder.
Pass `--case-file mac/tests/full_function_suite/numeric_chart_reuse_cases.json` to that folder's build command.
Those nine checks run separately from the default 526 cases.

## Confirmed retired sources and fields

The Mac source returns Excel `#N/A` for registered retired financial modes and explicitly obsolete default definitions.
The registry records the status **Not Supported**, a reason and a suggested replacement.
A `401`, `403`, `404` or timeout alone does not establish retirement.

A passing retired-mode case has raw status `PASS` and interpreted status `NOT_SUPPORTED`.
It proves the specified guard behavior and zero measured fetchURL bridge attempts. It establishes no financial feed.
Mixed option wrappers retain their supported Yahoo mode.
`RCHGetYahooQuotes` retains supplied CSV URLs with empty `pItems`, except URLs that match a registered retired route.
Generic transport/extraction functions retain their declared result types and ordinary `pError` behavior.

The [shipped status CSV](../../NotSupported.csv) contains six source-mode rows and 288 explicitly obsolete default-field rows.
The local **SMF: Source Status** command shows the currently loaded registry and field definitions in a filtered Excel table.
Valid custom definitions can replace obsolete defaults. Metadata selectors such as `SOURCE` and `DEFINITION` remain readable.

From the repository root, check the CSV and embedded registry without changing files:

```sh
python3 mac/build_not_supported.py
```

The bulk 288-definition case asserts every listed default through an explicit native loop.
It remains one case, with two additional field 625 worksheet/IFNA cases. It does not count as 288 worksheet formulas or financial values.
The status-table and registry Subs remain auxiliary checks, outside the 121 public Function rows.

## FRED API key for live checks

The Mac FRED function uses the official API. Its live checks need your own FRED API key.
Request a key at https://fredaccount.stlouisfed.org/apikeys.
Save only the key in a plain-text file named `smf-FRED-API-Key.txt` beside the disposable `SMF-Full-Suite.xlsm`.
Excel may ask you to select that file and grant access before the first FRED formula returns.
The installed add-in needs a separate local key file beside its own `SMF-Mac.xlam`.

Git ignores this filename in every folder. Release bundles contain no key file.
Do not put the key in a formula, case specification, report or commit.
The reports redact API-key query values. Keep raw workbooks and raw JSON private.
Without a valid readable key file, live FRED calls return `pError`; they cannot establish a working FRED connection.
Controlled FRED parser cases do not need a key.

## 4. Run in native Excel

1. Unload the production SMF add-in for this isolated run. Leave the checked-in package unchanged.
2. Install the copied helper with the run folder's `install_helper.command` if the reviewed helper is not already installed.
3. Open the run folder's `SMF-Full-Suite.xlsm` as a normal workbook. It must have `IsAddin=False` and a saved `.xlsm` filename.
4. Import only `modSMFFunctionSuite.bas` through the VBA editor. The workbook already contains the 51 packaged core components.
5. Save, close and reopen that workbook. Allow its macros when Excel asks. VBA project object-model access can remain OFF for manual import.
6. In Excel's Macro dialog, run **SMFRunFullSuite**. This entry point takes no arguments.

The runner checks the saved workbook, expected version and run metadata before executing cases.
It creates the input/results/control/journal sheets and saves each case's state before and after the call.
The cases check both arrays returned directly to VBA and formulas evaluated through worksheet callers.
Live cases contact their actual providers; provider pages can fail, change, require access or return obsolete-status text.
No live result is replaced with a fixture answer.

Use **SMFTestStop** to stop scheduling cases, and **SMFTestResume** to continue the saved run.
Resume preserves an unfinished `RUNNING` row instead of retrying it automatically.
Wait for `TestControl` to show `DONE`, then save the workbook before reading results.
If Excel exits or execution stops, retain the workbook and read its partial results. Unfinished rows remain unfinished.

## 5. Optional: record separate provider HTTP evidence

Use a new folder within this run:

```sh
SMF_HTTP_DIR="$SMF_SUITE_DIR/http-evidence"
python3 mac/tests/full_function_suite/run_datasource_http.py \
  --output "$SMF_HTTP_DIR" \
  --inventory "$SMF_SUITE_DIR/datasource_cases.json" \
  --metadata "$SMF_SUITE_DIR/datasource_inventory_meta.json" \
  --support-registry "$SMF_SUITE_DIR/not-supported-registry.json"
```

This command makes bounded requests to active inventoried destinations and saves responses, headers and request metadata.
It checks the supplied registry before curl or a Yahoo authentication handshake.
Registered retired routes produce **Not Supported** records with `request_performed=false`, `network_requests=0` and no response body.
Those records are skipped requests, without evidence of provider availability.
Its endpoint checks are separate from native extraction. An HTTP `200` does not establish that an SMF financial field parsed correctly.
Use evidence from this run; do not reuse an old ignored HTTP folder.

## 6. Read the saved native results

```sh
python3 mac/tests/full_function_suite/read_results.py \
  --suite "$SMF_SUITE_DIR" \
  --workbook "$SMF_SUITE_DIR/SMF-Full-Suite.xlsm"
```

This reads saved cells without running Excel or recalculating formulas.
It writes raw local `native-results.private.json` with file mode `0600` by default. The exact workbook and suite paths are required.

## 7. Create the report

If step 5 produced HTTP evidence:

```sh
python3 mac/tests/full_function_suite/make_report.py \
  --suite "$SMF_SUITE_DIR" \
  --workbook "$SMF_SUITE_DIR/SMF-Full-Suite.xlsm" \
  --http-dir "$SMF_HTTP_DIR"
```

Omit `--http-dir` if no separate HTTP run was made.
The report defaults to `test-results.md` in the explicit suite folder. It writes redacted `native-results.json` and per-function, per-case and per-datasource CSV files there.
It preserves setup, call, assertion, cleanup and unfinished outcomes. It does not turn a returned data error into a pass.
Provider-unavailable classifications require the native case's recorded HTTP attribution and measured transport evidence.
The runner preserves these columns in `FunctionResults`:

| Columns | Saved values | Meaning |
|---|---|---|
| `J:L` | Last HTTP status, URL and error | Raw diagnostics can describe validation or cached state |
| `N:Q` | Reset marker, case ID, run ID and HTTP observation | Associates a later URL diagnostic with the current case |
| `R` | Fetch call delta | Counts entries into `smfMacFetch` during the case |
| `S` | FetchURL bridge delta | Counts AppleScriptTask calls to the `fetchURL` handler during the case |

The runner records counter deltas after its intentional reset marker.
A zero `S` delta proves that no fetchURL bridge call occurred during that case.
A positive bridge delta records an attempt. It does not establish a successful wire response or count redirects/authentication requests individually.
A saved HTTP(S) URL with zero bridge attempts cannot establish a provider request.
The report keeps guard, provider-data and case-bound HTTP counts separate.
The separate HTTP probes appear in their own table and never override a native case's classification.

The suite covers the selected public-function calls, not every private function, numbered field, argument combination or provider field.
The earlier [0.6 full-suite report](../../test-results/0.6-dev-20261108-final/README.md) records the previous build.
Read the report from your new run for current results.
Earlier passing counts do not establish the outcome of this fresh run.

## Finish and retain evidence

Stop the fixture server with **Ctrl-C** in its terminal. Restore the normal add-in and calculation settings after native execution.
Keep generated workbooks, response bodies, headers and raw JSON in the private run folder, outside Git.
Retain reviewed, redacted reports and CSV summaries under `mac/test-results` when documenting a run.
Do not check in raw provider captures or local user paths.
Shared JSON, Markdown and CSV outputs redact API keys and anonymous Yahoo cookie/crumb values. Keep the raw workbook and `native-results.private.json` local. Review reports before sharing them.
