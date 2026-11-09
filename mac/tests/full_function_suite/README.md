# Run the SMF 0.5 native suite from a fresh checkout

These tools prepare and read a disposable Excel test workbook. They do not change the packaged add-in or run Excel themselves.
The reviewed set contains 239 cases for 121 public functions, including nine worksheet row/column regression cases.
The checked-in datasource inventory contains 35 domains. It does not enumerate arbitrary caller-supplied URLs.

Run the commands below from the repository root. Use Python 3 and native Excel on Mac.
The three UTC conversion cases assume the Mac timezone is `America/Los_Angeles`.
The short `fnWait` case also requires local time before `23:59:58`.
A fresh checkout still uses the current Mac's Excel installation and installed helper. First-time setup on another Mac remains unverified.

## 1. Prepare a private run folder

```sh
SMF_SUITE_DIR="$(mktemp -d "${TMPDIR:-/tmp}/smf-full-suite.XXXXXX")"
python3 mac/tests/full_function_suite/prepare_suite.py --suite "$SMF_SUITE_DIR"
```

The preparer uses tracked inputs, including `mac/built/0.5-dev/SMF-Mac.xlam`.
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
  --expected-version 3.0.2024.07.30-mac.0.5-dev
```

The generator writes `modSMFFunctionSuite.bas`, `test-cases.json`, `function-inventory.json` and `run-manifest.json`.
It verifies public-function coverage and records the fixture URL, case hashes and expected version.
Generation does not execute any case. A case is not a pass until native Excel saves its completed result.

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
  --metadata "$SMF_SUITE_DIR/datasource_inventory_meta.json"
```

This command makes bounded requests to the inventoried destinations and saves responses, headers and request metadata.
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
Provider-unavailable classifications require the native case's recorded HTTP attribution in columns `N:Q`.
The separate HTTP probes appear in their own table and never override a native case's classification.

The suite covers the selected public-function calls, not every private function, numbered field, argument combination or provider field.
See the [final full-suite report](../../test-results/0.5-dev-20261108/README.md) for current failed cases and their evidence.
Earlier passing counts do not establish the outcome of this fresh run.

## Finish and retain evidence

Stop the fixture server with **Ctrl-C** in its terminal. Restore the normal add-in and calculation settings after native execution.
Keep generated workbooks, response bodies, headers and raw JSON in the private run folder, outside Git.
Retain reviewed, redacted reports and CSV summaries under `mac/test-results` when documenting a run.
Do not check in raw provider captures or local user paths.
Shared JSON, Markdown and CSV outputs redact anonymous Yahoo cookie/crumb values. Keep the raw workbook and `native-results.private.json` local. Review reports before sharing them.
