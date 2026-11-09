# Experimental Mac testing plan

The [personal fork](https://github.com/dnarain/SMF_Add-in) remains experimental.
The working branch is `dnarain/mac-native-support`. No pull request is open for this Mac port.
The recorded checks cover specific inputs. Further checks must establish workbook behavior and field correctness.
See the [full-suite results](test-results/0.5-dev-20261108/README.md) and [fresh-checkout instructions](tests/full_function_suite/README.md) for the current 239-case workflow.
Run the suite with the current Mac's Excel and installed helper. First-time installation on another Mac remains unverified.

## Workbook sessions

- Check copies of real workbooks and representative upstream samples.
- Verify formulas after normal calculation, an SMF cache clear, reopening, and a new Excel session.
- Repeat the checks across separate Excel sessions.
- Record the Excel build, Mac model, add-in version, inputs, returned values and failures.
- Require zero Excel exits or restarts during the recorded checks.
- Preserve the exact case and saved workbook after an unexpected exit.

## Provider values

- Compare each claimed field with the provider response captured during the same check.
- Record the retrieval time, currency, units and reporting period.
- If the provider supplies a timestamp, record it with the retrieval time.
- Verify the requested ticker and date range for historical data.
- Record whether each value came from SMF's cache or a new request.
- Document supported fields and unavailable providers separately.

## Failure diagnosis

- Compare an extraction against the captured payload before changing its URL or parsing rule.
- Separate invalid URLs, changed pages, access restrictions, parsing errors and Mac-specific failures in the result record.
- Keep an original failed result after a fix.
- Repeat the exact failed case after a cache clear and workbook reopening.
- Verify every supported evaluator path in native Excel.

## Installation and Excel controls

- Verify helper installation, repeated installation and the helper backup.
- Build from a fresh checkout without local diagnostic files.
- Verify add-in loading and all menu actions in separate workbooks.
- Verify comment text, local pictures and provider chart pictures.
- Record the owned worksheet picture/chart type, visibility and cleanup, and inspect the displayed chart.
- Verify option expirations, strikes and quotes against the same provider response.
- Check small output ranges and larger output ranges for options.

## Before reconsidering a pull request

- Run the source checks and all relevant native fixtures after changes.
- Make the supported fields, known failures and untested functions explicit.
- Ask a Windows maintainer to compile the project in 32-bit and 64-bit Excel.
- Verify Windows JSON fields and keys, Yahoo examples, and repeated menu startup and removal.
- Verify that Windows still reports version `3.0.2024.07.30`.
- Keep disposable test binaries, captured pages and raw diagnostic output outside Git.
- Retain the reviewed shareable add-in and examples under `mac/built`, with build information and checksums.
- Redact Yahoo cookie and crumb values before sharing reports.

The full-function scripts require explicit run paths, retained metadata and local evidence.
The recorded UTC fixtures assume `America/Los_Angeles` on the Mac.
The report scripts summarize saved evidence. They do not repeat the native Excel checks.
