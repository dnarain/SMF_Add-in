# Alternative Mac builder: native verification incomplete

The next task uses three manual replacements in the saved checkpoint.
The procedure is in [EVALUATE-WORKAROUND.md](EVALUATE-WORKAROUND.md).
The original native fixture groups and repaired menu startup passed after the historical attempts described here.
The revised formula modules still require native tests.
The current state is in [PORT-STATUS.md](PORT-STATUS.md).

## Historical builder attempts

One controlled native attempt completed all 49 staged-source imports. Excel disappeared after the log entered the first transport fixture group.

We did not verify a completed fixture group or a working add-in. The cause of the unexpected exit remains unknown.

A later two-module diagnostic passed basic VBA, a worksheet UDF, all six
transport-parser checks, and one Finviz request through `AppleScriptTask`.
See [diagnostics/RESULTS.md](diagnostics/RESULTS.md). These results do not verify
the full builder or explain its unexpected exit.

Keep VBA project access OFF during the investigation. Do not start another native build without explicit approval.

For the next compilation diagnostic, use `SMFMacPrepareCompile` and
`prepare_compile.applescript`. That entry point explicitly sets both
`runTests=False` and `saveOutput=False`. See
[diagnostics/COMPILE-PLAN.md](diagnostics/COMPILE-PLAN.md).
Preparation of its blank owner ended with another unexpected Excel exit while
importing only the builder module. The preparation wrapper did not run.
Native automation stopped after that exit; see
[diagnostics/EDITOR-IMPORT-EXIT.md](diagnostics/EDITOR-IMPORT-EXIT.md).
A subsequent inspection confirmed a successful import into the manual-copy
workbook. Its builder is saved, and the user subsequently approved its use.
See [diagnostics/MANUAL-IMPORT-RESULT.md](diagnostics/MANUAL-IMPORT-RESULT.md).
The controlled import-only attempt completed all 49 modules, restored Excel's
flags, and returned without target tests or saving. VBA project access is OFF
again. Explicit compilation remains pending; see
[diagnostics/IMPORT-ONLY-RESULT.md](diagnostics/IMPORT-ONLY-RESULT.md).

## Source files

The alternative entry point is `SMFMacBuildNoPrompt` in `build_SMFMac_noprompt.bas`.

The function has three arguments:

```vb
SMFMacBuildNoPrompt(projectFolder, runTests, saveOutput)
```

`projectFolder` is the absolute POSIX path of this project. `runTests` defaults to `True`. `saveOutput` defaults to `False`.

The builder requires the v0.2 `SMFMac.scpt` helper, including `stageModule` and `logBuildState`.

## Source imports

The builder sends the project path and one module filename to the `stageModule` handler. A line feed separates those two values.

The handler copies the source bytes into this private build folder:

```text
~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/SMF-Mac-Build/v0.2/
```

The handler returns an HFS path. The builder verifies the cache folder and filename before it imports that path.

The builder imports `Dictionary.cls` first, followed by 48 standard modules. Its explicit filename list matches the 49 core source files.

The builder does not enumerate the external source folder or invoke the source-file permission dialog.

The controlled native attempt completed these imports without an external source-file permission dialog.

## Default result

The builder creates a new workbook and imports the source files. It then runs the transport, HTML, JSON, and integration fixtures.

The result identifies the target workbook and reports `saved=False`, the import count, and the fixture results.

The target remains open for inspection. The builder creates no `.xlam` file with the default arguments.

The builder returns caught errors with a `FAILED:` prefix and the stage that failed. The result identifies an owned target if one exists.

The builder restores `DisplayAlerts`, `ScreenUpdating`, and `EnableEvents`. It does not change the VBA project-access setting.

## Optional output

If `saveOutput=True`, the builder requests an HFS path through `excelPath`. It then calls Excel `SaveAs` for this output:

```text
mac/out/SMF-Mac.xlam
```

Excel can display a write-permission or overwrite request in this optional phase. We did not verify native output behavior.

The builder reports Excel's actual `FullName`.

After `SaveAs`, verify that the reported file exists at the expected path.

## Checks completed without Excel

Run the static checker from the project folder:

```sh
python3 mac/tests/build_noprompt_static.py
```

The checker verifies the source-file list, staging import path, v0.2 cache, four fixture calls, optional output, and basic VBA procedure boundaries.

The Python checker reads source text. It does not verify native VBA compilation or fixture results.

The latest native log confirms 49 imports and the start of the transport fixtures. Native fixture completion and output-file creation remain unverified.

The builder requires temporary VBA project access for programmatic imports.

After a future approved test, turn that setting OFF.
