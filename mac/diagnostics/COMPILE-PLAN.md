# Full-project compilation check

Status: the controlled import completed and the full-source diagnostic is saved.
The agent repaired Dictionary as a class. The saved project contains 48 standard modules and one Dictionary class.
Project access is OFF. The full project's transport fixture returned `6`.
Debug > Compile returned without an error dialog for the selected target.
The HTML fixture initially failed at `preformatted CRLF`. After the newline fix, its native rerun returned `55`.
A native probe confirmed `vbCr=13`, `vbLf=10`, and a two-character `vbCrLf` containing 10 then 13.
The JSON fixture stopped with ScriptControl error 429 under a `#If Not Mac Then` guard.
The source now uses explicit `#If Mac = 0 Then` guards.
The agent replaced both corrected modules manually. `InitScriptEngine` returned without error.
The native JSON fixture returned `PASS: 84 JSON checks`; integration returned `5`.
All four native fixture groups passed, and the full-source workbook was saved.
Live numbered-field checks and final add-in saving and loading remain pending.
No final `.xlam` exists.
See `IMPORT-ONLY-RESULT.md` for checkpoint paths, hashes, and the exact assertion.

The two-module probe passed basic VBA, a worksheet formula, six response-parser
checks, and a Finviz request from Excel. The earlier full project exited before
its first fixture returned. The cause of that earlier exit remains unknown.

## Prepared entry point

`SMFMacPrepareCompile(projectFolder)` in `mac/build_SMFMac_noprompt.bas` calls:

```vb
SMFMacBuildNoPrompt(projectFolder, False, False)
```

The two explicit `False` arguments disable fixture execution and add-in saving.
The entry point imports the 49 source modules, restores Excel's display and
event settings, and returns its result to the caller.

The builder names its new VBA project `SMFDiag_YYYYMMDD_HHMMSS` and writes that
same identifier into the target workbook's A1. The result and persistent log
identify the owner workbook, target workbook, target VBA project, and component count.

The intended blank owner was `SMF-Mac-Compile-Owner-20261107-01.xlsm` in
`mac/diagnostics/out`. Its file was not persisted, and the builder import did
not complete. The passing two-module probe stays separate.

The verified saved owner is now
`SMF-Mac-Compile-Owner-Manual-20261107-01.xlsm`. It contains the two passing
probe modules plus `build_SMFMac_noprompt`; its size is 37,366 bytes.

## Original import-only sequence

This sequence describes the completed preparation plan. Do not repeat it for the current saved target.

1. Temporarily enable **Trust access to the VBA project object model** for
   automatic source imports. Leave the other macro-security options unchanged.
2. Invoke only `SMFMacPrepareCompile` through `mac/prepare_compile.applescript`.
   The controller requires the dedicated owner filename.
3. Require a successful return with 49 imported modules, all four fixture groups
   marked `not run`, and no add-in file created. Record the actual restored flags.
4. Turn temporary VBA project access OFF again before compiling.
5. Select the uniquely named `SMFDiag_...` target in the VBA Project Explorer.
6. Invoke **Debug > Compile VBAProject** directly. Record any error dialog and
   highlighted source line. Record the Excel process ID before and after.

Do not evaluate an Immediate-window expression or run a target macro before
the explicit Compile command. Either action can trigger implicit compilation.
The target was unsaved when the import-only wrapper returned.
Later native saves created `out/SMF-Mac-Full-Source-20261107-01.xlsm` and its class-repaired backup.
The later native observations are recorded in `IMPORT-ONLY-RESULT.md`.

Stop if Excel exits or displays an unexpected error. The source files and the
persistent log and saved checkpoints remain available.

## Limits

A successful import does not establish compilation. A successful compilation
does not establish correct worksheet results. Fixture execution and a saved,
loaded add-in are later checks.
