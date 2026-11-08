# Native Mac build handoff

## Current state: November 8, 2026 — table-cell fix

Excel loaded `out/SMF-Mac.xlam`, version `3.0.2024.07.30-mac.0.3-dev`.
The revised code passed 57 new native checks:

| Native function | Result |
| --- | --- |
| `SMFMacTableRecipeSelfTest()` | 40 passed |
| `SMFMacTableGuardSelfTest()` | 17 passed |
| `SMFMacRecipeSelfTest()` | 20 passed |
| `SMFMacIntegrationSelfTest()` | 25 passed, including the existing tag checks |

The exact IBD call, `RCHGetElementNumber("AAPL",2003,"Error")`, returned `Error` without terminating Excel.
The direct call, Immediate-window call, and worksheet formula returned the same result.
The worksheet formula also returned normally after cache clearing, reopening, and another cache clear.
The IBD page returned HTTP 200 but lacked the requested industry-group rank field.
The agent verified that the repaired call returned normally. The internal Excel crash cause remains unproved.

Finviz fields 4072 and 4074 still returned `Nov 02 AMC` and numeric `342.37`.
The agent verified the loaded add-in version, IBD result, Finviz earnings, and Finviz price in a separate workbook.
These values record the November 8 provider responses.

The direct dispatchers bypass `Evaluate` for 429 of 1,815 effective formula recipes.
That count includes 291 tag recipes and 138 table-cell recipes.
The table guard returns `pError` for another 35 compound table-cell recipes.
The agent did not verify the evaluator path for every other recipe.
The agent did not resume the full provider batch. Its 18 pending cases remain untested.

The saved files are:

| File | Bytes | SHA256 |
| --- | ---: | --- |
| `out/SMF-Mac.xlam` | 634,949 | `0950924e9856f6f1b6a7361adb55f3c166d5ac2e9a55b13102a89aee630fe2b8` |
| `out/SMF-Mac-TableCell-Checks-20261108.xlsx` | 11,024 | `7b113485d967241ad22434c0d1dcaaa7673518c19d56ca05ebb3fe60bc939a9a` |
| `diagnostics/out/SMF-Mac-TableCell-Fix-Passed-20261108.xlsm` | 636,387 | `480eb7acf9890ce1ba9147b7bf4dbf0fcc297098c8021d69bbff97b8029f41c1` |

The earlier `0.2-dev` add-in remains at `out/SMF-Mac-0.2-dev-before-tablecell-fix.xlam`.
It contains 612,850 bytes with SHA256:
`08141c7a2277d85862dce5aa3b7cc2a2a1a23b4fee9f59bcef2577c73751e177`.
See [TABLE-CELL-CRASH-FIX.md](TABLE-CELL-CRASH-FIX.md) for the source changes and focused native checks.
The sections that follow preserve earlier checkpoints and their status at that time.

## Earlier November 8 checkpoint: Finviz workaround

The three replacement modules ran in native Excel.
`SMFMacRecipeSelfTest()` returned `20`; `SMFMacIntegrationSelfTest()` returned `25`.
The Finviz request returned HTTP `200`, 303,111 normalized characters, and no helper error.
Direct extraction and `RCHGetElementNumber("AAPL",4072)` both returned `Nov 02 AMC`.
The worksheet earnings formula returned the same text; field 4074 returned numeric `342.37`.
Both formulas passed after the SMF cache-clear command and after closing and reopening the workbook.
The reopened workbook also passed an explicit `smfForceRecalculation` run.
Excel displayed no error or restart during these checks.
These values record the November 8 checks and are not ongoing market-data guarantees.

The tested workbook is `diagnostics/out/SMF-Mac-Workaround-Passed-20261108.xlsm`.
It contains 613,952 bytes with SHA256:
`54945e3d91da9b99aa22c97bf151ed71f08a7dc198be4db38bf729cf9e587fdf`.
The original November 7 checkpoint remains unchanged.

Excel saved the separate development add-in to `out/SMF-Mac.xlam`.
It contains 612,850 bytes with SHA256:
`08141c7a2277d85862dce5aa3b7cc2a2a1a23b4fee9f59bcef2577c73751e177`.
Independent inspection confirmed the add-in file type and all 51 VBA components.
Its 48 standard modules, Dictionary class, and two document modules match the tested workbook.
Dictionary retains its default-member attribute. All 14 definition files match `docs/Elements`.
The add-in contains no live worksheet formulas.

Excel's Add-ins dialog loaded `Smf-Mac` with its checkbox selected.
A separate blank workbook then returned the Mac version, HTML result `42`, earnings, and price through the add-in.
Excel saved those checks to `out/SMF-Mac-Checks-20261108.xlsx`.
The source test workbook was closed before the add-in checks, so its duplicate function names were absent.
No security setting was changed during this session.

These checks establish the recorded workaround results and basic add-in loading.
The other recipes, providers, charts, and remaining menu actions still require tests.
A full-project Debug > Compile command was not selected during this session.
The updated procedures compiled and ran through the native fixture and field calls.
The user canceled PR preparation; keep further work focused on the Mac add-in.

## November 7 checkpoint

The source version is `3.0.2024.07.30-mac.0.2-dev`.
No `mac/out/SMF-Mac.xlam` exists.
The full-project transport fixture returned `6` after Dictionary class repair.
Debug > Compile returned without an error dialog for the selected target project.
After a newline fix, the native HTML fixture returned `55`.
The numeric-guard fix resolved ScriptControl error 429 in the native JSON fixture.
The JSON fixture returned `PASS: 84 JSON checks`; integration returned `5`.
All four native fixture groups have passed in the full project.
No working add-in is verified.

The latest saved checkpoint contains 594,521 bytes with SHA256
`94045758b38f61083d0a83ec065825ad2c051f8170764fa4c1b85cd256925925`.
The repaired Mac menu startup passed before this save.
Independent inspection confirmed 48 standard modules, one Dictionary class, and two document modules.
Only `modMenu` changed from the preceding saved checkpoint.
The separate `SMF-Mac-Full-Source-20261107-menu-passed.xlsm` copy contains the same bytes.

Four target-local probes passed on the full Finviz fixture.
Direct live extraction returned the earnings text during the recorded check.
Two numbered-field lookups and one cross-project parser call ended in Excel exits.
The second numbered-field lookup had an initialized cache with the exact URL and 303,114 characters.
The restart mechanism remains unknown.
See `diagnostics/LIVE-PARSER-EXIT.md` for the exact sequence.

The next native task replaces three modules, then runs the direct recipe parser checks.
The replacement modules are `modMacCode`, `modGetElementNumber`, and `modMacIntegrationTests`.
The recipe group expects `20`; the revised integration group expects `25`.
These source changes remain outside the saved native workbook.
The older `SMFMacCheckEvalLocal` helper still expects five and must not be reused unchanged.
The procedure and expected results are in [EVALUATE-WORKAROUND.md](EVALUATE-WORKAROUND.md).

The Mac is locked, and the user cannot unlock it remotely.
The agent discarded the incomplete module replacement without saving it.
Project access stayed OFF. No add-in SaveAs ran.
The current blocker is native access to Excel for the import and checks.

The installed Excel version is now verified as 16.114.4, build 16.114.26110224.
Earlier 16.114.3 notes are outdated.
Earlier native attempts also ended in unexpected exits and restarts.
The user confirmed that they did not restart Excel during those earlier attempts.
The agent issued no quit or restart command during the recorded exits.

The user restored **Trust access to the VBA project object model** to OFF after
the latest crash. Keep that setting OFF for the minimal diagnostics.
The general macro-security options were not changed by the agent.

## Offline PR preparation

The branch is `dnarain/mac-native-support` and still starts at upstream commit
`a0daad05699b1a7e213cf72ee641f30fc615af46`.
No local commit, push, fork, or PR has been created.

The version constant now selects `3.0.2024.07.30-mac.0.2-dev` on Mac only.
Windows keeps the upstream version, `3.0.2024.07.30`.
The conditional-source checker confirmed both numeric Mac values, `1` and `-1`, and Windows `0`.
The saved native checkpoint remains unchanged.

`.gitignore` now includes the Mac Python checks and their bundled signature baseline.
It excludes generated import copies, output workbooks, captured pages, and local ZIP packages.
The selected source files still need review before staging a PR.
Do not stage every file under `mac/`.

`README.md` identifies shared Windows changes and the required Windows checks.
The new recipe checks, numbered-field formulas, example workbooks, and add-in loading still require Excel.

## Proposed upstream file list

The minimal PR would contain these 24 files:

```text
.gitattributes
.gitignore
code/Dictionary.cls
code/JsonConverter.bas
code/modGetElementNumber.bas
code/modGetYahooJSONData.bas
code/modJSONExtract.bas
code/modMacCode.bas
code/modMacHTML.bas
code/modMacIntegrationTests.bas
code/modMacJSONTests.bas
code/modMenu.bas
code/smfGetYahooHistoryCSV_.bas
code/smfUtilities_.bas
licenses/VBA-Dictionary-LICENSE.txt
mac/README.md
mac/SMFMac.applescript
mac/build_SMFMac_noprompt.bas
mac/install_helper.command
mac/tests/build_noprompt_static.py
mac/tests/html_static_checks.py
mac/tests/port_static_checks.py
mac/tests/transport_tests.py
mac/tests/upstream-signatures.json
```

Keep this handoff, session diagnostics, captured pages, checkpoint workbooks, generated copies, and ZIPs local.
The older build automation and checkpoint-specific packagers stay outside this proposed PR.
`mac/tests/vba_project_inspect.py` is an optional developer tool beyond the 24 files.
The proposed PR must remain a draft while native and Windows checks remain incomplete.

## Historical native work

An earlier preparation stopped on November 7 at 13:47:54 PST, while importing
only the builder module into a new blank owner. No wrapper function, full-source
import, or explicit compilation ran. macOS recorded exit status 0 but no requester
or fault signal. See `diagnostics/EDITOR-IMPORT-EXIT.md`.
At that point all native automation stopped. Later user-authorized work is
recorded below.

The user then asked the agent to inspect the open manual-copy workbook.
The builder was present under the collapsed Modules folder. The agent expanded
the folder, verified the safe entry point visually, and saved the workbook.
The saved owner is 37,366 bytes and contains the correct three modules.
See `diagnostics/MANUAL-IMPORT-RESULT.md`.
The user then approved temporary VBA project access. The controlled import-only
attempt completed all 49 modules into `SMFDiag_20261107_142858 (Sheet1)`.
The builder returned with tests and saving disabled, and all three Excel flags
restored to True. The agent restored VBA project access to OFF and verified it.
At that point the wrapper had not run a target fixture or explicit compilation command.
See `diagnostics/IMPORT-ONLY-RESULT.md`.
An earlier Mac lock interrupted editor inspection after the import.
After unlock, the target remained present. The agent saved its full source.
CFB inspection confirmed that Dictionary had imported as a standard module.
A native CRLF class import repaired it. Both saved checkpoints contain 48 standard
modules, one Dictionary class, and two document modules. The default-Item attribute survived.

Verified checkpoint: `diagnostics/out/SMF-Mac-Full-Source-20261107-01.xlsm`.
It contained 451,477 bytes, with SHA256
`4411fbaa524f572f6dbb9d4dbf040da8818bb17c4b7cc6b8b014d6867f4b7e40`.
The class-repaired backup contains 451,814 bytes, with SHA256
`58e77947a2a8d81382bf64dde105bab794b6313fe5791e4f7a7c280117d86bd4`.
An accidental cut in modMenu was undone. Its decompressed source matches the backup byte for byte.
Details are in `diagnostics/IMPORT-ONLY-RESULT.md`.

Project access stayed OFF during the later native checks.
The full project's transport fixture returned `6`.
The target's Debug > Compile command returned without an error dialog.
The HTML fixture stopped at `preformatted CRLF`: expected `[A<LF><CR>B]`,
actual `[A<LF><CR><LF><CR>B]`. The agent selected End.

A source-only HTML fix handles actual CR/LF pairs in either order and adds eight fixture cases.
The source has 55 cases.
After another lock, the user unlocked the Mac. A native Immediate-window probe confirmed:

```text
AscW(vbCr) = 13
AscW(vbLf) = 10
Len(vbCrLf) = 2
vbCrLf first character = 10
vbCrLf last character = 13
```

The agent replaced the HTML module manually. Its native fixture returned `55`.
The next JSON fixture stopped with runtime error 429 at the Windows ScriptControl constructor.
Its `#If Not Mac Then` guard was active despite the Mac execution path.
A native probe returned `Not 1 = -2` and `Not True = False`.
The source now uses `#If Mac = 0 Then` in `modJSONExtract.bas` and `smfUtilities_.bas`.
The checker covers Mac values `1` and `-1` and rejects the old guard.
The agent replaced those two modules manually. `InitScriptEngine` returned without error.
The JSON fixture returned `PASS: 84 JSON checks`; integration returned `5`.
The agent saved the full-source workbook after these checks.
Its latest size is 590,999 bytes and SHA256 is
`78c4f79a12f21882327e19acacf1a5afb10dc322b52108708b30baac32b3f72e`.
CFB inspection verified 48 standard modules, one Dictionary class, and two document modules.
All 49 source names and the hidden default-Item attribute remain present.
Decompressed HTML source matches the current source after line-ending normalization.
JSON and utility source match apart from VBE identifier casing.
Do not repeat the 49-module import or re-enable project access for this work.

The user authorized a minimal diagnostic workbook. That completed scope was
basic VBA execution, the isolated transport parser, and one Mac helper request.
Do not repeat the full 49-module build during these checks.

## Earlier full-project attempt

The v0.2 `SMFMac.scpt` helper was installed before this attempt.
`SMFMacBuildNoPrompt` staged and imported all 49 core source files.
The target workbook's reported `Name` was `Sheet1`.
The builder did not save an add-in.

The persistent log recorded these final lines for that earlier attempt:

```text
Imports completed; target=Sheet1; imported=49
Run the transport fixtures
```

Excel exited before `Application.Run` returned a fixture result.
The first fixture uses fixed strings and makes no HTTP or AppleScript request.
The last log entry does not distinguish VBA compilation from test execution.
The exact cause remains unknown.

No owned automation script remained after the stop.
No accessible Excel or Microsoft crash report was found in the checked report folders.

## Minimal diagnostics

See `diagnostics/README.md` and `diagnostics/SMFMacProbe.bas`.
Use a uniquely named, saved disposable `.xlsm` workbook.
Import the probe through the VBA editor's File > Import File command.
Run its basic function and check the same function in a worksheet formula.
Then import the unchanged `code/modMacCode.bas` and run its six response checks.

These manual imports do not require programmatic VBA project access.
After the local checks pass, make one isolated Finviz request through the same module.
Stop and record evidence if Excel exits or displays an error.

The basic function passed three checks directly and through a worksheet formula.
The unchanged transport module passed its six fixed-response checks.
VBA project access stayed OFF, and Excel stayed in process `67044`.
The same workbook fetched the Finviz AAPL page through `AppleScriptTask`.
It returned HTTP 200, 308,819 characters, Earnings and Market Cap markers,
and an empty last-error value.
The detailed record is `diagnostics/RESULTS.md`.
The diagnostic workbook was saved successfully at the path recorded there.
Its ZIP passed integrity checks and contains `xl/vbaProject.bin`.
The minimal attempt is complete. Native automation stopped after final readback.

## Completed controlled import and pending checks

`SMFMacPrepareCompile` completed the later import with tests and saving disabled.
The agent invoked it through the saved owner's Immediate window.
The external `mac/prepare_compile.applescript` controller was not used for that invocation.
The agent restored project access to OFF before the subsequent native checks.
The detailed record is `diagnostics/IMPORT-ONLY-RESULT.md`.

The current full-source diagnostic is saved. Do not repeat its 49-module import.
All four native fixture groups passed before the live parser investigation.
The live numbered-field and retained-HTML parser calls ended in Excel exits.
Four later target-local parser probes passed on the full Finviz fixture.
Add-in saving and loading remain pending.
The causes of the unexpected exits remain unverified.

## Earlier attempts and disposable files

The initial builder stopped with runtime error 1004 when VBA project access was OFF.
The old external-folder permission route then stalled and Excel restarted.
Do not repeat `build_SMFMac.bas` or its unsuccessful batch editor automation.

The user confirmed that the test and recovery workbooks contain disposable SMF tests.
No further recovery work is requested.
`mac/out/SMF-Mac-Editor-Build.xlsm` is a disposable partial build.
It is not a working add-in and is excluded from the package.

## Completed checks outside Excel

- The standalone helper passed 24 HTTP and temporary-file tests.
- The HTML entity table passed three independent reference checks.
- The conditional/dependency scanner passed 21 fixtures.
- The scanner checked 49 modules and 464 signatures across four numeric Mac configurations.
- The alternative builder passed its manifest and import-path checks.
- `git diff --check` passed.

The helper retrieved expected markers from Finviz, Yahoo history, Yahoo options,
and a Zacks page. The legacy FRED URL failed.
These requests did not run through worksheet formulas.

The explicit target Compile command returned without an error dialog, and the full-project transport fixture passed.
The native HTML fixture passed 55 checks after the newline fix.
The native JSON fixture passed 84 checks, and the integration fixture returned `5`.
The repaired menu startup passed; individual menu actions still require checks.
Live worksheet extraction, charts, add-in saving, and add-in loading remain unverified.
