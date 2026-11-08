# Mac port status

This record describes checks on November 7 and 8, 2026.

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
The agent recorded results for all 18 previously pending cases in the separate [remaining-provider report](REMAINING-PROVIDER-RESULTS-20261108.md).

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

The saved checkpoint contains 48 standard modules, one Dictionary class, and two document modules.
It contains 594,521 bytes, with SHA256
`94045758b38f61083d0a83ec065825ad2c051f8170764fa4c1b85cd256925925`.
Its repaired Mac menu ran without an error before the agent saved it.
Independent inspection confirmed that only `modMenu` changed from the preceding checkpoint.

All four original native fixture groups passed: transport `6`, HTML `55`, JSON `84`, and integration `5`.
Four separate, target-local probes also passed on the 308,827-character Finviz fixture.
Direct live `smfGetTagContent` extraction returned `Nov 02 AMC` during the recorded check.
That date is a test result from November 7, 2026.

Excel exited during two numbered-field lookups and one cross-project parser call.
The last numbered-field lookup had the exact Finviz URL in the initialized cache.
The restart mechanism remains unknown.
See `diagnostics/LIVE-PARSER-EXIT.md` for the process evidence and verification limits.

The new direct recipe parser covers 291 of 1,815 formula recipes, including 69 of 81 FinViz definitions.
The other 1,524 formula recipes retain `Evaluate`.
The three revised formula modules passed source review and static checks.
They did not run in native Excel yet.
The new recipe group expects `20`; the revised integration group expects `25`.

The Mac is locked, and the user cannot unlock it remotely.
The agent closed the incomplete replacement attempt without saving its changes.
The saved checkpoint therefore retains its complete code and the earlier formula modules.
VBA project access remains OFF. No add-in SaveAs ran, and no `.xlam` exists.
The next native procedure is in [EVALUATE-WORKAROUND.md](EVALUATE-WORKAROUND.md).

## Completed checks

- AppleScript compilation passed with Apple's `osacompile` command.
- The standalone AppleScript helper passed 24 local HTTP and temporary-file tests.
- The HTML entity table passed three independent reference checks.
- The static checks passed 21 parser and scanner fixtures.
- We compared 464 public function signatures across four numeric Mac configurations with the upstream source.
- We scanned the active Mac code for selected Windows dependencies.
- `git diff --check` passed.
- After class repair, the full-project transport fixture returned `6` in native Excel.
- The target project's Debug > Compile command returned without an error dialog.
- The full project's native HTML, JSON, and integration fixtures returned `55`, `PASS: 84 JSON checks`, and `5`.

The static checks do not replace the Excel VBA compiler or runtime tests.

## Live HTTP requests

The actual Mac helper retrieved these responses with the expected field markers:

| Source | Response | Check |
| --- | --- | --- |
| Finviz AAPL page | HTTP 200 | Earnings and Market Cap text present |
| Yahoo AAPL history | HTTP 200 | Timestamp and indicator data present |
| Yahoo AAPL options | HTTP 200 | Expiration dates and strikes present |
| Zacks AAPL page | HTTP 200 | AAPL text present |
| FRED DGS10 legacy data URL | Failed | HTTP/2 error, then timeout with HTTP/1.1 |

HTTP success does not establish that every SMF extraction rule returns the correct field.
The live request details are in `live-source-results.json`.

## Earlier native Excel checks

The installed Excel app is now verified as 16.114.4, build 16.114.26110224,
on this Apple Silicon Mac. Earlier 16.114.3 notes are outdated.
The builder macro was imported into a new disposable workbook.
Excel granted access to the project files, then denied programmatic VBA project access with runtime error 1004.
The user temporarily enabled VBA project access for the later build attempts.
The editor route imported one finish helper into a separate test project.
The initial editor import route did not complete the 49-module import.
Excel exited and reopened unexpectedly during these attempts. Its recovery copies were opened with macros disabled.
The user confirmed that these workbooks are disposable failed tests.
The source package excludes the disposable workbooks.
No `.xlam` file exists yet. The package contains source code and the builder.
Later native checks are recorded below. SMF field-extraction formulas remain unverified.

## Latest investigation

The user temporarily enabled VBA project access for the automatic builder.
The old builder stopped at an external source-folder access request.
Excel then exited and restarted unexpectedly.
No accessible Excel crash report was found.
All full-project automation stopped. The user turned VBA project access OFF again.

An alternative builder and helper changes are prepared as `mac.0.2-dev`.
The builder stages modules inside Excel's own cache and avoids the earlier folder-access request.
Mac field-definition reads also use the helper.
The revised helper passed standalone tests. The revised builder passed static checks.
The updated helper was installed for one controlled native attempt.
The alternative builder imported all 49 source modules.
The persistent log then recorded `Run the transport fixtures`, and Excel exited.
No fixture result was returned. The log does not distinguish compilation from execution.
The user restored VBA project access to OFF.
No fix for the unexpected restarts is verified.

The user authorized a separate minimal workbook to check basic VBA execution and
the transport parser. See `diagnostics/README.md`. Those checks do not load the full project.
The basic function passed three checks directly and through a worksheet formula.
The isolated `SMFMacTransportSelfTest` passed all six response checks.
VBA project access stayed OFF, and Excel remained in the same process.
The native helper HTTP check is recorded separately in `diagnostics/RESULTS.md`.
It returned HTTP 200 and 308,819 characters from the Finviz AAPL page, with
Earnings and Market Cap markers present and no helper error.
This check ran through `AppleScriptTask` from real Excel.

The next preparation failed during a single builder-module import into a new
blank owner. No new full-project import or explicit compilation occurred.
macOS recorded exit status 0 at 13:47:54 PST; the cause remains unknown.
VBA project access stayed OFF. All native automation stopped.
See `diagnostics/EDITOR-IMPORT-EXIT.md` for the exact evidence and limits.

A later inspection found the builder successfully imported in the separate
manual-copy workbook, hidden under a collapsed Modules folder. The agent
verified `SMFMacPrepareCompile` uses `False, False`, saved the workbook, and
confirmed its three modules on disk. No full-project import or explicit
compilation ran during that inspection.
See `diagnostics/MANUAL-IMPORT-RESULT.md`.

The user then approved temporary project access for the controlled import-only
attempt. All 49 modules imported into `SMFDiag_20261107_142858 (Sheet1)` with
51 total components. Tests and saving stayed disabled, and Excel's display and
event settings were restored. The agent turned project access OFF again and
verified it. The wrapper did not execute target fixtures or save an add-in.
See `diagnostics/IMPORT-ONLY-RESULT.md`.

An earlier lock interrupted editor inspection. After unlock, the target remained
present in the VBA editor. The agent saved the full-source diagnostic workbook.
CFB inspection found Dictionary classified as a standard module.
A native class import repaired it. The saved project has 48 standard modules,
one Dictionary class, and two document modules. Its default-Item attribute survived.

The verified checkpoint is `diagnostics/out/SMF-Mac-Full-Source-20261107-01.xlsm`.
It contained 451,477 bytes, with SHA256
`4411fbaa524f572f6dbb9d4dbf040da8818bb17c4b7cc6b8b014d6867f4b7e40`.
The class-repaired backup contains 451,814 bytes, with SHA256
`58e77947a2a8d81382bf64dde105bab794b6313fe5791e4f7a7c280117d86bd4`.
Both saved projects passed ZIP integrity and CFB component checks.
An accidental cut in modMenu was undone. Decompressed source matches the backup exactly.

VBA project access stayed OFF. The full project's transport fixture returned `6`.
The agent selected Debug > Compile for the target. Excel returned without an error dialog.
The HTML fixture then stopped at `preformatted CRLF`: expected
`[A<LF><CR>B]`, actual `[A<LF><CR><LF><CR>B]`.
The agent selected End.

A source-only HTML fix handles actual CR/LF pairs in either order and includes 55 fixture cases.
After a lock and manual unlock, a native probe confirmed `vbCr=13` and `vbLf=10`.
It confirmed that this Excel's two-character `vbCrLf` contains 10 followed by 13.
The agent replaced the HTML module manually. Its native fixture returned `55`.
The JSON fixture then stopped with runtime error 429 at the Windows ScriptControl constructor.
The constructor's `#If Not Mac Then` guard was active on this Mac.
The source uses explicit `#If Mac = 0 Then` guards instead.
The checker now tests both numeric Mac representations and rejects the old guard.
The agent replaced both corrected modules manually. `InitScriptEngine` returned without error.
The native JSON fixture returned `PASS: 84 JSON checks`; integration returned `5`.
All four native fixture groups have passed in the full project.

The latest saved full-source workbook contains 590,999 bytes, with SHA256
`78c4f79a12f21882327e19acacf1a5afb10dc322b52108708b30baac32b3f72e`.
Independent CFB inspection verified all 49 source names, the Dictionary class, and its default-Item attribute.
Saved HTML source matches the current source after line-ending normalization.
Saved JSON and utility source match apart from VBE identifier casing.
Live numbered-field parsing remains blocked by the exits recorded in `diagnostics/LIVE-PARSER-EXIT.md`.
Final add-in saving and loading remain pending.
No final add-in exists. See `diagnostics/IMPORT-ONLY-RESULT.md` for the exact evidence.

## Function families

| Functions | Port work | Current verification |
| --- | --- | --- |
| HTTP download and shared page cache | Mac transport and failure handling added | Native Finviz fetch passed with HTTP 200. Exact mode-4 cache entry and its 303,114-character value were read back |
| HTML-to-text and tag extraction | Pure VBA HTML helper and newline fix added | HTML fixture passed 55 checks. Target-local full-page probes and direct live earnings extraction passed. Cross-project parser exit remains unexplained |
| Yahoo JSON fields and key lists | Dictionary class repaired; numeric Mac guards corrected | Native JSON fixture passed 84 checks |
| Numbered fields, including Finviz and Zacks | Direct parser added for 291 simple recipes | Definition 4072 passed. Earlier AAPL 4072 calls ended in exits. Revised parser and RCH callers require native checks |
| Yahoo price history, dividends and splits | Existing calculations retained; direct Windows CSV helper adapted | History HTTP response checked; native arrays pending |
| Options expirations, strikes and quotes | Existing functions use the Mac transport | Options HTTP response checked; native arrays pending |
| FRED economic data | Existing function retained | Legacy URL failed during live requests |
| Technical indicators | Existing VBA retained | Native calculations pending |
| Charts, pictures and menus | Ten flat Mac menu buttons replace unsupported popup access | Menu startup passed and was saved. Individual actions, charts, and pictures require native checks |
| Other legacy provider recipes | Existing definitions retained | Source-by-source checks pending |

The package contains the upstream examples for further checks.
The port does not establish that every old provider recipe still works.
