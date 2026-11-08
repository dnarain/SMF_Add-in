# Live Finviz parser exits

This record covers native Excel checks on November 7, 2026.
Excel was version `16.114.4`, build `16.114.26110224`, on Apple Silicon.

## Current result

Excel exited during two numbered-field lookups and one cross-project parser call.
The cross-project parser call requested no HTTP and used no `Evaluate`.
Four direct target-local probes and a direct live extraction subsequently passed.
The second numbered-field lookup still ended in a restart with a matching initialized page cache.
The cause remains unknown.
No `mac/out/SMF-Mac.xlam` exists. No add-in SaveAs ran during these checks.

The agent saved the repaired Mac menu after those checks.
The direct recipe workaround remains source-only and requires native import and tests.
The Mac is locked, and the user cannot unlock it remotely.
The agent stopped native UI attempts and prepared the offline test bundle.
**Trust access to the VBA project object model** stayed OFF.
The agent changed no security settings during these checks.

## Preserved files

The saved full-source workbook remains intact:

```text
mac/diagnostics/out/SMF-Mac-Full-Source-20261107-01.xlsm
Bytes: 594521
SHA256: 94045758b38f61083d0a83ec065825ad2c051f8170764fa4c1b85cd256925925
```

Its saved project contains 48 standard modules, one Dictionary class, and two document modules.
The four native fixture groups passed before this save:

| Native fixture | Returned result |
| --- | --- |
| `SMFMacTransportSelfTest` | `6` |
| `smfMacHTMLSelfTest` | `55` |
| `SMFMacJSONSelfTest` | `PASS: 84 JSON checks` |
| `SMFMacIntegrationSelfTest` | `5` |

These fixture results do not establish working live extraction or a completed Mac port.

The agent imported `SMFMacFinish.bas` into the separate diagnostic owner only.
The saved owner contains four standard modules:

```text
mac/diagnostics/out/SMF-Mac-Compile-Owner-Manual-20261107-01.xlsm
Bytes: 48694
SHA256: b0e9d1fbb8514b4293b4f81982bd42a6f405860f928bf01e708a9f59796451b5
```

The finish helper did not change the 49 source modules in the full-source workbook.

## Observed sequence

All recorded times use PST on November 7, 2026.

| Step | Observed result |
| --- | --- |
| Native `RCHGetElementNumber("Definition",4072)` | Returned the correct Finviz earnings recipe after approximately 13 seconds |
| Native `RCHGetElementNumber("AAPL",4072)` | CUA timed out after 12.345 seconds. No earnings result returned |
| First exit | macOS recorded PID `76311` terminating at `16:13:02.211` and `QUITTING` at `16:13:02.273` |
| First replacement process | Excel PID `90827` started at `16:13:17`. The agent issued no restart command |
| Reopen of the saved full-source workbook | `Auto_Open` stopped with runtime error `438` in the `modMenu.Controls.Add` path |
| Separate `SMFMacCheckEvalLocal()` | Returned `PASS: qualified Evaluate returned 5; no HTTP call was requested.` |
| Separate `SMFMacCheckFetch()` | Returned HTTP `200`, HTML length `308827`, Earnings and Market Cap markers `True`, and an empty last error |
| Separate `SMFMacCheckParse()` | CUA timed out after 12.436 seconds. The next UI showed a blank `Book1` |
| Second replacement process | Excel PID `92617` started at `16:27:00`. The agent stopped native UI work |

No accessible Excel DiagnosticReport was found for these exits.
The last inspected application-log entries concerned TextInputUI and a WLM error.
Those entries do not identify the cause.

## Target-local fixture results after unlock

The agent imported `SMFMacParserProbe.bas` into the full-source target as a temporary diagnostic module.
The module was imported in memory without a save of the full-source workbook.
The subsequent numbered-field check restarted Excel and discarded that temporary module.
The 49-source baseline on disk remains unchanged at the size and hash recorded above.

Each function ran separately and returned PASS:

| Function | Native result |
| --- | --- |
| `SMFMacParserLoadFixture()` | Loaded `308827` characters |
| `SMFMacParserUppercase()` | Input length `308827`, uppercase length `308827`, `>EARNINGS` index `107260` |
| `SMFMacParserExtractRaw()` | Returned a raw fragment of `154` characters |
| `SMFMacParserExtractText()` | Returned `Nov 02 AMC`, length `10` |

The probe requested no HTTP and used no `Application.Run` or `Evaluate`.
The raw and text functions called `smfGetTagContent` directly within the target project.
These results establish that full-page case conversion and direct parsing succeeded on this fixture.
They do not establish successful owner-to-target transfer of the large HTML argument through `Application.Run`.
They also do not establish successful live `RCHGetElementNumber("AAPL",4072)` evaluation.
No result establishes that the earlier exits are fixed.

## What the checks establish

The definition lookup established that native Excel loaded the local recipe.
The direct fetch established that native `AppleScriptTask` retrieved the Finviz page with mode `4`.
The offline `Evaluate` check established that the qualified integration function returned `5` in that context.

`SMFMacCheckParse` passed the retained HTML to the target through `Application.Run`.
Its source requested no network call and used no `Evaluate`.
The exit does not identify which operation failed inside that call.
The source call includes large-string transfer, full-page `UCase`, tag searches, HTML stripping, and text conversion.

The original numbered-field lookup passed a short URL to `smfGetTagContent` through `Evaluate`.
Both source paths call the full-page parser.
Only `SMFMacCheckParse` passes the full HTML through `Application.Run`.
The target-local fixture checks subsequently passed full-page case conversion and direct extraction.
They do not explain the earlier large-argument `Application.Run` exit.

## Direct URL and prepared-cache results

The agent ran the following direct Immediate-window call in the full target:

```vb
? smfGetTagContent("https://finviz.com/quote.ashx?t=AAPL&p=d","TD",1,">Earnings",,,,1,,4)
```

It returned `Nov 02 AMC` in about 1.5 seconds. This call used the omitted arguments from the original recipe.

The agent then loaded `RCHGetElementNumber("Definition",4072)` successfully and fetched the page again after initialization.
Native readback showed `iInit=1`, blank `sWebCache`, the full target as `ActiveWorkbook`, and this first cache entry:

```text
aData(1,1): 4:https://finviz.com/quote.ashx?t=AAPL&p=d
Len(aData(1,2)): 303114
```

The next `RCHGetElementNumber("AAPL",4072)` call timed out after 12.693 seconds.
Excel showed a blank `Book1` afterward. Read-only process inspection found a replacement Excel process, PID `1953`, started at `17:54:32` PST on November 7.
The agent issued no restart command. The saved full-source workbook still had its original 590999-byte size and SHA256.
The matching initialized cache and successful direct extraction narrow this failure to the numbered-field formula evaluation path.
They do not identify an Excel crash mechanism.

## Menu repair and pending formula changes

The attempted `CommandBarPopup` declarations failed natively with runtime error `13` at the popup assignment.
The object created by Excel reported `TypeName=CommandBarControl` and `Type=10`, matching `msoControlPopup=10`.
The new `code/modMenu.bas` Mac branch creates ten flat, tagged Cell-menu buttons with workbook-qualified actions.
It retains the original nested Windows menu.
The agent imported this module and ran `Call modMenu.Auto_Open` in native Excel successfully, without an error dialog.
The agent saved the workbook and copied it to `SMF-Mac-Full-Source-20261107-menu-passed.xlsm`.
Independent DIR/MODULEOFFSET/OVBA inspection confirmed that saved `modMenu` matches the source, with all ten Mac button calls.
Only `modMenu` changed from the preceding checkpoint. The other 48 code components and both document modules remained byte-identical after decompression.
The Dictionary default-member attribute remains present. Individual menu actions still need a native click check.

The earlier `before-menu-fix` copy remains available with the previous 590999-byte size and SHA256 `78c4f79a12f21882327e19acacf1a5afb10dc322b52108708b30baac32b3f72e`.
The agent began replacing `modMacCode` but the control tool continued to show Excel's worksheet menus while the editor was visible.
Minimizing the worksheet briefly restored the editor's menus, but this did not resolve the repeated focus problem.
The agent closed the test workbook and explicitly chose Don't Save, discarding the incomplete removal and an empty temporary `Module1`.
The saved menu-passed checkpoint therefore retains its complete 49 code components. No formula-dispatch changes were saved natively.
The test workbook is now closed; the control tool last showed the editor for the unrelated Acrobat add-in, which the agent left untouched.

The new `smfMacTryTagContentRecipe` in `code/modMacCode.bas` recognizes a complete literal `smfGetTagContent(...)` call and calls VBA directly.
The two Mac formula branches in `code/modGetElementNumber.bas` try this helper before the existing `Evaluate` fallback.
The helper covers 291 of 1815 effective formula recipes. The other 1524 retain the previous evaluator path.
The helper covers 69 of 81 Finviz definitions; multiplication and nested formulas remain outside this change.
Windows evaluation is unchanged.

`SMFMacRecipeSelfTest()` now contains 20 checks, including omitted arguments, quoted commas, doubled quotes, rejected expressions, and typed overflow.
`SMFMacIntegrationSelfTest()` now expects 25 checks. Both revised functions still require native execution.
The older diagnostic-owner `SMFMacCheckEvalLocal()` still expects the previous integration result of five and must not be reused unchanged with this revision.
Source review and static checks passed; no native result yet establishes that the new dispatcher fixes the restart.

## Fixture file

The completed fixture probes read the public Finviz file at this path:

```text
mac/diagnostics/out/smf-elements-parser-fixture.txt
Bytes: 308899
SHA256: 2e62409cfe1d8ce2a1065f048020b58e1bec0cb0e6d37fc191451d2f30b7514f
```

The fixture does not contain private workbook data.
The agent later refreshed the source package during the offline preparation.
The separate test bundle contains this native checkpoint, the three replacement modules, and all 14 supplied element files.
The bundle README states that the checkpoint does not yet contain the workaround.
Both package manifests identify the native workaround test as pending or identify the absence of a native add-in.
