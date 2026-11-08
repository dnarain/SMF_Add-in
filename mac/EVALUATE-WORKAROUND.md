# Mac formula workaround: native checks passed

The revised modules passed native recipe `20` and integration `25` checks on November 8, 2026.
AAPL fields 4072 and 4074 passed in worksheet formulas, after a cache clear, and after reopening.
Excel also loaded `mac/out/SMF-Mac.xlam` and ran four checks from a separate workbook.
See `PORT-STATUS.md` for saved files, hashes, and remaining checks.

The test bundle described below retains the original, unpatched November 7 checkpoint.
Its manifest describes that older bundle; it does not describe the later installed add-in.
Keep the replacement procedure for reproducing the original trial.

## What the change does

`RCHGetElementNumber` reads an extraction recipe for each numbered field.
The previous Mac code sends formula recipes to Excel's `Evaluate` function.
Two earnings lookups ended in Excel restarts on that path.
The second lookup had a matching Finviz page in the initialized cache.
Direct downloads and direct `smfGetTagContent` calls returned the expected earnings text.
These results identify the failing path but do not establish the internal Excel fault.

The new `smfMacTryTagContentRecipe` reads complete, literal `smfGetTagContent(...)` calls.
It preserves omitted arguments and quoted text, then calls the VBA extractor directly.
The worksheet formula and numbered field remain the same.
Unsupported expressions retain the previous `Evaluate` path.

The change covers 291 of 1,815 formula recipes, including 69 of 81 FinViz definitions.
The other 1,524 formula recipes retain `Evaluate`.
Twelve FinViz definitions contain multiplication or nested calls and remain outside this change.
No test establishes that those other formulas work on this Mac.

## Files in the test bundle

The bundle contains the saved workbook before this workaround and three replacement modules:

| File | Purpose |
| --- | --- |
| `SMF-Mac-Full-Source-20261107-01.xlsm` | Saved native checkpoint with the repaired Mac menu |
| `replacement-modules/modMacCode.bas` | Direct recipe parser and existing Mac downloader |
| `replacement-modules/modGetElementNumber.bas` | Calls the direct parser before the existing evaluator |
| `replacement-modules/modMacIntegrationTests.bas` | Adds 20 recipe checks with synthetic HTML |
| `smf-elements-*.txt` | The 14 supplied definition and location files beside the workbook |
| `test-manifest.json` | File hashes, expected results, and verification limits |

The workbook does not contain the workaround until the three modules replace its existing modules.
The saved workbook contains 48 standard modules, one Dictionary class, and two document modules.
Its SHA256 before replacement is:

```text
94045758b38f61083d0a83ec065825ad2c051f8170764fa4c1b85cd256925925
```

The saved checkpoint passed the earlier transport, HTML, JSON, and integration groups.
Those groups returned `6`, `55`, `PASS: 84 JSON checks`, and `5`.
The repaired menu startup also returned without an error.
These earlier results do not verify the new recipe parser.

The existing `SMFMac.scpt` helper is already installed on the development Mac.
The test bundle does not reinstall that helper or change Excel security settings.

## Original module replacement procedure

1. Extract the test bundle into a separate folder.
2. Open the workbook from that folder in Excel.
3. Open the Visual Basic Editor.
4. Select project `SMFDiag_20261107_142858` for the test workbook.
5. Remove its existing `modMacCode` module.
6. Import `replacement-modules/modMacCode.bas` into that project.
7. Remove its existing `modGetElementNumber` module.
8. Import `replacement-modules/modGetElementNumber.bas` into that project.
9. Remove its existing `modMacIntegrationTests` module.
10. Import `replacement-modules/modMacIntegrationTests.bas` into that project.
11. Select **Debug > Compile** for that project.
12. Save the test workbook copy.

Keep **Trust access to the VBA project object model** OFF.
Manual editor imports do not require this setting.
Microsoft documents this [Mac import procedure](https://support.microsoft.com/en-us/excel/import-and-shape-data-in-excel-for-mac-power-query#author-and-transfer-power-query-vba-code):
**Project window > right-click an object > Import File**.
The editor also provides **File > Import File**.
The control tool did not reliably activate those menus during the earlier attempt.
The native operator must verify each module name and the selected workbook before removal.

The original checkpoint remains in the project diagnostics folder.
The test workbook copy provides a separate place for these changes.

## Run the native checks

Run each command separately in the Immediate window of the test project:

```vb
? SMFMacRecipeSelfTest()
```

The expected result is `20`.
The recipe checks use synthetic HTML and make no HTTP requests.

```vb
? SMFMacIntegrationSelfTest()
```

The expected result is `25`.
The integration group includes the previous checks and the 20 recipe checks.
If either result differs, stop before the live lookup.

Then activate the test workbook and run these commands separately:

```vb
? RCHGetElementNumber("Definition",4072)
? Len(smfGetWebPage("https://finviz.com/quote.ashx?t=AAPL&p=d",4,0))
? SMFMacHTTPStatus(), SMFMacLastError()
? smfGetTagContent("https://finviz.com/quote.ashx?t=AAPL&p=d","TD",1,">Earnings",,,,1,,4)
? RCHGetElementNumber("AAPL",4072)
```

The definition must name the Finviz earnings recipe.
The page fetch must return HTML with an earnings field.
The HTTP status must be `200`, with an empty error message.
The last command must return the same earnings text as the direct extractor.
The returned date can differ from the earlier saved fixture.

Reloading definitions clears the page cache.
The procedure loads Definition 4072 before the page fetch.

If these commands pass, enter this formula in a cell of the test workbook:

```excel
=RCHGetElementNumber("AAPL",4072)
```

Compare the cell result with the direct extractor result.
Then repeat the worksheet check after a cache clear and after reopening the workbook.
The test remains incomplete until those checks pass without a restart.

Keep this trial limited to fields 4072 and 4074.
The other recipes require separate verification.
The older `SMFMacCheckEvalLocal` diagnostic expects the previous integration result of five.
Do not use that diagnostic with the revised modules.

## Recorded evidence

The detailed native results are in `mac/diagnostics/LIVE-PARSER-EXIT.md` in the source package.
The source package also contains the earlier build records and the original examples.
The test bundle runs no installer and contains no automatic patch macro.
The recorded November 8 checks passed without the earlier restart.
These checks cover the tested recipe group and fields 4072 and 4074.
