# SMF native Mac development add-in

To try the built add-in, download the [working branch](https://github.com/dnarain/SMF_Add-in/tree/dnarain/mac-native-support) with **Code > Download ZIP**.
Unzip it and open [`mac/built/0.5-dev`](built/0.5-dev/README.md) for the add-in, three examples and installation instructions.
Keep the bundle files together. The build is experimental. See the [full-suite results](test-results/0.5-dev-20261108/README.md) for current failures and limits.

This Mac port is experimental work in the [dnarain/SMF_Add-in fork](https://github.com/dnarain/SMF_Add-in).
The working branch is `dnarain/mac-native-support`.
There is no pull request for this Mac port. Further workbook and provider checks come first.

The current development source version is `3.0.2024.07.30-mac.0.5-dev`.
The [Yahoo history row-count fix](YAHOO-HISTORY-ROWCOUNT-RESULTS-20261108.md) passed 46 native checks of the actual public function.
The Mac declaration now uses `Double` for the ninth argument, `pCols`; `pRows` remains `Integer`.
The Windows declaration retains `Integer` for `pCols`. All argument names, positions and defaults remain unchanged.
Excel built, loaded and restarted with the audited `0.5-dev` add-in.
`mac/out/SMF-Mac.xlam` and `mac/out/SMF-Mac-0.5-dev.xlam` are byte-identical: 673,519 bytes, SHA-256 `3b9423659992348920ee08bd32898a9b20187073116271b20e7d090a33514331`.
After a clean quit/restart, all eight fresh [history worksheet checks](out/SMF-Mac-History-Checks-20261108.xlsx) passed, with a live five-row, two-column result.
The version formula reported `mac.0.5-dev` and the correct `mac/out` directory.
The add-in remained checked, calculation was **Automatic**, and VBA project access was OFF.
All 14 adjacent element files match `docs/Elements/` byte for byte. This fix does not change the AppleScript helper.

The following records describe the earlier `0.4-dev` build.
The [data and picture repair run](DATA-PICTURE-REPAIR-RESULTS-20261108.md) has final results for all 59 selected cases.
The isolated workbook recorded 52 passes and seven failed provider data checks.
All 25 regression/offline cases, 13 live data cases and 14 image/chart cases passed.
Seven additional [derived-history cases](tests/full_function_suite/derived_history_cases.json) passed against cached AAPL observations without HTTP requests.
The two batches contain 66 cases: 59 PASS and seven provider CHECK_FAILED results.
The [seven external provider fields](DATA-PICTURE-REPAIRS-20261108.md#seven-fields-still-blocked-or-unavailable) remain blocked or unavailable.
The previous `0.3-dev` add-in is preserved at `mac/out/SMF-Mac-0.3-dev-before-data-picture-fix.xlam`.
The isolated workbook passed the 50 nested-recipe, 12 history-conversion and 25 captured-provider checks.
At that checkpoint, Excel built and loaded the valid `0.4-dev` file at `mac/out/SMF-Mac.xlam`.
It is 713,180 bytes, with SHA-256 `ad7cf4d0a79dea7b28a57f0deb6945118e172ecd5b57efb6adbd81d375cbc18e`.
The file has the correct Excel add-in content type. Its 51 core modules match the passed 66-case snapshot.
It contains no active diagnostic test modules, and its 14 adjacent definition files match `docs/Elements/` byte for byte.
The add-in checkbox is enabled, and the source repair workbook is closed.
Excel uses automatic calculation, the status bar is reset, and VBA project access is OFF.
After registration, a full Excel quit and relaunch loaded `0.4-dev` automatically without an alert.
The reopened [data and picture check workbook](out/SMF-Mac-Data-Picture-Checks-20261108.xlsx) returned all 11 PASS results.
The StockCharts picture and native three-point chart remained visible.
These worksheet checks overlap the earlier cases and do not increase the 66-case total.
That build returned `#VALUE!` for the checked worksheet calls with explicit positive `pRows`.
The installed `0.5-dev` add-in now passes the focused explicit-size worksheet controls.

The previously verified `0.3-dev` build made the following changes.
It adds direct table-cell recipe calls after an IBD lookup crashed Excel during the broader provider tests.
All 57 new parser and guard checks passed in native Excel.
The repaired IBD lookup returned normally after a page-cache clear and after reopening the workbook.
The direct dispatchers handle 429 of 1,815 formula recipes: 291 tag recipes and 138 table-cell recipes.
The table guard returns `pError` for another 35 compound table-cell recipes.
These recipe counts describe the verified `0.3-dev` build.
The evaluator paths for other recipes still require native checks.
See [the crash fix](TABLE-CELL-CRASH-FIX.md), [the recorded tests](TEST-RESULTS-20261108.md), and [the testing plan](TESTING-PLAN.md).

The earlier [remaining-provider checks](REMAINING-PROVIDER-RESULTS-20261108.md) recorded three passes, eleven failed data checks and four image assertion errors.
The supplemental local-image case also raised an assertion error.
The image cases returned `#VALUE!`; reading the comment fill raised error `445`.
The operator observed no Excel exit during that batch and the image rerun.

Excel built and loaded the `0.3-dev` add-in during the earlier recorded checks.
That valid binary is now preserved as `mac/out/SMF-Mac-0.3-dev-before-data-picture-fix.xlam`.
The complete 15-file rollback copy remains in `mac/out/rollback-0.3-dev-before-data-picture-20261108/`.
The previous add-in is preserved as `mac/out/SMF-Mac-0.2-dev-before-tablecell-fix.xlam`.
The separate add-in checks are saved in `mac/out/SMF-Mac-TableCell-Checks-20261108.xlsx`.
These generated binaries and test workbooks remain local. They are excluded from Git.

The previous add-in was built and loaded in native Excel earlier on November 8, 2026.
Its recipe and integration groups returned `20` and `25`.
The AAPL earnings and price formulas passed after a cache clear and after reopening the test workbook.
A separate workbook then called the loaded add-in and returned the Mac version, HTML result `42`, earnings, and price.
The saved checks are in `mac/out/SMF-Mac-Checks-20261108.xlsx`.

These checks verify the tested functions and fields.
Other providers, legacy recipes, charts, and individual menu actions still require checks.
The original transport, HTML, and JSON fixture groups passed on November 7.

That earlier add-in contains Mac version `3.0.2024.07.30-mac.0.2-dev`.
Windows retains the upstream version, `3.0.2024.07.30`.
The source starts from upstream commit `a0daad05699b1a7e213cf72ee641f30fc615af46`.

## Developer build route

The builder remains experimental. Earlier automatic attempts ended in Excel restarts.
A later import-only attempt loaded all 49 modules. The optional automatic add-in save remains unverified.

1. Install the helper with `mac/install_helper.command`.
2. Create `mac/out` and copy all `docs/Elements/*.txt` files into it.
3. Create a separate macro-enabled workbook and import `mac/build_SMFMac_noprompt.bas` through the VBA editor.
4. Temporarily enable **Trust access to the VBA project object model** for the automatic import.
5. Run `SMFMacPrepareCompile` with the absolute repository folder as its argument.
6. Restore VBA project access to OFF after the import returns or fails.
7. Select the reported target project and verify its 48 standard modules and one Dictionary class.
8. Select **Debug > Compile**, then save a test `.xlsm` workbook in `mac/out`.
9. Run the fixture commands below before any live worksheet lookup.

`SMFMacPrepareCompile` imports source without running target functions or saving an add-in.
The helper stages source inside Excel's private cache before each import.
The builder reports the target workbook, project name, imported count, and restored Excel settings.
Stop if Excel exits or the builder reports a failure.

After the fixtures pass, test numbered fields 4072 and 4074 through worksheet formulas.
The verified `0.3-dev` dispatchers handled 429 literal recipes without `Evaluate` and guarded another 35 compound table-cell recipes.
The current source also handles selected nested calls through `smfMacTryNestedRecipe`.
The dispatcher validates the full expression before calling the supported VBA functions.
Unsupported compound table-cell recipes return `pError`. Other expressions can still reach the previous evaluator and require separate native checks.
The 40 table-parser checks, 17 guard checks, 20 tag checks, and 25 integration checks passed natively.
Save and load an `.xlam` only after the workbook checks pass.
The optional builder save path is `mac/out/SMF-Mac.xlam`.
The builder does not create the output folder or copy element files.
The confirmed manual save used a fresh POSIX path in Excel's private cache, then `SaveCopyAs` to a previously absent output file.
Earlier HFS-path and existing-target attempts returned error `1004`. The cause of those failures remains undetermined.
The user authorized file access, and `GrantAccessToMultipleFiles` returned `True` before the successful save.

## Source changes

- The Mac downloader calls `AppleScriptTask` and macOS `curl`.
- Some Yahoo requests obtain an anonymous cookie and crumb. The Mac history adapter reads chart JSON and produces the expected CSV.
- The Mac HTML parser uses VBA and does not require Windows `HTMLDocument`.
- The Mac JSON parser uses the bundled `JsonConverter` and `Dictionary` class.
- Mac compiler branches exclude Windows networking and Script Control code.
- Decimal conversion preserves more than four decimal places.
- Failed HTTP requests do not enter the successful-page cache.
- Mac field-definition reads use the AppleScript helper.
- Supported literal `smfGetTagContent` and `RCHGetTableCell` recipes call VBA directly.
- `smfMacTryNestedRecipe` handles selected `DateValue`, `smfConvertData` and `smfStrExtr` calls around supported tag and Yahoo JSON calls.
- Unsupported compound table-cell recipes return `pError` before the evaluator runs.
- Mac `RCHCreateComment` uses an embedded worksheet picture or native numeric chart, with a cell note for text and diagnostics.
- The Mac menu uses eleven flat buttons, including **SMF: Show/hide cell chart**, because popup access failed natively.

Existing function names and numbered fields remain.
The [four repaired field recipes](DATA-PICTURE-REPAIRS-20261108.md#four-recipe-changes) use current URLs or field markers from the same providers.
The other field definitions retain their upstream URLs and extraction rules.
The Mac element-file diagnostic reports the filename with a Mac-helper marker.
It does not include the old file-modification date.

## Checks completed

The actual standalone helper passed 24 tests against temporary files and a local HTTP server.
These tests cover redirects, errors, Unicode, large pages, module-copy bytes, and element-file reads.
We compared 464 existing public signatures across four numeric Mac configurations with the upstream source.
The `0.5-dev` checks permit one exact Mac-only difference: `smfGetYahooHistory.pCols` is `Double`.
The checks still require the upstream Windows declaration and every other argument and default.
We examined active Mac branches for selected Windows dependencies.
The HTML entity data passed independent reference checks.

The original native transport, HTML, JSON, and integration groups returned `6`, `55`, `PASS: 84 JSON checks`, and `5`.
Four target-local probes passed on a full Finviz page.
The direct live extractor also returned the earnings text during that recorded check.
The new recipe and integration groups returned `20` and `25` in native Excel on November 8.
The table-parser and guard groups later returned `40` and `17`.
All ten `SMFTech` indicator cases passed against fixed input data.

Live HTTP requests returned expected markers from Finviz, Yahoo history, Yahoo options, and a Zacks page.
The old FRED data URL failed during those checks.

The recorded Finviz checks returned the earnings text and numeric price through fields 4072 and 4074.
The repaired IBD 2003 call returned `Error` normally because the captured page lacked the requested rank field.
The focused checks do not establish correct results for every provider or recipe.
The earlier fixture groups ran on November 7, in Excel 16.114.4 on Apple Silicon.

The seven additional cached-history checks cover `DateOfHigh`, `SMFHighBetween`, `SMFLowBetween`, `smfLastPrice`, `smfPricesByDates` and two `smfPricesBetween` ranges.
They use five AAPL observations from November 2–6, with explicit dates and expected prices or volumes.
The agent verified those calculations on that capture. Live retrieval and other date ranges still require checks.
The `smfPricesBetween` loop now skips observations after the requested end date before processing older rows.

## Helper installation

`install_helper.command` compiles the helper into this location:

```text
~/Library/Application Scripts/com.microsoft.Excel/SMFMac.scpt
```

The installer preserves an existing helper as a dated backup.
It does not change Excel security settings.

The helper's `stageModule` handler copies reviewed source into this private build cache:

```text
~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/SMF-Mac-Build/v0.2/
```

The builder still requires temporary VBA project access.
The alternative builder and its optional save phase need native verification.

## Use the development add-in

Keep `SMF-Mac.xlam` beside the `smf-elements-*.txt` files.
Register the file through Excel's **Tools > Excel Add-ins > Browse** command.
Turn temporary VBA project access OFF after the approved build.

The installation checks are:

```excel
=RCHGetElementNumber("Version")
=smfStripHTML("<b>42</b>")
=RCHGetElementNumber("AAPL",4072)
=RCHGetElementNumber("AAPL",4074)
```

The current verified local add-in returns `mac.0.5-dev` as part of its version.
A new source build must return its recorded source version.
The HTML check must return `42`.
Fields 4072 and 4074 request the Finviz earnings date and price.
Loading the add-in and retrieving live data are separate checks.

The checked history formulas returned close, volume and date data when `pRows` was omitted or zero.
This close-price example omits the row and column overrides:

```excel
=INDEX(smfGetYahooHistory("AAPL",DATE(2026,11,2),DATE(2026,11,7),"d","C",0),1,1)
```

The `0.4-dev` explicit-row failure and its controls remain in the local [worksheet probe workbook](diagnostics/out/SMF-Mac-Worksheet-Probes-20261108.xlsx).
The corrected `0.5-dev` source also passed this explicit five-row, one-column call in the diagnostic workbook:

```excel
=INDEX(smfGetYahooHistory("AAPL",DATE(2026,11,2),DATE(2026,11,7),"d","C",0,0,5,1),1,1)
```

The [46 native controls](YAHOO-HISTORY-ROWCOUNT-RESULTS-20261108.md) cover worksheet dimensions, invalid-period diagnostics, live prices/volumes, ordering and macro calls.
The packaged add-in also passed eight separate worksheet checks after a clean restart. Other argument combinations still require checks.

If Excel reports a missing add-in, inspect its entry in **Tools > Excel Add-ins**.
The recorded stale entry pointed to `/Applications/Microsoft Excel.app/SMF-Mac.xlam`.
Remove that missing entry with **Delete from list**, then use **Browse** to select the complete absolute path to the repository's `mac/out/SMF-Mac.xlam`.
Enable the add-in and check that `RCHGetElementNumber("Version")` reports `mac.0.5-dev` and that same directory.
Quit Excel through its menu and reopen the workbook to confirm the registration persists.

The last-request diagnostics are `SMFMacHTTPStatus()`, `SMFMacLastError()`, and `SMFMacLastURL()`.
`SMFMacLastFileError()` reports the last element-file read error.
These diagnostics describe session state. They do not give a separate status for each cell.

Run `smfForceRecalculation` to clear SMF's page cache.

## Examples and limits

The upstream examples remain in `docs/samples`.
Some examples contain the old Windows add-in path.
A future verified Mac build must update those links before example checks.

The Mac downloader retrieves HTTP responses. It does not execute website JavaScript.
Source websites can change their pages, require subscriptions, or block requests.
The port does not bypass those restrictions.
The 14 recorded image cases cover selected StockCharts, local-image and numeric-chart inputs, visibility and cleanup.
Additional chart presets, individual menu actions, and other legacy recipes require further native checks.
The fixed-data indicator checks do not establish correct behavior in every workbook.
`smfRSITargetPrices` remains unfixed. It still requests the obsolete Yahoo quote CSV service.
Its history call also uses the default `dohlcvufgxs` item list, although `smfGetYahooHistory` rejects `F` and `G`.

Retain upstream attribution and the dictionary MIT license in `licenses/`.

## Changes that also affect Windows

These changes need Windows checks before merge:

- `JsonConverter.ParseJson` rejects truncated JSON and trailing text after the root value.
- `modJSONExtract.InitScriptEngine` creates Windows Script Control through `CreateObject`.
- `modJSONExtract.GetKeys` and `smfJSONExtractKeys` use `Long` counters.
- `modMenu.smfDeleteFromCellMenu` removes tagged controls in reverse order.
- `Dictionary.cls` uses `Scripting.Dictionary` on Windows through a bundled wrapper.
- `RCHGetElementNumber` sends ADVFN formula definitions through the formula path while retaining the legacy ADVFN handler for other definitions.
- Fields 2727, 849, 19004 and 5196 use the four updated same-provider definitions in `docs/Elements/`.
- `smfGetOptionExpirations` stops adding expiration dates when it reaches the requested row count.
- `smfGetYahooHistory` skips blank CSV lines before reading their dates.
- `smfPricesBetween` skips rows after `pEndDate` and stops before `pBegDate` when reading newest-first history.

A Windows maintainer must compile the project in 32-bit and 64-bit Excel.
The maintainer must check JSON fields and keys, Yahoo examples, and repeated menu startup and removal.
The version formula must still return `3.0.2024.07.30` on Windows.
No Windows runtime check has run for this port.

## Developer checks

```sh
python3 mac/tests/port_static_checks.py
python3 mac/tests/html_static_checks.py
python3 mac/tests/build_noprompt_static.py
python3 mac/tests/transport_tests.py
```

The original fixture groups passed in the saved checkpoint.
The tag, table-cell, guard and revised integration groups also passed in native Excel.
Run these commands in the full-source diagnostic workbook's VBA Immediate window for future changes.
The clean packaging copy excludes diagnostic modules.

```vb
? SMFMacTransportSelfTest()
? smfMacHTMLSelfTest()
? SMFMacJSONSelfTest()
? SMFMacTableRecipeSelfTest()
? SMFMacTableGuardSelfTest()
? SMFMacRecipeSelfTest()
? SMFMacIntegrationSelfTest()
? SMFMacNestedRecipeSelfTest()
? SMFMacHistoryCSVSelfTest()
? SMFMacEmptyRecipeSelfTest()
```

The expected results are `6`, `55`, `PASS: 84 JSON checks`, `40`, `17`, `20`, `25`, `50`, `12`, and `5`.
The last three groups passed in the isolated `0.4-dev` test workbook.
At the `0.4-dev` checkpoint, that add-in's core modules matched the passed workbook.
The saved worksheet's 11 overlapping checks passed again after a clean Excel restart and workbook reopen.
The `0.5-dev` source passed the focused row/column controls, and its packaged add-in passed eight worksheet checks after a clean restart.
The tests do not establish correct results for every provider or recipe.

The [fresh-checkout suite instructions](tests/full_function_suite/README.md) prepare 239 cases from tracked files for 121 public functions.
The preparer, generator and reader use explicit run paths; no ignored run-state files are required.
The reader saves raw `native-results.private.json` locally. The reporter writes redacted `native-results.json`, Markdown and CSV summaries.
The report scripts read saved results and separate HTTP evidence. They do not run Excel or request market data.
The recorded UTC fixtures assume the Mac timezone `America/Los_Angeles`.
Replaying those checks requires the same timezone or explicit changes to the timezone expectations.
Raw diagnostic output, captured pages and disposable test workbooks remain excluded from Git.
The reviewed add-in and examples are tracked separately in `mac/built/0.5-dev` with build information and checksums.
Reports for sharing must redact anonymous Yahoo cookie and crumb values from URLs and response details.
