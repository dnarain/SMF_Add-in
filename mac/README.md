# SMF native Mac development add-in

This Mac port is experimental work in the [dnarain/SMF_Add-in fork](https://github.com/dnarain/SMF_Add-in).
The working branch is `dnarain/mac-native-support`.
There is no pull request for this Mac port. Further workbook and provider checks come first.

The current source version is `3.0.2024.07.30-mac.0.3-dev`.
It adds direct table-cell recipe calls after an IBD lookup crashed Excel during the broader provider tests.
All 57 new parser and guard checks passed in native Excel.
The repaired IBD lookup returned normally after a page-cache clear and after reopening the workbook.
The direct dispatchers handle 429 of 1,815 formula recipes: 291 tag recipes and 138 table-cell recipes.
The table guard returns `pError` for another 35 compound table-cell recipes.
The evaluator paths for other recipes still require native checks.
See [the crash fix](TABLE-CELL-CRASH-FIX.md), [the recorded tests](TEST-RESULTS-20261108.md), and [the testing plan](TESTING-PLAN.md).

The separate [remaining-provider checks](REMAINING-PROVIDER-RESULTS-20261108.md) recorded three passes, eleven failed data checks and four image assertion errors.
The supplemental local-image case also raised an assertion error.
The image cases returned `#VALUE!`; reading the comment fill raised error `445`.
The operator observed no Excel exit during that batch and the image rerun.

Excel built and loaded the revised local `mac/out/SMF-Mac.xlam` during the recorded checks.
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
The direct dispatchers handle 429 literal recipes without `Evaluate`.
The guard blocks 35 compound table-cell recipes and returns their `pError` value.
Other expressions retain the previous evaluator path and require separate native checks.
The 40 table-parser checks, 17 guard checks, 20 tag checks, and 25 integration checks passed natively.
Save and load an `.xlam` only after the workbook checks pass.
The optional builder save path is `mac/out/SMF-Mac.xlam`.
The builder does not create the output folder or copy element files.

## Source changes

- The Mac downloader calls `AppleScriptTask` and macOS `curl`.
- Yahoo requests obtain an anonymous cookie and crumb.
- The Mac HTML parser uses VBA and does not require Windows `HTMLDocument`.
- The Mac JSON parser uses the bundled `JsonConverter` and `Dictionary` class.
- Mac compiler branches exclude Windows networking and Script Control code.
- Decimal conversion preserves more than four decimal places.
- Failed HTTP requests do not enter the successful-page cache.
- Mac field-definition reads use the AppleScript helper.
- Supported literal `smfGetTagContent` and `RCHGetTableCell` recipes call VBA directly.
- Unsupported compound table-cell recipes return `pError` before the evaluator runs.
- The Mac menu uses ten flat buttons because popup access failed natively.

Existing function names and numbered fields remain.
The field recipes retain upstream URLs and extraction rules.
The Mac element-file diagnostic reports the filename with a Mac-helper marker.
It does not include the old file-modification date.

## Checks completed

The actual standalone helper passed 24 tests against temporary files and a local HTTP server.
These tests cover redirects, errors, Unicode, large pages, module-copy bytes, and element-file reads.
We compared 464 existing public signatures across four numeric Mac configurations with the upstream source.
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

The version must include `mac.0.3-dev`.
The HTML check must return `42`.
Fields 4072 and 4074 request the Finviz earnings date and price.
Loading the add-in and retrieving live data are separate checks.

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
Charts, pictures, individual menu actions, and other legacy recipes require further native checks.
The fixed-data indicator checks do not establish correct behavior in every workbook.

Retain upstream attribution and the dictionary MIT license in `licenses/`.

## Changes that also affect Windows

These changes need Windows checks before merge:

- `JsonConverter.ParseJson` rejects truncated JSON and trailing text after the root value.
- `modJSONExtract.InitScriptEngine` creates Windows Script Control through `CreateObject`.
- `modJSONExtract.GetKeys` and `smfJSONExtractKeys` use `Long` counters.
- `modMenu.smfDeleteFromCellMenu` removes tagged controls in reverse order.
- `Dictionary.cls` uses `Scripting.Dictionary` on Windows through a bundled wrapper.

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
Run these commands in the VBA Immediate window for future changes:

```vb
? SMFMacTransportSelfTest()
? smfMacHTMLSelfTest()
? SMFMacJSONSelfTest()
? SMFMacTableRecipeSelfTest()
? SMFMacTableGuardSelfTest()
? SMFMacRecipeSelfTest()
? SMFMacIntegrationSelfTest()
```

The expected results are `6`, `55`, `PASS: 84 JSON checks`, `40`, `17`, `20`, and `25`.
The tests do not establish correct results for every provider or recipe.

The full-function developer scripts require local run metadata, retained evidence and explicit file paths.
`build_suite.py` and `read_results.py` use ignored run-state files unless the selected route supplies paths.
The report scripts read saved results and separate HTTP evidence. They do not run Excel or request market data.
The recorded UTC fixtures assume the Mac timezone `America/Los_Angeles`.
Replaying those checks requires the same timezone or explicit changes to the timezone expectations.
Raw diagnostic output, captured pages, generated workbooks and add-ins remain excluded from Git.
Reports for sharing must redact anonymous Yahoo cookie and crumb values from URLs and response details.
