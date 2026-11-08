# IBD table-cell crash fix

The source fix prevents numbered table-cell recipes from entering Excel's `Evaluate` function on Mac.
The revised code passed native Excel checks on November 8, 2026.
The repaired IBD call returned normally from VBA, a worksheet, and the loaded add-in.
It also returned normally after clearing the page cache and after reopening the workbook.
Excel loaded the revised `mac/out/SMF-Mac.xlam`, version `3.0.2024.07.30-mac.0.3-dev`.

## Evidence

Excel terminated during test case 169 on November 8, 2026.
The saved case was `RCHGetElementNumber("AAPL",2003,"Error")`.
The test runner saved `RUNNING` before the call and did not save a result afterward.

Element 2003 in `docs/Elements/smf-elements-9.txt` contains:

```vb
=RCHGetTableCell("https://research.investors.com/stock-checkup/xxxx-~~~~~.aspx",1,">Industry Group Rank")
```

The previous direct dispatcher only accepted complete `smfGetTagContent(...)` calls.
`RCHGetElementNumber` therefore sent this table-cell recipe through `Evaluate`.
A direct `RCHGetTableCell` call against controlled HTML passed earlier.
Earlier Finviz tests also ended in restarts before direct tag dispatch was added.

These observations identify the failing test and evaluator path.
They do not prove Excel's internal crash cause.
macOS denied access to the Microsoft crash-report folders; no readable crash stack was available.

The separate IBD page request returned HTTP 200 and redirected to an AAPL page.
Its captured text lacked `Industry Group Rank`.
The repaired call can therefore return `Error` even when Excel stays running.

## Code changes

- `code/modMacCode.bas`: `smfMacTryTableCellRecipe` parses complete literal `RCHGetTableCell(...)` calls and calls VBA directly.
- The dispatcher preserves all 11 arguments, defaults, omitted values, quoted commas, doubled quotes and typed error values.
- The dispatcher accepts valid literal syntax before converting numeric values. Numeric overflow returns `pError` and cannot trigger evaluator fallback.
- `code/modGetElementNumber.bas`: both Mac formula branches try the tag and table dispatchers before evaluating other expressions.
- `MacRecipeCallsTableCell` detects remaining table-cell calls outside quoted strings. Unsupported table expressions return the outer `pError` immediately.
- `code/modMacIntegrationTests.bas`: `SMFMacTableRecipeSelfTest` checks 40 accepted, rejected and overflow cases with cached HTML.
- `code/modGetElementNumber.bas`: `SMFMacTableGuardSelfTest` checks 17 scanner and numbered-recipe cases without HTTP requests.
- The Mac development version becomes `3.0.2024.07.30-mac.0.3-dev`.

The direct dispatchers cover 429 of 1,815 effective formula recipes: 291 tag recipes and 138 table-cell recipes.
The table guard blocks another 35 compound recipes until direct support is added.
Other expressions keep their previous behavior. Their evaluator safety has not been established for every recipe.
Windows formula branches remain unchanged.

## Verification

| Native check | Result |
| --- | --- |
| `SMFMacTableRecipeSelfTest()` | 40 passed |
| `SMFMacTableGuardSelfTest()` | 17 passed |
| `SMFMacRecipeSelfTest()` | 20 passed |
| `SMFMacIntegrationSelfTest()` | 25 passed, including the existing tag fixtures |
| IBD 2003 from VBA and worksheet | Returned the string `Error`; Excel stayed running |
| Direct IBD table-cell comparison | Same `Error` result; HTTP 200; no HTTP error |
| Worksheet after page-cache clear | Same results; Excel stayed running |
| Worksheet after reopening and page-cache clear | Same results; Excel stayed running |
| Loaded add-in from a separate workbook | Version `0.3-dev`, IBD `Error`, Finviz earnings and price returned |

Finviz fields 4072 and 4074 returned `Nov 02 AMC` and `342.37` during this run.
These are dated observations from the provider.
The IBD error remains because the captured page lacks the requested rank field.
The focused checks verify the repaired call; they do not establish the internal Excel crash cause.

Source checks preserve the 464 existing upstream public signatures across four Mac conditional configurations.
Static checks do not establish native compilation or execution.

Run these commands separately after importing the three revised modules into an isolated full-source workbook:

```vb
? SMFMacTableRecipeSelfTest()
? SMFMacTableGuardSelfTest()
? SMFMacRecipeSelfTest()
? SMFMacIntegrationSelfTest()
```

Expected results are `40`, `17`, `20` and `25`.
All four groups must finish before another live lookup.

Then run the repaired call:

```vb
? RCHGetElementNumber("AAPL",2003,"Error")
```

Compare it with the direct call:

```vb
? RCHGetTableCell("https://research.investors.com/stock-checkup/xxxx-AAPL.aspx",1,">Industry Group Rank")
```

Both calls must return normally. A missing provider field must remain an error value.
Repeat the numbered lookup from a worksheet, after clearing the SMF page cache, and after reopening the workbook.
Check Finviz fields 4072 and 4074 again before creating a revised add-in.

Keep VBA project access off. Manual editor imports do not require it.
Do not rerun the full provider batch during this focused verification.

## Saved evidence

The original interrupted workbook is preserved at:

`mac/diagnostics/out/SMF-Mac-exhaustive-20261108-mqcw7cog/SMF-Mac-Function-Tests-crash-IBD2003.xlsm`

The current isolated fix workbook and replacement modules are listed in:

`mac/diagnostics/out/tablecell-fix-current.json`

The verified full-source checkpoint is `mac/diagnostics/out/SMF-Mac-TableCell-Fix-Passed-20261108.xlsm`.
The separate verification workbook is `mac/out/SMF-Mac-TableCell-Checks-20261108.xlsx`.
The previous add-in is preserved as `mac/out/SMF-Mac-0.2-dev-before-tablecell-fix.xlam`.
No VBA project security setting changed during these imports or checks.

The revised add-in has 48 standard modules, one Dictionary class and two document modules.
All 51 saved VBA source streams exactly match the passed full-source workbook.
All 14 adjacent definition files match the source files.

The revised add-in SHA256 is:

```text
0950924e9856f6f1b6a7361adb55f3c166d5ac2e9a55b13102a89aee630fe2b8
```

The focused checks did not resume the 18 remaining native cases from the earlier provider run.
