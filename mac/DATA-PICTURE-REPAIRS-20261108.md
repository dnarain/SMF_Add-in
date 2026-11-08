# Data and picture repairs — November 8, 2026

These repairs remain experimental in the [personal fork](https://github.com/dnarain/SMF_Add-in).
No pull request is open for this work.
Four same-provider recipes changed after direct requests exposed their current public fields.
`provider_recipe_checks.py` matched eight captured fragments to their expected field values.
The four repaired fields passed live calls in the isolated native Excel workbook.
The [saved 59-case results](DATA-PICTURE-REPAIR-RESULTS-20261108.md) contain 52 passes and seven failed provider data checks.
Seven additional [cached-history cases](tests/full_function_suite/derived_history_cases.json) passed.
The two batches contain 66 cases: 59 PASS and seven provider CHECK_FAILED results.
The `0.4-dev` add-in build remains pending after a file-permission save failure.

## Current verification checkpoint

The development source reports `3.0.2024.07.30-mac.0.4-dev`.
The last verified `0.3-dev` add-in remains preserved at `mac/out/SMF-Mac-0.3-dev-before-data-picture-fix.xlam`.
The agent imported eleven revised core modules into the isolated workbook across the repair and derived-history checks.
The agent has not built or loaded a `0.4-dev` add-in.
Excel returned error `1004` for file permission during the attempted add-in save.
The output at `mac/out/SMF-Mac.xlam` is zero bytes and must not be loaded.
File-access approval remains pending.
The previous add-in and its element files remain in the 15-file `mac/out/rollback-0.3-dev-before-data-picture-20261108/` copy.
The complete 66-case workbook snapshot remains saved separately.
The agent removed five diagnostic modules and eight test worksheets from the packaging copy, leaving one blank `Sheet1`.

The native image probe confirmed that `Shapes.AddPicture` embeds a decoded PNG from Excel's private cache.
The same call succeeded from a worksheet function.
The saved workbook retained the PNG after the helper deleted its temporary file.
A separate worksheet function created a native line chart with three saved data points.
The `RCHCreateComment` checks for StockCharts presets 4 and 2 returned embedded pictures of type `13`.
The numeric-chart check returned a native chart of type `3`.
All 14 selected image and chart cases passed, including local-file, visibility and cleanup checks.

The first native history fixture exposed a decimal-format mismatch: VBA emitted `.25` instead of `0.25`.
The source now adds the leading zero. The corrected module passed its 12-check native self-test.
The agent also corrected numeric-chart extrema for small and large positive values.
Three new cases check tiny positive values, equal tiny values and trillion-scale values.
The prepared focused runner contains 59 cases, including all 18 original pending cases.

The image helper passed 18 focused checks and all 24 existing helper checks.
The provider-fragment and source checks passed.
All 25 regression and offline cases passed in the native workbook, including the 50 nested-recipe and 25 provider-fixture checks.
All 13 selected live data cases passed, including six history and bounded-expiration repair cases.

The live repaired fields returned GuruFocus yield `0.0034`, Zacks target `338.68`, Netflix year `2010` and ADVFN date `2025-09-27`.
The seven unchanged external-provider checks failed to return their required data.
Four local-file image checks first failed because the generated formula used a relative file path.
After the harness path was corrected, those four cases passed on rerun.
The seven failed provider data checks remain failures; no source repair is claimed for them.

## Four recipe changes

The field IDs, provider labels and field meanings remain unchanged:

| Field | Source change | Captured-page expectation |
| --- | --- | --- |
| **2727 — GuruFocus dividend yield** | Read the div marked `(Forward)` on the [current dividend page](https://www.gurufocus.com/stock/AAPL/dividend). Stop before that label and convert the percentage to a fraction. | AAPL `0.0034`, MSFT `0.0078` |
| **849 — Zacks target price consensus** | Read the first data row under `Average Price Target` on the [current price-target page](https://www.zacks.com/stock/research/AAPL/price-target-stock-forecast). | AAPL `338.68`, MSFT `629.46` |
| **19004 — Netflix movie year** | Replace the DVD URL with the [Netflix title page](https://www.netflix.com/title/70111470). Read the first list item inside `data-uia="metadata"`. | *Iron Man 2*: `2010`. [*The Irishman*](https://www.netflix.com/title/80175798): `2019` |
| **5196 — ADVFN annual ending date** | Follow the [financials redirect](https://www.advfn.com/p.php?pid=financials&symbol=NASDAQ:AAPL), read `Year End`, remove that prefix and convert the date. | AAPL `2025-09-27`, MSFT `2026-06-30` |

The MSFT dividend capture has different trailing and forward yields. The fixture expects the forward value.
The Netflix checks use movie IDs. They do not establish an initial release year for a television series.
The ADVFN date is the latest annual date displayed in the captured provider page.
These values describe the November 8 captures. The agent must compare each live result with the corresponding provider response.

The definitions changed in these files:

- [smf-elements-9.txt](../docs/Elements/smf-elements-9.txt): field 2727
- [smf-elements-6.txt](../docs/Elements/smf-elements-6.txt): field 849
- [smf-elements-20.txt](../docs/Elements/smf-elements-20.txt): field 19004
- [smf-elements-7.txt](../docs/Elements/smf-elements-7.txt): field 5196.

The ADVFN handler skips formula definitions while preserving the `AdvFN-A` and `AdvFN-Q` labels.
The four shared definition changes also affect Windows. Windows runtime checks remain pending.

## Seven fields still blocked or unavailable

The agent retained these recipes without substituting another provider:

| Field | Exact source observation | Remaining limit |
| --- | --- | --- |
| **19902 — TD Ameritrade ex-dividend date estimate** | The [legacy research host](http://research.tdameritrade.com/public/stocks/overview/overview.asp?symbol=AAPL) failed DNS resolution over HTTP and HTTPS. | No replacement TD research field was established. |
| **702 — Barchart overall opinion** | The [opinion URL](https://www.barchart.com/stocks/quotes/AAPL/opinion) returned HTTP `202`, an empty body and `x-amzn-waf-action: challenge`. | The request did not expose the opinion field. |
| **995 — ETFInvestmentOutlook new highs** | The [configured URL](http://www.etfinvestmentoutlook.com/rank_breadth_new_highs_lows_avg_d.php) redirected to [Avenues Financial](https://avenuesfinancial.com/), an accounting-services website. | The response has no ETF breadth table. |
| **1084 — ETFScreen relative strength** | The [configured ETF table](https://www.etfscreen.com/etfperf.php?wl=0&s=Rtn-1d%7Cdesc&t=6&d=i&ftS=yes&ftL=yes&vFf=dolVol21&vFl=gt&vFv=0&udc=default&d=i) returned a domain-parking page. | The response has no relative-strength table. |
| **13626 — Reuters trailing P/E** | Both the [legacy financial-highlights URL](https://www.reuters.com/finance/stocks/financial-highlights/AAPL.O) and [current metrics URL](https://www.reuters.com/markets/companies/AAPL.O/key-metrics/) returned HTTP `401` with DataDome protection. | The requests did not expose the P/E field. |
| **20000 — MSN StockScouter rating** | The [legacy MSN host](http://investing.money.msn.com/investments/stock-price?Symbol=AAPL) failed DNS resolution. | No current same-provider StockScouter field was established. |
| **1116 — Barchart NYSE new highs** | The [configured momentum API](https://core-api.barchart.com/v1/momentum/get?raw=1&fields=exchange,high52w,low52w,newHighs,newLows,advancingVolume,unchangedVolume,decliningVolume,advancingIssues,percentAdvancingIssues,unchangedIssues,percentUnchangedIssues,decliningIssues,percentDecliningIssues&exchanges=NASDAQ,NYSE,OTC,AMEX) timed out while connecting to port `443`. | No response was available to verify its field path. A prior independent request also timed out. |

Schwab reported completion of the Ameritrade client conversion in [May 2024](https://pressroom.aboutschwab.com/press-releases/press-release/2024/Schwab-Reports-Monthly-Activity-Highlights-73055b4cb/default.aspx).
That corporate transition does not identify a replacement for field 19902.
The agent did not bypass access challenges or add account credentials.
The Barchart API timeout does not establish that its data field is retired.

## Mac execution and history changes

The bounded [nested dispatcher](../code/modMacCode.bas) accepts selected literal calls without `Evaluate`.
Its leaves are `smfGetTagContent` and `smfGetYahooJSONField`.
Its wrappers are `DateValue`, `smfConvertData` and `smfStrExtr`.
The dispatcher validates the complete syntax before executing a provider call.
The 50 nested-recipe checks passed in the isolated native workbook.

The [Yahoo history adapter](../code/smfGetYahooHistoryCSV_.bas) reads Yahoo chart JSON and produces the CSV expected by existing history functions.
It retains the requested ticker, date bounds, interval and history event.
A direct [Yahoo chart request](https://query1.finance.yahoo.com/v8/finance/chart/AAPL?period1=1793577600&period2=1794009600&interval=1d&events=div%2Csplits&includeAdjustedClose=true) returned HTTP `200` for the November 2–6 period.
`SMFMacHistoryCSVSelfTest` compares dates, prices, adjusted close, volume, dividends, splits, nulls, ordering and invalid inputs.
The 12 history-conversion checks passed in the isolated native workbook.

Seven additional native checks used the cached November 2–6 AAPL capture without HTTP requests:

- `DateOfHigh` returned the highest-high date, November 5.
- `SMFHighBetween` and `SMFLowBetween` returned the expected extreme, its date, range-start open and range-end close.
- `smfLastPrice` returned the November 5 adjusted close while skipping the newer observation.
- `smfPricesByDates` returned the expected November 2 and November 6 closes.
- `smfPricesBetween` returned the expected November 3–6 and November 3–5 summaries.

`smfPricesBetween` returned volumes `273152900` and `228476600` for the two ranges; November 2 supplied the previous close.
The earlier-ending range first exposed a loop error in [smfPricesBetween_.bas](../code/smfPricesBetween_.bas).
The loop now skips rows after `pEndDate` and stops before `pBegDate` while reading newest-first history.
The two-line fix also affects Windows. Windows runtime checks remain pending.
These seven cases cover the captured observations and explicit inputs. They do not verify every date range or live request.

`smfRSITargetPrices` remains unfixed. It still requests the obsolete Yahoo quote CSV service.
Its history call inherits `pItems="dohlcvufgxs"`, which includes the unsupported `F` and `G` items.

## Picture findings

The operator recorded these earlier native Excel probes before adding the private image-cache path:

- Comment `Fill`, `Shadow` and `Line` operations returned error `445`.
- A standalone `Shapes.AddPicture` probe with a one-pixel image returned error `1004`.
- `Pictures.Insert` worked outside the worksheet UDF. That probe does not verify the worksheet formula path.
- A worksheet UDF created a `ChartObject` successfully.

The current Mac path uses owned worksheet objects for charts and pictures.
It avoids the unsupported comment-fill operations and retains a cell note for text and diagnostics.
The numeric-chart path uses `ChartObjects` instead of the external numeric-chart service.
The completed 14-case image group checks selected StockCharts, local-image and numeric-chart calls, cleanup and visibility.
Individual menu-click checks and other chart presets still require verification.
The standalone probes do not establish that every `RCHCreateComment` mode works.

## Native verification checkpoint

| Test group | Expected checks | Current status |
| --- | ---: | --- |
| `SMFMacNestedRecipeSelfTest()` | 50 | Passed in the isolated native workbook |
| `SMFMacHistoryCSVSelfTest()` | 12 | Passed in the isolated native workbook |
| `SMFMacProviderRecipeSelfTest()` | 25 | Passed in the isolated native workbook |
| [Focused repair cases](tests/full_function_suite/repair_cases.json) | 23 cases | Passed in the isolated native workbook |
| Full selected repair run | 59 cases | Complete: 52 PASS, 7 CHECK_FAILED |
| [Additional cached-history cases](tests/full_function_suite/derived_history_cases.json) | 7 cases | Complete: 7 PASS |
| Combined two batches | 66 cases | Complete: 59 PASS, 7 CHECK_FAILED |

These counts describe different test sets. The 23 repair cases include the nested and history self-tests and six live calls.
The seven cached-history cases are separate from the 59-case report.
`SMFMacProviderRecipeSelfTest` checks eight captures through direct and numbered calls, eight missing-field errors and unchanged HTTP state.
It belongs in an isolated diagnostic workbook, not the production add-in.

## Evidence

The public fragments contain field data and markup without scripts, cookies or credentials:

- [Eight captured fragments and expected values](tests/provider_recipe_fixtures/manifest.json)
- [Independent fragment checks](tests/provider_recipe_checks.py)
- [Cached native fixture module](tests/provider_recipe_tests.bas)
- [Seven derived-history case specifications](tests/full_function_suite/derived_history_cases.json).

The full captures remain local and excluded from Git. These relative links resolve in a checkout that retains the evidence:

- [Provider observations and seven blockers](diagnostics/out/SMF-Mac-provider-repair-20261108-pm2746de/provider-assessment.json)
- [Initial public requests](diagnostics/out/SMF-Mac-provider-repair-20261108-pm2746de/results.json)
- [Additional provider samples](diagnostics/out/SMF-Mac-provider-repair-20261108-pm2746de/additional-samples.json)
- [Second movie sample](diagnostics/out/SMF-Mac-provider-repair-20261108-pm2746de/additional-movie-sample.json)
- [Exact recipe changes](diagnostics/out/SMF-Mac-provider-repair-20261108-pm2746de/recipe-changes.json).
- [Seven saved derived-history results](diagnostics/out/SMF-Mac-derived-history-20261108-4afd47e4/native-results.json).
