# SMF Mac test results — November 8, 2026 — original interrupted run

## Later remaining-provider run

The separate [remaining-provider report](REMAINING-PROVIDER-RESULTS-20261108.md) records three passes, eleven failed checks and four image assertion errors among eighteen original cases.
The supplemental local-image control also raised an assertion error.
The image rerun retained `#VALUE!`; reading the comment fill raised error `445` in the result-check phase.
The operator observed no Excel exit during that batch and its image rerun.

## Original interrupted run record

Later the same day, the targeted IBD fix passed native verification and the revised add-in was built and loaded.
All 57 new table-parser and guard checks passed. The existing tag and integration checks also passed.
IBD 2003 returned Error normally from VBA, a worksheet, and the loaded add-in, including after reopening and recalculation.
The provider field remains unavailable. See [the verified crash fix](TABLE-CELL-CRASH-FIX.md).
The counts below preserve the earlier interrupted run. At that checkpoint, the 18 remaining native cases had not been resumed. Later results are linked above.

Excel crashed during case 169, `RCHGetElementNumber("AAPL",2003,"Error")`, the IBD Industry Group Rank recipe. The runner saved `RUNNING` before calling it. The code then sent that recipe through `Evaluate`. The crash's internal cause has not been established. Excel restarted with blank Book1, and the batch stopped. One case was interrupted and 18 remain unrun.

Observed results:

- All 111 public functions received a primary check. 85 passed, 25 failed the expected-value check, and `DateOfHigh` raised a type error. These counts include local tests and offline option-response fixtures.
- All 52 JSON, utility and Mac helper cases passed. All 10 `SMFTech` indicator cases passed.
- Finviz returned AAPL earnings text and price. Yahoo returned live option strikes and a bid of 3.9 for the AAPL November 20, 2026 $350 call through both the option quote function and its wrapper.
- The four-row `smfGetOptionExpirations` call returned raw dates before formatting option symbols. The planned 50-row comparison did not run.
- Yahoo CSV-history calls returned HTTP 401. `RCHGetYahooHistory2` also contains a hardcoded MMM 2017 URL that ignores the requested ticker and dates.
- The text-only comment formula returned `#VALUE!`. The live picture-comment cases did not run.
- The separate HTTP run completed 58 requests across 35 configured data, chart and helper hosts. Some HTTP 200 responses contained unrelated pages or parked domains. The source table preserves those failures.

The working add-in and earlier checkpoints were not changed during that interrupted run. The later fix replaced the add-in with the verified 0.3-dev build and preserved a rollback copy. VBA project access stayed off.

The native evidence records 168 completed cases. 19 planned cases are unfinished or missing. 56 CHECK_FAILED, 111 PASS, 18 PENDING, 1 RUNNING, 1 RUNTIME_ERROR.

The suite includes one primary test for each of 111 public functions: 25 CHECK_FAILED, 85 PASS, 1 RUNTIME_ERROR. Additional cases test indicators, live providers and numbered elements.

A PASS applies to the specific input and check in that row. Offline option responses establish parsing against supplied data. Live extraction rows establish whether that call returned the requested data during this run.

## Results by test mode

| Test mode | Case results |
|---|---|
| Legacy option path | 4 CHECK_FAILED |
| Live HTTP/content | 21 CHECK_FAILED, 9 PASS |
| Live comment picture | 3 PENDING |
| Live data extraction | 30 CHECK_FAILED, 15 PASS, 14 PENDING, 1 RUNNING, 1 RUNTIME_ERROR |
| Local option formatting | 2 PASS |
| Local picture | 1 PENDING |
| Local/core | 1 CHECK_FAILED, 74 PASS |
| Offline provider parsing | 11 PASS |

## Public functions

Each row shows the primary test. The last column reports extra live extraction or picture cases for the same function.

| Function | Primary result | Test mode | Other live cases |
|---|---|---|---|
| `ConvertToIso` | PASS | Local/core |  |
| `ConvertToJson` | PASS | Local/core |  |
| `ConvertToUtc` | PASS | Local/core |  |
| `DateOfHigh` | RUNTIME_ERROR (phase unrecorded) | Live data extraction |  |
| `DecodeJsonString` | PASS | Local/core |  |
| `fnWait` | PASS | Local/core |  |
| `GetKeys` | PASS | Local/core |  |
| `GetObjectProperty` | PASS | Local/core |  |
| `GetProperty` | PASS | Local/core |  |
| `ParseIso` | PASS | Local/core |  |
| `ParseJson` | PASS | Local/core |  |
| `ParseUtc` | PASS | Local/core |  |
| `RCHCreateComment` | CHECK_FAILED | Local/core | 3 PENDING |
| `RCHExtractData` | PASS | Local/core |  |
| `RCHGetElementNumber` | PASS | Live data extraction | 1 CHECK_FAILED, 5 PASS, 12 PENDING, 1 RUNNING |
| `RCHGetHTMLTable` | PASS | Local/core |  |
| `RCHGetTableCell` | PASS | Local/core |  |
| `RCHGetURLData` | PASS | Local/core |  |
| `RCHGetURLData1` | PASS | Local/core |  |
| `RCHGetURLData1Mac` | PASS | Local/core |  |
| `RCHGetURLData2` | PASS | Local/core |  |
| `RCHGetURLData3` | PASS | Local/core |  |
| `RCHGetURLData4` | PASS | Local/core |  |
| `RCHGetWebData` | PASS | Local/core |  |
| `RCHGetYahooHistory` | CHECK_FAILED | Live data extraction |  |
| `RCHGetYahooHistory2` | CHECK_FAILED | Live data extraction |  |
| `RCHGetYahooQuotes` | CHECK_FAILED | Live data extraction |  |
| `RCHImageTable` | PASS | Local/core |  |
| `smfBuildYahooTimeSeriesURL` | PASS | Live data extraction |  |
| `smfCDec` | PASS | Local/core |  |
| `smfChkJSONFieldType` | PASS | Local/core |  |
| `smfConvertData` | PASS | Local/core |  |
| `smfConvertYahooTicker` | PASS | Local/core |  |
| `smfDate2Unix` | PASS | Local/core |  |
| `smfEval` | PASS | Local/core |  |
| `smfEvaluateTwice` | PASS | Local/core |  |
| `smfGet888OptionQuote` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetAConstants` | PASS | Local/core |  |
| `smfGetAData` | PASS | Local/core |  |
| `smfGetADVFNElement` | CHECK_FAILED | Live data extraction |  |
| `smfGetaFiles` | PASS | Local/core |  |
| `smfGetAParms` | PASS | Local/core |  |
| `smfGetBarChartOptionQuote` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetBigChartsOptionQuote` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetCSVFile` | PASS | Local/core |  |
| `smfGetEconData` | CHECK_FAILED | Live data extraction |  |
| `smfGetGoogleHistoryCSV` | CHECK_FAILED | Live data extraction |  |
| `smfGetGoogleOptionQuote` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetGuruFocusCSVItem` | CHECK_FAILED | Live data extraction |  |
| `smfGetGuruFocusData` | PASS | Live data extraction |  |
| `smfGetGuruFocusItem` | CHECK_FAILED | Live data extraction |  |
| `smfGetMorningstarHistoryCSV` | CHECK_FAILED | Live data extraction |  |
| `smfGetMSNHistory` | CHECK_FAILED | Live data extraction |  |
| `smfGetMSNOptionQuote` | CHECK_FAILED | Legacy option path |  |
| `smfGetMWOptionQuote` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetNASDAQOptionQuote` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetOptionExpirations` | PASS | Offline provider parsing | 1 CHECK_FAILED, 1 PENDING |
| `smfGetOptionExpiry` | PASS | Local option formatting |  |
| `smfGetOptionQuotes` | PASS | Local option formatting | 1 PASS |
| `smfGetOptionStrikes` | PASS | Offline provider parsing | 1 PASS |
| `smfGetOX2OptionQuote` | CHECK_FAILED | Legacy option path |  |
| `smfGetOX3OptionQuote` | CHECK_FAILED | Legacy option path |  |
| `smfGetOXOptionQuote` | CHECK_FAILED | Legacy option path |  |
| `smfGetTagContent` | PASS | Local/core |  |
| `smfGetWebPage` | PASS | Local/core |  |
| `smfGetYahooHistory` | CHECK_FAILED | Live data extraction |  |
| `smfGetYahooHistoryCSV` | CHECK_FAILED | Live data extraction |  |
| `smfGetYahooHistoryCSVData` | CHECK_FAILED | Live data extraction |  |
| `smfGetYahooJSONData` | PASS | Live data extraction |  |
| `smfGetYahooJSONField` | PASS | Live data extraction |  |
| `smfGetYahooOptionQuote` | PASS | Offline provider parsing | 1 PASS |
| `smfGetYahooOptionQuote1` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetYahooOptionQuote2` | PASS | Offline provider parsing | 1 CHECK_FAILED |
| `smfGetYahooPortfolioView` | PASS | Live data extraction | 1 PENDING |
| `SMFHighBetween` | CHECK_FAILED | Live data extraction |  |
| `smfHTMLDecode` | PASS | Local/core |  |
| `smfInCellChart` | PASS | Local/core |  |
| `smfJoin` | PASS | Local/core |  |
| `smfJSONExtractField` | PASS | Local/core |  |
| `smfJSONExtractField_x64` | PASS | Local/core |  |
| `smfJSONExtractKeys` | PASS | Local/core |  |
| `smfLastPrice` | CHECK_FAILED | Live data extraction |  |
| `smfLogInternetCalls` | PASS | Local/core |  |
| `SMFLowBetween` | CHECK_FAILED | Live data extraction |  |
| `SMFMacElementsSelfTest` | PASS | Local/core |  |
| `smfMacFetch` | PASS | Local/core |  |
| `smfMacHTMLSelfTest` | PASS | Local/core |  |
| `SMFMacHTTPStatus` | PASS | Local/core |  |
| `SMFMacIntegrationSelfTest` | PASS | Local/core |  |
| `SMFMacJSONSelfTest` | PASS | Local/core |  |
| `SMFMacLastError` | PASS | Local/core |  |
| `SMFMacLastFileError` | PASS | Local/core |  |
| `SMFMacLastURL` | PASS | Local/core |  |
| `smfMacReadElementLocation` | PASS | Local/core |  |
| `smfMacReadElementText` | PASS | Local/core |  |
| `smfMacReadResponse` | PASS | Local/core |  |
| `SMFMacRecipeSelfTest` | PASS | Local/core |  |
| `smfMacStripHTML` | PASS | Local/core |  |
| `SMFMacTransportSelfTest` | PASS | Local/core |  |
| `smfMacTryTagContentRecipe` | PASS | Local/core |  |
| `smfPricesBetween` | CHECK_FAILED | Live data extraction |  |
| `smfPricesByDates` | CHECK_FAILED | Live data extraction |  |
| `smfRSITargetPrices` | CHECK_FAILED | Live data extraction |  |
| `smfStrExtr` | PASS | Local/core |  |
| `smfStripHTML` | PASS | Local/core |  |
| `SMFTech` | PASS | Local/core |  |
| `smfUnix2Date` | PASS | Local/core |  |
| `smfUnix2DateStr` | PASS | Local/core |  |
| `smfWord` | PASS | Local/core |  |
| `smfYahooAPIData` | CHECK_FAILED | Live data extraction |  |
| `smfYahooCodeDesc` | PASS | Live data extraction |  |

## Data sources

The separate network run made 58 requests across 35 provider groups. 11 provider groups met the HTTP/content checks; 24 did not. The native extraction column reports Excel calls separately.

An HTTP 200 response can contain a login page, a parked domain or an unrelated page. HTTP/content checks do not establish that SMF extracted a financial value.

| Source | HTTP | HTTP/content check | Native extraction or picture cases |
|---|---|---|---|
| bigcharts.marketwatch.com | 403 | FAILED | 1 CHECK_FAILED |
| c.stockcharts.com | No response (curl 28) | FAILED | 1 PENDING |
| ca.advfn.com | 200 | FAILED | No extraction case mapped |
| climbermel.github.io | 200 | PASS | No extraction case mapped |
| core-api.barchart.com | No response (curl 28) | FAILED | 1 CHECK_FAILED, 1 PENDING |
| data.moneycentral.msn.com | No response (curl 6) | FAILED | No extraction case mapped |
| download.finance.yahoo.com | No response (curl 6) | FAILED | No extraction case mapped |
| dvd.netflix.com | 200 | FAILED | 1 PENDING |
| fc.yahoo.com | 404 | PASS | No extraction case mapped |
| finance.google.com | 200 | PASS | 1 CHECK_FAILED |
| finance.yahoo.com | 200 | PASS | 2 CHECK_FAILED, 1 PASS |
| finviz.com | 200 | PASS | 1 PASS |
| fred.stlouisfed.org | No response (curl 28) | FAILED | No extraction case mapped |
| investing.money.msn.com | No response (curl 6) | FAILED | 1 PENDING |
| ogres-crypt.com | 404 | FAILED | 1 PENDING |
| oic.ivolatility.com | 404 | FAILED | 1 CHECK_FAILED |
| performance.morningstar.com | No response (curl 28) | FAILED | No extraction case mapped |
| query.yahooapis.com | No response (curl 6) | FAILED | No extraction case mapped |
| query1.finance.yahoo.com | 200 | PASS | 1 CHECK_FAILED, 3 PASS, 2 PENDING |
| query2.finance.yahoo.com | 200 | PASS | 1 PASS |
| quotes.freerealtime.com | 200 | PASS | 1 PASS |
| research.investors.com | 200 | PASS | 1 RUNNING |
| research.tdameritrade.com | No response (curl 6) | FAILED | 1 PENDING |
| stockcharts.com | 200 | PASS | 1 PENDING |
| www.advfn.com | 200 | FAILED | 1 PENDING |
| www.barchart.com | 202 | FAILED | 1 PENDING |
| www.etfinvestmentoutlook.com | 200 | FAILED | 1 PENDING |
| www.etfscreen.com | 200 | FAILED | 1 PENDING |
| www.google.com | 200 | FAILED | 1 CHECK_FAILED |
| www.gurufocus.com | 200 | PASS | 1 PENDING |
| www.marketwatch.com | 401 | FAILED | 1 CHECK_FAILED |
| www.nasdaq.com | 403 | FAILED | 1 CHECK_FAILED |
| www.reuters.com | 401 | FAILED | 1 PENDING |
| www.stockscores.com | 200 | FAILED | 1 PENDING |
| www.zacks.com | 200 | FAILED | 1 PENDING |

## Cases that did not pass

`CHECK_FAILED` means the returned value did not meet the expected check. `CALL_ERROR` means the call raised an exception. `CHECK_ERROR` or `ASSERTION_ERROR` means the assertion could not inspect the result. `SETUP_ERROR` means the test could not prepare its inputs. Older runtime rows may have no recorded phase.

| Case | Result | Actual value or error |
|---|---|---|
| `RCHCreateComment` | CHECK_FAILED | Excel error Error 2015 |
| `smfGetADVFNElement` | CHECK_FAILED | String: Error; length=5 |
| `smfGetEconData` | CHECK_FAILED | String: Error; length=5 |
| `smfGetGuruFocusItem` | CHECK_FAILED | String: Error; length=5 |
| `smfGetGuruFocusCSVItem` | CHECK_FAILED | String: ; length=0 |
| `smfGetMSNHistory` | CHECK_FAILED | Array  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12 |
| `RCHGetYahooHistory` | CHECK_FAILED | Array Date \|  \|  \|  \|  \|  \| AdjClose \|  \|  \|  \|  \|  \|  cells=12 |
| `RCHGetYahooHistory2` | CHECK_FAILED | Array Error -- HTTP 401 \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12 |
| `RCHGetYahooQuotes` | CHECK_FAILED | Array Error -- HTTP 401 \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=42 |
| `smfYahooAPIData` | CHECK_FAILED | String: Error; length=5 |
| `DateOfHigh` | RUNTIME_ERROR (phase unrecorded) | Empty; 13: DateOfHigh: Type mismatch |
| `SMFHighBetween` | CHECK_FAILED | Array 0 \|  \| 0 \| 0 \|  cells=4 |
| `SMFLowBetween` | CHECK_FAILED | Array 999999 \|  \| 0 \| 0 \|  cells=4 |
| `smfLastPrice` | CHECK_FAILED | Integer: 0 |
| `smfGetGoogleHistoryCSV` | CHECK_FAILED | Array Error -- Request failed (3): curl: (3) URL rejected: Malformed input to a URL function \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=36 |
| `smfGetMorningstarHistoryCSV` | CHECK_FAILED | Array Error -- Request failed (28): curl: (28) Failed to connect to performance.morningstar.com port 80 af \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12 |
| `smfGetYahooHistoryCSV` | CHECK_FAILED | Array Error -- HTTP 401 \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=42 |
| `smfGetYahooHistoryCSVData` | CHECK_FAILED | String: Error -- HTTP 401; length=17 |
| `smfGetYahooHistory` | CHECK_FAILED | Array Date \|  \|  \|  \|  \|  \| AdjClose \|  \|  \|  \|  \|  \|  cells=12 |
| `smfPricesBetween` | CHECK_FAILED | Array  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=13 |
| `smfPricesByDates` | CHECK_FAILED | Array [Excel error] \| [Excel error] \|  cells=2 |
| `smfRSITargetPrices` | CHECK_FAILED | Array Error \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12 |
| `smfGetMSNOptionQuote` | CHECK_FAILED | String: Obsolete; length=8 |
| `smfGetOXOptionQuote` | CHECK_FAILED | String: Obsolete -- OptionsXpress is no longer a valid data source; length=58 |
| `smfGetOX2OptionQuote` | CHECK_FAILED | String: Obsolete -- OptionsXpress is no longer a valid data source; length=58 |
| `smfGetOX3OptionQuote` | CHECK_FAILED | String: Obsolete -- OptionsXpress is no longer a valid data source; length=58 |
| `LiveSource_bigcharts_marketwatch_com` | CHECK_FAILED | String: Error -- HTTP 403; length=17 |
| `Live_smfGetBigChartsOptionQuote` | CHECK_FAILED | String: Error; length=5 |
| `LiveSource_core_api_barchart_com` | CHECK_FAILED | String: Error -- Request failed (28): curl: (28) Failed to connect to core-api.barchart.com port 443 after 8005 ms: Timeout was reached / ; length=128 |
| `Live_smfGetBarChartOptionQuote` | CHECK_FAILED | String: ; length=0 |
| `LiveSource_data_moneycentral_msn_com` | CHECK_FAILED | String: Error -- Request failed (6): curl: (6) Could not resolve host: data.moneycentral.msn.com / ; length=89 |
| `LiveSource_download_finance_yahoo_com` | CHECK_FAILED | String: Error -- Request failed (6): curl: (6) Could not resolve host: download.finance.yahoo.com / ; length=90 |
| `LiveSource_dvd_netflix_com` | CHECK_FAILED | String: <!doctype html><html lang="en" class=""><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/><meta charSet="utf-8"/><meta http-equiv="X-UA-Compatible" content="IE=edge"/><meta http-equiv="origin-trial" data-feature="EME Extension - Policy Check" data-expires="2018-11-26" content="Aob+++752GiUzm1RNSIkM9TINnQDxTlxz02v8hFJK/uGO2h... |
| `Live_smfGetYahooOptionQuote1` | CHECK_FAILED | String: Error; length=5 |
| `Live_smfGetYahooOptionQuote2` | CHECK_FAILED | String: Error; length=5 |
| `LiveSource_fred_stlouisfed_org` | CHECK_FAILED | String: Error -- Request failed (28): curl: (28) Operation timed out after 25006 milliseconds with 0 bytes received / ; length=108 |
| `LiveSource_investing_money_msn_com` | CHECK_FAILED | String: Error -- Request failed (6): curl: (6) Could not resolve host: investing.money.msn.com / ; length=87 |
| `LiveSource_oic_ivolatility_com` | CHECK_FAILED | String: Error -- HTTP 404; length=17 |
| `Live_smfGet888OptionQuote` | CHECK_FAILED | String: Error; length=5 |
| `LiveSource_performance_morningstar_com` | CHECK_FAILED | String: Error -- Request failed (28): curl: (28) Failed to connect to performance.morningstar.com port 80 after 8005 ms: Timeout was reached / ; length=133 |
| `LiveSource_query_yahooapis_com` | CHECK_FAILED | String: Error -- Request failed (6): curl: (6) Could not resolve host: query.yahooapis.com / ; length=83 |
| `Live_smfGetOptionExpirations` | CHECK_FAILED | Array 11/9/26 \| 11/11/26 \| 11/13/26 \| 11/20/26 \|  cells=4 |
| `LiveSource_research_tdameritrade_com` | CHECK_FAILED | String: Error -- Request failed (6): curl: (6) Could not resolve host: research.tdameritrade.com / ; length=89 |
| `LiveSource_www_advfn_com` | CHECK_FAILED | String: <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd"> / <html lang="en" xmlns="http://www.w3.org/1999/xhtml"> / <head> / <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /> / <meta http-equiv="X-UA-Compatible" content="IE=edge" /> / <meta name="format-detection" content="telephone=no" /> / <me... |
| `LiveSource_www_barchart_com` | CHECK_FAILED | String: ; length=0 |
| `LiveSource_www_etfinvestmentoutlook_com` | CHECK_FAILED | String: <!doctype html> / <html lang="en-US" prefix="og: https://ogp.me/ns#"> / <head> / 	<meta charset="UTF-8"> / <script> / var gform;gform\|\|(document.addEventListener("gform_main_scripts_loaded",function(){gform.scriptsLoaded=!0}),document.addEventListener("gform/theme/scripts_loaded",function(){gform.themeScriptsLoaded=!0}),window.addEventListener("... |
| `LiveSource_www_etfscreen_com` | CHECK_FAILED | String: <!DOCTYPE html><html lang="en"  data-adblockkey=MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBANnylWw2vLY4hUn9w06zQKbhKBfvjFUCsdFlb6TdQhxb9RXWXuI4t31c+o8fYOv/s8q1LGPga3DE1L/tHU4LENMCAwEAAQ==_Gdh2PARqVtW/WK8warrh1EbbzTiq++73cTKnKcsLSDMUBwnP8Lk+TWBveKaPzyMLxAUzN5NT43l3VVXXC9zfig==><head><meta charset="utf-8"><title>etfscreen.com&nbsp;-&nbsp;etfscreen Resources an... |
| `LiveSource_www_google_com` | CHECK_FAILED | String: <!doctype html><html lang="en-US" dir="ltr"><head><base href="https://www.google.com/finance/"><link rel="preconnect" href="//www.gstatic.com"><meta name="referrer" content="origin"><script nonce="Ke69tuVns9q83NE-t6LVdA">window['ppConfig'] = {productName: 'FinHubUi', deleteIsEnforced:  true , sealIsEnforced:  true , heartbeatRate:  0.5 , periodicR... |
| `Live_smfGetGoogleOptionQuote` | CHECK_FAILED | String: ; length=0 |
| `LiveSource_www_marketwatch_com` | CHECK_FAILED | String: Error -- HTTP 401; length=17 |
| `Live_smfGetMWOptionQuote` | CHECK_FAILED | String: Error; length=5 |
| `LiveSource_www_nasdaq_com` | CHECK_FAILED | String: Error -- HTTP 403; length=17 |
| `Live_smfGetNASDAQOptionQuote` | CHECK_FAILED | String: Error; length=5 |
| `LiveSource_www_reuters_com` | CHECK_FAILED | String: Error -- HTTP 401; length=17 |
| `LiveSource_www_stockscores_com` | CHECK_FAILED | String: <!--include virtual="dbini.asp"--> /  / _x000D_ / <!DOCTYPE html>_x000D_ / <html dir="ltr" lang="en-US">_x000D_ / <head>_x000D_ /     _x000D_ / 	<meta http-equiv="content-type" content="text/html; charset=utf-8" />_x000D_ / 	<meta name="author" content="Stockscores" />_x000D_ / _x000D_ / 	<!-- Stylesheets_x000D_ / 	================================... |
| `LiveSource_www_zacks_com` | CHECK_FAILED | String:  /  /  / <!DOCTYPE html> / <html lang="en"> /  / <!-- Faster page load hack --> / <!--[if IE]><![endif]--> /  / <head> / <!-- Begin: head --> / <!-- z2_head_module_v2 start --> /  / <meta charset="utf-8" /> /  / <style> / *{margin:0 auto;padding:0;outline:0 none}.clearfix{zoom:1}.clearfix:after{content:".";display:block;height:0;clear:both;visibil... |
| `Extract_finance_google_com` | CHECK_FAILED | String: Error; length=5 |

## Unfinished cases

A saved RUNNING row has no completed result. The interrupted row does not count as a pass or a completed test.

| Case | Saved status |
|---|---|
| `Extract_research_investors_com` | INTERRUPTED (saved RUNNING) |
| `Extract_research_tdameritrade_com` | PENDING |
| `Extract_www_barchart_com` | PENDING |
| `Extract_www_etfinvestmentoutlook_com` | PENDING |
| `Extract_www_etfscreen_com` | PENDING |
| `Extract_www_gurufocus_com` | PENDING |
| `Extract_www_reuters_com` | PENDING |
| `Extract_www_stockscores_com` | PENDING |
| `Extract_www_zacks_com` | PENDING |
| `Extract_investing_money_msn_com` | PENDING |
| `Extract_dvd_netflix_com` | PENDING |
| `Extract_core-api_barchart_com` | PENDING |
| `Extract_www_advfn_com` | PENDING |
| `Extract_query1_finance_yahoo_com` | PENDING |
| `ImageComment_c_stockcharts_com` | PENDING |
| `ImageComment_stockcharts_com` | PENDING |
| `ImageComment_ogres-crypt_com` | PENDING |
| `ImageComment_loopback` | PENDING |
| `Live_smfGetOptionExpirations_50Rows` | PENDING |

## Evidence and limits

- `function-results.csv` lists all public functions, the primary results and recorded diagnostics.
- `all-case-results.csv` lists every native row, its call expression, result, error and recorded HTTP fields.
- `datasource-results.csv` lists the separate provider HTTP/content checks and matched native extraction cases.
- `native-results.json` comes from the saved Excel workbook. The HTTP run has separate response bodies, headers, redirect URLs and hashes.
- The last HTTP URL/status can belong to setup, a cached lookup or an earlier case. The report preserves those fields as diagnostics and does not automatically assign a provider failure from them.
- The primary tests and later extra cases ran in batches. Earlier rows can come from an earlier harness revision; saved checkpoints preserve those observations.
- The suite does not directly test every private function, every numbered element or every provider field. A returned price or option quote is a dated observation.

The local evidence files are excluded from Git. Their links resolve in a checkout that retains those files.

- [Excel test workbook](diagnostics/out/SMF-Mac-exhaustive-20261108-mqcw7cog/SMF-Mac-Function-Tests.xlsm)
- [Native result rows](diagnostics/out/SMF-Mac-exhaustive-20261108-mqcw7cog/native-results.json)
- [Executed case specifications](diagnostics/out/SMF-Mac-exhaustive-20261108-mqcw7cog/test-cases.json)
- [Public HTTP results](diagnostics/out/datasource-http-20261108-100559/results.json)
- [Provider assessment](diagnostics/out/datasource-http-20261108-100559/provider-assessment.json)
