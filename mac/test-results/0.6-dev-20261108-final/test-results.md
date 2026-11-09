# SMF Mac native suite results

Run `smf-0.6-full-final-20261108` (saved state `DONE`) records 292 planned cases across 121 public functions and 35 inventoried source groups.

Raw results: 54 CHECK_FAILED, 238 PASS. Interpreted results: 1 CODE_FAILURE, 4 OBSOLETE, 238 PASS, 29 PROVIDER_UNAVAILABLE, 20 UNCLASSIFIED_FAIL.

Expected version: `3.0.2024.07.30-mac.0.6-dev`. Observed version: `Stock Market Functions add-in, Version 3.0.2024.07.30-mac.0.6-dev ([final-run]; Macintosh (ARM) Version 27.0.1 (Build 26A434); 16.114; ; Local; 1)`. Version verified: True.

Workbook SHA256: `14ed9c3af4e9733deefbde7fac3129dcea888006cd8e47a189b763289132e5a6`. Saved VBA SHA256: `31066426d87e287f55c29c3df295f4061b47c7cd295bc48da229e3b86e67bec9`.

A PASS applies to the recorded input/assertion. Many live assertions require only a nonempty or positive result. They do not independently verify the provider's field accuracy, meaning, or freshness. Endpoint-content PASS, offline fixture parsing, error responses, and obsolete responses do not establish a working financial feed.

48 provider-data/image checks passed, including 2 image checks and any cached values returned by the selected calls. This is not a count of fresh provider requests.

This run uses a fresh checkout on the same Mac. The saved VBA hash identifies the owner workbook, including its imported runner; it does not claim that the packaged add-in binary was tested without modification or that a fresh macOS account was tested.

Case specification SHA256: `75196b800984e07692881ed64dff7d3bd498e7bb9d7dbcc6fab6e6dd1c53086f`. Saved workbook case hash verified: True. Recorded source fingerprint matches the manifest: True.

## Public functions

The primary row uses explicit case metadata, then a matching case ID, then the first selected direct call to that function. The CSV records the selection reason and every case; primary status alone does not summarize all behavior.

| Function | Primary raw | Primary interpretation | All cases |
|---|---|---|---|
| `ConvertToIso` | PASS | PASS | 1 PASS |
| `ConvertToJson` | PASS | PASS | 1 PASS |
| `ConvertToUtc` | PASS | PASS | 1 PASS |
| `DateOfHigh` | PASS | PASS | 2 PASS |
| `DecodeJsonString` | PASS | PASS | 1 PASS |
| `fnWait` | PASS | PASS | 1 PASS |
| `GetKeys` | PASS | PASS | 1 PASS |
| `GetObjectProperty` | PASS | PASS | 1 PASS |
| `GetProperty` | PASS | PASS | 1 PASS |
| `ParseIso` | PASS | PASS | 1 PASS |
| `ParseJson` | PASS | PASS | 1 PASS |
| `ParseUtc` | PASS | PASS | 1 PASS |
| `RCHCreateComment` | PASS | PASS | 15 PASS |
| `RCHExtractData` | PASS | PASS | 1 PASS |
| `RCHGetElementNumber` | PASS | PASS | 11 PASS, 4 PROVIDER_UNAVAILABLE, 5 UNCLASSIFIED_FAIL |
| `RCHGetHTMLTable` | PASS | PASS | 1 PASS |
| `RCHGetTableCell` | PASS | PASS | 1 PASS |
| `RCHGetURLData` | PASS | PASS | 1 PASS |
| `RCHGetURLData1` | PASS | PASS | 1 PASS |
| `RCHGetURLData1Mac` | PASS | PASS | 1 PASS |
| `RCHGetURLData2` | PASS | PASS | 1 PASS |
| `RCHGetURLData3` | PASS | PASS | 1 PASS |
| `RCHGetURLData4` | PASS | PASS | 1 PASS |
| `RCHGetWebData` | PASS | PASS | 1 PASS |
| `RCHGetYahooHistory` | PASS | PASS | 2 PASS |
| `RCHGetYahooHistory2` | PASS | PASS | 6 PASS |
| `RCHGetYahooQuotes` | PASS | PASS | 2 PASS |
| `RCHImageTable` | PASS | PASS | 1 PASS |
| `smfBuildYahooTimeSeriesURL` | PASS | PASS | 1 PASS |
| `smfCDec` | PASS | PASS | 1 PASS |
| `smfChkJSONFieldType` | PASS | PASS | 1 PASS |
| `smfConvertData` | PASS | PASS | 1 PASS |
| `smfConvertYahooTicker` | PASS | PASS | 1 PASS |
| `smfDate2Unix` | PASS | PASS | 1 PASS |
| `smfEval` | PASS | PASS | 1 PASS |
| `smfEvaluateTwice` | PASS | PASS | 1 PASS |
| `smfGet888OptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetAConstants` | PASS | PASS | 1 PASS |
| `smfGetAData` | PASS | PASS | 1 PASS |
| `smfGetADVFNElement` | CHECK_FAILED | UNCLASSIFIED_FAIL | 1 UNCLASSIFIED_FAIL |
| `smfGetaFiles` | PASS | PASS | 1 PASS |
| `smfGetAParms` | PASS | PASS | 1 PASS |
| `smfGetBarChartOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetBigChartsOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetCSVFile` | PASS | PASS | 1 PASS |
| `smfGetEconData` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfGetGoogleHistoryCSV` | CHECK_FAILED | CODE_FAILURE | 1 CODE_FAILURE |
| `smfGetGoogleOptionQuote` | PASS | PASS | 1 PASS, 1 UNCLASSIFIED_FAIL |
| `smfGetGuruFocusCSVItem` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfGetGuruFocusData` | PASS | PASS | 1 PASS |
| `smfGetGuruFocusItem` | CHECK_FAILED | UNCLASSIFIED_FAIL | 1 UNCLASSIFIED_FAIL |
| `smfGetMorningstarHistoryCSV` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfGetMSNHistory` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfGetMSNOptionQuote` | CHECK_FAILED | OBSOLETE | 1 OBSOLETE |
| `smfGetMWOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetNASDAQOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetOptionExpirations` | PASS | PASS | 4 PASS |
| `smfGetOptionExpiry` | PASS | PASS | 1 PASS |
| `smfGetOptionQuotes` | PASS | PASS | 1 PASS, 1 UNCLASSIFIED_FAIL |
| `smfGetOptionStrikes` | PASS | PASS | 2 PASS |
| `smfGetOX2OptionQuote` | CHECK_FAILED | OBSOLETE | 1 OBSOLETE |
| `smfGetOX3OptionQuote` | CHECK_FAILED | OBSOLETE | 1 OBSOLETE |
| `smfGetOXOptionQuote` | CHECK_FAILED | OBSOLETE | 1 OBSOLETE |
| `smfGetTagContent` | PASS | PASS | 1 PASS |
| `smfGetWebPage` | PASS | PASS | 1 PASS |
| `smfGetYahooHistory` | PASS | PASS | 33 PASS |
| `smfGetYahooHistoryCSV` | PASS | PASS | 2 PASS |
| `smfGetYahooHistoryCSVData` | PASS | PASS | 2 PASS |
| `smfGetYahooJSONData` | PASS | PASS | 1 PASS |
| `smfGetYahooJSONField` | PASS | PASS | 1 PASS |
| `smfGetYahooOptionQuote` | PASS | PASS | 1 PASS, 1 UNCLASSIFIED_FAIL |
| `smfGetYahooOptionQuote1` | PASS | PASS | 34 PASS, 1 UNCLASSIFIED_FAIL |
| `smfGetYahooOptionQuote2` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetYahooPortfolioView` | PASS | PASS | 2 PASS |
| `SMFHighBetween` | PASS | PASS | 2 PASS |
| `smfHTMLDecode` | PASS | PASS | 1 PASS |
| `smfInCellChart` | PASS | PASS | 1 PASS |
| `smfJoin` | PASS | PASS | 1 PASS |
| `smfJSONExtractField` | PASS | PASS | 1 PASS |
| `smfJSONExtractField_x64` | PASS | PASS | 1 PASS |
| `smfJSONExtractKeys` | PASS | PASS | 1 PASS |
| `smfLastPrice` | PASS | PASS | 2 PASS |
| `smfLogInternetCalls` | PASS | PASS | 1 PASS |
| `SMFLowBetween` | PASS | PASS | 2 PASS |
| `SMFMacElementsSelfTest` | PASS | PASS | 1 PASS |
| `SMFMacEmptyRecipeSelfTest` | PASS | PASS | 1 PASS |
| `smfMacFetch` | PASS | PASS | 9 PASS, 14 PROVIDER_UNAVAILABLE, 8 UNCLASSIFIED_FAIL |
| `SMFMacHistoryCSVSelfTest` | PASS | PASS | 1 PASS |
| `smfMacHTMLSelfTest` | PASS | PASS | 1 PASS |
| `SMFMacHTTPStatus` | PASS | PASS | 1 PASS |
| `SMFMacIntegrationSelfTest` | PASS | PASS | 1 PASS |
| `SMFMacJSONSelfTest` | PASS | PASS | 1 PASS |
| `SMFMacLastError` | PASS | PASS | 1 PASS |
| `SMFMacLastFileError` | PASS | PASS | 1 PASS |
| `SMFMacLastURL` | PASS | PASS | 1 PASS |
| `SMFMacNestedRecipeSelfTest` | PASS | PASS | 1 PASS |
| `smfMacReadElementLocation` | PASS | PASS | 1 PASS |
| `smfMacReadElementText` | PASS | PASS | 1 PASS |
| `smfMacReadResponse` | PASS | PASS | 1 PASS |
| `SMFMacRecipeSelfTest` | PASS | PASS | 1 PASS |
| `smfMacStripHTML` | PASS | PASS | 1 PASS |
| `SMFMacTableGuardSelfTest` | PASS | PASS | 1 PASS |
| `SMFMacTableRecipeSelfTest` | PASS | PASS | 1 PASS |
| `SMFMacTransportSelfTest` | PASS | PASS | 1 PASS |
| `smfMacTryNestedRecipe` | PASS | PASS | 2 PASS |
| `smfMacTryTableCellRecipe` | PASS | PASS | 1 PASS |
| `smfMacTryTagContentRecipe` | PASS | PASS | 1 PASS |
| `smfMacYahooChartCSV` | PASS | PASS | 1 PASS |
| `smfMacYahooDownloadCSV` | PASS | PASS | 2 PASS |
| `smfMacYahooDownloadURL` | PASS | PASS | 2 PASS |
| `smfPricesBetween` | PASS | PASS | 3 PASS |
| `smfPricesByDates` | PASS | PASS | 2 PASS |
| `smfRSITargetPrices` | CHECK_FAILED | UNCLASSIFIED_FAIL | 1 UNCLASSIFIED_FAIL |
| `smfStrExtr` | PASS | PASS | 1 PASS |
| `smfStripHTML` | PASS | PASS | 1 PASS |
| `SMFTech` | PASS | PASS | 10 PASS |
| `smfUnix2Date` | PASS | PASS | 1 PASS |
| `smfUnix2DateStr` | PASS | PASS | 1 PASS |
| `smfWord` | PASS | PASS | 1 PASS |
| `smfYahooAPIData` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfYahooCodeDesc` | PASS | PASS | 1 PASS |

## Source groups

| Source | Raw cases | Interpreted cases | Endpoint-content passes | Provider-data/image checks passed |
|---|---|---|---|---|
| bigcharts.marketwatch.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| c.stockcharts.com | No cases | No cases | 0 | 0 |
| climbermel.github.io | 1 PASS | 1 PASS | 1 | 0 |
| core-api.barchart.com | 3 CHECK_FAILED | 3 PROVIDER_UNAVAILABLE | 0 | 0 |
| data.moneycentral.msn.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| download.finance.yahoo.com | 1 CHECK_FAILED | 1 PROVIDER_UNAVAILABLE | 0 | 0 |
| dvd.netflix.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 0 | 1 |
| finance.google.com | 2 CHECK_FAILED, 1 PASS | 1 CODE_FAILURE, 1 PASS, 1 UNCLASSIFIED_FAIL | 1 | 0 |
| finance.yahoo.com | 2 CHECK_FAILED, 1 PASS | 1 PASS, 2 PROVIDER_UNAVAILABLE | 0 | 1 |
| finviz.com | 3 PASS | 3 PASS | 1 | 2 |
| fred.stlouisfed.org | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| investing.money.msn.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| ogres-crypt.com | No cases | No cases | 0 | 0 |
| oic.ivolatility.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| performance.morningstar.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| query.yahooapis.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| query1.finance.yahoo.com | 4 CHECK_FAILED, 28 PASS | 28 PASS, 4 UNCLASSIFIED_FAIL | 1 | 27 |
| query2.finance.yahoo.com | 2 PASS | 2 PASS | 1 | 1 |
| quotes.freerealtime.com | 2 PASS | 2 PASS | 1 | 1 |
| research.investors.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 1 | 0 |
| research.tdameritrade.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| stockcharts.com | 2 PASS | 2 PASS | 0 | 2 |
| www.advfn.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 0 | 1 |
| www.barchart.com | 2 CHECK_FAILED | 2 UNCLASSIFIED_FAIL | 0 | 0 |
| www.etfinvestmentoutlook.com | 2 CHECK_FAILED | 2 UNCLASSIFIED_FAIL | 0 | 0 |
| www.etfscreen.com | 2 CHECK_FAILED | 2 UNCLASSIFIED_FAIL | 0 | 0 |
| www.google.com | 2 CHECK_FAILED | 2 UNCLASSIFIED_FAIL | 0 | 0 |
| www.gurufocus.com | 2 CHECK_FAILED, 3 PASS | 3 PASS, 1 PROVIDER_UNAVAILABLE, 1 UNCLASSIFIED_FAIL | 1 | 2 |
| www.marketwatch.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| www.nasdaq.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| www.reuters.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| www.stockscores.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 0 | 1 |
| www.zacks.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 0 | 1 |
| fc.yahoo.com | No cases | No cases | 0 | 0 |
| ca.advfn.com | 1 CHECK_FAILED | 1 UNCLASSIFIED_FAIL | 0 | 0 |

## Cases without a working result

| Case | Raw | Interpretation | Reason / recorded result |
|---|---|---|---|
| `smfGetADVFNElement` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `smfGetEconData` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error; length=5;  |
| `smfGetGuruFocusItem` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `smfGetGuruFocusCSVItem` | CHECK_FAILED | PROVIDER_UNAVAILABLE | The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct; String: ; length=0;  |
| `smfGetMSNHistory` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; Array  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12;  |
| `smfYahooAPIData` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error; length=5;  |
| `smfGetGoogleHistoryCSV` | CHECK_FAILED | CODE_FAILURE | The case supplied or constructed an invalid request URL (same-case curl3/malformed-URL evidence); this is not evidence of a provider outage; Array Error -- Request failed (3): curl: (3) URL rejected: Malformed input to a URL function \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=36;  |
| `smfGetMorningstarHistoryCSV` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; Array Error -- Request failed (28): curl: (28) Failed to connect to performance.morningstar.com port 80 af \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12;  |
| `smfRSITargetPrices` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; Array Error \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12;  |
| `smfGetMSNOptionQuote` | CHECK_FAILED | OBSOLETE | The case identifies an obsolete response; it does not establish a working feed; String: Obsolete; length=8;  |
| `smfGetOXOptionQuote` | CHECK_FAILED | OBSOLETE | The case identifies an obsolete response; it does not establish a working feed; String: Obsolete -- OptionsXpress is no longer a valid data source; length=58;  |
| `smfGetOX2OptionQuote` | CHECK_FAILED | OBSOLETE | The case identifies an obsolete response; it does not establish a working feed; String: Obsolete -- OptionsXpress is no longer a valid data source; length=58;  |
| `smfGetOX3OptionQuote` | CHECK_FAILED | OBSOLETE | The case identifies an obsolete response; it does not establish a working feed; String: Obsolete -- OptionsXpress is no longer a valid data source; length=58;  |
| `LiveSource_bigcharts_marketwatch_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error -- HTTP 403; length=17;  |
| `Live_smfGetBigChartsOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error; length=5;  |
| `LiveSource_core_api_barchart_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (28): curl: (28) Failed to connect to core-api.barchart.com port 443 after 8003 ms: Timeout was reached / ; length=128;  |
| `Live_smfGetBarChartOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: ; length=0;  |
| `LiveSource_data_moneycentral_msn_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: data.moneycentral.msn.com / ; length=89;  |
| `LiveSource_download_finance_yahoo_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: download.finance.yahoo.com / ; length=90;  |
| `LiveSource_dvd_netflix_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!doctype html><html lang="en" class=""><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/><meta charSet="utf-8"/><meta http-equiv="X-UA-Compatible" content="IE=edge"/><meta http-equiv="origin-trial" data-feature="EME Extension - Policy Check" data-expires="2018-11-26" content="Aob+++752GiUzm1RNSIkM9TINnQDxTlxz02v8hFJK/uGO2hmXnJqH8c/ZpI05b2nLsHDhGO3Ce2z... |
| `LiveSource_finance_yahoo_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct; String: Error -- HTTP 404; length=17;  |
| `Live_smfGetYahooOptionQuote1` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; Double: 0;  |
| `Live_smfGetYahooOptionQuote2` | CHECK_FAILED | PROVIDER_UNAVAILABLE | The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct; String: Error; length=5;  |
| `LiveSource_fred_stlouisfed_org` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (28): curl: (28) Operation timed out after 25010 milliseconds with 0 bytes received / ; length=108;  |
| `LiveSource_investing_money_msn_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: investing.money.msn.com / ; length=87;  |
| `LiveSource_oic_ivolatility_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct; String: Error -- HTTP 404; length=17;  |
| `Live_smfGet888OptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct; String: Error; length=5;  |
| `LiveSource_performance_morningstar_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (28): curl: (28) Failed to connect to performance.morningstar.com port 80 after 8010 ms: Timeout was reached / ; length=133;  |
| `LiveSource_query_yahooapis_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: query.yahooapis.com / ; length=83;  |
| `Live_smfGetOptionQuotes` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; Array 0 \|  cells=1;  |
| `Live_smfGetYahooOptionQuote` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; Double: 0;  |
| `LiveSource_research_tdameritrade_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: research.tdameritrade.com / ; length=89;  |
| `LiveSource_www_advfn_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd"> / <html lang="en" xmlns="http://www.w3.org/1999/xhtml"> / <head> / <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /> / <meta http-equiv="X-UA-Compatible" content="IE=edge" /> / <meta name="format-detection" content="telephone=no" /> / <meta name='impact-site-verifica... |
| `LiveSource_www_barchart_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: ; length=0;  |
| `LiveSource_www_etfinvestmentoutlook_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!doctype html> / <html lang="en-US" prefix="og: https://ogp.me/ns#"> / <head> / 	<meta charset="UTF-8"> / <script> / var gform;gform\|\|(document.addEventListener("gform_main_scripts_loaded",function(){gform.scriptsLoaded=!0}),document.addEventListener("gform/theme/scripts_loaded",function(){gform.themeScriptsLoaded=!0}),window.addEventListener("DOMContentLoaded",function(){... |
| `LiveSource_www_etfscreen_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!DOCTYPE html><html lang="en"  data-adblockkey=MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBANnylWw2vLY4hUn9w06zQKbhKBfvjFUCsdFlb6TdQhxb9RXWXuI4t31c+o8fYOv/s8q1LGPga3DE1L/tHU4LENMCAwEAAQ==_Gdh2PARqVtW/WK8warrh1EbbzTiq++73cTKnKcsLSDMUBwnP8Lk+TWBveKaPzyMLxAUzN5NT43l3VVXXC9zfig==><head><meta charset="utf-8"><title>etfscreen.com&nbsp;-&nbsp;etfscreen Resources and Information.</title><meta n... |
| `LiveSource_www_google_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!doctype html><html lang="en-US" dir="ltr"><head><base href="https://www.google.com/finance/"><link rel="preconnect" href="//www.gstatic.com"><meta name="referrer" content="origin"><script nonce="BF9ZS5jwWZT8U0K2KmuU6A">window['ppConfig'] = {productName: 'FinHubUi', deleteIsEnforced:  true , sealIsEnforced:  true , heartbeatRate:  0.5 , periodicReportingRateMillis:  60000.0 ... |
| `Live_smfGetGoogleOptionQuote` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: ; length=0;  |
| `LiveSource_www_marketwatch_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 401; String: Error -- HTTP 401; length=17;  |
| `Live_smfGetMWOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 401; String: Error; length=5;  |
| `LiveSource_www_nasdaq_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error -- HTTP 403; length=17;  |
| `Live_smfGetNASDAQOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error; length=5;  |
| `LiveSource_www_reuters_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 401; String: Error -- HTTP 401; length=17;  |
| `LiveSource_www_stockscores_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!--include virtual="dbini.asp"--> /  / _x000D_ / <!DOCTYPE html>_x000D_ / <html dir="ltr" lang="en-US">_x000D_ / <head>_x000D_ /     _x000D_ / 	<meta http-equiv="content-type" content="text/html; charset=utf-8" />_x000D_ / 	<meta name="author" content="Stockscores" />_x000D_ / _x000D_ / 	<!-- Stylesheets_x000D_ / 	============================================= -->_x000D_ / 	<... |
| `LiveSource_www_zacks_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String:  /  /  / <!DOCTYPE html> / <html lang="en"> /  / <!-- Faster page load hack --> / <!--[if IE]><![endif]--> /  / <head> / <!-- Begin: head --> / <!-- z2_head_module_v2 start --> /  / <meta charset="utf-8" /> /  / <style> / *{margin: 0 auto;padding: 0;outline: 0 none}.clearfix{zoom: 1}.clearfix:after{content:".";display:block;height: 0;clear:both;visibility:hidden}html{font-size: 100... |
| `Extract_finance_google_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `Extract_research_investors_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `Extract_research_tdameritrade_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error; length=5;  |
| `Extract_www_barchart_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `Extract_www_etfinvestmentoutlook_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `Extract_www_etfscreen_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `Extract_www_reuters_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 401; String: Error; length=5;  |
| `Extract_investing_money_msn_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: ; length=0;  |
| `Extract_core-api_barchart_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Not Found; length=9;  |

## Coverage limits

Inventoried public .bas Functions and selected cases only; Subs, Dictionary members, private functions, and every possible argument/provider field are not claimed covered.

32 selected cases enter formulas in a worksheet: 32 PASS. `worksheet_entry` is independent of fixture/live categories; those categories do not replace evidence of actual formula assignment.

`uncovered-call-patterns.csv` records 151 missing primary, worksheet, live-extraction, or unfinished call patterns. `report-data.json` lists each function's exact expressions and all source endpoints.

Provider-unavailable classifications require matching case/run/reset fields and a same-case HTTP failure or a captured direct-response rejection marker. The label describes the requested endpoint during this run; it does not establish a provider-wide outage or prove that the requested URL is correct. HTTP 200 plus an extraction error remains unclassified when the response content was not captured. Last-HTTP diagnostics alone are not causal evidence.

`native-results.private.json` preserves raw native evidence with owner-only file permissions. Shared native/report JSON, Markdown and CSV redact anonymous Yahoo crumbs and cookies; they do not publish those session values.

The report reads the saved workbook only. It does not rerun Excel, retry providers, test a fresh Mac, or infer coverage for unbounded caller URLs.
