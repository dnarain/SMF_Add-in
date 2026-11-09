# SMF Mac native suite results

Run `smf-0.7-google-full-run4` (saved state `DONE`) records 538 planned cases across 121 public functions and 36 inventoried source groups.

Raw results: 36 CHECK_FAILED, 502 PASS. Interpreted results: 44 NOT_SUPPORTED, 458 PASS, 22 PROVIDER_UNAVAILABLE, 14 UNCLASSIFIED_FAIL.

Expected version: `3.0.2024.07.30-mac.0.7-dev`. Observed version: `Stock Market Functions add-in, Version 3.0.2024.07.30-mac.0.7-dev ({{PRIVATE_NATIVE_RUN}}; Macintosh (ARM) Version 27.0.1 (Build 26A434); 16.114; ; Local; 1)`. Version verified: True.

Workbook SHA256: `392139b5c9de34eaf7399b8860e4b2203ca9c09b8d9eeed64ff557ea0bf8fd93`. Saved VBA SHA256: `7d309445abba5da5fe4bdded4e4652d47b992fc41b1158423e6cfc38d0df0b83`.

Columns R:S record text-fetch counter changes across setup, call, checks and cleanup. A zero fetchURL-bridge delta excludes that text-fetch path. Image downloads use a separate handler. Positive counts still require matching case/run markers. Later calls can overwrite the last diagnostic in that case.

A PASS applies to the recorded input/assertion. Many live assertions require only a nonempty or positive result. They do not independently verify the provider's field accuracy, meaning, or freshness. Endpoint-content PASS, offline fixture parsing, error responses, obsolete responses, and NOT_SUPPORTED guard results do not establish a working financial feed.

63 provider-data/image checks passed, including 2 image checks and any cached values returned by the selected calls. This is not a count of fresh provider requests.

The operator ran this saved source workbook on one Mac. The VBA hash identifies that workbook and its imported runner. This report does not establish a fresh macOS account or an unmodified packaged add-in.

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
| `RCHGetElementNumber` | PASS | PASS | 3 NOT_SUPPORTED, 11 PASS, 4 PROVIDER_UNAVAILABLE, 5 UNCLASSIFIED_FAIL |
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
| `RCHGetYahooQuotes` | PASS | PASS | 3 NOT_SUPPORTED, 3 PASS |
| `RCHImageTable` | PASS | PASS | 1 PASS |
| `smfBuildYahooTimeSeriesURL` | PASS | PASS | 1 PASS |
| `smfCDec` | PASS | PASS | 1 PASS |
| `smfChkJSONFieldType` | PASS | PASS | 1 PASS |
| `smfConvertData` | PASS | PASS | 1 PASS |
| `smfConvertYahooTicker` | PASS | PASS | 1 PASS |
| `smfDate2Unix` | PASS | PASS | 1 PASS |
| `smfEval` | PASS | PASS | 1 PASS |
| `smfEvaluateTwice` | PASS | PASS | 1 PASS |
| `smfGet888OptionQuote` | PASS | NOT_SUPPORTED | 2 NOT_SUPPORTED |
| `smfGetAConstants` | PASS | PASS | 1 PASS |
| `smfGetAData` | PASS | PASS | 1 PASS |
| `smfGetADVFNElement` | PASS | PASS | 37 PASS |
| `smfGetaFiles` | PASS | PASS | 1 PASS |
| `smfGetAParms` | PASS | PASS | 1 PASS |
| `smfGetBarChartOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetBigChartsOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetCSVFile` | PASS | PASS | 5 PASS |
| `smfGetEconData` | PASS | PASS | 67 PASS |
| `smfGetGoogleHistoryCSV` | CHECK_FAILED | UNCLASSIFIED_FAIL | 8 PASS, 1 UNCLASSIFIED_FAIL |
| `smfGetGoogleOptionQuote` | PASS | NOT_SUPPORTED | 2 NOT_SUPPORTED |
| `smfGetGuruFocusCSVItem` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfGetGuruFocusData` | PASS | PASS | 1 PASS |
| `smfGetGuruFocusItem` | CHECK_FAILED | UNCLASSIFIED_FAIL | 1 UNCLASSIFIED_FAIL |
| `smfGetMorningstarHistoryCSV` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfGetMSNHistory` | CHECK_FAILED | PROVIDER_UNAVAILABLE | 1 PROVIDER_UNAVAILABLE |
| `smfGetMSNOptionQuote` | PASS | NOT_SUPPORTED | 2 NOT_SUPPORTED |
| `smfGetMWOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetNASDAQOptionQuote` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetOptionExpirations` | PASS | PASS | 3 NOT_SUPPORTED, 6 PASS |
| `smfGetOptionExpiry` | PASS | PASS | 1 PASS |
| `smfGetOptionQuotes` | PASS | PASS | 7 NOT_SUPPORTED, 3 PASS |
| `smfGetOptionStrikes` | PASS | PASS | 3 NOT_SUPPORTED, 4 PASS |
| `smfGetOX2OptionQuote` | PASS | NOT_SUPPORTED | 2 NOT_SUPPORTED |
| `smfGetOX3OptionQuote` | PASS | NOT_SUPPORTED | 2 NOT_SUPPORTED |
| `smfGetOXOptionQuote` | PASS | NOT_SUPPORTED | 2 NOT_SUPPORTED |
| `smfGetTagContent` | PASS | PASS | 3 PASS |
| `smfGetWebPage` | PASS | PASS | 1 NOT_SUPPORTED, 1 PASS |
| `smfGetYahooHistory` | PASS | PASS | 33 PASS |
| `smfGetYahooHistoryCSV` | PASS | PASS | 2 PASS |
| `smfGetYahooHistoryCSVData` | PASS | PASS | 2 PASS |
| `smfGetYahooJSONData` | PASS | PASS | 1 PASS |
| `smfGetYahooJSONField` | PASS | PASS | 1 PASS |
| `smfGetYahooOptionQuote` | PASS | PASS | 2 PASS |
| `smfGetYahooOptionQuote1` | PASS | PASS | 35 PASS |
| `smfGetYahooOptionQuote2` | PASS | PASS | 1 PASS, 1 PROVIDER_UNAVAILABLE |
| `smfGetYahooPortfolioView` | PASS | PASS | 2 PASS |
| `SMFHighBetween` | PASS | PASS | 2 PASS |
| `smfHTMLDecode` | PASS | PASS | 3 PASS |
| `smfInCellChart` | PASS | PASS | 1 PASS |
| `smfJoin` | PASS | PASS | 1 PASS |
| `smfJSONExtractField` | PASS | PASS | 1 PASS |
| `smfJSONExtractField_x64` | PASS | PASS | 1 PASS |
| `smfJSONExtractKeys` | PASS | PASS | 1 PASS |
| `smfLastPrice` | PASS | PASS | 2 PASS |
| `smfLogInternetCalls` | PASS | PASS | 1 PASS |
| `SMFLowBetween` | PASS | PASS | 2 PASS |
| `SMFMacElementsSelfTest` | PASS | PASS | 2 PASS |
| `SMFMacEmptyRecipeSelfTest` | PASS | PASS | 1 PASS |
| `smfMacFetch` | PASS | PASS | 10 NOT_SUPPORTED, 22 PASS, 10 PROVIDER_UNAVAILABLE, 7 UNCLASSIFIED_FAIL |
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
| `smfMacReadResponse` | PASS | PASS | 9 PASS |
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
| `smfRSITargetPrices` | PASS | PASS | 65 PASS |
| `smfStrExtr` | PASS | PASS | 1 PASS |
| `smfStripHTML` | PASS | PASS | 1 PASS |
| `SMFTech` | PASS | PASS | 10 PASS |
| `smfUnix2Date` | PASS | PASS | 1 PASS |
| `smfUnix2DateStr` | PASS | PASS | 1 PASS |
| `smfWord` | PASS | PASS | 1 PASS |
| `smfYahooAPIData` | PASS | NOT_SUPPORTED | 2 NOT_SUPPORTED |
| `smfYahooCodeDesc` | PASS | PASS | 1 PASS |

## Source groups

| Source | Raw cases | Interpreted cases | Endpoint-content passes | Provider-data/image checks passed |
|---|---|---|---|---|
| bigcharts.marketwatch.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| c.stockcharts.com | No cases | No cases | 0 | 0 |
| climbermel.github.io | 1 PASS | 1 PASS | 1 | 0 |
| core-api.barchart.com | 3 CHECK_FAILED | 3 PROVIDER_UNAVAILABLE | 0 | 0 |
| data.moneycentral.msn.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| download.finance.yahoo.com | 1 PASS | 1 NOT_SUPPORTED | 0 | 0 |
| dvd.netflix.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 0 | 1 |
| finance.google.com | 2 CHECK_FAILED, 4 PASS | 4 PASS, 2 UNCLASSIFIED_FAIL | 4 | 0 |
| finance.yahoo.com | 1 CHECK_FAILED, 2 PASS | 2 PASS, 1 PROVIDER_UNAVAILABLE | 1 | 1 |
| finviz.com | 3 PASS | 3 PASS | 1 | 2 |
| fred.stlouisfed.org | 1 CHECK_FAILED | 1 PROVIDER_UNAVAILABLE | 0 | 0 |
| investing.money.msn.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| ogres-crypt.com | No cases | No cases | 0 | 0 |
| oic.ivolatility.com | 2 PASS | 2 NOT_SUPPORTED | 0 | 0 |
| performance.morningstar.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| query.yahooapis.com | 1 PASS | 1 NOT_SUPPORTED | 0 | 0 |
| query1.finance.yahoo.com | 34 PASS | 34 PASS | 1 | 33 |
| query2.finance.yahoo.com | 2 PASS | 2 PASS | 1 | 1 |
| quotes.freerealtime.com | 2 PASS | 2 PASS | 1 | 1 |
| research.investors.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 1 | 0 |
| research.tdameritrade.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| stockcharts.com | 2 PASS | 2 PASS | 0 | 2 |
| www.advfn.com | 1 CHECK_FAILED, 2 PASS | 2 PASS, 1 UNCLASSIFIED_FAIL | 0 | 2 |
| www.barchart.com | 2 CHECK_FAILED | 2 UNCLASSIFIED_FAIL | 0 | 0 |
| www.etfinvestmentoutlook.com | 2 CHECK_FAILED | 2 UNCLASSIFIED_FAIL | 0 | 0 |
| www.etfscreen.com | 2 CHECK_FAILED | 2 UNCLASSIFIED_FAIL | 0 | 0 |
| www.google.com | 2 PASS | 2 NOT_SUPPORTED | 0 | 0 |
| www.gurufocus.com | 2 CHECK_FAILED, 3 PASS | 3 PASS, 1 PROVIDER_UNAVAILABLE, 1 UNCLASSIFIED_FAIL | 1 | 2 |
| www.marketwatch.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| www.nasdaq.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| www.reuters.com | 2 CHECK_FAILED | 2 PROVIDER_UNAVAILABLE | 0 | 0 |
| www.stockscores.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 0 | 1 |
| www.zacks.com | 1 CHECK_FAILED, 1 PASS | 1 PASS, 1 UNCLASSIFIED_FAIL | 0 | 1 |
| fc.yahoo.com | No cases | No cases | 0 | 0 |
| ca.advfn.com | 3 PASS | 3 PASS | 0 | 3 |
| api.stlouisfed.org | 6 PASS | 6 PASS | 0 | 5 |

## Separate HTTP diagnostics

The explicitly selected HTTP directory contains 58 probes: 18 0, 24 200, 1 202, 6 401, 2 403, 7 404 (HTTP status counts). These requests are separate observations; they do not change native case interpretations, and HTTP 200 does not establish content or extraction success.

| Probe | Source IDs | HTTP | curl exit | Requested URL |
|---|---|---|---|---|
| 1 | bigcharts.marketwatch.com | 403 | 0 | http://bigcharts.marketwatch.com/quickchart/options.asp?showAll=True&symb=AAPL |
| 2 | c.stockcharts.com | 0 | 28 | https://c.stockcharts.com/c-sc/sc?s=AAPL&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806 |
| 3 | climbermel.github.io | 200 | 0 | https://climbermel.github.io/SMF_Add-in/Elements/smf-elements-4.txt |
| 4 | core-api.barchart.com | 0 | 28 | https://core-api.barchart.com/v1/momentum/get?raw=1&fields=exchange,high52w,low52w,newHighs,newLows,advancingVolume,unchangedVolume,decliningVolume,advancingIssues,percentAdvancingIssues,unchangedIssues,percentUnchang... |
| 5 | core-api.barchart.com | 0 | 28 | https://core-api.barchart.com/v1/options/chain?fields=optionType,strikePrice,lastPrice,percentFromLast,priceChange,percentChange,bidPrice,midpoint,askPrice,theoretical,volatility,delta,gamma,rho,theta,vega,volume,open... |
| 6 | data.moneycentral.msn.com | 0 | 6 | http://data.moneycentral.msn.com/scripts/chrtsrv.dll?C1=2&C2=&FileDownload=&C9=0&Symbol=AAPL&C1=11&C6=2026&C7=11&C8=2026 |
| 7 | download.finance.yahoo.com | 0 | None | https://download.finance.yahoo.com/d/quotes.csv?s=AAPL&f=sl1&e=.ignore |
| 8 | dvd.netflix.com | 200 | 0 | http://dvd.netflix.com/Movie/60010110 |
| 9 | finance.google.com | 200 | 0 | https://finance.google.com/finance?fstype=ii&q=AAPL |
| 10 | google_history_csv | 200 | 0 | http://finance.google.com/finance/historical?output=csv&q=AAPL&startdate=Nov%205%2C%202026&enddate=Nov%206%2C%202026 |
| 11 | finance.yahoo.com | 200 | 0 | https://finance.yahoo.com/quote/AAPL261120C00350000 |
| 12 | finance.yahoo.com | 404 | 0 | http://finance.yahoo.com/q/op?s=AAPL&date=1795132800 |
| 13 | finviz.com | 200 | 0 | https://finviz.com/quote.ashx?t=AAPL&p=d |
| 14 | fred.stlouisfed.org | 0 | 28 | https://fred.stlouisfed.org/data/DGS10 |
| 15 | investing.money.msn.com | 0 | 6 | http://investing.money.msn.com/investments/stock-price?Symbol=AAPL |
| 16 | ogres-crypt.com | 404 | 0 | http://ogres-crypt.com/php/chart.php?d=1,50,98 |
| 17 | oic.ivolatility.com | 0 | None | http://oic.ivolatility.com/oic_adv_options.j?exp_date=-1&ticker=AAPL |
| 18 | performance.morningstar.com | 0 | 28 | http://performance.morningstar.com/perform/Performance/stock/exportStockPrice.action?t=AAPL&pd=custom&freq=d&sd=11/05/2026&ed=11/06/2026 |
| 19 | query.yahooapis.com | 0 | None | http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20%28%22AAPL%22%29&diagnostics=false&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys |
| 20 | query1.finance.yahoo.com | 200 | 0 | https://query1.finance.yahoo.com/v10/finance/quoteSummary/AAPL?modules=price |
| 21 | query1.finance.yahoo.com | 200 | 0 | https://query1.finance.yahoo.com/v7/finance/options/AAPL? |
| 22 | query1.finance.yahoo.com, yahoo_options_20261120 | 200 | 0 | https://query1.finance.yahoo.com/v7/finance/options/AAPL?date=1795132800 |
| 23 | query1.finance.yahoo.com | 200 | 0 | https://query1.finance.yahoo.com/v10/finance/quoteSummary/AAPL261120C00350000?modules=summaryDetail |
| 24 | yahoo_history_AAPL | 401 | 0 | https://query1.finance.yahoo.com/v7/finance/download/AAPL?period1=1793836800&period2=1794009600&interval=1d&events=history |
| 25 | yahoo_history_hardcoded_MMM | 401 | 0 | https://query1.finance.yahoo.com/v7/finance/download/MMM?period1=1493610466&period2=1496202466&interval=1d&events=history&crumb= |
| 26 | yahoo_quote_v7 | 200 | 0 | https://query1.finance.yahoo.com/v7/finance/quote?fields=symbol,longName,shortName,regularMarketPrice,regularMarketTime,regularMarketChange,regularMarketDayHigh,regularMarketDayLow,regularMarketPrice,regularMarketOpen... |
| 27 | yahoo_quote_v6_constant | 404 | 0 | https://query1.finance.yahoo.com/v6/finance/quote?fields=symbol,longName,shortName,regularMarketPrice,regularMarketTime,regularMarketChange,regularMarketDayHigh,regularMarketDayLow,regularMarketPrice,regularMarketOpen... |
| 28 | yahoo_news | 404 | 0 | https://query1.finance.yahoo.com/v2/finance/news?count=20&symbols=AAPL&start=0 |
| 29 | yahoo_timeseries | 200 | 0 | https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/AAPL?merge=false&padTimeSeries=true&period1=493590046&period2=1794164399&type=annualTotalRevenue&lang=en-US&region=US |
| 30 | query2.finance.yahoo.com | 200 | 0 | https://query2.finance.yahoo.com/v10/finance/quoteSummary/AAPL?modules=assetProfile |
| 31 | quotes.freerealtime.com | 200 | 0 | https://quotes.freerealtime.com/stock-scouter?title=AAPL |
| 32 | research.investors.com | 200 | 0 | https://research.investors.com/stock-checkup/xxxx-AAPL.aspx |
| 33 | research.tdameritrade.com | 0 | 6 | http://research.tdameritrade.com/public/stocks/overview/overview.asp?symbol=AAPL |
| 34 | stockcharts.com | 200 | 0 | http://stockcharts.com/c-sc/sc?s=AAPL&p=D&b=3&g=0&id=t08330678207&r=4815 |
| 35 | www.advfn.com | 200 | 0 | https://www.advfn.com/stock-market/NASDAQ/AAPL/financials?&mode=annual_reports&btn=start_date |
| 36 | advfn_discover_exchange | 200 | 0 | https://www.advfn.com/p.php?pid=financials&mode=quarterly_reports&symbol=AAPL |
| 37 | advfn_chart | 404 | 0 | https://www.advfn.com/p.php?pid=financialgraphs&a0=13&a1=[redacted] |
| 38 | www.barchart.com | 202 | 0 | https://www.barchart.com/stocks/quotes/AAPL/opinion |
| 39 | www.etfinvestmentoutlook.com | 200 | 0 | http://www.etfinvestmentoutlook.com/rank_breadth_new_highs_lows_avg_d.php |
| 40 | www.etfscreen.com | 200 | 0 | https://www.etfscreen.com/rsftrends.php?wl=0&s=default&t=6&d=i&ftS=yes&ftL=yes&vFf=dolVol21&vFl=gt&vFv=0&d=i&rs=1 |
| 41 | www.google.com | 0 | None | http://www.google.com/finance/option_chain?output=json&q=AAPL&expd=20&expm=11&expy=2026 |
| 42 | www.gurufocus.com | 200 | 0 | https://www.gurufocus.com/financials/AAPL |
| 43 | gurufocus_csv | 404 | 0 | https://www.gurufocus.com/download_financials_in_CSV.php?symbol=AAPL |
| 44 | www.marketwatch.com | 401 | 0 | http://www.marketwatch.com/investing/stock/AAPL/options?countrycode=US&showAll=True |
| 45 | www.marketwatch.com | 401 | 0 | http://www.marketwatch.com/investing/fund/AAPL/options?countrycode=US&showAll=True |
| 46 | www.marketwatch.com | 401 | 0 | http://www.marketwatch.com/investing/index/AAPL/options?countrycode=US&showAll=True |
| 47 | www.nasdaq.com | 403 | 0 | http://www.nasdaq.com/symbol/aapl/option-chain/261120C00350000-aapl-Call |
| 48 | www.reuters.com | 401 | 0 | https://www.reuters.com/finance/stocks/financial-highlights/AAPL |
| 49 | www.stockscores.com | 200 | 0 | https://www.stockscores.com/charts/charts?ticker=AAPL |
| 50 | www.zacks.com | 200 | 0 | https://www.zacks.com/stock/quote/AAPL |
| 51 | fc.yahoo.com | 404 | 0 | https://fc.yahoo.com |
| 52 | ca.advfn.com | 200 | 0 | https://ca.advfn.com/stock-market/NASDAQ/AAPL/financials?&mode=annual_reports&btn=start_date |
| 53 | unreachable:smfGetMSNOptionQuote | 0 | None | http://investing.money.msn.com/investments/equity-options/?symbol=AAPL&optionsdate=11/20/2026 |
| 54 | unreachable:smfGetOXOptionQuote | 0 | None | https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=0&lstMarket=0&ChainType=3&lstMonths=11/20/2026;1&Symbol=AAPL |
| 55 | unreachable:smfGetOX2OptionQuote | 0 | None | https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=All&lstMarket=0&ChainType=14&lstMonths=11/20/2026;7&Symbol=AAPL |
| 56 | unreachable:smfGetOX3OptionQuote | 0 | None | https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=All&lstMarket=0&lstMonths=11/20/2026;7&Symbol=AAPL |
| 57 | unreachable:smfGetOptionExpirations | 0 | None | https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=All&lstMarket=0&ChainType=14&lstMonths=0&Symbol=AAPL |
| 58 | unreachable:smfGetOptionStrikes | 0 | None | https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=0&lstMarket=0&Symbol=AAPL&lstMonths=11/20/2026;7 |

## Cases without a working result

| Case | Raw | Interpretation | Reason / recorded result |
|---|---|---|---|
| `smfGetGuruFocusItem` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: Error; length=5;  |
| `smfGetGuruFocusCSVItem` | CHECK_FAILED | PROVIDER_UNAVAILABLE | The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct; String: ; length=0;  |
| `smfGetMSNHistory` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; Array  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12;  |
| `smfYahooAPIData` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042;  |
| `smfGetGoogleHistoryCSV` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; Array <!doctype html><html lang="en-US" dir="ltr"><head><base href="https://www.google.com/finance/"><link \| var n=typeof Object.defineProperties=="function"?Object.defineProperty:function(a \|  \|  \|  \|  \|  deleteIsEnforced:  true  \| b \|  \|  \|  \|  \|  sealIsEnforced:  true  \| c){if(a==Array.prototype\|\|a==Object.prototype)return a;a[b]=c.value;return a};function p(a){a=["... |
| `smfGetMorningstarHistoryCSV` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; Array Error -- Request failed (28): curl: (28) Failed to connect to performance.morningstar.com port 80 af \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12;  |
| `smfGet888OptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `smfGetGoogleOptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `smfGetMSNOptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `smfGetOXOptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `smfGetOX2OptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `smfGetOX3OptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Worksheet_smfGetMSNOptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Worksheet_smfGetOXOptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Worksheet_smfGetOX2OptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Worksheet_smfGetOX3OptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_YahooCSVLegacy` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array [Excel error] \| [Excel error] \|  cells=2; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_YahooCSVRange` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array [Excel error] \| [Excel error] \| [Excel error] \| [Excel error] \|  cells=4; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_YahooYQLFeedF` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Expirations_OX` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Strikes_OX` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Expirations_G` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Strikes_G` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Expirations_8` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Strikes_8` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Wrapper_MSN` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array [Excel error] \| [Excel error] \| [Excel error] \| [Excel error] \|  cells=4; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Wrapper_OX3` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array [Excel error] \| [Excel error] \| [Excel error] \| [Excel error] \|  cells=4; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Wrapper_G` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array [Excel error] \| [Excel error] \| [Excel error] \| [Excel error] \|  cells=4; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_Wrapper_8` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array [Excel error] \| [Excel error] \| [Excel error] \| [Excel error] \|  cells=4; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_WrapperCodes` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \|  cells=18; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_URL_msn-options` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- SMF marks the MSN option quote function obsolete.; length=75; Fetch-entry delta=1; fetchURL-bridge delta=0;  |
| `NotSupported_URL_optionsxpress-options` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- SMF records that optionsXpress is no longer a valid data source.; length=90; Fetch-entry delta=1; fetchURL-bridge delta=0;  |
| `NotSupported_URL_yahoo-quote-csv` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- SMF explicitly marks the old ticker/quote-code CSV mode obsolete.; length=91; Fetch-entry delta=1; fetchURL-bridge delta=0;  |
| `NotSupported_URL_yahoo-yql` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- Yahoo retired query.yahooapis.com YQL on 2019-01-03, as recorded in the SMF change log.; length=113; Fetch-entry delta=1; fetchURL-bridge delta=0;  |
| `NotSupported_URL_google-option-chain` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- The SMF option-strike manual explicitly labels the Google option source obsolete.; length=107; Fetch-entry delta=1; fetchURL-bridge delta=0;  |
| `NotSupported_URL_oic-legacy-options` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- The SMF option-strike manual explicitly labels source 8 obsolete; these calls use that same legacy endpoint.; length=134; Fetch-entry delta=1; fetchURL-bridge delta=0;  |
| `NotSupported_WorksheetWrapperNamed` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array Bid Price \| [Excel error] \| [Excel error] \| Ask Price \| [Excel error] \| [Excel error] \|  cells=6; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_WorksheetWrapperCodes` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \| [Excel error] \| [Excel error] \| Bid Price \|  cells=18; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_PrimedRetiredPage` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- Yahoo retired query.yahooapis.com YQL on 2019-01-03, as recorded in the SMF change log.; length=113; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_YahooCSVRetiredURL` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Array [Excel error] \| [Excel error] \|  cells=2; Fetch-entry delta=0; fetchURL-bridge delta=0;  |
| `NotSupported_All288DefaultFields` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Long: 288;  |
| `NotSupported_FieldWorksheetNA` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042;  |
| `NotSupported_FieldIFNA` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Not Supported; length=13;  |
| `LiveSource_bigcharts_marketwatch_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error -- HTTP 403; length=17;  |
| `Live_smfGetBigChartsOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error; length=5;  |
| `LiveSource_core_api_barchart_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (28): curl: (28) Failed to connect to core-api.barchart.com port 443 after 8009 ms: Timeout was reached / ; length=128;  |
| `Live_smfGetBarChartOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: ; length=0;  |
| `LiveSource_data_moneycentral_msn_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: data.moneycentral.msn.com / ; length=89;  |
| `LiveSource_download_finance_yahoo_com` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- SMF explicitly marks the old ticker/quote-code CSV mode obsolete.; length=91;  |
| `LiveSource_dvd_netflix_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!doctype html><html lang="en" class=""><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/><meta charSet="utf-8"/><meta http-equiv="X-UA-Compatible" content="IE=edge"/><meta http-equiv="origin-trial" data-feature="EME Extension - Policy Check" data-expires="2018-11-26" content="Aob+++752GiUzm1RNSIkM9TINnQDxTlxz02v8hFJK/uGO2hmXnJqH8c/ZpI05b2nLsHDhGO3Ce2z... |
| `Live_smfGetYahooOptionQuote2` | CHECK_FAILED | PROVIDER_UNAVAILABLE | The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct; String: Error; length=5;  |
| `LiveSource_fred_stlouisfed_org` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (28): curl: (28) Operation timed out after 25008 milliseconds with 0 bytes received / ; length=108;  |
| `LiveSource_investing_money_msn_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: investing.money.msn.com / ; length=87;  |
| `LiveSource_oic_ivolatility_com` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- The SMF option-strike manual explicitly labels source 8 obsolete; these calls use that same legacy endpoint.; length=134;  |
| `Live_smfGet888OptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042;  |
| `LiveSource_performance_morningstar_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (28): curl: (28) Failed to connect to performance.morningstar.com port 80 after 8010 ms: Timeout was reached / ; length=133;  |
| `LiveSource_query_yahooapis_com` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- Yahoo retired query.yahooapis.com YQL on 2019-01-03, as recorded in the SMF change log.; length=113;  |
| `LiveSource_research_tdameritrade_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request recorded a transport failure; String: Error -- Request failed (6): curl: (6) Could not resolve host: research.tdameritrade.com / ; length=89;  |
| `LiveSource_www_advfn_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd"> / <html lang="en" xmlns="http://www.w3.org/1999/xhtml"> / <head> / <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /> / <meta http-equiv="X-UA-Compatible" content="IE=edge" /> / <meta name="format-detection" content="telephone=no" /> / <meta name='impact-site-verifica... |
| `LiveSource_www_barchart_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: ; length=0;  |
| `LiveSource_www_etfinvestmentoutlook_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!doctype html> / <html lang="en-US" prefix="og: https://ogp.me/ns#"> / <head> / 	<meta charset="UTF-8"> / <script> / var gform;gform\|\|(document.addEventListener("gform_main_scripts_loaded",function(){gform.scriptsLoaded=!0}),document.addEventListener("gform/theme/scripts_loaded",function(){gform.themeScriptsLoaded=!0}),window.addEventListener("DOMContentLoaded",function(){... |
| `LiveSource_www_etfscreen_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!DOCTYPE html><html lang="en"  data-adblockkey=MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBANnylWw2vLY4hUn9w06zQKbhKBfvjFUCsdFlb6TdQhxb9RXWXuI4t31c+o8fYOv/s8q1LGPga3DE1L/tHU4LENMCAwEAAQ==_Gdh2PARqVtW/WK8warrh1EbbzTiq++73cTKnKcsLSDMUBwnP8Lk+TWBveKaPzyMLxAUzN5NT43l3VVXXC9zfig==><head><meta charset="utf-8"><title>etfscreen.com&nbsp;-&nbsp;etfscreen Resources and Information.</title><meta n... |
| `LiveSource_www_google_com` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; String: Error -- Not Supported -- The SMF option-strike manual explicitly labels the Google option source obsolete.; length=107;  |
| `Live_smfGetGoogleOptionQuote` | PASS | NOT_SUPPORTED | The registered retired-mode guard passed; this does not establish a financial feed; Excel error Error 2042;  |
| `LiveSource_www_marketwatch_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 401; String: Error -- HTTP 401; length=17;  |
| `Live_smfGetMWOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 401; String: Error; length=5;  |
| `LiveSource_www_nasdaq_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error -- HTTP 403; length=17;  |
| `Live_smfGetNASDAQOptionQuote` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 403; String: Error; length=5;  |
| `LiveSource_www_reuters_com` | CHECK_FAILED | PROVIDER_UNAVAILABLE | Same-case HTTP request returned 401; String: Error -- HTTP 401; length=17;  |
| `LiveSource_www_stockscores_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String: <!--include virtual="dbini.asp"--> /  / _x000D_ / <!DOCTYPE html>_x000D_ / <html dir="ltr" lang="en-US">_x000D_ / <head>_x000D_ /     _x000D_ / 	<meta http-equiv="content-type" content="text/html; charset=utf-8" />_x000D_ / 	<meta name="author" content="Stockscores" />_x000D_ / _x000D_ / 	<!-- Stylesheets_x000D_ / 	============================================= -->_x000D_ / 	<... |
| `LiveSource_www_zacks_com` | CHECK_FAILED | UNCLASSIFIED_FAIL | Insufficient evidence to distinguish a code/provider failure; Request observed after this case's reset marker; String:  /  /  / <!DOCTYPE html> / <html lang="en"> /  / <!-- Faster page load hack --> / <!--[if IE]><![endif]--> /  / <head> / <!-- Begin: head --> / <!-- z2_head_module_v2 start --> /  / <meta charset="utf-8" /> /  / <style> / *{margin:0 auto;padding:0;outline:0 none}.clearfix{zoom:1}.clearfix:after{content:".";display:block;height:0;clear:both;visibility:hidden}html{font-size:100... |
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

51 selected cases enter formulas in a worksheet: 51 PASS. `worksheet_entry` is independent of fixture/live categories; those categories do not replace evidence of actual formula assignment.

`uncovered-call-patterns.csv` records 146 missing primary, worksheet, live-extraction, or unfinished call patterns. `report-data.json` lists each function's exact expressions and all source endpoints.

Provider-unavailable classifications require matching case/run/reset fields and a same-case HTTP failure or a captured direct-response rejection marker. The label describes the requested endpoint during this run; it does not establish a provider-wide outage or prove that the requested URL is correct. HTTP 200 plus an extraction error remains unclassified when the response content was not captured. Last-HTTP diagnostics alone are not causal evidence.

`native-results.private.json` preserves raw native evidence with owner-only file permissions. Shared native/report JSON, Markdown and CSV redact anonymous Yahoo crumbs and cookies; they do not publish those session values.

The report reads the saved workbook only. It does not rerun Excel, retry providers, test a fresh Mac, or infer coverage for unbounded caller URLs.
