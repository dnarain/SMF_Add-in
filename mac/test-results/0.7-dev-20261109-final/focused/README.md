# SMF Mac native suite results

Run `smf-0.7-google-focused-run4` (saved state `DONE`) records 246 planned cases across 121 public functions and 36 inventoried source groups.

Raw results: 246 PASS. Interpreted results: 31 NOT_SUPPORTED, 215 PASS.

Expected version: `3.0.2024.07.30-mac.0.7-dev`. Observed version: `Stock Market Functions add-in, Version 3.0.2024.07.30-mac.0.7-dev ({{PRIVATE_NATIVE_RUN}}; Macintosh (ARM) Version 27.0.1 (Build 26A434); 16.114; ; Local; 1)`. Version verified: True.

Workbook SHA256: `f083b700d602ce34ff9dbf4cd757b70837580098fb300a820665d22acd7155da`. Saved VBA SHA256: `261fcd9081e2880dbccf70b4dbff87f58b9a6b637bbe4e7d78825a42a4f7c5ac`.

Columns R:S record text-fetch counter changes across setup, call, checks and cleanup. A zero fetchURL-bridge delta excludes that text-fetch path. Image downloads use a separate handler. Positive counts still require matching case/run markers. Later calls can overwrite the last diagnostic in that case.

A PASS applies to the recorded input/assertion. Many live assertions require only a nonempty or positive result. They do not independently verify the provider's field accuracy, meaning, or freshness. Endpoint-content PASS, offline fixture parsing, error responses, obsolete responses, and NOT_SUPPORTED guard results do not establish a working financial feed.

9 provider-data/image checks passed, including 0 image checks and any cached values returned by the selected calls. This is not a count of fresh provider requests.

The operator ran this saved source workbook on one Mac. The VBA hash identifies that workbook and its imported runner. This report does not establish a fresh macOS account or an unmodified packaged add-in.

## Public functions

The primary row uses explicit case metadata, then a matching case ID, then the first selected direct call to that function. The CSV records the selection reason and every case; primary status alone does not summarize all behavior.

| Function | Primary raw | Primary interpretation | All cases |
|---|---|---|---|
| `ConvertToIso` | MISSING | UNTESTED | No cases |
| `ConvertToJson` | MISSING | UNTESTED | No cases |
| `ConvertToUtc` | MISSING | UNTESTED | No cases |
| `DateOfHigh` | MISSING | UNTESTED | No cases |
| `DecodeJsonString` | MISSING | UNTESTED | No cases |
| `fnWait` | MISSING | UNTESTED | No cases |
| `GetKeys` | MISSING | UNTESTED | No cases |
| `GetObjectProperty` | MISSING | UNTESTED | No cases |
| `GetProperty` | MISSING | UNTESTED | No cases |
| `ParseIso` | MISSING | UNTESTED | No cases |
| `ParseJson` | MISSING | UNTESTED | No cases |
| `ParseUtc` | MISSING | UNTESTED | No cases |
| `RCHCreateComment` | MISSING | UNTESTED | No cases |
| `RCHExtractData` | MISSING | UNTESTED | No cases |
| `RCHGetElementNumber` | MISSING | UNTESTED | 3 NOT_SUPPORTED |
| `RCHGetHTMLTable` | MISSING | UNTESTED | No cases |
| `RCHGetTableCell` | MISSING | UNTESTED | No cases |
| `RCHGetURLData` | MISSING | UNTESTED | No cases |
| `RCHGetURLData1` | MISSING | UNTESTED | No cases |
| `RCHGetURLData1Mac` | MISSING | UNTESTED | No cases |
| `RCHGetURLData2` | MISSING | UNTESTED | No cases |
| `RCHGetURLData3` | MISSING | UNTESTED | No cases |
| `RCHGetURLData4` | MISSING | UNTESTED | No cases |
| `RCHGetWebData` | MISSING | UNTESTED | No cases |
| `RCHGetYahooHistory` | MISSING | UNTESTED | No cases |
| `RCHGetYahooHistory2` | MISSING | UNTESTED | No cases |
| `RCHGetYahooQuotes` | PASS | NOT_SUPPORTED | 3 NOT_SUPPORTED, 1 PASS |
| `RCHImageTable` | MISSING | UNTESTED | No cases |
| `smfBuildYahooTimeSeriesURL` | MISSING | UNTESTED | No cases |
| `smfCDec` | MISSING | UNTESTED | No cases |
| `smfChkJSONFieldType` | MISSING | UNTESTED | No cases |
| `smfConvertData` | MISSING | UNTESTED | No cases |
| `smfConvertYahooTicker` | MISSING | UNTESTED | No cases |
| `smfDate2Unix` | MISSING | UNTESTED | No cases |
| `smfEval` | MISSING | UNTESTED | No cases |
| `smfEvaluateTwice` | MISSING | UNTESTED | No cases |
| `smfGet888OptionQuote` | MISSING | UNTESTED | No cases |
| `smfGetAConstants` | MISSING | UNTESTED | No cases |
| `smfGetAData` | MISSING | UNTESTED | No cases |
| `smfGetADVFNElement` | PASS | PASS | 36 PASS |
| `smfGetaFiles` | MISSING | UNTESTED | No cases |
| `smfGetAParms` | MISSING | UNTESTED | No cases |
| `smfGetBarChartOptionQuote` | MISSING | UNTESTED | No cases |
| `smfGetBigChartsOptionQuote` | MISSING | UNTESTED | No cases |
| `smfGetCSVFile` | PASS | PASS | 4 PASS |
| `smfGetEconData` | PASS | PASS | 66 PASS |
| `smfGetGoogleHistoryCSV` | PASS | PASS | 8 PASS |
| `smfGetGoogleOptionQuote` | MISSING | UNTESTED | No cases |
| `smfGetGuruFocusCSVItem` | MISSING | UNTESTED | No cases |
| `smfGetGuruFocusData` | MISSING | UNTESTED | No cases |
| `smfGetGuruFocusItem` | MISSING | UNTESTED | No cases |
| `smfGetMorningstarHistoryCSV` | MISSING | UNTESTED | No cases |
| `smfGetMSNHistory` | MISSING | UNTESTED | No cases |
| `smfGetMSNOptionQuote` | MISSING | UNTESTED | 1 NOT_SUPPORTED |
| `smfGetMWOptionQuote` | MISSING | UNTESTED | No cases |
| `smfGetNASDAQOptionQuote` | MISSING | UNTESTED | No cases |
| `smfGetOptionExpirations` | PASS | NOT_SUPPORTED | 3 NOT_SUPPORTED, 2 PASS |
| `smfGetOptionExpiry` | MISSING | UNTESTED | No cases |
| `smfGetOptionQuotes` | PASS | NOT_SUPPORTED | 7 NOT_SUPPORTED, 1 PASS |
| `smfGetOptionStrikes` | PASS | NOT_SUPPORTED | 3 NOT_SUPPORTED, 2 PASS |
| `smfGetOX2OptionQuote` | MISSING | UNTESTED | 1 NOT_SUPPORTED |
| `smfGetOX3OptionQuote` | MISSING | UNTESTED | 1 NOT_SUPPORTED |
| `smfGetOXOptionQuote` | MISSING | UNTESTED | 1 NOT_SUPPORTED |
| `smfGetTagContent` | PASS | PASS | 2 PASS |
| `smfGetWebPage` | PASS | NOT_SUPPORTED | 1 NOT_SUPPORTED |
| `smfGetYahooHistory` | MISSING | UNTESTED | No cases |
| `smfGetYahooHistoryCSV` | MISSING | UNTESTED | No cases |
| `smfGetYahooHistoryCSVData` | MISSING | UNTESTED | No cases |
| `smfGetYahooJSONData` | MISSING | UNTESTED | No cases |
| `smfGetYahooJSONField` | MISSING | UNTESTED | No cases |
| `smfGetYahooOptionQuote` | MISSING | UNTESTED | No cases |
| `smfGetYahooOptionQuote1` | MISSING | UNTESTED | No cases |
| `smfGetYahooOptionQuote2` | MISSING | UNTESTED | No cases |
| `smfGetYahooPortfolioView` | MISSING | UNTESTED | No cases |
| `SMFHighBetween` | MISSING | UNTESTED | No cases |
| `smfHTMLDecode` | PASS | PASS | 2 PASS |
| `smfInCellChart` | MISSING | UNTESTED | No cases |
| `smfJoin` | MISSING | UNTESTED | No cases |
| `smfJSONExtractField` | MISSING | UNTESTED | No cases |
| `smfJSONExtractField_x64` | MISSING | UNTESTED | No cases |
| `smfJSONExtractKeys` | MISSING | UNTESTED | No cases |
| `smfLastPrice` | MISSING | UNTESTED | No cases |
| `smfLogInternetCalls` | MISSING | UNTESTED | No cases |
| `SMFLowBetween` | MISSING | UNTESTED | No cases |
| `SMFMacElementsSelfTest` | PASS | PASS | 1 PASS |
| `SMFMacEmptyRecipeSelfTest` | MISSING | UNTESTED | No cases |
| `smfMacFetch` | PASS | PASS | 6 NOT_SUPPORTED, 12 PASS |
| `SMFMacHistoryCSVSelfTest` | MISSING | UNTESTED | No cases |
| `smfMacHTMLSelfTest` | MISSING | UNTESTED | No cases |
| `SMFMacHTTPStatus` | MISSING | UNTESTED | No cases |
| `SMFMacIntegrationSelfTest` | MISSING | UNTESTED | No cases |
| `SMFMacJSONSelfTest` | MISSING | UNTESTED | No cases |
| `SMFMacLastError` | MISSING | UNTESTED | No cases |
| `SMFMacLastFileError` | MISSING | UNTESTED | No cases |
| `SMFMacLastURL` | MISSING | UNTESTED | No cases |
| `SMFMacNestedRecipeSelfTest` | MISSING | UNTESTED | No cases |
| `smfMacReadElementLocation` | MISSING | UNTESTED | No cases |
| `smfMacReadElementText` | MISSING | UNTESTED | No cases |
| `smfMacReadResponse` | PASS | PASS | 8 PASS |
| `SMFMacRecipeSelfTest` | MISSING | UNTESTED | No cases |
| `smfMacStripHTML` | MISSING | UNTESTED | No cases |
| `SMFMacTableGuardSelfTest` | MISSING | UNTESTED | No cases |
| `SMFMacTableRecipeSelfTest` | MISSING | UNTESTED | No cases |
| `SMFMacTransportSelfTest` | MISSING | UNTESTED | No cases |
| `smfMacTryNestedRecipe` | MISSING | UNTESTED | No cases |
| `smfMacTryTableCellRecipe` | MISSING | UNTESTED | No cases |
| `smfMacTryTagContentRecipe` | MISSING | UNTESTED | No cases |
| `smfMacYahooChartCSV` | MISSING | UNTESTED | No cases |
| `smfMacYahooDownloadCSV` | MISSING | UNTESTED | No cases |
| `smfMacYahooDownloadURL` | MISSING | UNTESTED | No cases |
| `smfPricesBetween` | MISSING | UNTESTED | No cases |
| `smfPricesByDates` | MISSING | UNTESTED | No cases |
| `smfRSITargetPrices` | PASS | PASS | 64 PASS |
| `smfStrExtr` | MISSING | UNTESTED | No cases |
| `smfStripHTML` | MISSING | UNTESTED | No cases |
| `SMFTech` | MISSING | UNTESTED | No cases |
| `smfUnix2Date` | MISSING | UNTESTED | No cases |
| `smfUnix2DateStr` | MISSING | UNTESTED | No cases |
| `smfWord` | MISSING | UNTESTED | No cases |
| `smfYahooAPIData` | PASS | NOT_SUPPORTED | 1 NOT_SUPPORTED |
| `smfYahooCodeDesc` | MISSING | UNTESTED | No cases |

## Source groups

| Source | Raw cases | Interpreted cases | Endpoint-content passes | Provider-data/image checks passed |
|---|---|---|---|---|
| bigcharts.marketwatch.com | No cases | No cases | 0 | 0 |
| c.stockcharts.com | No cases | No cases | 0 | 0 |
| climbermel.github.io | No cases | No cases | 0 | 0 |
| core-api.barchart.com | No cases | No cases | 0 | 0 |
| data.moneycentral.msn.com | No cases | No cases | 0 | 0 |
| download.finance.yahoo.com | No cases | No cases | 0 | 0 |
| dvd.netflix.com | No cases | No cases | 0 | 0 |
| finance.google.com | 3 PASS | 3 PASS | 3 | 0 |
| finance.yahoo.com | No cases | No cases | 0 | 0 |
| finviz.com | No cases | No cases | 0 | 0 |
| fred.stlouisfed.org | No cases | No cases | 0 | 0 |
| investing.money.msn.com | No cases | No cases | 0 | 0 |
| ogres-crypt.com | No cases | No cases | 0 | 0 |
| oic.ivolatility.com | No cases | No cases | 0 | 0 |
| performance.morningstar.com | No cases | No cases | 0 | 0 |
| query.yahooapis.com | No cases | No cases | 0 | 0 |
| query1.finance.yahoo.com | 2 PASS | 2 PASS | 0 | 2 |
| query2.finance.yahoo.com | No cases | No cases | 0 | 0 |
| quotes.freerealtime.com | No cases | No cases | 0 | 0 |
| research.investors.com | No cases | No cases | 0 | 0 |
| research.tdameritrade.com | No cases | No cases | 0 | 0 |
| stockcharts.com | No cases | No cases | 0 | 0 |
| www.advfn.com | 1 PASS | 1 PASS | 0 | 1 |
| www.barchart.com | No cases | No cases | 0 | 0 |
| www.etfinvestmentoutlook.com | No cases | No cases | 0 | 0 |
| www.etfscreen.com | No cases | No cases | 0 | 0 |
| www.google.com | No cases | No cases | 0 | 0 |
| www.gurufocus.com | No cases | No cases | 0 | 0 |
| www.marketwatch.com | No cases | No cases | 0 | 0 |
| www.nasdaq.com | No cases | No cases | 0 | 0 |
| www.reuters.com | No cases | No cases | 0 | 0 |
| www.stockscores.com | No cases | No cases | 0 | 0 |
| www.zacks.com | No cases | No cases | 0 | 0 |
| fc.yahoo.com | No cases | No cases | 0 | 0 |
| ca.advfn.com | 2 PASS | 2 PASS | 0 | 2 |
| api.stlouisfed.org | 5 PASS | 5 PASS | 0 | 4 |

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

## Coverage limits

Inventoried public .bas Functions and selected cases only; Subs, Dictionary members, private functions, and every possible argument/provider field are not claimed covered.

19 selected cases enter formulas in a worksheet: 19 PASS. `worksheet_entry` is independent of fixture/live categories; those categories do not replace evidence of actual formula assignment.

`uncovered-call-patterns.csv` records 298 missing primary, worksheet, live-extraction, or unfinished call patterns. `report-data.json` lists each function's exact expressions and all source endpoints.

Provider-unavailable classifications require matching case/run/reset fields and a same-case HTTP failure or a captured direct-response rejection marker. The label describes the requested endpoint during this run; it does not establish a provider-wide outage or prove that the requested URL is correct. HTTP 200 plus an extraction error remains unclassified when the response content was not captured. Last-HTTP diagnostics alone are not causal evidence.

`native-results.private.json` preserves raw native evidence with owner-only file permissions. Shared native/report JSON, Markdown and CSV redact anonymous Yahoo crumbs and cookies; they do not publish those session values.

The report reads the saved workbook only. It does not rerun Excel, retry providers, test a fresh Mac, or infer coverage for unbounded caller URLs.
