# SMF Mac 0.6 function results

This table records the final four-module 0.6 candidate. Nine separate fresh-call reuse cases also passed.
The [final report](README.md) records stable cold reopens and the preserved 10/11 data/picture example comparison.

This reference covers all 121 public functions in the completed 0.6 native run.
The run recorded 292 selected cases: 238 PASS and 54 CHECK_FAILED. Each result applies to its selected inputs.
All 32 worksheet-entry cases passed. Most worksheet entry paths and argument combinations remain untested.

## Results at a glance

| Function result | Functions |
|---|---:|
| Listed data call passed | 17 |
| Limited | 2 |
| Partial by field | 1 |
| Live call failed | 19 |
| Obsolete | 4 |
| Local or controlled check passed | 77 |
| URL dependent | 1 |

The categories use every case, including failed live calls after passing controlled parser checks.
The five new live option comparisons accepted matching numeric zero. They do not establish fresh positive trading quotes.
Passing data/image assertions can use cached responses. Financial accuracy and freshness need independent comparisons.

## Default fields and Yahoo compatibility

Mac `smfGetYahooHistory` uses `dohlcvu` for omitted and empty field lists. All 33 selected history cases passed.
Mac `smfGetYahooOptionQuote1` uses the existing Yahoo JSON function. It preserves fractional percentages and formatted expiration strings.
Its 34 passing cases include four live numeric-zero comparisons and a matching expiration string. Its separate positive-bid case failed.
`RCHGetYahooQuotes` still passed CSV-URL mode only. Its old ticker/quote-code mode remains obsolete.
`smfGetGuruFocusData` passed company-name item 1 only. Historical ratios and protected fields remain unverified.

## Functions with selected data/image results

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `DateOfHigh` | Listed data call passed | Selected AAPL date-of-high checks passed using Yahoo history. | 2 / 0 |
| `RCHCreateComment` | Listed data call passed | All 15 selected default-suite cases passed. Nine separate fresh-call numeric-reuse cases also passed for this candidate. The README records the separate cold-add-in reopen result. | 15 / 0 |
| `RCHGetElementNumber` | Partial by field | 11 case assertions passed and 9 failed. Results depend on the numbered field. The field table gives each result. | 11 / 9 |
| `RCHGetYahooHistory` | Listed data call passed | Selected AAPL daily history with explicit fields returned dated prices and volume. | 2 / 0 |
| `RCHGetYahooHistory2` | Listed data call passed | Selected AAPL/IBM history, date order, adjustment and output-size checks passed. | 6 / 0 |
| `RCHGetYahooQuotes` | Limited | Only a supplied history-CSV URL with pItems="" passed. The source labels the old ticker/quote-code mode obsolete. | 2 / 0 |
| `smfGetGuruFocusData` | Limited | Only item 1, company name, passed. Financial-ratio and historical-metric branches were not established. No GuruFocus login/token integration exists. | 1 / 0 |
| `smfGetOptionExpirations` | Listed data call passed | Yahoo mode returned AAPL call expirations. Four-row and 50-row output checks passed. Other providers were not established. | 4 / 0 |
| `smfGetOptionStrikes` | Listed data call passed | Yahoo mode returned strikes for the selected AAPL expiry. Other provider modes were not established. | 2 / 0 |
| `smfGetYahooHistory` | Listed data call passed | All 33 selected cases passed. On Mac, omitted/empty pItems use dohlcvu: Date/Open/High/Low/AdjClose/Volume/Close. Explicit lists, dimensions and four genuine array-entered worksheet cases passed. Windows defaults remain unchanged. | 33 / 0 |
| `smfGetYahooHistoryCSV` | Listed data call passed | Selected AAPL/IBM history-CSV calls returned dated OHLC and volume data. | 2 / 0 |
| `smfGetYahooHistoryCSVData` | Listed data call passed | The supplied Yahoo history URL returned CSV with the expected dates/header. The omitted URL default was not tested. | 2 / 0 |
| `smfGetYahooJSONData` | Listed data call passed | AAPL financialData.targetMeanPrice returned a positive value. Other modules and fields were not established. | 1 / 0 |
| `smfGetYahooJSONField` | Listed data call passed | AAPL price.regularMarketPrice.raw returned a positive value. Other JSON paths were not established. | 1 / 0 |
| `smfGetYahooPortfolioView` | Listed data call passed | AAPL symbol and last-price fields passed. Other portfolio columns were not established. | 2 / 0 |
| `SMFHighBetween` | Listed data call passed | Selected AAPL high, its date and range prices passed using Yahoo history. | 2 / 0 |
| `smfLastPrice` | Listed data call passed | The selected dated AAPL closing price matched the available Yahoo-history reference. | 2 / 0 |
| `SMFLowBetween` | Listed data call passed | Selected AAPL low, its date and range prices passed using Yahoo history. | 2 / 0 |
| `smfPricesBetween` | Listed data call passed | Selected AAPL range summaries passed, including an ending date before the newest cached observation. | 3 / 0 |
| `smfPricesByDates` | Listed data call passed | Selected AAPL closing prices for two dates passed using Yahoo history. | 2 / 0 |

## Functions whose live calls failed

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `smfGet888OptionQuote` | Live call failed | Saved-input checks passed; the live OIC/iVolatility option request returned HTTP 404. | 1 / 1 |
| `smfGetADVFNElement` | Live call failed | The selected annual Year End Date lookup returned Error after HTTP 200. Numbered field 5196 passed separately. | 0 / 1 |
| `smfGetBarChartOptionQuote` | Live call failed | Saved-input checks passed; the live Barchart API request timed out. | 1 / 1 |
| `smfGetBigChartsOptionQuote` | Live call failed | Saved-input checks passed; the live BigCharts option request returned HTTP 403. | 1 / 1 |
| `smfGetEconData` | Live call failed | The selected FRED DGS10 request timed out. No positive series result was recorded. | 0 / 1 |
| `smfGetGoogleHistoryCSV` | Live call failed | The constructed URL contains unescaped date spaces. curl rejected the URL before HTTP. | 0 / 1 |
| `smfGetGoogleOptionQuote` | Live call failed | Saved-input checks passed; the live Google request returned HTTP 200 but no usable option quote. | 1 / 1 |
| `smfGetGuruFocusCSVItem` | Live call failed | The selected revenue CSV URL returned HTTP 404; the function returned no revenue value. | 0 / 1 |
| `smfGetGuruFocusItem` | Live call failed | The selected AAPL TTM financial item returned Error after HTTP 200. Old page selectors need review. | 0 / 1 |
| `smfGetMorningstarHistoryCSV` | Live call failed | The selected request timed out connecting to the configured Morningstar history endpoint. | 0 / 1 |
| `smfGetMSNHistory` | Live call failed | The configured legacy MSN hostname failed DNS resolution; no history data returned. | 0 / 1 |
| `smfGetMWOptionQuote` | Live call failed | Saved-input checks passed; the live MarketWatch option request returned HTTP 401. | 1 / 1 |
| `smfGetNASDAQOptionQuote` | Live call failed | Saved-input checks passed; the live Nasdaq option request returned HTTP 403. | 1 / 1 |
| `smfGetOptionQuotes` | Live call failed | The local contract-symbol/strike check passed. The live Yahoo bid returned numeric 0 after HTTP 200 and failed its positive-bid assertion. Other provider modes remain unverified. | 1 / 1 |
| `smfGetYahooOptionQuote` | Live call failed | The controlled bid-parser check passed. The live AAPL call bid returned numeric 0 after HTTP 200 and failed its positive-bid assertion. Other quote fields and contracts remain unverified. | 1 / 1 |
| `smfGetYahooOptionQuote1` | Live call failed | Mac now uses Yahoo JSON with the same five arguments. The 29 controlled checks and five live same-response comparisons passed. Live bid/ask/put/percent values were 0. Expiry was the string 2026-11-20. The separate positive-bid check failed on 0 after HTTP 200. A fresh usable positive quote was not established. | 34 / 1 |
| `smfGetYahooOptionQuote2` | Live call failed | Saved-input checks passed; the live legacy Yahoo HTML option request returned HTTP 404. | 1 / 1 |
| `smfRSITargetPrices` | Live call failed | No usable target returned. The Mac history-default repair removes the inherited F/G defect. The old quote-CSV dependency remains. This case's final recorded request was Yahoo history HTTP 200. That status does not isolate the cause. | 0 / 1 |
| `smfYahooAPIData` | Live call failed | The legacy YQL hostname query.yahooapis.com failed DNS resolution; no data returned. | 0 / 1 |

## Functions that return obsolete text

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `smfGetMSNOptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |
| `smfGetOX2OptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |
| `smfGetOX3OptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |
| `smfGetOXOptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |

## Local calculations and controlled checks

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `ConvertToIso` | Local or controlled check passed | Selected check: 2024-01-15T20: 00: 00.000Z; requires the verified America/Los_Angeles host timezone | 1 / 0 |
| `ConvertToJson` | Local or controlled check passed | Selected check: [7,false,null] | 1 / 0 |
| `ConvertToUtc` | Local or controlled check passed | Selected check: 2024-01-15 20: 00: 00 UTC; requires the verified America/Los_Angeles host timezone | 1 / 0 |
| `DecodeJsonString` | Local or controlled check passed | Selected check: Collection with two Dictionary items named first and second | 1 / 0 |
| `fnWait` | Local or controlled check passed | Selected check: Empty return after at least0.9seconds | 1 / 0 |
| `GetKeys` | Local or controlled check passed | Selected check: Native array indexed 0..1 containing 0 and 01 | 1 / 0 |
| `GetObjectProperty` | Local or controlled check passed | Selected check: Collection containing 3 and 9 in that order | 1 / 0 |
| `GetProperty` | Local or controlled check passed | Selected check: Boolean False | 1 / 0 |
| `ParseIso` | Local or controlled check passed | Selected check: 2024-02-29 00: 00: 00 | 1 / 0 |
| `ParseJson` | Local or controlled check passed | Selected check: Dictionary with n=7, Boolean False, Null, and identifier text 1234567890123456789 | 1 / 0 |
| `ParseUtc` | Local or controlled check passed | Selected check: 2024-01-15 12: 00: 00 PST; requires the verified America/Los_Angeles host timezone | 1 / 0 |
| `RCHExtractData` | Local or controlled check passed | Selected check: 42 as a numeric result. | 1 / 0 |
| `RCHGetHTMLTable` | Local or controlled check passed | Selected check: The 2-by-2 table Label/42 and Earnings/Nov 20 AMC. | 1 / 0 |
| `RCHGetTableCell` | Local or controlled check passed | Selected check: 42 as a numeric result. | 1 / 0 |
| `RCHGetURLData` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response | 1 / 0 |
| `RCHGetURLData1` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response | 1 / 0 |
| `RCHGetURLData1Mac` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response | 1 / 0 |
| `RCHGetURLData2` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response | 1 / 0 |
| `RCHGetURLData3` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response | 1 / 0 |
| `RCHGetURLData4` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response | 1 / 0 |
| `RCHGetWebData` | Local or controlled check passed | Selected check: Label | 1 / 0 |
| `RCHImageTable` | Local or controlled check passed | Generated HTML image-table rows for two tickers. This function did not download or display the images. | 1 / 0 |
| `smfBuildYahooTimeSeriesURL` | Local or controlled check passed | Built the selected Yahoo fundamentals URL. No provider data request was made by this case. | 1 / 0 |
| `smfCDec` | Local or controlled check passed | Selected check: Double0.123456789 | 1 / 0 |
| `smfChkJSONFieldType` | Local or controlled check passed | Selected check: Numeric 1 | 1 / 0 |
| `smfConvertData` | Local or controlled check passed | Selected check: Double0.125 | 1 / 0 |
| `smfConvertYahooTicker` | Local or controlled check passed | Selected check: CA:CMG | 1 / 0 |
| `smfDate2Unix` | Local or controlled check passed | Selected check: 86400 | 1 / 0 |
| `smfEval` | Local or controlled check passed | Selected check: 7 | 1 / 0 |
| `smfEvaluateTwice` | Local or controlled check passed | Selected check: 3 | 1 / 0 |
| `smfGetAConstants` | Local or controlled check passed | Selected check: suite-fixed-constant | 1 / 0 |
| `smfGetAData` | Local or controlled check passed | Selected check: 32767 x characters | 1 / 0 |
| `smfGetaFiles` | Local or controlled check passed | Selected check: A source-file record naming smf-elements-4.txt. | 1 / 0 |
| `smfGetAParms` | Local or controlled check passed | Selected check: The initialized FinViz definition containing the Earnings search marker. | 1 / 0 |
| `smfGetCSVFile` | Local or controlled check passed | Selected check: Date/Close/Volume headers, then 100.5/1000 and 102.25/1200. | 1 / 0 |
| `smfGetOptionExpiry` | Local or controlled check passed | The local November 2026 monthly-expiration calculation returned November 20. This function does not fetch an option chain. | 1 / 0 |
| `smfGetTagContent` | Local or controlled check passed | Selected check: Nov 20 AMC | 1 / 0 |
| `smfGetWebPage` | Local or controlled check passed | Selected check: Exact fixture HTML stored under mode0 URL key | 1 / 0 |
| `smfHTMLDecode` | Local or controlled check passed | Selected check: <b>"A&B"</b> | 1 / 0 |
| `smfInCellChart` | Local or controlled check passed | Line mode created four worksheet shapes from local sample data. Other chart modes were not established. | 1 / 0 |
| `smfJoin` | Local or controlled check passed | Selected check: 1, 2, 3, 4, 5 | 1 / 0 |
| `smfJSONExtractField` | Local or controlled check passed | Selected check: Numeric 7 from the second array item | 1 / 0 |
| `smfJSONExtractField_x64` | Local or controlled check passed | Selected check: Boolean False from the first array item | 1 / 0 |
| `smfJSONExtractKeys` | Local or controlled check passed | Selected check: One-column array with rows 0, 1, and #N/A | 1 / 0 |
| `smfLogInternetCalls` | Local or controlled check passed | Selected check: Logging off; sLog=N | 1 / 0 |
| `SMFMacElementsSelfTest` | Local or controlled check passed | Selected check: 3 element parsing checks. | 1 / 0 |
| `SMFMacEmptyRecipeSelfTest` | Local or controlled check passed | Selected check: 5 checks: empty values and empty error fallbacks in both Mac formula branches, plus unchanged HTTP state; private element state restored. | 1 / 0 |
| `SMFMacHistoryCSVSelfTest` | Local or controlled check passed | Selected check: 12 assertions; no live requests. | 1 / 0 |
| `smfMacHTMLSelfTest` | Local or controlled check passed | Selected check: 55 HTML checks | 1 / 0 |
| `SMFMacHTTPStatus` | Local or controlled check passed | Selected check: 207 | 1 / 0 |
| `SMFMacIntegrationSelfTest` | Local or controlled check passed | Selected check: 25 integration checks | 1 / 0 |
| `SMFMacJSONSelfTest` | Local or controlled check passed | Selected check: PASS: 84 JSON checks | 1 / 0 |
| `SMFMacLastError` | Local or controlled check passed | Selected check: HTTP 403 | 1 / 0 |
| `SMFMacLastFileError` | Local or controlled check passed | Selected check: empty string | 1 / 0 |
| `SMFMacLastURL` | Local or controlled check passed | Selected check: fixture://recorded-url | 1 / 0 |
| `SMFMacNestedRecipeSelfTest` | Local or controlled check passed | Selected check: 50 assertions; no live requests. | 1 / 0 |
| `smfMacReadElementLocation` | Local or controlled check passed | Selected check: Local followed by one LF | 1 / 0 |
| `smfMacReadElementText` | Local or controlled check passed | Selected check: 19998;Fixture;Value;=1+2 followed by one LF | 1 / 0 |
| `smfMacReadResponse` | Local or controlled check passed | Selected check: A<LF>B; status200; no transport error | 1 / 0 |
| `SMFMacRecipeSelfTest` | Local or controlled check passed | Selected check: 20 zero-HTTP recipe checks | 1 / 0 |
| `smfMacStripHTML` | Local or controlled check passed | Selected check: 42&7 | 1 / 0 |
| `SMFMacTableGuardSelfTest` | Local or controlled check passed | Selected check: 17 table guard checks | 1 / 0 |
| `SMFMacTableRecipeSelfTest` | Local or controlled check passed | Selected check: 40 table-recipe checks | 1 / 0 |
| `SMFMacTransportSelfTest` | Local or controlled check passed | Selected check: 6 deterministic transport checks | 1 / 0 |
| `smfMacTryNestedRecipe` | Local or controlled check passed | Selected check: True, Double42 in the output argument, and unchanged HTTP state. | 2 / 0 |
| `smfMacTryTableCellRecipe` | Local or controlled check passed | Selected check: True with numeric parsed value42; no HTTP request | 1 / 0 |
| `smfMacTryTagContentRecipe` | Local or controlled check passed | Selected check: True; parsed result numeric42 | 1 / 0 |
| `smfMacYahooChartCSV` | Local or controlled check passed | Selected check: Date,Open,High,Low,Close,Adj Close,Volume; Nov2 null open/adjusted/volume remain null, not0. | 1 / 0 |
| `smfMacYahooDownloadCSV` | Local or controlled check passed | Direct reversed-date and cached-CSV checks passed. Passed live CSV wrapper calls also used this helper indirectly. | 2 / 0 |
| `smfMacYahooDownloadURL` | Local or controlled check passed | Selected check: True for the exact Yahoo host; unchanged HTTP state. | 2 / 0 |
| `smfStrExtr` | Local or controlled check passed | Selected check: Double123.45 | 1 / 0 |
| `smfStripHTML` | Local or controlled check passed | Selected check: 7&8 | 1 / 0 |
| `SMFTech` | Local or controlled check passed | Ten local indicator cases passed: SMA, EMA, ATR, CCI, MACD, OBV, ROC, RSI, ADL and STO. It needs supplied price data. | 10 / 0 |
| `smfUnix2Date` | Local or controlled check passed | Selected check: 1970-01-02 00: 00: 00 as Date | 1 / 0 |
| `smfUnix2DateStr` | Local or controlled check passed | Selected check: 1970-01-02 01: 01: 01 | 1 / 0 |
| `smfWord` | Local or controlled check passed | Selected check: gamma | 1 / 0 |
| `smfYahooCodeDesc` | Local or controlled check passed | The local l1 code lookup returned Last Trade. It does not retrieve a quote. | 1 / 0 |

## Download helper

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `smfMacFetch` | URL dependent | The controlled fetch passed. Of 30 external page-content checks, 8 passed and 22 failed. Returned HTTP errors do not establish a downloader defect. | 9 / 22 |

## Numbered fields for RCHGetElementNumber

The 20 selected cases cover 19 unique ticker/field combinations. Field 4074 has two cases.
These outcomes apply to the selected assertions. Other numbered fields require separate checks.

| Field | Ticker or identifier | Data requested | Result | Qualification |
|---|---|---|---|---|
| 543 | `AAPL` | FreeRealTime StockScouter rating | Passed selected check |  |
| 702 | `AAPL` | Barchart overall opinion | Failed selected check |  |
| 785 | `AAPL` | Stockscores 20-day average volume | Passed selected check | The data-field and provider page-marker checks are separate assertions. |
| 849 | `AAPL` | Zacks analyst target price | Passed selected check | The data-field and provider page-marker checks are separate assertions. |
| 980 | `AAPL` | Yahoo 10-day average volume | Passed selected check | Positive-number assertion only. Units and accuracy were not independently checked. |
| 995 | `SPY` | ETF Investment Outlook new-high percentage | Failed selected check |  |
| 1084 | `SPY` | ETFScreen relative-strength rank | Failed selected check |  |
| 1116 | `NYSE` | Barchart NYSE new-high count | Failed selected check |  |
| 2003 | `AAPL` | IBD industry-group rank | Failed selected check |  |
| 2727 | `AAPL` | GuruFocus forward dividend yield | Passed selected check | Public dividend page. Does not establish authenticated GuruFocus API access. |
| 2747 | `NASDAQ:AAPL` | Google quarterly revenue | Failed selected check |  |
| 4072 | `AAPL` | Finviz earnings date | Passed selected check |  |
| 4074 | `AAPL` | Finviz share price | Passed selected check | 2 cases. 1 recorded case-bound HTTP request. |
| 5196 | `NASDAQ:AAPL` | ADVFN annual statement-ending date | Passed selected check | This recipe passed its valid-date assertion. The final lookup returned 2026-09-26. The cold example retained 2025-09-27 and failed that comparison. |
| 13626 | `AAPL.O` | Reuters trailing P/E | Failed selected check |  |
| 13865 | `AAPL` | Yahoo sector | Passed selected check |  |
| 19004 | `70111470` | Netflix movie release year | Passed selected check | Nonfinancial, year-only assertion. Movie identity was not verified. |
| 19902 | `AAPL` | TD Ameritrade ex-dividend date | Failed selected check |  |
| 20000 | `AAPL` | Legacy MSN StockScouter rating | Failed selected check |  |

## Scope and evidence

All 121 function names appear once in the five function tables. The CSV retains every exact case ID.
The 0.5 reference remains unchanged. This reference uses the 0.6 saved results and adds no provider requests.
The run used the same Mac, Excel installation and helper. Another Mac and a fresh macOS account remain untested.
Public Subs, private functions, Dictionary methods, most fields and Windows runtime behavior remain outside this inventory.

- [Function CSV](function-status.csv)
- [All saved case results](all-case-results.csv)
- [Saved report data](report-data.json)
- [Full results](README.md)
