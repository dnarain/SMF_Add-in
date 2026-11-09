# SMF Mac function results

This list covers all 121 public functions in the saved November 8, 2026, Excel run of `3.0.2024.07.30-mac.0.5-dev`.
The run completed 239 selected cases. Each result applies to the listed call and inputs.
Passing data checks include cached results and two StockCharts image checks. The checks do not independently establish financial accuracy or freshness.

## Results at a glance

| Function result | Functions |
|---|---:|
| Listed data call passed | 18 |
| Limited | 3 |
| Partial by field | 1 |
| Live call failed | 17 |
| Obsolete | 4 |
| Local or controlled check passed | 77 |
| URL dependent | 1 |

`Limited` identifies a known unsupported mode/default or a company-name-only check.
`Live call failed` takes priority over a passing saved-input parser test. Eight option functions passed local parsing but failed their live requests.
`Local or controlled check passed` covers calculations, parsing, local charts, diagnostics and fixture downloads. It does not establish an external feed.

## Functions that returned data or chart images

The table includes three limited functions and the numbered-field dispatcher. Their limits are stated beside each function.
The generic option functions passed in Yahoo mode. Other provider modes are not covered by those passes.

**Default-field warning:** `smfGetYahooHistory` passed explicit field lists. Omitting `pItems` selects `dohlcvufgxs`, whose `F/G` characters the price-history validator rejects.
This default defect follows from [the declaration](../../../code/smfGetYahooHistory_.bas#L11) and [the validator](../../../code/smfGetYahooHistory_.bas#L180). The default call was not a native case.
`RCHGetYahooHistory` and `RCHGetYahooHistory2` use different defaults.

**Quote-mode warning:** `RCHGetYahooQuotes` passed supplied history-CSV URLs with `pItems=""`.
The source marks its old ticker/quote-code mode [obsolete](../../../code/modGetYahooQuotes.bas#L46).

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `DateOfHigh` | Listed data call passed | Selected AAPL date-of-high checks passed using Yahoo history. | 2 / 0 |
| `RCHCreateComment` | Listed data call passed | Text comments, local charts/images and two StockCharts image calls passed. All 15 selected calls passed. | 15 / 0 |
| `RCHGetElementNumber` | Partial by field | 11 case assertions passed and nine failed. Status depends on the numbered field. See the field table. | 11 / 9 |
| `RCHGetYahooHistory` | Listed data call passed | Selected AAPL daily history with explicit fields returned dated prices and volume. | 2 / 0 |
| `RCHGetYahooHistory2` | Listed data call passed | Selected AAPL/IBM history, date order, adjustment and output-size checks passed. | 6 / 0 |
| `RCHGetYahooQuotes` | Limited | Only a supplied history-CSV URL with pItems="" passed. The source labels the old ticker/quote-code mode obsolete. | 2 / 0 |
| `smfGetGuruFocusData` | Limited | Only item 1, company name, passed. Financial-ratio and historical-metric branches were not established. No GuruFocus login/token integration exists. | 1 / 0 |
| `smfGetOptionExpirations` | Listed data call passed | Yahoo mode returned AAPL call expirations. Four-row and 50-row output checks passed. Other providers were not established. | 4 / 0 |
| `smfGetOptionQuotes` | Listed data call passed | Yahoo mode returned a bid for the selected AAPL call. Other quote fields and provider modes were not established. | 2 / 0 |
| `smfGetOptionStrikes` | Listed data call passed | Yahoo mode returned strikes for the selected AAPL expiry. Other provider modes were not established. | 2 / 0 |
| `smfGetYahooHistory` | Limited | Explicit field lists DC, DOHLUCV and DAUV passed. The omitted daily-history field-list default contains unsupported F/G, by source inspection. | 13 / 0 |
| `smfGetYahooHistoryCSV` | Listed data call passed | Selected AAPL/IBM history-CSV calls returned dated OHLC and volume data. | 2 / 0 |
| `smfGetYahooHistoryCSVData` | Listed data call passed | The supplied Yahoo history URL returned CSV with the expected dates/header. The omitted URL default was not tested. | 2 / 0 |
| `smfGetYahooJSONData` | Listed data call passed | AAPL financialData.targetMeanPrice returned a positive value. Other modules and fields were not established. | 1 / 0 |
| `smfGetYahooJSONField` | Listed data call passed | AAPL price.regularMarketPrice.raw returned a positive value. Other JSON paths were not established. | 1 / 0 |
| `smfGetYahooOptionQuote` | Listed data call passed | The selected AAPL call bid passed. Ask, last price, other quote fields and other contracts need checks. | 2 / 0 |
| `smfGetYahooPortfolioView` | Listed data call passed | AAPL symbol and last-price fields passed. Other portfolio columns were not established. | 2 / 0 |
| `SMFHighBetween` | Listed data call passed | Selected AAPL high, its date and range prices passed using Yahoo history. | 2 / 0 |
| `smfLastPrice` | Listed data call passed | The selected dated AAPL closing price matched the available Yahoo-history reference. | 2 / 0 |
| `SMFLowBetween` | Listed data call passed | Selected AAPL low, its date and range prices passed using Yahoo history. | 2 / 0 |
| `smfPricesBetween` | Listed data call passed | Selected AAPL range summaries passed, including an ending date before the newest cached observation. | 3 / 0 |
| `smfPricesByDates` | Listed data call passed | Selected AAPL closing prices for two dates passed using Yahoo history. | 2 / 0 |

## Functions whose live calls failed

A failed request can reflect the URL, provider access, transport or parsing. The result does not establish a provider-wide outage.
Four option functions that return obsolete text are listed in the next table.

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
| `smfGetYahooOptionQuote1` | Live call failed | Saved-input checks passed; the live legacy Yahoo HTML option request returned HTTP 404. | 1 / 1 |
| `smfGetYahooOptionQuote2` | Live call failed | Saved-input checks passed; the live legacy Yahoo HTML option request returned HTTP 404. | 1 / 1 |
| `smfRSITargetPrices` | Live call failed | The legacy Yahoo quote dependency failed DNS resolution. The history call also inherits the unsupported F/G defaults. | 0 / 1 |
| `smfYahooAPIData` | Live call failed | The legacy YQL hostname query.yahooapis.com failed DNS resolution; no data returned. | 0 / 1 |

## Functions that return obsolete text

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `smfGetMSNOptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |
| `smfGetOX2OptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |
| `smfGetOX3OptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |
| `smfGetOXOptionQuote` | Obsolete | Returns obsolete-status text before making a provider request. | 0 / 1 |

## Local calculations and controlled checks

All selected direct checks in this section passed. Some helpers also run indirectly inside the data functions.
The UTC conversion cases used the `America/Los_Angeles` timezone.
Functions such as `SMFTech` and `smfInCellChart` operate on supplied data and do not require a financial-data provider.

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `ConvertToIso` | Local or controlled check passed | Selected check: 2024-01-15T20:00:00.000Z; requires the verified America/Los_Angeles host timezone | 1 / 0 |
| `ConvertToJson` | Local or controlled check passed | Selected check: [7,false,null] | 1 / 0 |
| `ConvertToUtc` | Local or controlled check passed | Selected check: 2024-01-15 20:00:00 UTC; requires the verified America/Los_Angeles host timezone | 1 / 0 |
| `DecodeJsonString` | Local or controlled check passed | Selected check: Collection with two Dictionary items named first and second | 1 / 0 |
| `fnWait` | Local or controlled check passed | Selected check: Empty return after at least0.9seconds | 1 / 0 |
| `GetKeys` | Local or controlled check passed | Selected check: Native array indexed 0..1 containing 0 and 01 | 1 / 0 |
| `GetObjectProperty` | Local or controlled check passed | Selected check: Collection containing 3 and 9 in that order | 1 / 0 |
| `GetProperty` | Local or controlled check passed | Selected check: Boolean False | 1 / 0 |
| `ParseIso` | Local or controlled check passed | Selected check: 2024-02-29 00:00:00 | 1 / 0 |
| `ParseJson` | Local or controlled check passed | Selected check: Dictionary with n=7, Boolean False, Null, and identifier text 1234567890123456789 | 1 / 0 |
| `ParseUtc` | Local or controlled check passed | Selected check: 2024-01-15 12:00:00 PST; requires the verified America/Los_Angeles host timezone | 1 / 0 |
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
| `smfJoin` | Local or controlled check passed | Selected check: 1,2,3,4,5 | 1 / 0 |
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
| `smfUnix2Date` | Local or controlled check passed | Selected check: 1970-01-02 00:00:00 as Date | 1 / 0 |
| `smfUnix2DateStr` | Local or controlled check passed | Selected check: 1970-01-02 01:01:01 | 1 / 0 |
| `smfWord` | Local or controlled check passed | Selected check: gamma | 1 / 0 |
| `smfYahooCodeDesc` | Local or controlled check passed | The local l1 code lookup returned Last Trade. It does not retrieve a quote. | 1 / 0 |

## Download helper

| Function | Result | Scope of the result | Cases passed / failed |
|---|---|---|---|
| `smfMacFetch` | URL dependent | Controlled fetch passed. Of 30 external page-content checks, eight passed and 22 failed. Returned HTTP errors do not establish a downloader defect. | 9 / 22 |

## Numbered fields for RCHGetElementNumber

Zacks, Stockscores, Finviz and the GuruFocus dividend field appear here because the caller uses `RCHGetElementNumber`.
The 20 selected cases cover 19 unique ticker/field combinations. Finviz price field 4074 was checked twice.
Each other numbered field still needs its own check.

| Field | Ticker or identifier | Data requested | Result | Qualification |
|---|---|---|---|
| 543 | `AAPL` | FreeRealTime StockScouter rating | Passed selected check |  |
| 702 | `AAPL` | Barchart overall opinion | Failed selected check |  |
| 785 | `AAPL` | Stockscores 20-day average volume | Passed selected check | The data-field case passed even though a separate page-marker check failed. |
| 849 | `AAPL` | Zacks analyst target price | Passed selected check | The data-field case passed even though a separate page-marker check failed. |
| 980 | `AAPL` | Yahoo 10-day average volume | Passed selected check | Positive-number assertion only. Units and accuracy were not independently checked. |
| 995 | `SPY` | ETF Investment Outlook new-high percentage | Failed selected check |  |
| 1084 | `SPY` | ETFScreen relative-strength rank | Failed selected check |  |
| 1116 | `NYSE` | Barchart NYSE new-high count | Failed selected check |  |
| 2003 | `AAPL` | IBD industry-group rank | Failed selected check |  |
| 2727 | `AAPL` | GuruFocus forward dividend yield | Passed selected check | Public dividend page. Does not establish authenticated GuruFocus API access. |
| 2747 | `NASDAQ:AAPL` | Google quarterly revenue | Failed selected check |  |
| 4072 | `AAPL` | Finviz earnings date | Passed selected check |  |
| 4074 | `AAPL` | Finviz share price | Passed selected check | Two cases: cached result and a later fresh-request result. |
| 5196 | `NASDAQ:AAPL` | ADVFN annual statement-ending date | Passed selected check | This recipe passed while smfGetADVFNElement failed through a different route. |
| 13626 | `AAPL.O` | Reuters trailing P/E | Failed selected check |  |
| 13865 | `AAPL` | Yahoo sector | Passed selected check |  |
| 19004 | `70111470` | Netflix movie release year | Passed selected check | Nonfinancial, year-only assertion. Movie identity was not verified. |
| 19902 | `AAPL` | TD Ameritrade ex-dividend date | Failed selected check |  |
| 20000 | `AAPL` | Legacy MSN StockScouter rating | Failed selected check |  |

## Scope and evidence

All 121 function names appear once in the five function tables. The CSV lists each function and its exact case IDs.
The function categories use every selected case, not only the primary case.
No new provider requests or native Excel runs were made to produce this list.
New Mac installation, most argument combinations, most numbered fields and Windows runtime behavior remain untested.
VBA Subs and Dictionary methods are outside the 121-function inventory.

- [Function CSV](function-status.csv)
- [All saved case results](all-case-results.csv)
- [Saved report data](report-data.json)
- [Full results and remaining gaps](README.md)
