# SMF Mac 0.7-dev function results

This reference uses saved run `smf-0.7-google-full-run4` and exact version `3.0.2024.07.30-mac.0.7-dev`.
The completed run covers 121 public functions and 538 selected cases.
Raw case results: 36 CHECK_FAILED, 502 PASS.
51 cases enter a worksheet formula. Their saved results are `{"PASS": 51}`.

Every function appears once in the function tables. The CSV retains every exact case ID.
Provider assertions can accept cached values or numeric zero. They do not establish financial accuracy or fresh positive quotes.
Controlled fixtures establish the specified local behavior. They do not establish current provider access.
Not Supported means a registered retired mode returned its expected guard result. It establishes no financial feed.
The report records 63 passing provider-data/image assertions, including cached values.
Measured fetchURL bridge attempts and case-bound HTTP observations are separate counts.
A recorded zero bridge delta means no text-fetch bridge attempt. Image downloads use a separate handler.

## Results at a glance

| Function result | Functions |
|---|---:|
| Listed data call passed | 23 |
| Limited | 3 |
| Partial by field | 1 |
| Live call failed | 10 |
| Not Supported | 7 |
| Local or controlled check passed | 76 |
| URL dependent | 1 |

## Selected data/image results

| Function | Result | Scope and limits | Remaining work | Passed / failed |
|---|---|---|---|---|
| `DateOfHigh` | Listed data call passed | 2 selected cases passed and 0 failed. A real date in the latest completed trading week, 2026-11-02 through 2026-11-06. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `RCHCreateComment` | Listed data call passed | 15 selected cases passed and 0 failed. Selected note, picture and numeric-chart cases are listed in the CSV. Cold reopen and chart visibility need their separate receipts. This report alone does not establish that lifecycle. | Other inputs and worksheet paths remain outside these selected cases. | 15 / 0 |
| `RCHGetElementNumber` | Partial by field | 14 selected cases passed and 9 failed. Results depend on the numbered field. 3 retired-mode guard assertions passed; they establish no financial feed. The field table counts only literal financial ticker/field inputs. Bulk definition assertions and metadata checks are reported separately. Most financial values remain untested. | Repair failing numbered-field recipes individually. Add cases for other definitions. | 14 / 9 |
| `RCHGetYahooHistory` | Listed data call passed | 2 selected cases passed and 0 failed. A Date header and a dated positive adjusted close within the requested week, bounded to 6 rows and 2 columns. AAPL 2026-11-06 first data row; exact snapshot prices/volume within1e-7; 5 data rows. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `RCHGetYahooHistory2` | Listed data call passed | 6 selected cases passed and 0 failed. AAPL history with a Date header and a positive close dated within 2026-11-02 through 2026-11-06. Old MMM data is not a pass. IBM 2026-11-06 first data row; exact snapshot prices/volume within1e-7; 5 data rows. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 6 / 0 |
| `RCHGetYahooQuotes` | Limited | 6 selected cases passed and 0 failed. Supplied CSV-URL calls with empty pItems are checked separately. 3 retired-mode guard assertions passed; they establish no financial feed. The old ticker/quote-code mode and registered retired CSV URLs return #N/A on Mac when the guard cases pass. The retained CSV parser does not restore the retired quote feed. | Keep the established mode explicit. Add separate cases before claiming other fields or modes. | 6 / 0 |
| `smfGetADVFNElement` | Listed data call passed | 33 of 33 controlled cases passed. 4 of 4 selected live cases passed. Mac repair covers the latest annual Year End Date only. Earlier periods and other metrics retain legacy code. Exchange-qualified symbols specify the security. | Earlier annual periods, quarterly periods and other ADVFN metrics need repairs for those providers. | 37 / 0 |
| `smfGetEconData` | Listed data call passed | 61 of 61 controlled cases passed. 6 of 6 selected live cases passed. Mac calls the official FRED API with a local key and file access. Exact dates, missing observations and metadata have separate cases. Other series remain untested. | Other series, dates and metadata need separate cases. Another Mac account and key-file access remain separate installation checks. | 67 / 0 |
| `smfGetGuruFocusData` | Limited | 1 selected cases passed and 0 failed. Selected calls request item 1, company name. Historical ratios and protected fields remain unverified. No GuruFocus login/token integration is established. | Keep the established mode explicit. Add separate cases before claiming other fields or modes. | 1 / 0 |
| `smfGetOptionExpirations` | Listed data call passed | 9 selected cases passed and 0 failed. 3 retired-mode guard assertions passed; they establish no financial feed. Exactly4 formatted AAPL $350.00 Call labels; no raw date rows; HTTP200. At least one available AAPL $350 Call expiration and HTTP200; the function has no expiration-date filter Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 9 / 0 |
| `smfGetOptionQuotes` | Listed data call passed | 9 of 9 controlled cases passed. 1 of 1 selected live cases passed. 7 retired-mode guard assertions passed; they establish no financial feed. The selected Yahoo mode and contracts do not establish other provider modes or fields. Valid zero remains distinct from missing or null. | Other inputs and worksheet paths remain outside these selected cases. | 10 / 0 |
| `smfGetOptionStrikes` | Listed data call passed | 7 selected cases passed and 0 failed. 3 retired-mode guard assertions passed; they establish no financial feed. A positive numeric strike in the returned array and HTTP200; the function has no strike-price input Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 7 / 0 |
| `smfGetWebPage` | Limited | Selected check: Exact fixture HTML stored under mode0 URL key 1 retired-mode guard assertions passed; they establish no financial feed. Other inputs, fields and worksheet calls require separate cases. | Keep the established mode explicit. Add separate cases before claiming other fields or modes. | 2 / 0 |
| `smfGetYahooHistory` | Listed data call passed | 33 selected cases passed and 0 failed. Mac omitted/empty fields use dohlcvu. The field order is Date/Open/High/Low/AdjClose/Volume/Close. Selected dimensions and worksheet cases do not cover every argument combination. | Other inputs and worksheet paths remain outside these selected cases. | 33 / 0 |
| `smfGetYahooHistoryCSV` | Listed data call passed | 2 selected cases passed and 0 failed. A Date header, an in-week observation, and a positive Close in a 6-by-7 CSV result. IBM 2026-11-02 first data row; exact snapshot prices/volume within1e-7; 5 data rows. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfGetYahooHistoryCSVData` | Listed data call passed | 2 selected cases passed and 0 failed. Raw Yahoo CSV beginning Date, and containing both requested boundary dates. HTML/JSON/error responses do not pass. HTTP200; standard CSV header and terminal LF; Nov2‚Äì6; first close338.32000732421875, adjusted same, volume64757700. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfGetYahooJSONData` | Listed data call passed | 1 selected cases passed and 0 failed. A positive numeric targetMeanPrice; Not found and Error are not passes. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetYahooJSONField` | Listed data call passed | 1 selected cases passed and 0 failed. A positive numeric regularMarketPrice; Not Found, an authorization response, or a JSON parser failure is not a pass. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetYahooOptionQuote` | Listed data call passed | 1 of 1 controlled cases passed. 1 of 1 selected live cases passed. The selected Yahoo mode and contracts do not establish other provider modes or fields. Valid zero remains distinct from missing or null. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfGetYahooOptionQuote1` | Listed data call passed | 29 of 29 controlled cases passed. 6 of 6 selected live cases passed. Mac delegates to Yahoo JSON. Percentages remain fractions and expiration remains formatted text. Equality with a zero response does not establish a usable positive quote. | Other inputs and worksheet paths remain outside these selected cases. | 35 / 0 |
| `smfGetYahooPortfolioView` | Listed data call passed | 2 selected cases passed and 0 failed. Symbol/Last Price headers, AAPL in the data row, and a positive price; -- placeholders do not pass. Headers Symbol/Last Price and one AAPL row with a positive share price. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `SMFHighBetween` | Listed data call passed | 2 selected cases passed and 0 failed. A positive high, its date within the requested week, and positive start/end prices in the 1-by-4 result. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfLastPrice` | Listed data call passed | 2 selected cases passed and 0 failed. A positive adjusted close equal to an available AAPL history close dated 2026-11-06. Missing reference data makes this comparison unverified, never passed. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `SMFLowBetween` | Listed data call passed | 2 selected cases passed and 0 failed. A positive low below the sentinel 999999, its date within the week, and positive start/end prices. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfPricesBetween` | Listed data call passed | 3 selected cases passed and 0 failed. Open date 2026-11-02, close date 2026-11-06, positive open/close/volume, and high at least low. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 3 / 0 |
| `smfPricesByDates` | Listed data call passed | 2 selected cases passed and 0 failed. A two-element positive-price vector; its 2026-11-06 element equals available history for that date. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfRSITargetPrices` | Listed data call passed | 62 of 62 controlled cases passed. 3 of 3 selected live cases passed. Mac uses 21 prior closes for Wilder RSI(2), then appends the quote once by exchange session. Flat input uses RSI 50 with unavailable targets. Optional bid/ask can remain unavailable. | Other tickers and market sessions need separate comparisons. Live range checks do not independently verify each target. | 65 / 0 |

## Failed selected calls

| Function | Result | Scope and limits | Remaining work | Passed / failed |
|---|---|---|---|---|
| `smfGetBarChartOptionQuote` | Live call failed | 1 selected cases passed and 1 failed. Live_smfGetBarChartOptionQuote: CHECK_FAILED, String: ; length=0 [PROVIDER_UNAVAILABLE] Other inputs, fields and worksheet calls require separate cases. | Inspect the failed assertions and HTTP evidence from that case. Establish permitted access before changing provider extraction. | 1 / 1 |
| `smfGetBigChartsOptionQuote` | Live call failed | 1 selected cases passed and 1 failed. Live_smfGetBigChartsOptionQuote: CHECK_FAILED, HTTP 403, String: Error; length=5 [PROVIDER_UNAVAILABLE] Other inputs, fields and worksheet calls require separate cases. | Inspect the failed assertions and HTTP evidence from that case. Establish permitted access before changing provider extraction. | 1 / 1 |
| `smfGetGoogleHistoryCSV` | Live call failed | 8 selected cases passed and 1 failed. smfGetGoogleHistoryCSV: CHECK_FAILED, HTTP 200, Array <!doctype html><html lang="en-US" dir="ltr"><head><base href="https://www.google.com/finance/"><link \| var n=typeof Object.definePrope [UNCLASSIFIED_FAIL] Other inputs, fields and worksheet calls require separate cases. | Establish an available historical-data source or endpoint; the encoded legacy URL now returns Google Finance HTML instead of CSV. | 8 / 1 |
| `smfGetGuruFocusCSVItem` | Live call failed | 0 selected cases passed and 1 failed. smfGetGuruFocusCSVItem: CHECK_FAILED, HTTP 404, String: ; length=0 [PROVIDER_UNAVAILABLE] Other inputs, fields and worksheet calls require separate cases. | Inspect the failed assertions and HTTP evidence from that case. Establish permitted access before changing provider extraction. | 0 / 1 |
| `smfGetGuruFocusItem` | Live call failed | 0 selected cases passed and 1 failed. smfGetGuruFocusItem: CHECK_FAILED, HTTP 200, String: Error; length=5 [UNCLASSIFIED_FAIL] Other inputs, fields and worksheet calls require separate cases. | Inspect the failed assertions and HTTP evidence from that case. Establish permitted access before changing provider extraction. | 0 / 1 |
| `smfGetMorningstarHistoryCSV` | Live call failed | 0 selected cases passed and 1 failed. smfGetMorningstarHistoryCSV: CHECK_FAILED, Array Error -- Request failed (28): curl: (28) Failed to connect to performance.morningstar.com port 80 af \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|   [PROVIDER_UNAVAILABLE] Other inputs, fields and worksheet calls require separate cases. | Inspect the failed assertions and HTTP evidence from that case. Establish permitted access before changing provider extraction. | 0 / 1 |
| `smfGetMSNHistory` | Live call failed | 0 selected cases passed and 1 failed. smfGetMSNHistory: CHECK_FAILED, Array  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  \|  cells=12 [PROVIDER_UNAVAILABLE] Other inputs, fields and worksheet calls require separate cases. | Choose and document a replacement for the configured legacy history endpoint. | 0 / 1 |
| `smfGetMWOptionQuote` | Live call failed | 1 selected cases passed and 1 failed. Live_smfGetMWOptionQuote: CHECK_FAILED, HTTP 401, String: Error; length=5 [PROVIDER_UNAVAILABLE] Other inputs, fields and worksheet calls require separate cases. | Inspect the failed assertions and HTTP evidence from that case. Establish permitted access before changing provider extraction. | 1 / 1 |
| `smfGetNASDAQOptionQuote` | Live call failed | 1 selected cases passed and 1 failed. Live_smfGetNASDAQOptionQuote: CHECK_FAILED, HTTP 403, String: Error; length=5 [PROVIDER_UNAVAILABLE] Other inputs, fields and worksheet calls require separate cases. | Inspect the failed assertions and HTTP evidence from that case. Establish permitted access before changing provider extraction. | 1 / 1 |
| `smfGetYahooOptionQuote2` | Live call failed | 1 of 1 controlled cases passed. 0 of 1 selected live cases passed. The legacy HTML path supports extra item codes and a Date expiration. A 12-code alias does not preserve all those behaviors. | Replace legacy HTML extraction. Preserve the extra item codes, percentage units and Date expiration. | 1 / 1 |

## Retired modes: Not Supported

| Function | Result | Scope and limits | Remaining work | Passed / failed |
|---|---|---|---|---|
| `smfGet888OptionQuote` | Not Supported | 2 selected cases passed and 0 failed. 2 retired-mode guard assertions passed; they establish no financial feed. Mac returns typed #N/A for the registered retired mode (oic-legacy-options). No replacement financial source is established. | Choose and document a supported replacement before restoring financial results. Keep the existing #N/A guard until then. | 2 / 0 |
| `smfGetGoogleOptionQuote` | Not Supported | 2 selected cases passed and 0 failed. 2 retired-mode guard assertions passed; they establish no financial feed. Mac returns typed #N/A for the registered retired mode (google-option-chain). No replacement financial source is established. | Choose and document a supported replacement before restoring financial results. Keep the existing #N/A guard until then. | 2 / 0 |
| `smfGetMSNOptionQuote` | Not Supported | 2 selected cases passed and 0 failed. 2 retired-mode guard assertions passed; they establish no financial feed. Mac returns typed #N/A for the registered retired mode (msn-options). No replacement financial source is established. | Choose and document a supported replacement before restoring financial results. Keep the existing #N/A guard until then. | 2 / 0 |
| `smfGetOX2OptionQuote` | Not Supported | 2 selected cases passed and 0 failed. 2 retired-mode guard assertions passed; they establish no financial feed. Mac returns typed #N/A for the registered retired mode (optionsxpress-options). No replacement financial source is established. | Choose and document a supported replacement before restoring financial results. Keep the existing #N/A guard until then. | 2 / 0 |
| `smfGetOX3OptionQuote` | Not Supported | 2 selected cases passed and 0 failed. 2 retired-mode guard assertions passed; they establish no financial feed. Mac returns typed #N/A for the registered retired mode (optionsxpress-options). No replacement financial source is established. | Choose and document a supported replacement before restoring financial results. Keep the existing #N/A guard until then. | 2 / 0 |
| `smfGetOXOptionQuote` | Not Supported | 2 selected cases passed and 0 failed. 2 retired-mode guard assertions passed; they establish no financial feed. Mac returns typed #N/A for the registered retired mode (optionsxpress-options). No replacement financial source is established. | Choose and document a supported replacement before restoring financial results. Keep the existing #N/A guard until then. | 2 / 0 |
| `smfYahooAPIData` | Not Supported | 2 selected cases passed and 0 failed. 2 retired-mode guard assertions passed; they establish no financial feed. Mac returns typed #N/A for the registered retired mode (yahoo-yql). No replacement financial source is established. | Choose and document a supported replacement before restoring financial results. Keep the existing #N/A guard until then. | 2 / 0 |

## Local calculations and controlled cases

| Function | Result | Scope and limits | Remaining work | Passed / failed |
|---|---|---|---|---|
| `ConvertToIso` | Local or controlled check passed | Selected check: 2024-01-15T20:00:00.000Z; requires the verified America/Los_Angeles host timezone Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `ConvertToJson` | Local or controlled check passed | Selected check: [7,false,null] Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `ConvertToUtc` | Local or controlled check passed | Selected check: 2024-01-15 20:00:00 UTC; requires the verified America/Los_Angeles host timezone Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `DecodeJsonString` | Local or controlled check passed | Selected check: Collection with two Dictionary items named first and second Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `fnWait` | Local or controlled check passed | Selected check: Empty return after at least0.9seconds Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `GetKeys` | Local or controlled check passed | Selected check: Native array indexed 0..1 containing 0 and 01 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `GetObjectProperty` | Local or controlled check passed | Selected check: Collection containing 3 and 9 in that order Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `GetProperty` | Local or controlled check passed | Selected check: Boolean False Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `ParseIso` | Local or controlled check passed | Selected check: 2024-02-29 00:00:00 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `ParseJson` | Local or controlled check passed | Selected check: Dictionary with n=7, Boolean False, Null, and identifier text 1234567890123456789 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `ParseUtc` | Local or controlled check passed | Selected check: 2024-01-15 12:00:00 PST; requires the verified America/Los_Angeles host timezone Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHExtractData` | Local or controlled check passed | Selected check: 42 as a numeric result. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetHTMLTable` | Local or controlled check passed | Selected check: The 2-by-2 table Label/42 and Earnings/Nov 20 AMC. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetTableCell` | Local or controlled check passed | Selected check: 42 as a numeric result. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetURLData` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetURLData1` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetURLData1Mac` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetURLData2` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetURLData3` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetURLData4` | Local or controlled check passed | Selected check: HTTP 200; exact body SMF fixture response Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHGetWebData` | Local or controlled check passed | Selected check: Label Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `RCHImageTable` | Local or controlled check passed | Generated HTML image-table rows for two tickers. This function did not download or display the images. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfBuildYahooTimeSeriesURL` | Local or controlled check passed | Built the selected Yahoo fundamentals URL. No provider data request was made by this case. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfCDec` | Local or controlled check passed | Selected check: Double0.123456789 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfChkJSONFieldType` | Local or controlled check passed | Selected check: Numeric 1 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfConvertData` | Local or controlled check passed | Selected check: Double0.125 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfConvertYahooTicker` | Local or controlled check passed | Selected check: CA:CMG Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfDate2Unix` | Local or controlled check passed | Selected check: 86400 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfEval` | Local or controlled check passed | Selected check: 7 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfEvaluateTwice` | Local or controlled check passed | Selected check: 3 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetAConstants` | Local or controlled check passed | Selected check: suite-fixed-constant Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetAData` | Local or controlled check passed | Selected check: 32767 x characters Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetaFiles` | Local or controlled check passed | Selected check: A source-file record naming smf-elements-4.txt. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetAParms` | Local or controlled check passed | Selected check: The initialized FinViz definition containing the Earnings search marker. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetCSVFile` | Local or controlled check passed | Selected check: Date/Close/Volume headers, then 100.5/1000 and 102.25/1200. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 5 / 0 |
| `smfGetOptionExpiry` | Local or controlled check passed | 1 selected cases passed and 0 failed. This function calculates calendar expiration dates locally. It does not retrieve a provider option chain or establish available contracts. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfGetTagContent` | Local or controlled check passed | Selected check: Nov 20 AMC Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 3 / 0 |
| `smfHTMLDecode` | Local or controlled check passed | Selected check: <b>"A&B"</b> Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 3 / 0 |
| `smfInCellChart` | Local or controlled check passed | Line mode created four worksheet shapes from local sample data. Other chart modes were not established. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfJoin` | Local or controlled check passed | Selected check: 1,2,3,4,5 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfJSONExtractField` | Local or controlled check passed | Selected check: Numeric 7 from the second array item Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfJSONExtractField_x64` | Local or controlled check passed | Selected check: Boolean False from the first array item Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfJSONExtractKeys` | Local or controlled check passed | Selected check: One-column array with rows 0, 1, and #N/A Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfLogInternetCalls` | Local or controlled check passed | Selected check: Logging off; sLog=N Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacElementsSelfTest` | Local or controlled check passed | Selected check: 19 checks (3 element parser checks and 16 retirement checks). Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `SMFMacEmptyRecipeSelfTest` | Local or controlled check passed | Selected check: 5 checks: empty values and empty error fallbacks in both Mac formula branches, plus unchanged HTTP state; private element state restored. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacHistoryCSVSelfTest` | Local or controlled check passed | Selected check: 12 assertions; no live requests. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacHTMLSelfTest` | Local or controlled check passed | Selected check: 55 HTML checks Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacHTTPStatus` | Local or controlled check passed | Selected check: 207 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacIntegrationSelfTest` | Local or controlled check passed | Selected check: 25 integration checks Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacJSONSelfTest` | Local or controlled check passed | Selected check: PASS: 84 JSON checks Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacLastError` | Local or controlled check passed | Selected check: HTTP 403 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacLastFileError` | Local or controlled check passed | Selected check: empty string Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacLastURL` | Local or controlled check passed | Selected check: fixture://recorded-url Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacNestedRecipeSelfTest` | Local or controlled check passed | Selected check: 50 assertions; no live requests. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacReadElementLocation` | Local or controlled check passed | Selected check: Local followed by one LF Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacReadElementText` | Local or controlled check passed | Selected check: 19998;Fixture;Value;=1+2 followed by one LF Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacReadResponse` | Local or controlled check passed | 9 of 9 controlled cases passed. Cases cover response envelopes and credential redaction. Synthetic transport responses do not establish external provider access. | Other inputs and worksheet paths remain outside these selected cases. | 9 / 0 |
| `SMFMacRecipeSelfTest` | Local or controlled check passed | Selected check: 20 zero-HTTP recipe checks Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacStripHTML` | Local or controlled check passed | Selected check: 42&7 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacTableGuardSelfTest` | Local or controlled check passed | Selected check: 17 table guard checks Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacTableRecipeSelfTest` | Local or controlled check passed | Selected check: 40 table-recipe checks Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFMacTransportSelfTest` | Local or controlled check passed | Selected check: 6 deterministic transport checks Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacTryNestedRecipe` | Local or controlled check passed | Selected check: True, Double42 in the output argument, and unchanged HTTP state. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfMacTryTableCellRecipe` | Local or controlled check passed | Selected check: True with numeric parsed value42; no HTTP request Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacTryTagContentRecipe` | Local or controlled check passed | Selected check: True; parsed result numeric42 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacYahooChartCSV` | Local or controlled check passed | Selected check: Date,Open,High,Low,Close,Adj Close,Volume; Nov2 null open/adjusted/volume remain null, not0. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfMacYahooDownloadCSV` | Local or controlled check passed | Direct reversed-date and cached-CSV checks passed. Passed live CSV wrapper calls also used this helper indirectly. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfMacYahooDownloadURL` | Local or controlled check passed | Selected check: True for the exact Yahoo host; unchanged HTTP state. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 2 / 0 |
| `smfStrExtr` | Local or controlled check passed | Selected check: Double123.45 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfStripHTML` | Local or controlled check passed | Selected check: 7&8 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `SMFTech` | Local or controlled check passed | Ten local indicator cases passed: SMA, EMA, ATR, CCI, MACD, OBV, ROC, RSI, ADL and STO. It needs supplied price data. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 10 / 0 |
| `smfUnix2Date` | Local or controlled check passed | Selected check: 1970-01-02 00:00:00 as Date Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfUnix2DateStr` | Local or controlled check passed | Selected check: 1970-01-02 01:01:01 Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfWord` | Local or controlled check passed | Selected check: gamma Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |
| `smfYahooCodeDesc` | Local or controlled check passed | The local l1 code lookup returned Last Trade. It does not retrieve a quote. Other inputs, fields and worksheet calls require separate cases. | Other inputs and worksheet paths remain outside these selected cases. | 1 / 0 |

## Download helper

| Function | Result | Scope and limits | Remaining work | Passed / failed |
|---|---|---|---|---|
| `smfMacFetch` | URL dependent | 23 of 23 controlled cases passed. 9 of 26 external page-content checks passed. Endpoint-content failures do not identify a transport defect. Invalid-method redaction fixtures send no provider request. | Keep endpoint-content failures separate from transport errors. Additional request modes need separate cases. | 32 / 17 |

## Numbered fields for RCHGetElementNumber

22 literal financial-input cases cover 20 distinct ticker/field combinations.
Retired-field guards can pass without returning a financial value. Other financial values remain outside these cases.

| Field | Ticker or identifier | Data requested | Result | Qualification |
|---|---|---|---|---|
| 543 | `AAPL` | FreeRealTime StockScouter rating | Passed selected check |  |
| 625 | `AAPL` | Numbered field 625 | Not Supported | 2 of 2 assertions establish the retired-field guard, including any IFNA display. They establish no financial value. |
| 702 | `AAPL` | Barchart overall opinion | Failed selected check | Extract_www_barchart_com: CHECK_FAILED, HTTP 202, String: Error; length=5 [UNCLASSIFIED_FAIL] |
| 785 | `AAPL` | Stockscores 20-day average volume | Passed selected check | The data-field and provider page-marker checks are separate assertions. |
| 849 | `AAPL` | Zacks analyst target price | Passed selected check | The data-field and provider page-marker checks are separate assertions. |
| 980 | `AAPL` | Yahoo 10-day average volume | Passed selected check | Positive-number assertion only. Units and accuracy were not independently checked. |
| 995 | `SPY` | ETF Investment Outlook new-high percentage | Failed selected check | Extract_www_etfinvestmentoutlook_com: CHECK_FAILED, HTTP 200, String: Error; length=5 [UNCLASSIFIED_FAIL] |
| 1084 | `SPY` | ETFScreen relative-strength rank | Failed selected check | Extract_www_etfscreen_com: CHECK_FAILED, HTTP 200, String: Error; length=5 [UNCLASSIFIED_FAIL] |
| 1116 | `NYSE` | Barchart NYSE new-high count | Failed selected check | Extract_core-api_barchart_com: CHECK_FAILED, String: Not Found; length=9 [PROVIDER_UNAVAILABLE] |
| 2003 | `AAPL` | IBD industry-group rank | Failed selected check | Extract_research_investors_com: CHECK_FAILED, HTTP 200, String: Error; length=5 [UNCLASSIFIED_FAIL] |
| 2727 | `AAPL` | GuruFocus forward dividend yield | Passed selected check | Public dividend page. Does not establish authenticated GuruFocus API access. |
| 2747 | `NASDAQ:AAPL` | Google quarterly revenue | Failed selected check | Extract_finance_google_com: CHECK_FAILED, HTTP 200, String: Error; length=5 [UNCLASSIFIED_FAIL] |
| 4072 | `AAPL` | Finviz earnings date | Passed selected check |  |
| 4074 | `AAPL` | Finviz share price | Passed selected check | 2 cases. 1 recorded a case-bound HTTP request. |
| 5196 | `NASDAQ:AAPL` | ADVFN annual statement-ending date | Passed selected check | Valid-date assertion only. Saved result: Date: 9/26/26. Cold-example comparisons require separate evidence. |
| 13626 | `AAPL.O` | Reuters trailing P/E | Failed selected check | Extract_www_reuters_com: CHECK_FAILED, HTTP 401, String: Error; length=5 [PROVIDER_UNAVAILABLE] |
| 13865 | `AAPL` | Yahoo sector | Passed selected check |  |
| 19004 | `70111470` | Netflix movie release year | Passed selected check | Nonfinancial, year-only assertion. Movie identity was not verified. |
| 19902 | `AAPL` | TD Ameritrade ex-dividend date | Failed selected check | Extract_research_tdameritrade_com: CHECK_FAILED, String: Error; length=5 [PROVIDER_UNAVAILABLE] |
| 20000 | `AAPL` | Legacy MSN StockScouter rating | Failed selected check | Extract_investing_money_msn_com: CHECK_FAILED, String: ; length=0 [PROVIDER_UNAVAILABLE] |

### Explicit bulk definition assertions

- `NotSupported_All288DefaultFields`: PASS / NOT_SUPPORTED; one explicit native loop asserts 288 listed definitions.
  This is one bulk case. It does not establish that many separate worksheet formulas or working financial fields.

## Additional support Sub checks

These cases remain outside the public Function inventory and establish no financial feed.

| Case | Sub | Raw result | Interpretation |
|---|---|---|---|
| `NotSupported_RegistryUnknown` | `smfUnsupportedValue` | PASS | PASS |
| `NotSupported_RegistryKnown` | `smfUnsupportedValue` | PASS | PASS |
| `NotSupported_StatusRows` | `smfUnsupportedRows` | PASS | PASS |
| `NotSupported_UninitializedStatus` | `smfUnsupportedRows` | PASS | PASS |
| `NotSupported_StatusWorkbook` | `smfShowSourceStatus` | PASS | PASS |
| `NotSupported_URLAuthorityBoundaries` | `smfUnsupportedURL` | PASS | PASS |

## Evidence and limits

Saved workbook SHA256: `392139b5c9de34eaf7399b8860e4b2203ca9c09b8d9eeed64ff557ea0bf8fd93`.
Saved VBA SHA256: `7d309445abba5da5fe4bdded4e4652d47b992fc41b1158423e6cfc38d0df0b83`.
Source commit supplied for this reference: `e7d788b4f0871e94dbe6e13131ca8d9b7ff0eb8a`.
Prior reference: `32a04d90c4fca556c638362a29f7a4f7aa3f466e:mac/test-results/0.6-dev-20261108-final/FUNCTION-STATUS.md`.
The report verifies its case hashes and source fingerprint. A supplied commit ID is not a separate native measurement.
This generator makes no provider requests and leaves the earlier references unchanged.
Public Subs, private functions, Dictionary methods and Windows runtime behavior remain outside this inventory.
Cold reopens, installation and another Mac account require their separate receipts.

- [Function CSV](function-status.csv)
- [All saved case results](all-case-results.csv)
- [Saved report data](report-data.json)
- [Full results](README.md)
