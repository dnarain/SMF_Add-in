# Data and picture repair results — November 8, 2026

This report records the saved test-workbook results for the `0.4-dev` source changes.
The later add-in build, restart and saved-worksheet checks are recorded below.
The following tables retain the original 59-case results.

The selected run contains 59 cases. The harness recorded final statuses for 59.
52 cases passed. 0 cases remain unfinished.
Recorded results: 7 CHECK_FAILED, 52 PASS.

The case count includes grouped self-tests. Their internal assertion counts are not added to the case count.
Some negative regressions expect `Error`, `None` or rejection. Their expected behavior appears in the result tables.

## Results by group

| Group | Cases | Result statuses |
| --- | ---: | --- |
| Regressions and offline fixtures | 25 | 25 PASS |
| Live data checks | 13 | 13 PASS |
| Previously blocked provider checks | 7 | 7 CHECK_FAILED |
| Image and chart checks | 14 | 14 PASS |

## Seven fields without a source repair

The seven definitions below retain their original providers. No source repair is claimed for them.
The independent request observations and this run's native outcomes are separate records.

| Field | Prior external observation | Native result |
| --- | --- | --- |
| 19902 — TD Ameritrade | Legacy host failed DNS resolution. | 1 CHECK_FAILED |
| 702 — Barchart opinion | HTTP 202, empty body, WAF challenge. | 1 CHECK_FAILED |
| 995 — ETFInvestmentOutlook | Redirect to an unrelated accounting-services website. | 1 CHECK_FAILED |
| 1084 — ETFScreen | Domain-parking page instead of the ETF table. | 1 CHECK_FAILED |
| 13626 — Reuters | Legacy and current routes returned HTTP 401 with DataDome protection. | 1 CHECK_FAILED |
| 20000 — MSN StockScouter | Legacy host failed DNS resolution. | 1 CHECK_FAILED |
| 1116 — Barchart momentum API | Connection to port 443 timed out. | 1 CHECK_FAILED |

An observed PASS can establish a successful request for that case. It does not establish a code repair for a prior external blocker.

## Regressions and offline fixtures

| Case | Status | Expected result | Returned value and diagnostics |
| --- | --- | --- | --- |
| `SMFMacTransportSelfTest` | PASS | 6 deterministic transport checks | Long: 6 |
| `smfMacHTMLSelfTest` | PASS | 55 HTML checks | Long: 55 |
| `SMFMacJSONSelfTest` | PASS | PASS: 84 JSON checks | String: PASS: 84 JSON checks; length=20 |
| `SMFMacRecipeSelfTest` | PASS | 20 zero-HTTP recipe checks | Long: 20 |
| `SMFMacTableRecipeSelfTest` | PASS | 40 table-recipe checks | Long: 40 |
| `SMFMacTableGuardSelfTest` | PASS | 17 table guard checks | Long: 17 |
| `SMFMacIntegrationSelfTest` | PASS | 25 integration checks | Long: 25 |
| `Repair_SMFMacNestedRecipeSelfTest` | PASS | 50 assertions; no live requests. | Long: 50 |
| `Repair_SMFMacHistoryCSVSelfTest` | PASS | 12 assertions; no live requests. | Long: 12 |
| `SMFMacProviderRecipeSelfTest` | PASS | 25 checks:8 direct extractions,8 numbered extractions,8 missing-field Errors and unchanged HTTP state. | Long: 25 |
| `Repair_History_AAPL_Snapshot` | PASS | AAPL 2026-11-06 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 11/6/26 \| 11/5/26 \| 11/4/26 \| 11/3/26 \| 11/2/26 \| Open \| 338.730010986328 \| 353.890014648438 \| 341.570007324219 \| 325.579986572266 \| 342.670013427734 \| High \| 345.790008544922 \| 355.380004882812 \| 354.5 \|  cells=42 |
| `Repair_History2_AAPL_Snapshot` | PASS | AAPL 2026-11-06 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 11/6/26 \| 11/5/26 \| 11/4/26 \| 11/3/26 \| 11/2/26 \| Open \| 338.730010986328 \| 353.890014648438 \| 341.570007324219 \| 325.579986572266 \| 342.670013427734 \| High \| 345.790008544922 \| 355.380004882812 \| 354.5 \|  cells=42 |
| `Repair_History_AAPL_Ascending` | PASS | AAPL 2026-11-02 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 11/2/26 \| 11/3/26 \| 11/4/26 \| 11/5/26 \| 11/6/26 \| Open \| 342.670013427734 \| 325.579986572266 \| 341.570007324219 \| 353.890014648438 \| 338.730010986328 \| High \| 343.799987792969 \| 334.5 \| 354.5 \|  cells=42 |
| `Repair_History2_AAPL_Ascending` | PASS | AAPL 2026-11-02 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 11/2/26 \| 11/3/26 \| 11/4/26 \| 11/5/26 \| 11/6/26 \| Open \| 342.670013427734 \| 325.579986572266 \| 341.570007324219 \| 353.890014648438 \| 338.730010986328 \| High \| 343.799987792969 \| 334.5 \| 354.5 \|  cells=42 |
| `Repair_History_IBM_AdjustedAlias` | PASS | IBM 2026-11-06 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 11/6/26 \| 11/5/26 \| 11/4/26 \| 11/3/26 \| 11/2/26 \| AdjClose \| 245 \| 247.110000610352 \| 254.320007324219 \| 250.039993286133 \| 252.509994506836 \| Close \| 245 \| 247.110000610352 \| 254.320007324219 \|  cells=24 |
| `Repair_History2_AdjustmentRatio` | PASS | Adjusted OHLC uses adjusted/close0.5; volume stays2345/1234; adjusted close stays55.5/50.5; newest first. | Array Date \| 11/3/26 \| 11/2/26 \| Adj. Open \| 55 \| 50 \| Adj. High \| 57.5 \| 52.5 \| Adj. Low \| 52.5 \| 47.5 \| Adj. Close \| 55.5 \| 50.5 \| Volume \|  cells=21 |
| `Repair_History_ReversedDates_NoHTTP` | PASS | Error identifying reversed dates; HTTP status, URL and error unchanged. | String: Error -- Yahoo history: Starting date must precede ending date; length=62 |
| `Repair_History_NullValues_CSV` | PASS | Date,Open,High,Low,Close,Adj Close,Volume; Nov2 null open/adjusted/volume remain null, not0. | String: Date,Open,High,Low,Close,Adj Close,Volume / 2026-11-02,null,105,95,101,null,null / ; length=79 |
| `Repair_RCH_LiteralEmpty_NoHTTP` | PASS | 5 checks: empty values and empty error fallbacks in both Mac formula branches, plus unchanged HTTP state; private element state restored. | Long: 5 |
| `Repair_History2_NoHeader_Bounded` | PASS | 3x2 array: Nov6/5/4 closes342.3699951171875,344.1700134277344,349.510009765625. | Array 11/6/26 \| 11/5/26 \| 11/4/26 \| 342.369995117188 \| 344.170013427734 \| 349.510009765625 \|  cells=6 |
| `Repair_NestedDispatcher_Direct` | PASS | True, Double42 in the output argument, and unchanged HTTP state. | Boolean: True |
| `Repair_NestedDispatcher_RejectWhole` | PASS | False, untouched output, and unchanged HTTP state. | Boolean: False |
| `Repair_DownloadURL_AcceptYahoo` | PASS | True for the exact Yahoo host; unchanged HTTP state. | Boolean: True |
| `Repair_DownloadURL_RejectSpoof` | PASS | False for the lookalike host; unchanged HTTP state. | Boolean: False |
| `Repair_DownloadCSV_AAPL_Snapshot` | PASS | Standard CSV with five dated AAPL rows, exact captured close/adjusted/volume, terminal LF, and no HTTP request. | String: Date,Open,High,Low,Close,Adj Close,Volume / 2026-11-02,342.6700134277344,343.79998779296875,335.82000732421875,338.32000732421875,338.32000732421875,64757700 / 2026-11-03,325.5799865722656,334.5,325.3299865722656,331.8999938964844,331.8999938964844,63160000 / 2026-11-04,341.57000732421875,354.5,340.989990234375,349.510009765625,349.510009765625,89637700 / 2026-11-05,353.8900146484375,355.3800048828125,342.05999755859375,344.1700134277344,344.1700134277344,75678900 / 2026-11-06,338.730... |

## Live data checks

| Case | Status | Expected result | Returned value and diagnostics |
| --- | --- | --- | --- |
| `Repair_RCHHistory_AAPL_Live` | PASS | AAPL 2026-11-06 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 11/6/26 \| 11/5/26 \| 11/4/26 \| 11/3/26 \| 11/2/26 \| Open \| 338.730010986328 \| 353.890014648438 \| 341.570007324219 \| 325.579986572266 \| 342.670013427734 \| High \| 345.790008544922 \| 355.380004882812 \| 354.5 \|  cells=42 |
| `Repair_History2_IBM_Live` | PASS | IBM 2026-11-06 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 11/6/26 \| 11/5/26 \| 11/4/26 \| 11/3/26 \| 11/2/26 \| Open \| 242.919998168945 \| 255.919998168945 \| 255 \| 249.559997558594 \| 258.119995117188 \| High \| 246.339996337891 \| 256.179992675781 \| 258.589996337891 \|  cells=42 |
| `Repair_HistoryCSV_IBM_Live` | PASS | IBM 2026-11-02 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 2026-11-02 \| 2026-11-03 \| 2026-11-04 \| 2026-11-05 \| 2026-11-06 \| Open \| 258.119995117188 \| 249.559997558594 \| 255 \| 255.919998168945 \| 242.919998168945 \| High \| 259.869995117188 \| 250.949996948242 \| 258.589996337891 \|  cells=42 |
| `Repair_YahooQuotes_CSVMode_AAPL_Live` | PASS | AAPL 2026-11-02 first data row; exact snapshot prices/volume within1e-7; 5 data rows. | Array Date \| 2026-11-02 \| 2026-11-03 \| 2026-11-04 \| 2026-11-05 \| 2026-11-06 \| Open \| 342.670013427734 \| 325.579986572266 \| 341.570007324219 \| 353.890014648438 \| 338.730010986328 \| High \| 343.799987792969 \| 334.5 \| 354.5 \|  cells=42 |
| `Repair_HistoryCSVData_AAPL_Live` | PASS | HTTP200; standard CSV header and terminal LF; Nov2‚Äì6; first close338.32000732421875, adjusted same, volume64757700. | String: Date,Open,High,Low,Close,Adj Close,Volume / 2026-11-02,342.6700134277344,343.79998779296875,335.82000732421875,338.32000732421875,338.32000732421875,64757700 / 2026-11-03,325.5799865722656,334.5,325.3299865722656,331.8999938964844,331.8999938964844,63160000 / 2026-11-04,341.57000732421875,354.5,340.989990234375,349.510009765625,349.510009765625,89637700 / 2026-11-05,353.8900146484375,355.3800048828125,342.05999755859375,344.1700134277344,344.1700134277344,75678900 / 2026-11-06,338.730... |
| `Repair_OptionExpirations_FourRows` | PASS | Exactly4 formatted AAPL $350.00 Call labels; no raw date rows; HTTP200. | Array AAPL 11/9 2026 $350.00 Call \| AAPL 11/11 2026 $350.00 Call \| AAPL 11/13 2026 $350.00 Call \| AAPL 11/16 2026 $350.00 Call \|  cells=4 |
| `Extract_www_gurufocus_com` | PASS | Positive AAPL forward dividend yield. | Double: 0.0034 |
| `Extract_www_stockscores_com` | PASS | Positive AAPL twenty-day average volume. | Double: 48601150 |
| `Extract_www_zacks_com` | PASS | Positive AAPL consensus analyst target price. | Double: 338.68 |
| `Extract_dvd_netflix_com` | PASS | Movie release year2010 for Iron Man2 (Netflix ID70111470), extracted from the legacy DVD page. | Double: 2010 |
| `Extract_www_advfn_com` | PASS | An actual annual statement-ending date from ADVFN; exchange-qualified NASDAQ:AAPL is supplied. | Date: 9/27/25 |
| `Extract_query1_finance_yahoo_com` | PASS | Headers Symbol/Last Price and one AAPL row with a positive share price. | Array Symbol \| AAPL \| Last Price \| 342.37 \|  cells=4 |
| `Live_smfGetOptionExpirations_50Rows` | PASS | At least one available AAPL $350 Call expiration and HTTP200; the function has no expiration-date filter | Array AAPL 11/9 2026 $350.00 Call \| AAPL 11/11 2026 $350.00 Call \| AAPL 11/13 2026 $350.00 Call \| AAPL 11/16 2026 $350.00 Call \| AAPL 11/18 2026 $350.00 Call \| AAPL 11/20 2026 $350.00 Call \| AAPL 11/27 2026 $350.00 Call \| AAPL 12/4 2026 $350.00 Call \| AAPL 12/11 2026 $350.00 Call \| AAPL 12/18 2026 $350.00 Call \| AAPL 12/24 2026 $350.00 Call \| AAPL 1/15 2027 $350.00 Call \| AAPL 2/19 2027 $350.00 Call \| AAPL 3/19 2027 $350.00 Call \| AAPL 4/16 2027 $350.00 Call \| AAPL 6/17 2027 $3... |

## Previously blocked provider checks

| Case | Status | Expected result | Returned value and diagnostics |
| --- | --- | --- | --- |
| `Extract_research_tdameritrade_com` | CHECK_FAILED | An AAPL ex-dividend date from the original TD Ameritrade recipe. | String: Error; length=5 |
| `Extract_www_barchart_com` | CHECK_FAILED | AAPL overall opinion containing Buy, Sell or Hold. | String: Error; length=5 |
| `Extract_www_etfinvestmentoutlook_com` | CHECK_FAILED | Numeric SPY new-high percentage in the configured breadth table. | String: Error; length=5 |
| `Extract_www_etfscreen_com` | CHECK_FAILED | A numeric SPY current relative-strength rank between 0 and 100. | String: Error; length=5 |
| `Extract_www_reuters_com` | CHECK_FAILED | Positive AAPL trailing P/E; the Reuters instrument identifier is AAPL.O. | String: Error; length=5 |
| `Extract_investing_money_msn_com` | CHECK_FAILED | An AAPL StockScouter rating from the custom MSN recipe. | String: ; length=0 |
| `Extract_core-api_barchart_com` | CHECK_FAILED | A numeric nonnegative NYSE count of new 52-week highs; constant 1 supplies the Barchart URL. | String: Not Found; length=9 |

## Image and chart checks

| Case | Status | Expected result | Returned value and diagnostics |
| --- | --- | --- | --- |
| `ImageComment_local_file` | PASS | Chart result and a newly embedded Mac picture tied to J20, with the requested visible state. | String: Chart; length=5; Mac Shape.Type=13; Visible=-1; Metadata=SMFCommentPicture/1 / J20 / mac/diagnostics/out/SMF-Mac-data-picture-repair-20261108-k78nfdd7/fixture.png |
| `MacComment_numeric_constant` | PASS | Chart result and a newly embedded Mac line chart tied to J20, with the requested visible state. | String: Chart; length=5; Mac Shape.Type=3; Visible=-1; Metadata=SMFCommentPicture/1 / J20 / SMF local numeric chart: 2,2,2; Series=50,50,50 |
| `MacComment_numeric_nonpositive` | PASS | Chart result and a newly embedded Mac line chart tied to J20, with the requested visible state. | String: Chart; length=5; Mac Shape.Type=3; Visible=-1; Metadata=SMFCommentPicture/1 / J20 / SMF local numeric chart: -1,0,2,4; Series=0,0,1 |
| `MacComment_toggle_visibility` | PASS | Chart result and a newly embedded Mac picture tied to J20, with the requested hidden state. | String: Chart; length=5; Mac Shape.Type=13; Visible=0; Metadata=SMFCommentPicture/1 / J20 / mac/diagnostics/out/SMF-Mac-data-picture-repair-20261108-k78nfdd7/fixture.png |
| `MacComment_text_only` | PASS | Chart result, the requested text note, and no SMF picture object. | String: Chart; length=5; Note text=SMF picture fixture |
| `MacComment_clear_none` | PASS | None result after a valid picture; the previous picture and note must both be absent. | String: None; length=4; Previous Mac picture and note removed |
| `MacComment_failed_update_clears_picture` | PASS | Error result for HTTP200 text, no previous picture, and a note identifying the invalid image response. | String: Error; length=5; Failure note=SMF chart unavailable: Downloading the image: Image request failed (-2700): Response is not a PNG, JPEG, or GIF image |
| `ImageComment_c_stockcharts_com` | PASS | Chart result and a newly embedded Mac picture tied to J20, with the requested hidden state. | String: Chart; length=5; Mac Shape.Type=13; Visible=0; Metadata=SMFCommentPicture/1 / J20 / https://stockcharts.com/c-sc/sc?s=AAPL&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806 |
| `ImageComment_stockcharts_com` | PASS | Chart result and a newly embedded Mac picture tied to J20, with the requested hidden state. | String: Chart; length=5; Mac Shape.Type=13; Visible=0; Metadata=SMFCommentPicture/1 / J20 / https://stockcharts.com/pnf/chart?c=AAPL,P |
| `ImageComment_ogres-crypt_com` | PASS | Chart result and a newly embedded Mac line chart tied to J20, with the requested hidden state. The one series must contain 1,50,98. | String: Chart; length=5; Mac Shape.Type=3; Visible=0; Metadata=SMFCommentPicture/1 / J20 / SMF local numeric chart: 1,2,3; Series=1,50,98 |
| `ImageComment_loopback` | PASS | Chart result and a newly embedded Mac picture tied to J20, with the requested hidden state. | String: Chart; length=5; Mac Shape.Type=13; Visible=0; Metadata=SMFCommentPicture/1 / J20 / http://127.0.0.1:53225/image.png |
| `MacComment_numeric_tiny_positive` | PASS | Chart result with one local numeric series containing exactly 1,98. | String: Chart; length=5; Mac Shape.Type=3; Visible=-1; Metadata=SMFCommentPicture/1 / J20 / SMF local numeric chart: 0.001,0.002; Series=1,98 |
| `MacComment_numeric_tiny_constant` | PASS | Chart result with one local numeric series containing exactly 50,50. | String: Chart; length=5; Mac Shape.Type=3; Visible=-1; Metadata=SMFCommentPicture/1 / J20 / SMF local numeric chart: 0.005,0.005; Series=50,50 |
| `MacComment_numeric_large_positive` | PASS | Chart result with one local numeric series containing exactly 1,98. | String: Chart; length=5; Mac Shape.Type=3; Visible=-1; Metadata=SMFCommentPicture/1 / J20 / SMF local numeric chart: 1000000000000,2000000000000; Series=1,98 |

## Loaded add-in and saved worksheet

Excel built and loaded `3.0.2024.07.30-mac.0.4-dev` from `mac/out/SMF-Mac.xlam`.
The file is 713,180 bytes, with SHA-256 `ad7cf4d0a79dea7b28a57f0deb6945118e172ecd5b57efb6adbd81d375cbc18e`.
Its 51 core modules match the passed snapshot, and its 14 adjacent definition files match `docs/Elements/` byte for byte.

The first restart reported a format error for the registered add-in. Opening the exact file succeeded.
The agent registered that exact file again, quit Excel through its menu and relaunched it.
Excel then loaded `0.4-dev` automatically without an alert. The cause of the first alert remains undetermined.
The reopened [data and picture check workbook](out/SMF-Mac-Data-Picture-Checks-20261108.xlsx) returned all 11 PASS results.
The StockCharts PNG and native three-point chart remained visible.
The saved workbook contains one embedded PNG, one chart and no cached error cells.

Seven separate [derived-history cases](tests/full_function_suite/derived_history_cases.json) also passed.
The 59-case and seven-case batches contain 66 unique cases: 59 PASS and seven provider CHECK_FAILED results.
The 11 worksheet checks overlap those cases and do not increase that total.
The seven provider failures in the tables above remain unresolved.

## Worksheet row-count limitation in the `0.4-dev` build

The checked worksheet calls returned close (`C`), volume (`V`) and date (`D`) data with omitted or zero `pRows`.
An invalid period returned a diagnostic with default row settings.
An explicit positive `pRows` still returned `#VALUE!` after a fresh Excel launch.
An invalid-period control also returned `#VALUE!` before HTTP retrieval.
The native macro history cases with explicit row counts passed.
The local [worksheet probe workbook](diagnostics/out/SMF-Mac-Worksheet-Probes-20261108.xlsx) preserves the exact controls.

This checked formula omits the row and column overrides:

```excel
=INDEX(smfGetYahooHistory("AAPL",DATE(2026,11,2),DATE(2026,11,7),"d","C",0),1,1)
```

The later [`0.5-dev` source fix](YAHOO-HISTORY-ROWCOUNT-RESULTS-20261108.md) passed 46 native controls of the actual public function.
It uses `Double` for Mac `pCols` and tests converted local dimensions before caller sizing.
The earlier tables and their 59-case count remain unchanged. The seven external provider failures remain unresolved.
The audited `0.5-dev` add-in subsequently passed eight fresh [worksheet checks](out/SMF-Mac-History-Checks-20261108.xlsx) and a live `5x2` result after a clean Excel restart.
Its version formula reported the correct output directory and `mac.0.5-dev`.
The 46 focused controls and eight follow-up checks remain separate from this report's original counts.

## Scope and evidence

- A PASS applies to the selected inputs and expected result. It does not verify every field on a provider.
- The image cases check returned text, object type, visibility, metadata or plotted values according to each case.
- A picture object does not establish the chart ticker, date, prices or quote freshness.
- The last HTTP URL/status can belong to an earlier request. The local CSV preserves those fields as diagnostics.
- `SETUP_ERROR` records preparation failure. `ASSERTION_ERROR` records a failure inside the result check.
- `CALL_ERROR` records a call-stage exception. The recorded error identifies the failing operation.
- Saved RUNNING, PENDING and missing rows have no completed result. They do not count as passes.
- Windows runtime verification remains separate and pending.
- Historical native rows and earlier reports remain unchanged.

The full evidence files remain local under `mac/diagnostics/out` and are excluded from Git.
The public report redacts cookie/crumb values and user paths. It omits raw HTML responses.

- [Native result rows](diagnostics/out/SMF-Mac-data-picture-repair-20261108-k78nfdd7/native-results.json)
- [Selected case specifications](diagnostics/out/SMF-Mac-data-picture-repair-20261108-k78nfdd7/test-cases.json)
- [Full result CSV](diagnostics/out/SMF-Mac-data-picture-repair-20261108-k78nfdd7/repair-results.csv)
- [Run summary](diagnostics/out/SMF-Mac-data-picture-repair-20261108-k78nfdd7/repair-summary.json)
- [Source repair notes](DATA-PICTURE-REPAIRS-20261108.md)

- [Saved test workbook](diagnostics/out/SMF-Mac-data-picture-repair-20261108-k78nfdd7/SMF-Mac-Repair.xlsm)
