# Remaining provider results — November 8, 2026

This run uses the isolated full-source workbook with core version `3.0.2024.07.30-mac.0.3-dev`.
The selected cases include 18 original pending checks and 1 supplemental local-image control.

The harness recorded final statuses for 18 of 18 primary cases.
3 primary cases passed. 0 primary cases remain unfinished.
Primary results: 4 ASSERTION_ERROR, 11 CHECK_FAILED, 3 PASS.

The report counts a returned `Error` as a failed data result.
`SETUP_ERROR` records an error during preparation. `ASSERTION_ERROR` records an error in the result check.
`CALL_ERROR` records an exception during the call. Unfinished rows do not count as passes.

## Image rerun

The first attempt recorded five `SETUP_ERROR` results for picture cases.
The agent repeated those cases after the harness retained the J20 value before reading the picture fill.
All five reruns retained `#VALUE!` and error `445` in the `CHECK` phase at `J20.Comment.Shape.Fill`.
These cases include three provider images, a loopback image and a valid local PNG.
The getter returned `Object doesn't support this action`. These results do not establish chart-provider HTTP failures.
The operator observed no Excel exit during the 19-case batch and five-case image rerun.

## The original pending checks

| Case | Provider or control | Result | Returned value and image details |
| --- | --- | --- | --- |
| `Extract_research_tdameritrade_com` | research.tdameritrade.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_www_barchart_com` | www.barchart.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_www_etfinvestmentoutlook_com` | www.etfinvestmentoutlook.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_www_etfscreen_com` | www.etfscreen.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_www_gurufocus_com` | www.gurufocus.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_www_reuters_com` | www.reuters.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_www_stockscores_com` | www.stockscores.com | PASS | Double: 48601150 |
| `Extract_www_zacks_com` | www.zacks.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_investing_money_msn_com` | investing.money.msn.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_dvd_netflix_com` | dvd.netflix.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_core-api_barchart_com` | core-api.barchart.com | CHECK_FAILED | String: Error; length=5; Configured Barchart URL=https://core-api.barchart.com/v1/momentum/get?raw=1&fields=exchange,high52w,low52w,newHighs,newLows,advancingVolume,unchangedVolume,decliningVolume,advancingIssues,percentAdvancingIssues,unchangedIssues,percentUnchangedIssues,decliningIssues,percentDecliningIssues&exchanges=NASDAQ,NYSE,OTC,AMEX |
| `Extract_www_advfn_com` | www.advfn.com | CHECK_FAILED | String: Error; length=5 |
| `Extract_query1_finance_yahoo_com` | query1.finance.yahoo.com | PASS | Array Symbol \| AAPL \| Last Price \| 342.37 \|  cells=4 |
| `ImageComment_c_stockcharts_com` | c.stockcharts.com | ASSERTION_ERROR | Excel error Error 2015; Reading J20.Comment.Shape.Fill; 445: ImageComment_c_stockcharts_com [CHECK]: Object doesn't support this action |
| `ImageComment_stockcharts_com` | stockcharts.com | ASSERTION_ERROR | Excel error Error 2015; Reading J20.Comment.Shape.Fill; 445: ImageComment_stockcharts_com [CHECK]: Object doesn't support this action |
| `ImageComment_ogres-crypt_com` | ogres-crypt.com | ASSERTION_ERROR | Excel error Error 2015; Reading J20.Comment.Shape.Fill; 445: ImageComment_ogres-crypt_com [CHECK]: Object doesn't support this action |
| `ImageComment_loopback` | loopback | ASSERTION_ERROR | Excel error Error 2015; Reading J20.Comment.Shape.Fill; 445: ImageComment_loopback [CHECK]: Object doesn't support this action |
| `Live_smfGetOptionExpirations_50Rows` | query1.finance.yahoo.com | PASS | Array AAPL 11/9 2026 $350.00 Call \| AAPL 11/11 2026 $350.00 Call \| AAPL 11/13 2026 $350.00 Call \| AAPL 11/16 2026 $350.00 Call \| AAPL 11/18 2026 $350.00 Call \| AAPL 11/20 2026 $350.00 Call \| AAPL 11/27 2026 $350.00 Call \| AAPL 12/4 2026 $350.00 Call \| AAPL 12/11 2026 $350.00 Call \| AAPL 12/18 2026 $350.00 Call \| AAPL 12/24 2026 $350.00 Call \| AAPL 1/15 2027 $350.00 Call \| AAPL 2/19 2027 $350.00 Call \| AAPL 3/19 2027 $350.00 Call \... |

## Supplemental local-image control

This control is separate from the original pending checks.
Supplemental results: 1 ASSERTION_ERROR.

| Case | Provider or control | Result | Returned value and image details |
| --- | --- | --- | --- |
| `ImageComment_local_file` | local_file | ASSERTION_ERROR | Excel error Error 2015; Reading J20.Comment.Shape.Fill; 445: ImageComment_local_file [CHECK]: Object doesn't support this action |

## Expected values and diagnostics

| Case | Expected result | Recorded error |
| --- | --- | --- |
| `Extract_research_tdameritrade_com` | An AAPL ex-dividend date from the original TD Ameritrade recipe. |  |
| `Extract_www_barchart_com` | AAPL overall opinion containing Buy, Sell or Hold. |  |
| `Extract_www_etfinvestmentoutlook_com` | Numeric SPY new-high percentage in the configured breadth table. |  |
| `Extract_www_etfscreen_com` | A numeric SPY current relative-strength rank between 0 and 100. |  |
| `Extract_www_gurufocus_com` | Positive AAPL forward dividend yield. |  |
| `Extract_www_reuters_com` | Positive AAPL trailing P/E; the Reuters instrument identifier is AAPL.O. |  |
| `Extract_www_stockscores_com` | Positive AAPL twenty-day average volume. |  |
| `Extract_www_zacks_com` | Positive AAPL consensus analyst target price. |  |
| `Extract_investing_money_msn_com` | An AAPL StockScouter rating from the custom MSN recipe. |  |
| `Extract_dvd_netflix_com` | Movie release year2010 for Iron Man2 (Netflix ID70111470), extracted from the legacy DVD page. |  |
| `Extract_core-api_barchart_com` | A numeric nonnegative NYSE count of new 52-week highs; constant 1 supplies the Barchart URL. |  |
| `Extract_www_advfn_com` | An actual annual statement-ending date from ADVFN; exchange-qualified NASDAQ:AAPL is supplied. |  |
| `Extract_query1_finance_yahoo_com` | Headers Symbol/Last Price and one AAPL row with a positive share price. |  |
| `ImageComment_c_stockcharts_com` | Cell result Chart, a newly created comment and Comment.Shape.Fill.Type = msoFillPicture (6). | 445: ImageComment_c_stockcharts_com [CHECK]: Object doesn't support this action |
| `ImageComment_stockcharts_com` | Cell result Chart, a newly created comment and Comment.Shape.Fill.Type = msoFillPicture (6). | 445: ImageComment_stockcharts_com [CHECK]: Object doesn't support this action |
| `ImageComment_ogres-crypt_com` | Cell result Chart, a newly created comment and Comment.Shape.Fill.Type = msoFillPicture (6). | 445: ImageComment_ogres-crypt_com [CHECK]: Object doesn't support this action |
| `ImageComment_loopback` | Cell result Chart, a newly created comment and Comment.Shape.Fill.Type = msoFillPicture (6). | 445: ImageComment_loopback [CHECK]: Object doesn't support this action |
| `Live_smfGetOptionExpirations_50Rows` | An AAPL future expiration formatted with $350.00 Call, returned normally in the explicit50-row result; HTTP200. This call does not validate strike availability. |  |
| `ImageComment_local_file` | Cell result Chart, a newly created comment and Comment.Shape.Fill.Type = msoFillPicture (6). | 445: ImageComment_local_file [CHECK]: Object doesn't support this action |

## Interpretation limits

- The harness requires returned text `Chart` and `Comment.Shape.Fill.Type = 6` for a picture case.
- The harness does not inspect the picture ticker, date, prices or quote freshness.
- Excel loads comment pictures through `Shape.Fill.UserPicture`. Those requests do not update the SMF HTTP diagnostic fields.
- The SMF last HTTP URL, status and error can come from an earlier request. The CSV preserves them as recorded diagnostics.
- A PASS applies to the selected field or function call. This run does not verify every field on that provider.
- The report does not change the original interrupted workbook or its saved results.
- Netflix lists movie `70111470` as *Iron Man 2* (2010). The new expected year follows [Netflix's title page](https://www.netflix.com/title/70111470).

## Saved evidence

These evidence files remain local under `mac/diagnostics/out`. They are excluded from Git.
The relative links resolve in a checkout that retains the evidence.

- [Native result rows](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/native-results.json)
- [Executed case specifications](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/test-cases.json)
- [Parent run and selected IDs](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/parent-run.json)
- [Full result CSV](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/remaining-results.csv)
- [Run summary](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/run-summary.json)

- [Saved Excel workbook](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/SMF-Mac-Remaining-Providers.xlsm)

First-attempt derived records remain local:

- [Summary before the image rerun](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/run-summary-before-image-rerun.json)
- [CSV before the image rerun](diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo/remaining-results-before-image-rerun.csv)
