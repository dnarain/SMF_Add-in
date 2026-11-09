# Provider changes in the native Mac port

This file records changes from upstream commit `a0daad05699b1a7e213cf72ee641f30fc615af46` through Mac source commit `c4852c1bd3d20a04b435dff62141270bca550af1`.
It also records the Google-history URL repair, which passed its isolated native Excel checks on November 9, 2026.

The downloadable experimental bundle remains [0.6-dev](mac/built/0.6-dev/README.md).
The local 0.7 checkpoint saved **526 cases: 490 PASS and 36 CHECK_FAILED**, including 50 passing worksheet-entry cases.
Those counts include controlled fixtures and expected retired-mode errors. They do not count working providers or independently verified financial fields.
A later runner audit found a tab-containing ADVFN fixture that Excel imported as spaces.
The Google-history repair passed12 isolated URL/CSV regression checks. Its separate financial-data assertion still failed because Google returned HTML. The corrected fixture and final combined source require new full/focused runs before the next package is verified.

URL values in angle brackets are placeholders. `<API_KEY>` never represents a checked-in credential.
The table describes configured requests, before provider redirects.

## Financial data and page extraction

| Provider / functions or fields | Old URL or route | New Mac URL or route | Status, reason and native evidence |
|---|---|---|---|
| **Finviz** — fields 4072 earnings and 4074 price | `https://finviz.com/quote.ashx?t=<ticker>&p=d` | Same URL | The fork adds the Mac downloader and direct execution of supported field recipes. It does not change these two recipes' provider or URL. Selected earnings/price calls passed historically; the 526 checkpoint returned AAPL price 350.69. |
| **GuruFocus** — field 2727, forward dividend yield | `https://www.gurufocus.com/dividend/<ticker>` | `https://www.gurufocus.com/stock/<ticker>/dividend` | Reads the current `(Forward)` div instead of the old h1 marker. The field remains a percentage fraction. Captured fixtures and selected live AAPL/MSFT calls passed earlier; the 526 AAPL check returned 0.0034. Other GuruFocus fields and authenticated access remain separate. |
| **Zacks** — field 849, target-price consensus | `https://www.zacks.com/stock/research/<ticker>/brokerage-recommendations` | `https://www.zacks.com/stock/research/<ticker>/price-target-stock-forecast` | The table label changes from `Average Target Price` to `Average Price Target`. Selected captures and live extraction passed; the 526 AAPL check returned 338.96. This does not establish all Zacks fields. |
| **Netflix** — field 19004, movie release year | `http://dvd.netflix.com/Movie/<title_id>` | `https://www.netflix.com/title/<title_id>` | Reads the first metadata list item instead of the DVD-page year span. The 526 check for title 70111470 returned 2010. Other DVD recipes remain unchanged; this repair does not establish television-series dates. |
| **ADVFN** — field 5196, latest annual ending date | An `AdvFN-A` numbered-field lookup through the legacy financial-table extractor | `https://www.advfn.com/p.php?pid=financials&symbol=<symbol>` | The field recipe now extracts `Year End` and converts its text to a date. This stays with ADVFN. The 526 exchange-qualified AAPL field check returned 2026-09-26. |
| **ADVFN** — `smfGetADVFNElement`, latest annual Year End Date only | `https://<prefix>.advfn.com/stock-market/<exchange>/<symbol>/financials?&mode=annual_reports&btn=start_date` | `https://<prefix>.advfn.com/p.php?pid=financials&symbol=<symbol>` | Keeps the configured ADVFN regional host and reads a recognized annual-table header. It validates the displayed date and rightmost fiscal year. Only annual period, cell1 and the recognized year-end labels use this path. Two selected live calls returned 2026-09-26; controlled cases also passed in the 526 checkpoint. Earlier periods and other metrics retain the previous extractor. |
| **Yahoo history** — `smfGetYahooHistory`, `smfGetYahooHistoryCSV`, `smfGetYahooHistoryCSVData`, and dependent history functions | `https://query1.finance.yahoo.com/v7/finance/download/<ticker>?period1=<unix_start>&period2=<unix_end>&interval=<interval>&events=<event>`; query2 download URLs are also recognized | `https://query1.finance.yahoo.com/v8/finance/chart/<ticker>?period1=<unix_start>&period2=<unix_end>&interval=<interval>&events=div%2Csplits&includeAdjustedClose=true` | The Mac adapter converts chart JSON back to the CSV format the callers expect. The ticker, range, interval and requested event remain caller inputs. History/date/volume, array sizing and dependent-function checks passed for selected inputs. The 526 run records 33 passing cases under `smfGetYahooHistory`; this does not establish every historical event or ticker. |
| **Yahoo history** — `RCHGetYahooHistory` / `RCHGetYahooHistory2` compatibility route | The active upstream `RCHGetYahooHistory2` request hard-codes `.../v7/finance/download/MMM?period1=1493610466&period2=1496202466&interval=1d&events=history&crumb=` | Builds a download request from the actual ticker/dates/period, then uses the same v8 chart adapter | Removes that hard-coded example from the Mac path and corrects ordering/column handling. Selected native compatibility/history checks passed. The Windows branch retains its upstream request. |
| **Yahoo options** — `smfGetYahooOptionQuote1` | `https://finance.yahoo.com/quote/<option_symbol>` HTML | `https://query1.finance.yahoo.com/v10/finance/quoteSummary/<option_symbol>?modules=summaryDetail` or `?modules=price` | The Mac compatibility function delegates to the existing JSON-based `smfGetYahooOptionQuote`. The provider remains Yahoo. Percentage output remains a fraction; expiration remains Yahoo-formatted text. Controlled field/argument cases passed; the 526 live bid check returned 6.45. Other contracts and market sessions need separate checks. |
| **Yahoo JSON and option chains** — existing JSON functions, expiration/strike functions | Existing `.../v10/finance/quoteSummary/<symbol>?modules=<module>` and `.../v7/finance/options/<ticker>?date=<unix_expiry>` routes | Same routes | The Mac port changes transport/JSON parsing and cache handling. These URLs were already in upstream code. Selected quote, expiration and strike checks passed; this does not establish every JSON module or every option price. Retired named modes are listed separately below. |
| **Yahoo anonymous session setup** — helper used by selected Yahoo JSON calls | HTML cookie/crumb discovery through `https://finance.yahoo.com/lookup?s=%7B0%7D` | `https://fc.yahoo.com`, then `https://query1.finance.yahoo.com/v1/test/getcrumb` | Uses temporary anonymous cookies and a crumb. These are authentication helpers, not financial-data replacements. They do not require a personal Yahoo login. Public reports redact session values. |
| **FRED** — `smfGetEconData` | `https://fred.stlouisfed.org/data/<series_id>` | `https://api.stlouisfed.org/fred/series/observations?series_id=<series_id>&file_type=json&observation_start=<date>&observation_end=<date>&sort_order=desc&limit=1&api_key=<API_KEY>` | The Mac path uses the official JSON API. Exact dates, `.` missing observations and caller error values remain explicit. Five selected live checks, including DGS10/DGS2 and metadata, passed in the 526 checkpoint. The user supplies a private key file beside the loaded workbook/add-in. |
| **FRED metadata** — `smfGetEconData` text selectors | The same public `/data/<series_id>` HTML page | `https://api.stlouisfed.org/fred/series?series_id=<series_id>&file_type=json&api_key=<API_KEY>`; `/series/release?series_id=<series_id>&file_type=json&api_key=<API_KEY>`; `/release/sources?release_id=<release_id>&file_type=json&api_key=<API_KEY>` | Keeps the requested metadata selector while reading API fields. The source, units and controlled missing-field/newline cases have separate checks. No key is included in the bundle, formulas or this file. |
| **Yahoo RSI inputs** — `smfRSITargetPrices` | `RCHGetYahooQuotes` uses retired `https://download.finance.yahoo.com/d/quotes.csv?s=<ticker>&f=sl1baoghvd1t1&e=.ignore`, alongside the Yahoo download-history route | `https://query1.finance.yahoo.com/v8/finance/chart/<ticker>?range=3mo&interval=1d&events=div%2Csplits&includeAdjustedClose=true`; optional bid/ask use `https://query1.finance.yahoo.com/v7/finance/quote?fields=symbol,regularMarketTime,bid,ask&formatted=false&symbols=<ticker>` | Removes the retired quote-CSV dependency on Mac. Calculates Wilder RSI(2) from 21 prior unadjusted closes and appends the quote once by exchange session. The 526 checkpoint records 65 passing cases, including two live core calls. Those core live calls do not establish optional bid/ask availability or independently verify every target. |
| **Google history** — `smfGetGoogleHistoryCSV` | `http://finance.google.com/finance/historical?output=csv&q=<ticker>&startdate=<MMM d, yyyy>&enddate=<MMM d, yyyy>` with literal spaces/comma in the query | `http://finance.google.com/finance/historical?output=csv&q=<encoded_ticker>&startdate=<encoded_date>&enddate=<encoded_date>` | The 526 case failed with a same-case malformed-URL/curl3 error. The proposed repair percent-encodes ticker/date values. The existing `pPeriod` argument remains unused. New native tests remain pending. A separate encoded request reached HTTP 200 after redirects but returned the current Finance HTML page, not historical CSV. No replacement feed is established. |

The four shared definition changes are fields 2727, 849, 19004 and 5196.
Their text files also load on Windows. Most transport/function changes use Mac compiler branches; Windows runtime verification remains pending.
The Finviz URLs did not change.

## Chart and picture sources

| Provider / function | Old URL | New Mac URL or behavior | Status / native evidence |
|---|---|---|---|
| **StockCharts**, `RCHCreateComment` choice 1 | `https://c.stockcharts.com/c-sc/sc?chart=<ticker>,uu[h,a]daclyyay[pb50!b200!f][vc60][iue12,26,9!lc20]` | `https://stockcharts.com/c-sc/sc?chart=<ticker>,uu[h,a]daclyyay[pb50!b200!f][vc60][iue12,26,9!lc20]` | Same provider and preset; Mac host changes. This preset needs its own live check. |
| **StockCharts**, choice 2 | `https://stockcharts.com/def/servlet/SharpChartv05.ServletDriver?chart=<ticker>,pltad[pa][da][f!3!!]&pnf=y` | `https://stockcharts.com/pnf/chart?c=<ticker>,P` | Replaces the old point-and-figure servlet route. Selected native image checks passed. |
| **StockCharts**, choice 3 | `https://c.stockcharts.com/c-sc/sc?chart=<ticker>,uu[305,a]dacayaci[pb20!b50][dc]` | `https://stockcharts.com/c-sc/sc?chart=<ticker>,uu[305,a]dacayaci[pb20!b50][dc]` | Same preset on the Mac host. This preset needs its own live check. |
| **StockCharts**, choice 4 | `https://c.stockcharts.com/c-sc/sc?s=<ticker>&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806` | `https://stockcharts.com/c-sc/sc?s=<ticker>&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806` | Selected native PNG, visibility and reopen checks passed in earlier builds. The old code comment mentions Business Week, but the actual URL is StockCharts. |
| **StockCharts**, choice 5 | `https://c.stockcharts.com/c-sc/sc?s=<ticker>&p=D&yr=0&mn=6&dy=0&i=t39628903145&r=9933` | `https://stockcharts.com/c-sc/sc?s=<ticker>&p=D&yr=0&mn=6&dy=0&i=t39628903145&r=9933` | Same provider/preset on the Mac host. This preset needs its own live check. |
| **Ogre's Crypt**, choice 98 | `http://ogres-crypt.com/php/chart.php?d=<normalized_values>` | No HTTP URL: an embedded native Excel line chart | The Mac code renders the numeric data locally. Selected normalization/reuse/visibility cases passed. This does not declare the external website retired. |
| **ADVFN**, choice 97 | `https://www.advfn.com/p.php?pid=financialgraphs&a0=<n>&a1=<n>&a2=<n>&a3=<n>&a4=<n>` | Same URL | Mac embedding differs from Windows comment-picture fills. No provider replacement or selected live success for this preset is claimed. |
| **Caller-supplied pictures**, choice 99 | Caller-supplied URL or file | Same caller-supplied source, through the Mac image helper | Selected remote/local image cases passed. This is not an exhaustive list of image hosts. |

Mac pictures use embedded worksheet objects plus a text note.
Numeric charts use no image-download request. Image downloads use a separate helper from text-fetch bridge counters.

## Confirmed retired modes: no automatic replacement

The [registry](mac/not_supported.json) records six groups. The Mac financial functions return typed Excel `#N/A` for those registered modes.
The status is **Not Supported**, with its reason and suggested alternatives in [NotSupported.csv](mac/NotSupported.csv).
Suggested alternatives do not establish that a replacement was implemented for the retired mode.

| Provider / affected functions | Old URL or mode | New URL | Status / reason |
|---|---|---|---|
| **MSN options** — `smfGetMSNOptionQuote`, relevant `smfGetOptionQuotes` mode | `http://investing.money.msn.com/investments/equity-options/?symbol=<ticker>...` | None | Upstream marks this quote function obsolete from 2014-12-23. Mac returns `#N/A` before a financial request. |
| **OptionsXpress** — `smfGetOXOptionQuote`, `smfGetOX2OptionQuote`, `smfGetOX3OptionQuote`, and their wrapper/expiration/strike modes | `https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?<chain_parameters>` | None | Upstream records OptionsXpress as no longer a valid source from 2017-10-10. Mac returns `#N/A`. |
| **Yahoo legacy quote CSV** — `RCHGetYahooQuotes` ticker/quote-code mode | `https://download.finance.yahoo.com/d/quotes.csv?s=<tickers>&f=s<codes>&e=.ignore`; the old function also had configurable server variants | None for that mode | Upstream marks the code mode obsolete. Supplied CSV-URL mode with empty `pItems` remains available, except registered retired URLs. |
| **Yahoo YQL** — `smfYahooAPIData`, feeds A–F | `https://query.yahooapis.com/v1/public/yql?<query>` for feed A; `http://query.yahooapis.com/v1/public/yql?<query>` for feeds B–F | None | The upstream change log records YQL retirement on 2019-01-03. All Mac YQL feeds return `#N/A`. |
| **Google legacy option chain** — `smfGetGoogleOptionQuote` and G/GOOGLE wrapper/expiration/strike modes | `http://www.google.com/finance/option_chain?output=json&q=<ticker>...` | None | The option-strike manual explicitly calls this mode obsolete. This entry concerns option chains, not all Google Finance services or the separate history repair. |
| **OIC legacy option chain** — `smfGet888OptionQuote` and 8/888 modes | `http://oic.ivolatility.com/oic_adv_options.j?exp_date=<expiry>&ticker=<ticker>` | None | The manual marks the configured legacy source 8 obsolete. Other iVolatility services remain outside this guard. |

The native 526 checkpoint records 44 expected Not Supported guard cases.
A passing guard proves the specified error behavior and measured absence of a text-fetch bridge attempt. It does not establish a working financial feed.
Mixed wrappers retain supported modes. The existing B/N expiration/strike aliases still use Yahoo; this fork does not establish working Barchart/Nasdaq feeds through those aliases.

### Exact registered URL routes

The page/financial download guards match these ten host/path pairs, for HTTP or HTTPS.
They compare host and path without case sensitivity; query parameters do not change the match.
Other paths and lookalike domains remain outside these exact rules.

| Registry group | Host | Exact path |
|---|---|---|
| msn-options | `investing.money.msn.com` | `/investments/equity-options/` |
| msn-options | `investing.money.msn.com` | `/investments/equity-options` |
| optionsxpress-options | `www.optionsxpress.com` | `/OXNetTools/Chains/index.aspx` |
| optionsxpress-options | `optionsxpress.com` | `/OXNetTools/Chains/index.aspx` |
| yahoo-quote-csv | `download.finance.yahoo.com` | `/d/quotes.csv` |
| yahoo-yql | `query.yahooapis.com` | `/v1/public/yql` |
| yahoo-yql | `query.yahooapis.com` | `/v1/public/yql/` |
| google-option-chain | `www.google.com` | `/finance/option_chain` |
| google-option-chain | `finance.google.com` | `/finance/option_chain` |
| oic-legacy-options | `oic.ivolatility.com` | `/oic_adv_options.j` |

### 288 obsolete default fields

The status CSV contains six provider-mode rows and 288 explicitly obsolete default-definition rows:294 data rows in total.
Those288 rows are numbered fields, not 288 providers.
The runtime guard follows the effective loaded definition, so a valid custom replacement can supply a working recipe.
Metadata selectors remain readable. Premium-only and uncertain recipes remain outside the approved obsolete-literal list.
The native bulk assertion is one case over the listed 288 definitions; field 625 also has worksheet/IFNA cases.

## Unresolved access and unmodified sources

HTTP401/403, a WAF challenge, a timeout or a missing marker does not by itself establish retirement.
Reuters, Barchart, Nasdaq, MarketWatch and other protected/unavailable paths were not silently replaced with another provider.
The earlier [data-repair notes](mac/DATA-PICTURE-REPAIRS-20261108.md#seven-fields-still-blocked-or-unavailable) retain their observations and unchanged recipes.
A tested candidate URL, such as Reuters key-metrics, is not a configured replacement unless the code or definition changed.

ETFScreen separately [announces that it shut down after August 22, 2025](https://www.etfscreen.com/index.html).
Its failed relative-strength recipe has no replacement in this fork; that announcement is separate from the six guarded source groups.
The earlier ETFInvestmentOutlook capture redirected to an unrelated accounting site. No replacement or owner-confirmed retirement is inferred from that capture alone.

## Evidence and source files

The URL comparison uses upstream `a0daad0` and the Mac history through `c4852c1`, plus the Google query-encoding change tested in an isolated native13-case run:

- [Four shared field definitions](mac/DATA-PICTURE-REPAIRS-20261108.md#four-recipe-changes): [GuruFocus](docs/Elements/smf-elements-9.txt), [Zacks](docs/Elements/smf-elements-6.txt), [Netflix](docs/Elements/smf-elements-20.txt), [ADVFN](docs/Elements/smf-elements-7.txt)
- [Unchanged Finviz definitions](docs/Elements/smf-elements-4.txt) and [Mac transport/parser](code/modMacCode.bas)
- [Yahoo history adapter](code/smfGetYahooHistoryCSV_.bas), [history callers](code/smfGetYahooHistory_.bas), [compatibility history](code/modGetYahooHistory.bas)
- [Yahoo option functions](code/modGetOptionQuotes.bas), [JSON endpoint builder](code/modGetYahooJSONData.bas), [session/image helper](mac/SMFMac.applescript)
- [FRED API adapter](code/modGetEconData.bas), [RSI inputs](code/modRSITargetPrices.bas), [ADVFN year-end parser](code/modGetAdvFNElement.bas)
- [Image/chart URL choices](code/smfCreateComment_.bas), [Google-history function](code/smfGetGoogleHistoryCSV_.bas)
- [Retirement registry](mac/not_supported.json), [generated status CSV](mac/NotSupported.csv), [registry/CSV generator](mac/build_not_supported.py)

The 526 checkpoint's saved workbook SHA256 is `59bca5fbf8ebd7c9e982703c69f1da04c1ad10778c0e334e233d5abf384983ac`.
Its source fingerprint is `7d0079bf7a1c6787dabd22d0a8b406a8eb3d9d7fa9b13367fd0d07f7e09d3212`.
Keep that checkpoint separate from the corrected-fixture/final combined runs and the next package's native receipts.

The isolated Google run saved12 PASS and1 CHECK_FAILED. The failed case is the financial-data assertion. Its saved workbook SHA256 is `8cb967bcfad9c4e3a7b5dc6f5a11b723f1815164241b12f82741f029559d4f6f`; source fingerprint `e8d82e4755d8d74dda22219cdf9f16f0d46a0b77b70cba3ea88b23ee6e3faaac`. The URL checks prove request construction and CSV/error handling, without establishing a working Google history feed.
