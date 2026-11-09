# Provider changes in the native Mac port

This file records changes from upstream commit `a0daad05699b1a7e213cf72ee641f30fc615af46` through Mac source commit `e7d788b4f0871e94dbe6e13131ca8d9b7ff0eb8a`.

The experimental [0.7-dev bundle](mac/built/0.7-dev/README.md) includes the final combined source.
The [full native run](mac/test-results/0.7-dev-20261109-final/full/README.md) recorded **538 cases: 502 PASS and 36 CHECK_FAILED**, including 51 passing worksheet-entry cases.
The separate focused native run recorded **246 PASS**, including 19 passing worksheet-entry cases.
Those counts include controlled fixtures and expected retired-mode errors. They do not count working providers or independently verified financial fields.
The final runs used the corrected ADVFN fixture and the repaired Google-history URL. Saved-workbook source audits passed.
Google's 12 URL/CSV regression checks passed. Its separate financial-data assertion still failed because Google returned Finance HTML instead of historical CSV.
Earlier 526-case observations below remain labeled as earlier evidence. The 0.6-dev bundle remains available for rollback.

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
| **ADVFN** — `smfGetADVFNElement`, latest annual Year End Date only | `https://<prefix>.advfn.com/stock-market/<exchange>/<symbol>/financials?&mode=annual_reports&btn=start_date` | `https://<prefix>.advfn.com/p.php?pid=financials&symbol=<symbol>` | Keeps the configured ADVFN regional host and reads a recognized annual-table header. It validates the displayed date and rightmost fiscal year. Only annual period, cell 1 and the recognized year-end labels use this path. Two selected live calls returned 2026-09-26; controlled cases also passed in the 526 checkpoint. Earlier periods and other metrics retain the previous extractor. |
| **Yahoo history** — `smfGetYahooHistory`, `smfGetYahooHistoryCSV`, `smfGetYahooHistoryCSVData`, and dependent history functions | `https://query1.finance.yahoo.com/v7/finance/download/<ticker>?period1=<unix_start>&period2=<unix_end>&interval=<interval>&events=<event>`; query2 download URLs are also recognized | `https://query1.finance.yahoo.com/v8/finance/chart/<ticker>?period1=<unix_start>&period2=<unix_end>&interval=<interval>&events=div%2Csplits&includeAdjustedClose=true` | The Mac adapter converts chart JSON back to the CSV format the callers expect. The ticker, range, interval and requested event remain caller inputs. History/date/volume, array sizing and dependent-function checks passed for selected inputs. The 526 run records 33 passing cases under `smfGetYahooHistory`; this does not establish every historical event or ticker. |
| **Yahoo history** — `RCHGetYahooHistory` / `RCHGetYahooHistory2` compatibility route | The active upstream `RCHGetYahooHistory2` request hard-codes `.../v7/finance/download/MMM?period1=1493610466&period2=1496202466&interval=1d&events=history&crumb=` | Builds a download request from the actual ticker/dates/period, then uses the same v8 chart adapter | Removes that hard-coded example from the Mac path and corrects ordering/column handling. Selected native compatibility/history checks passed. The Windows branch retains its upstream request. |
| **Yahoo options** — `smfGetYahooOptionQuote1` | `https://finance.yahoo.com/quote/<option_symbol>` HTML | `https://query1.finance.yahoo.com/v10/finance/quoteSummary/<option_symbol>?modules=summaryDetail` or `?modules=price` | The Mac compatibility function delegates to the existing JSON-based `smfGetYahooOptionQuote`. The provider remains Yahoo. Percentage output remains a fraction; expiration remains Yahoo-formatted text. Controlled field/argument cases passed; the 526 live bid check returned 6.45. Other contracts and market sessions need separate checks. |
| **Yahoo JSON and option chains** — existing JSON functions, expiration/strike functions | Existing `.../v10/finance/quoteSummary/<symbol>?modules=<module>` and `.../v7/finance/options/<ticker>?date=<unix_expiry>` routes | Same routes | The Mac port changes transport/JSON parsing and cache handling. These URLs were already in upstream code. Selected quote, expiration and strike checks passed; this does not establish every JSON module or every option price. Retired named modes are listed separately below. |
| **Yahoo anonymous session setup** — helper used by selected Yahoo JSON calls | HTML cookie/crumb discovery through `https://finance.yahoo.com/lookup?s=%7B0%7D` | `https://fc.yahoo.com`, then `https://query1.finance.yahoo.com/v1/test/getcrumb` | Uses temporary anonymous cookies and a crumb. These are authentication helpers, not financial-data replacements. They do not require a personal Yahoo login. Public reports redact session values. |
| **FRED** — `smfGetEconData` | `https://fred.stlouisfed.org/data/<series_id>` | `https://api.stlouisfed.org/fred/series/observations?series_id=<series_id>&file_type=json&observation_start=<date>&observation_end=<date>&sort_order=desc&limit=1&api_key=<API_KEY>` | The Mac path uses the official JSON API. Exact dates, `.` missing observations and caller error values remain explicit. Five selected live checks, including DGS10/DGS2 and metadata, passed in the 526 checkpoint. The user supplies a private key file beside the loaded workbook/add-in. |
| **FRED metadata** — `smfGetEconData` text selectors | The same public `/data/<series_id>` HTML page | `https://api.stlouisfed.org/fred/series?series_id=<series_id>&file_type=json&api_key=<API_KEY>`; `/series/release?series_id=<series_id>&file_type=json&api_key=<API_KEY>`; `/release/sources?release_id=<release_id>&file_type=json&api_key=<API_KEY>` | Keeps the requested metadata selector while reading API fields. The source, units and controlled missing-field/newline cases have separate checks. No key is included in the bundle, formulas or this file. |
| **Yahoo RSI inputs** — `smfRSITargetPrices` | `RCHGetYahooQuotes` uses retired `https://download.finance.yahoo.com/d/quotes.csv?s=<ticker>&f=sl1baoghvd1t1&e=.ignore`, alongside the Yahoo download-history route | `https://query1.finance.yahoo.com/v8/finance/chart/<ticker>?range=3mo&interval=1d&events=div%2Csplits&includeAdjustedClose=true`; optional bid/ask use `https://query1.finance.yahoo.com/v7/finance/quote?fields=symbol,regularMarketTime,bid,ask&formatted=false&symbols=<ticker>` | Removes the retired quote-CSV dependency on Mac. Calculates Wilder RSI(2) from 21 prior unadjusted closes and appends the quote once by exchange session. The 526 checkpoint records 65 passing cases, including two live core calls. Those core live calls do not establish optional bid/ask availability or independently verify every target. |
| **Google history** — `smfGetGoogleHistoryCSV` | `http://finance.google.com/finance/historical?output=csv&q=<ticker>&startdate=<MMM d, yyyy>&enddate=<MMM d, yyyy>` with literal spaces/comma in the query | `http://finance.google.com/finance/historical?output=csv&q=<encoded_ticker>&startdate=<encoded_date>&enddate=<encoded_date>` | The 526 case failed with a same-case malformed-URL/curl3 error. The repair percent-encodes ticker/date values. The existing `pPeriod` argument remains unused. All 12 native regression checks passed in the final source. The financial-data assertion still failed: the encoded request reached HTTP 200 after redirects but returned the current Finance HTML page, not historical CSV. No replacement feed is established. |

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

The status CSV contains six provider-mode rows and 288 explicitly obsolete default-definition rows: 294 data rows in total.
Those 288 rows are numbered fields, not 288 providers.
The runtime guard follows the effective loaded definition, so a valid custom replacement can supply a working recipe.
Metadata selectors remain readable. Premium-only and uncertain recipes remain outside the approved obsolete-literal list.
The native bulk assertion is one case over the listed 288 definitions; field 625 also has worksheet/IFNA cases.

## Current unavailable or unverified provider data

The final combined Excel run on November 9, 2026 saved 538 cases: 502 PASS and 36 CHECK_FAILED. All 51 worksheet-entry cases passed. The source commit is `e7d788b4f0871e94dbe6e13131ca8d9b7ff0eb8a`.
The table below lists the problematic paths, including providers with selected working fields. It does not label every service from the same provider unavailable.
HTTP 200 means a page responded. It does not establish that SMF extracted the requested financial data. HTTP 401/403 alone does not establish a subscription requirement.

| Provider | Unavailable or unverified data / URL path | Observed result | Selected working path or limit |
|---|---|---|---|
| **Barchart** | `core-api.barchart.com/v1/options/chain`, `/v1/momentum/get`; `www.barchart.com/stocks/quotes/<ticker>/opinion` | API timeouts; opinion page HTTP 202 without the expected data | No working native option/momentum extraction was established through these paths. The existing B/N expiry/strike aliases use Yahoo. |
| **BigCharts / MarketWatch** | `bigcharts.marketwatch.com/quickchart/options.asp`; `www.marketwatch.com/investing/stock/<ticker>/options` and `/investing/index/<ticker>/options` | BigCharts HTTP 403; MarketWatch HTTP 401 | Option data unavailable in the tested requests. |
| **MSN Money** | `data.moneycentral.msn.com/scripts/chrtsrv.dll`; `investing.money.msn.com/investments/stock-price` | Both legacy hosts failed DNS resolution | These history/stock-price paths remain separate from the guarded retired MSN option mode. |
| **Morningstar** | `performance.morningstar.com/perform/Performance/stock/exportStockPrice.action` | Connection timeout | Historical CSV data unavailable in the tested call. |
| **TD Ameritrade** | `research.tdameritrade.com/public/stocks/overview/overview.asp` | Legacy host failed DNS resolution | Stock-overview extraction unavailable in the tested call. |
| **Nasdaq** | `www.nasdaq.com/symbol/<ticker>/option-chain/<contract>` | HTTP 403 | Legacy option-chain extraction unavailable. Yahoo redirects used by some aliases are separate. |
| **Reuters** | `www.reuters.com/finance/stocks/financial-highlights/<ticker>` | HTTP 401 | Selected financial-field extraction failed. No replacement Reuters feed is configured. |
| **Google Finance** | `finance.google.com/finance/historical`; `finance.google.com/finance?fstype=ii&q=<ticker>` | HTTP 200 HTML, without the requested historical CSV or financial-statement data | The URL-encoding repair passed 12 native regression checks. The historical-data assertion still failed. |
| **ETFScreen** | `www.etfscreen.com/rsftrends.php`; `/etfperf.php` | Requested ranking/relative-strength data missing; parked/content pages | The official index announces shutdown after August 22, 2025. This provider is not yet in the six guarded source groups. |
| **ETF Investment Outlook** | `www.etfinvestmentoutlook.com/rank_breadth_new_highs_lows_avg_d.php` | HTTP 200 with unrelated content; ranking extraction failed | No replacement feed was implemented. A redirect does not establish owner-confirmed retirement. |
| **IBD / Investors.com** | `research.investors.com/stock-checkup/xxxx-<ticker>.aspx` | Page responded, but the requested stock-checkup field returned Error | Page access alone does not establish working ratings/stock-checkup extraction. |
| **GuruFocus** | `www.gurufocus.com/financials/<ticker>`; `/download_financials_in_CSV.php?symbol=<ticker>` | Revenue-per-share lookup returned Error on HTTP 200; financial CSV HTTP 404 | Selected forward-dividend-yield calls through `/stock/<ticker>/dividend` passed. |
| **Yahoo** | `finance.yahoo.com/q/op`; direct `/v7/finance/download/<ticker>` CSV probes | Legacy HTML option lookup HTTP 404; direct historical CSV probes HTTP 401 | Selected JSON quote/option calls and the v8-chart history adapter passed. These results do not establish every contract or JSON field. |
| **FRED** | `fred.stlouisfed.org/data/<series_id>` | Old public download route timed out | The official `api.stlouisfed.org/fred/` path passed selected observation/metadata checks with the private API key. |
| **ADVFN** | General legacy annual-financials page content/marker check | HTTP 200 page did not pass that generic content check; other financial fields remain unverified | The repaired latest annual year-end lookup passed selected live and controlled cases. |
| **Zacks** | `www.zacks.com/stock/quote/<ticker>` generic page-content check | HTTP 200 page failed the expected marker check | Selected field 849 price-target extraction through `/stock/research/<ticker>/price-target-stock-forecast` passed. |
| **StockScores** | `www.stockscores.com/charts/charts?ticker=<ticker>` generic page-content check | HTTP 200 page failed the expected marker check | Selected field 785, **20 Day Avg Vol**, passed. This was a volume result, not verification of a stock score. |
| **Netflix** | `dvd.netflix.com/Movie/<title_id>` | Legacy DVD-page request did not supply the expected movie-year content | The replacement `www.netflix.com/title/<title_id>` recipe passed the selected movie-year check. |

These 36 failed checks include repeated requests and functions for the same provider. They are not 36 distinct failed providers. The 44 passing Not Supported cases are counted separately from these failures.

## Unresolved access and unmodified sources

HTTP 401/403, a WAF challenge, a timeout or a missing marker does not by itself establish retirement.
Reuters, Barchart, Nasdaq, MarketWatch and other protected/unavailable paths were not silently replaced with another provider.
The earlier [data-repair notes](mac/DATA-PICTURE-REPAIRS-20261108.md#seven-fields-still-blocked-or-unavailable) retain their observations and unchanged recipes.
A tested candidate URL, such as Reuters key-metrics, is not a configured replacement unless the code or definition changed.

ETFScreen separately [announces that it shut down after August 22, 2025](https://www.etfscreen.com/index.html).
Its failed relative-strength recipe has no replacement in this fork; that announcement is separate from the six guarded source groups.
The earlier ETFInvestmentOutlook capture redirected to an unrelated accounting site. No replacement or owner-confirmed retirement is inferred from that capture alone.

## Evidence and source files

The URL comparison uses upstream `a0daad0` and the Mac history through `c4852c1`, plus the Google query-encoding change tested in an isolated native 13-case run:

- [Four shared field definitions](mac/DATA-PICTURE-REPAIRS-20261108.md#four-recipe-changes): [GuruFocus](docs/Elements/smf-elements-9.txt), [Zacks](docs/Elements/smf-elements-6.txt), [Netflix](docs/Elements/smf-elements-20.txt), [ADVFN](docs/Elements/smf-elements-7.txt)
- [Unchanged Finviz definitions](docs/Elements/smf-elements-4.txt) and [Mac transport/parser](code/modMacCode.bas)
- [Yahoo history adapter](code/smfGetYahooHistoryCSV_.bas), [history callers](code/smfGetYahooHistory_.bas), [compatibility history](code/modGetYahooHistory.bas)
- [Yahoo option functions](code/modGetOptionQuotes.bas), [JSON endpoint builder](code/modGetYahooJSONData.bas), [session/image helper](mac/SMFMac.applescript)
- [FRED API adapter](code/modGetEconData.bas), [RSI inputs](code/modRSITargetPrices.bas), [ADVFN year-end parser](code/modGetAdvFNElement.bas)
- [Image/chart URL choices](code/smfCreateComment_.bas), [Google-history function](code/smfGetGoogleHistoryCSV_.bas)
- [Retirement registry](mac/not_supported.json), [generated status CSV](mac/NotSupported.csv), [registry/CSV generator](mac/build_not_supported.py)

The 526 checkpoint's saved workbook SHA256 is `59bca5fbf8ebd7c9e982703c69f1da04c1ad10778c0e334e233d5abf384983ac`.
Its source fingerprint is `7d0079bf7a1c6787dabd22d0a8b406a8eb3d9d7fa9b13367fd0d07f7e09d3212`.
Keep that earlier checkpoint separate from the final 538-case and 246-case runs and the 0.7 package receipts.

The isolated Google run saved12 PASS and1 CHECK_FAILED. The failed case is the financial-data assertion. Its saved workbook SHA256 is `8cb967bcfad9c4e3a7b5dc6f5a11b723f1815164241b12f82741f029559d4f6f`; source fingerprint `e8d82e4755d8d74dda22219cdf9f16f0d46a0b77b70cba3ea88b23ee6e3faaac`. The URL checks prove request construction and CSV/error handling, without establishing a working Google history feed.

The final538-case workbook SHA256 is `392139b5c9de34eaf7399b8860e4b2203ca9c09b8d9eeed64ff557ea0bf8fd93`; source commit `e7d788b4f0871e94dbe6e13131ca8d9b7ff0eb8a`. Its full-owner source audit passed: all 51 core modules match the clean workbook and all executable test strings match the generated runner.
