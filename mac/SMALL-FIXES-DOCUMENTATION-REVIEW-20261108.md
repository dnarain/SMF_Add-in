# SMF documentation review and small Mac repairs

The documentation describes several workflows that matter when interpreting the Mac tests.
It also contains old examples and Windows request methods that the Mac downloader cannot reproduce.
This review compares those instructions with the saved 0.5 tests and the two Mac 0.6 repairs.
The repairs change history defaults and `smfGetYahooOptionQuote1` retrieval.
Windows active code remains unchanged. Windows runtime checks were not performed.

## Native repair checks

Excel completed the first focused 53-case run with 49 passes and four failed assertions.
All 33 added option-compatibility cases passed. The four failed assertions belonged to the history array-caller cases.
`Range.Value2` returned numeric date serials, while the test helper `DateNumber` rejected those numbers.
For example, November 6, 2026 appeared as the `Double` value `46332`.

Only those four assertions changed. They now require `Double` values and compare their serials with `DateSerial`.
They reject text, errors and empty cells. The production function and shared helper did not change.
Headers, prices, volume, dimensions and existing array-entry assertions remain checked.

The completed full run shows all 20 added history-default cases and all 33 added option-compatibility cases passing.
The history set contains 16 offline cases and four live cases. The option set contains 28 controlled cases and five live cases.
The first full run completed 292 cases: 238 PASS and 54 CHECK_FAILED. It recorded zero native phase errors or unfinished cases.
All 32 worksheet-entry cases passed within that batch. The separate positive-bid checks still failed on numeric zero.
The initial failures remain preserved in the local `SMF-Focused-0.6-first-run.xlsm` and `focused-first-run-results.private.json`.
Raw workbooks and private result files remain local.

These selected checks do not establish every ticker, provider field, locale or Excel version.
The option numeric assertions include locale-dependent conversions; the native results apply to the tested Mac settings.

## Later cold-reopen failure

Two subsequent cold-reopen checks of the data-and-picture example restarted Excel.
The diagnostic logs record ungraceful exits and terminal chart-insertion failures.
The new process log records `Office.Charting.InsertChartE2o`, `Success=false`, HRESULT `-2147221504` and `ChartType=51` immediately before its final chart error.
The logs do not provide a decoded crash stack that identifies the exact VBA line.

A control copy disabled only the B16 numeric-chart formula and kept A16 live.
Excel reopened that copy with Automatic calculation and Ready status.
The revised chart-reuse candidate passed nine focused native cases. The final 292 rerun completed 238 PASS and 54 CHECK_FAILED. Cold-add-in reopens completed with the preserved 10/11 data/picture result.
The first numeric-reuse focus recorded five PASS and four failed assertions. Its unchanged-formula calls could retain cached results.
The revised cases change return markers, mutate state and force calculation before checking reuse. The original focus remains preserved.
The two small repairs and the documentation review are complete. The experimental 0.6 build is prepared for fork testing during the chart investigation.
The installed 0.5 add-in remains unchanged and is unloaded for the investigation.

The first 292-case results remain unchanged. The later reopen failures require separate records and a later candidate needs separate native checks.

## History fields and output ranges

The [history parameter table](../docs/Documentation/doc-links/smfGetYahooHistory-Function_files/smfGetYahooHistory-Function.htm#L56) documents `DOHLCVU` as the default field list.
The 0.5 source instead declares `dohlcvufgxs`. Its price-history validator rejects `F/G`.
Mac 0.6 uses `dohlcvu` for both omitted and explicitly empty field lists.
The native checks confirm the documented column order: Date, Open, High, Low, AdjClose, Volume, Close.
`C` supplies adjusted close; `U` supplies unadjusted close.
The cached fixture gives adjusted and unadjusted closes different values, so swapped close columns fail the check.

The same document contains conflicting older examples. One example lists twelve default items, including `F/G`.
It also calls period `s` semiannual, while the current code uses `S` for splits.
The parameter table and those examples cannot all describe the current implementation.

The `RCHGetYahooHistory` manual also conflicts with its own example.
Its parameter table gives `DOHLCVA` and labels `U` adjusted; its seven-field example lists `DOHLCVU`.
The current wrapper defaults to `DOHLCVU`, with `U` mapped to unadjusted close.
This repair preserves the existing Windows behavior.

The [array-entry guide](<../docs/Documentation/doc-links/arrayEnterFormula-files/Array Entered Formulas for Excel.htm>) uses a preselected output range.
The history instructions explain that oversized ranges contain blanks and undersized ranges truncate returned data.
The original guide uses Excel 2010 and Windows shortcuts.
The Mac checks include two cached 3-by-7 and two live 6-by-7 `FormulaArray` callers.
The live checks omit row and column arguments and verify `HasArray=True` and the exact 6-by-7 `CurrentArray` range.

Most calls in the earlier 239-case suite used VBA; market-array cases generally supplied explicit dimensions.
Its nine worksheet-history regressions mainly check invalid-period diagnostics, `None` and caller dimensions.
The repair cases add seven worksheet-history calls, including the four multi-cell array callers.
The generator's existing `worksheet_history_case_count` counts only the nine legacy row-count cases.
There are now 16 history worksheet cases across those two sets.
Other documented examples still require checks.

An incorrect output range can explain one-cell results or truncation.
It cannot explain a recorded HTTP 401, 403, 404, failed DNS lookup or malformed request URL.

## Numbered fields and their sources

[`RCHGetElementNumber`](../docs/Documentation/doc-links/RCHGetElementNumber-Function_files/RCHGetElementNumber-Function.htm) retrieves a numbered field from its configured source.
These calls apply the documented diagnostics to Finviz field 4072; the manual examples use field 941:

```excel
=RCHGetElementNumber("Source", 4072)
=RCHGetElementNumber("Element", 4072)
=RCHGetElementNumber("Web Page", 4072)
=RCHGetElementNumber("ElementFile", 4072)
=RCHGetElementNumber("P-Type", 4072)
```

The configured definition determines the provider, extraction method and requested field.
For formula-based definitions, the recorded web-page value can contain the extraction expression.
A successful field does not establish the other fields from that provider.

The old documentation's market-cap examples use different numeric scales for different providers.
A positive-number check therefore does not establish correct units or financial meaning.
The next value checks must compare the ticker, period, units and returned field with the corresponding provider response.

## Cache clearing and recalculation

The [force-recalculation instructions](../docs/Documentation/doc-links/smfForceRecalculation-Macro_files/smfForceRecalculation-Macro.htm) describe three actions:

- Purge the saved web pages.
- Reload the numbered-field definitions.
- Recalculate Excel.

`smfForceRecalculation` performs those actions in `code/smfUtilities_.bas`.
The documentation says add-in macros can require typing the macro name in Excel's Macro dialog.
It also describes the context-menu and toolbar routes.

A normal formula recalculation can reuse SMF's cached page.
The legacy Yahoo quote function also documents a separate `NOW()` refresh argument.
That argument does not provide account authentication.

Several live option and numbered-field test cases explicitly cleared the page cache before the call.
Other passing cases reused data already retrieved during the run. Their passing results do not each establish a new request.

## Legacy Yahoo quote mode

The [Yahoo quote documentation](../docs/Documentation/doc-links/RCHGetYahooQuotes-Function_files/RCHGetYahooQuotes-Function.htm#L11) explicitly marks `RCHGetYahooQuotes` obsolete.
The documentation index still shows an old ticker/code example.
The current source permits supplied URL/CSV mode, which passed the selected history tests.
Restoring ticker/code mode is separate compatibility work. Array entry and refresh settings do not restore its unavailable endpoint.

## Yahoo option fields, percentages and expiration text

The [option guide](../docs/Documentation/doc-links/smfGetOptionQuotes-Function_files/smfGetOptionQuotes-Function.htm) describes `%` as percent change and `x` as expiration date.
Its introduction names Yahoo and says most other sources are no longer accessible.
Its older source tables still contain those legacy providers; they do not establish current availability.

On Mac, `smfGetYahooOptionQuote1` now calls `smfGetYahooOptionQuote` with the same five arguments.
That function reads Yahoo JSON through `smfGetYahooJSONData`:

| Item | Yahoo field | Value handling |
|---|---|---|
| `%` | `price.regularMarketChangePercent.raw` | Preserve the numeric fraction. |
| `X` | `summaryDetail.expireDate.fmt` | Return Yahoo's formatted date text. |

The saved native checks contain these public values:

| Case | Input or same-response reference | Returned value |
|---|---|---|
| Controlled percent fixture | `raw=-0.125`, `fmt="-12.50%"` | Numeric `-0.125` |
| Controlled expiration fixture | `raw=1795132800`, `fmt="2026-11-20"` | Text `2026-11-20` |
| Live AAPL November 20 $350 call, percent | Same-response JSON raw value `0` | Numeric `0` |
| Live AAPL November 20 $350 call, expiration | Same-response formatted value `2026-11-20` | Text `2026-11-20` |

The nonzero controlled percent fixture detects an extra division or multiplication by 100.
The zero live percent value cannot establish that scaling by itself.
The live call bid, call ask and put bid checks also returned numeric zero matching their JSON references.
These checks accept zero and reject missing values; they do not independently verify market prices or freshness.
The expiration check verifies the date represented by the text; it does not promise a VBA `Date` result or one universal text format.
The controlled fixtures contain test values, not captured market quotes.

## Documented browser request modes

The [table-cell request types](../docs/Documentation/doc-links/RCHGetTableCell-Function_files/RCHGetTableCell-Function.htm#L104) describe Windows engines:

| Type | Documented engine | Current Mac behavior |
|---|---|---|
| 0 | XMLHTTP GET | HTTP GET through the Mac helper |
| 1 | Internet Explorer object | HTTP GET through the Mac helper |
| 2 | HTMLDocument | HTTP GET through the Mac helper |
| 3 | XMLHTTP POST | HTTP POST through the Mac helper |
| 4 | WinHTTP with Yahoo crumb | Yahoo cookie/crumb request through the Mac helper |

Types 1 and 2 do not execute browser JavaScript or import a Safari/Chrome login on Mac.
This limitation can affect pages that require browser rendering or an authenticated session.
The request-type parameter is not a GuruFocus username or API-token setting.

## Published manuals fetched and hash-checked

We fetched these eight published inner manuals. Each request returned HTTP 200.
Every downloaded body matched the corresponding repository file byte for byte and by SHA256.
The download record identifies these eight pages:

1. [smfGetYahooHistory](../docs/Documentation/doc-links/smfGetYahooHistory-Function_files/smfGetYahooHistory-Function.htm): fields, defaults, dates, dimensions and array entry.
2. [smfGetYahooOptionQuote](../docs/Documentation/doc-links/smfGetYahooOptionQuote-Function_files/smfGetYahooOptionQuote-Function.htm): five arguments, option items and older examples.
3. [Array Entered Formulas for Excel](<../docs/Documentation/doc-links/arrayEnterFormula-files/Array Entered Formulas for Excel.htm>): selecting and entering a complete array range.
4. [smfGetYahooJSONData](../docs/Documentation/doc-links/smfGetYahooJSONData-Function_files/smfGetYahooJSONData-Function.htm): modules, fields and raw/formatted values.
5. [smfGetOptionQuotes](../docs/Documentation/doc-links/smfGetOptionQuotes-Function_files/smfGetOptionQuotes-Function.htm): option items, source choices and array dimensions.
6. [smfGetOptionExpirations](../docs/Documentation/doc-links/smfGetOptionExpirations_files/smfGetOptionExpirations.htm): expiration sources and output formats.
7. [smfGetOptionStrikes](../docs/Documentation/doc-links/smfGetOptionStrikes-Function_files/smfGetOptionStrikes-Function.htm): expiration-specific strikes and output ranges.
8. [smfGetYahooPortfolioView](../docs/Documentation/doc-links/smfGetYahooPortfolioView-Function_files/smfGetYahooPortfolioView-Function.htm): batch quotes, field codes and displayed values.

## Related repository manuals reviewed

We also reviewed these five manuals from the repository for the related behavior described in this note:

- [RCHGetYahooHistory](../docs/Documentation/doc-links/RCHGetYahooHistory-Function_files/RCHGetYahooHistory-Function.htm): compatibility fields and range behavior
- [RCHGetElementNumber](../docs/Documentation/doc-links/RCHGetElementNumber-Function_files/RCHGetElementNumber-Function.htm): source diagnostics, field definitions and numeric scales
- [smfForceRecalculation](../docs/Documentation/doc-links/smfForceRecalculation-Macro_files/smfForceRecalculation-Macro.htm): cached pages, definition reload and recalculation
- [RCHGetYahooQuotes](../docs/Documentation/doc-links/RCHGetYahooQuotes-Function_files/RCHGetYahooQuotes-Function.htm): obsolete ticker/code mode and refresh arguments
- [RCHGetTableCell](../docs/Documentation/doc-links/RCHGetTableCell-Function_files/RCHGetTableCell-Function.htm): extraction parameters and request types.

The saved download record does not include online fetch/hash comparisons for these five additional pages.
The review provenance records each online comparison and separately identifies these local-only reviews.

## Remaining checks

This review did not execute every example in those manuals.
The manuals contain contradictory defaults, old provider URLs and Windows-specific instructions.
Future checks must compare each requested ticker, date, units and field with the provider response.
A case requiring a new response must clear the page cache; cached cases must remain identified separately.

The native results support the two selected Mac repairs.
They do not establish that worksheet settings can restore failed legacy endpoints or authenticated data access.
The later chart-reopen failures block a 0.6 release until the reviewed public-package copy and repository copy is complete.
The nine focused reuse cases and the unselected plot-above-note probe passed.

The [current chart recheck status](CHART-RECHECK-STATUS-20261108.md) records the revised candidate and the remaining gates.
