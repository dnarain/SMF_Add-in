# Full-suite failure triage — run 1, November 8, 2026

Run: `smf-0.5-full-run-20261108-run1`

Run ID: `smf-0.5-30fe5d1421284455a9e787ea32bd9ede`

Run 1 saved 239 completed cases: **184 PASS, 52 CHECK_FAILED and 3 CLEANUP_ERROR**.
This remains the first-run record. Final run 2 saved **187 PASS and 52 CHECK_FAILED**, with no unfinished cases or cleanup errors.
The same 13 primary cases failed in both runs. Only the three harness cleanup cases changed to PASS.
The table below covers 13 failed primary cases named after public functions. Other live-provider and numbered-field cases can also fail for functions that passed local fixtures.

The Google history URL and RSI history defaults contain source defects. The remaining observations identify the exact failed requests or missing parser markers.

## Nine data-function cases

| Function | Source | Observed failure and established limit |
| --- | --- | --- |
| `smfGetGoogleHistoryCSV` | [source:75](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/smfGetGoogleHistoryCSV_.bas#L75) | Date parameters contain unencoded spaces. Native row 100 and HTTP probe 10 both failed with curl error 3 before an HTTP response. Encoding is required; the old CSV route still needs verification. |
| `smfRSITargetPrices` | [source:50](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modRSITargetPrices.bas#L50) | The quote CSV host failed DNS resolution. Its history call at line 52 also omits the item list: the default includes F/G, which the history validator rejects. Restoring the quote request alone would leave that argument defect. |
| `smfGetADVFNElement` | [source:59](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetAdvFNElement.bas#L59) | Returned Error after HTTP 200. Same-URL probe 52 lacks the required Select start date marker and the requested Year End Date label. Current financial-page parsing needs review; field 5196 passed separately. |
| `smfGetGuruFocusItem` | [source:62](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetGuruFocusItem.bas#L62) | Returned Error after HTTP 200. Probe 42 lacks the old id=Rf, yesttm and normal_pershare markers. This establishes a page/parser mismatch; availability through JavaScript or paid access remains undetermined. Field 2727 passed separately. |
| `smfGetGuruFocusCSVItem` | [source:34](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/smfGetGuruFocusCSVItem_.bas#L34) | The configured CSV URL returned HTTP 404 in native execution and probe 43. The function returned an empty string. |
| `smfGetMSNHistory` | [source:73](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetMSNHistory.bas#L73) | The configured data.moneycentral.msn.com host failed DNS resolution in both native execution and probe 6. No history payload was available. |
| `smfYahooAPIData` | [source:24](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modYahooAPIData.bas#L24) | The configured query.yahooapis.com host failed DNS resolution; probe 19 agreed. No YQL data was available. |
| `smfGetMorningstarHistoryCSV` | [source:82](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/smfGetMorningStarCSV_.bas#L82) | Connection to the configured performance.morningstar.com route timed out. No payload was available to assess its CSV parser. |
| `smfGetEconData` | [source:27](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetEconData.bas#L27) | The configured FRED request timed out receiving data. The cause and parser correctness remain undetermined. |

## Four intentionally obsolete entry points

| Function | Source | Native result |
| --- | --- | --- |
| `smfGetMSNOptionQuote` | [source:707](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetOptionQuotes.bas#L707) | Obsolete |
| `smfGetOXOptionQuote` | [source:1087](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetOptionQuotes.bas#L1087) | Obsolete — OptionsXpress is no longer a valid data source |
| `smfGetOX2OptionQuote` | [source:1191](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetOptionQuotes.bas#L1191) | Obsolete — OptionsXpress is no longer a valid data source |
| `smfGetOX3OptionQuote` | [source:1279](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/code/modGetOptionQuotes.bas#L1279) | Obsolete — OptionsXpress is no longer a valid data source |

These four functions return before making a provider request. Each function exits with an obsolete-status string before the Mac downloader runs.
Other legacy option routes also failed: Google returned HTML instead of an option chain, and the two Yahoo HTML routes returned 404.
The Yahoo JSON bid, strike and expiration cases passed, so those failures must remain specific to the called route.

## Interpretation limits

- The hardcoded MMM history URL is in the Windows branch. The active Mac `RCHGetYahooHistory2` used AAPL and passed.
- HTTP 200 did not establish working data: the captures include a Netflix home page, an unrelated accounting redirect and an ETFScreen parking page.
- The Zacks and Stockscores endpoint-marker checks failed while their numbered-field extractions passed. Neither result establishes every field on that provider.
- MarketWatch, Nasdaq and Reuters returned 401/403 for the tested requests. Those statuses do not establish a global provider outage or the required remedy.
- The three cleanup errors affected `MacComment_text_only`, `MacComment_clear_none` and `MacComment_failed_update_clears_picture`. Run 2 passed all three after the harness cleanup correction; retain the original run 1 rows as historical evidence.

This triage used the saved native case attribution and the existing 58-probe HTTP record. No new network request or source change was made.
