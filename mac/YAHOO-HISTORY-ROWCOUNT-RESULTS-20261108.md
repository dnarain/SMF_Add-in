# Yahoo history row/column fix — November 8, 2026

The actual public `smfGetYahooHistory` passed 46 native checks in an isolated full-source workbook.
The source version is `3.0.2024.07.30-mac.0.5-dev`.
The audited add-in also passed eight fresh worksheet checks after a clean Excel restart, with a live `5x2` result.
The port remains experimental in the personal fork. No pull request is open.

## Source change

[smfGetYahooHistory_.bas](../code/smfGetYahooHistory_.bas) changes the ninth Mac argument, `pCols`, from `Integer` to `Double`.
`pRows` remains `Integer`. Argument names, order, defaults and the remaining declared types stay unchanged.
The Windows branch retains `pCols As Integer`; no Windows runtime check has run.
The function still assigns `pCols` to local `iCols As Integer` before allocating the result array.
The zero-dimension check now uses the converted local `iRows` and `iCols`, preserving the tested caller-sizing behavior.

Native scalar controls read the ninth `Integer` argument as `-32768` for the tested positive inputs.
The `Double` controls preserved the requested column values, and the actual public function then passed the checks below.
These observations support the selected source fix. They do not establish a general Excel argument-passing cause across Mac versions.
Out-of-range or nonnumeric column inputs were not verified. Their failure timing can differ after changing the public type.

## Native results

| Group | Checks | Result |
| --- | ---: | --- |
| Offline worksheet controls | 28 | 28 matched |
| Live worksheet controls | 12 | 12 matched |
| Direct macro controls | 6 | 6 matched |
| Total | 46 | 46 matched |

The 28 offline checks intentionally expect an invalid-period diagnostic, blank trailing cells or `None`.
The 12 live checks used AAPL inputs dated November 2–7 and returned five trading rows from November 2–6.
The live checks covered `DC`, `DV`, close-only output, oldest/latest values, ascending order and array dimensions.
The six macro checks returned Variant arrays (`VarType = 8204`) with the recorded bounds.
Their `Application.Caller` type was `Error`; this describes the macro context, not a failed test.

The earlier 66-case repair/history batch remains 59 PASS and seven provider CHECK_FAILED results.
The 46 focused controls are reported separately. No combined unique-case total is claimed.
The seven external provider failures remain unresolved.

## Worksheet controls

| Group | Control | Expected | Recorded | Result |
| --- | --- | --- | --- | --- |
| Offline | `bad:5,1:first` | Error on period: bad | Error on period: bad | Matched |
| Offline | `bad:5,1:rows` | 5 | 5 | Matched |
| Offline | `bad:5,1:cols` | 1 | 1 | Matched |
| Offline | `bad:5,1:last` | (empty) | (empty) | Matched |
| Offline | `bad:5,2:first` | Error on period: bad | Error on period: bad | Matched |
| Offline | `bad:5,2:rows` | 5 | 5 | Matched |
| Offline | `bad:5,2:cols` | 2 | 2 | Matched |
| Offline | `bad:5,2:last` | (empty) | (empty) | Matched |
| Offline | `bad:5,0:first` | Error on period: bad | Error on period: bad | Matched |
| Offline | `bad:5,0:rows` | 2 | 2 | Matched |
| Offline | `bad:5,0:cols` | 1 | 1 | Matched |
| Offline | `bad:5,0:last` | (empty) | (empty) | Matched |
| Offline | `bad:1,1:first` | Error on period: bad | Error on period: bad | Matched |
| Offline | `bad:1,1:rows` | 2 | 2 | Matched |
| Offline | `bad:1,1:cols` | 1 | 1 | Matched |
| Offline | `bad:1,1:last` | (empty) | (empty) | Matched |
| Offline | `bad:0,0:first` | Error on period: bad | Error on period: bad | Matched |
| Offline | `bad:0,0:rows` | 2 | 2 | Matched |
| Offline | `bad:0,0:cols` | 1 | 1 | Matched |
| Offline | `bad:0,0:last` | (empty) | (empty) | Matched |
| Offline | `None:first` | None | None | Matched |
| Offline | `None:rows` | 1 | 1 | Matched |
| Offline | `None:cols` | 1 | 1 | Matched |
| Offline | `string-zero:first` | Error on period: bad | Error on period: bad | Matched |
| Offline | `string-zero:rows` | 2 | 2 | Matched |
| Offline | `string-zero:cols` | 1 | 1 | Matched |
| Offline | `fractional-cols:0.4` | 2x1:Error on period: bad | 2x1:Error on period: bad | Matched |
| Offline | `fractional-cols:1.5` | 5x2:Error on period: bad | 5x2:Error on period: bad | Matched |
| Live | `DC:latest-date` | 46332 | 46332 | Matched |
| Live | `DC:latest-close` | 342.3699951171875 | 342.3699951171875 | Matched |
| Live | `DC:oldest-date` | 46328 | 46328 | Matched |
| Live | `DC:oldest-close` | 338.32000732421875 | 338.32000732421875 | Matched |
| Live | `DV:latest-volume` | 44676300 | 44676300 | Matched |
| Live | `DV:oldest-volume` | 64757700 | 64757700 | Matched |
| Live | `C:latest-close` | 342.3699951171875 | 342.3699951171875 | Matched |
| Live | `DC:ascending-first-date` | 46328 | 46328 | Matched |
| Live | `DC:ascending-last-date` | 46332 | 46332 | Matched |
| Live | `DV:ascending-first-volume` | 64757700 | 64757700 | Matched |
| Live | `DC:rows` | 5 | 5 | Matched |
| Live | `DV:cols` | 2 | 2 | Matched |

For fractional column controls, `0.4` converted to local zero and used the scalar caller: `2x1`.
The `1.5` control converted to two columns and returned `5x2`. Integer column counts remain the normal input.

## Direct macro controls

| Control | Bounds | First value | Last value | Result |
| --- | --- | --- | --- | --- |
| `bad:5,1` | `1..5,1..1` | Error on period: bad | (empty) | Matched |
| `bad:5,2` | `1..5,1..2` | Error on period: bad | (empty) | Matched |
| `bad:1,1` | `1..2,1..1` | Error on period: bad | (empty) | Matched |
| `bad:0,0` | `1..8,1..2` | Error on period: bad | (empty) | Matched |
| `string-zero:5,0` | `1..5,1..2` | Error on period: bad | (empty) | Matched |
| `None:5,1` | `1..1,1..1` | None | None | Matched |

Default macro dimensions depend on the caller context. The table records the tested context, including `8x2` for `0,0`.

## Explicit-size example

This worksheet call passed in the diagnostic workbook with the corrected public function:

```excel
=INDEX(smfGetYahooHistory("AAPL",DATE(2026,11,2),DATE(2026,11,7),"d","C",0,0,5,1),1,1)
```

It returned `342.3699951171875` for the newest close in the requested period.
The installed `0.5-dev` add-in now passes the checked explicit-size worksheet calls.

## Evidence and remaining checks

The native workbooks and readback remain local under ignored `mac/diagnostics/out`.
The files below contain the actual public-function run, not a renamed candidate function.

- [Passed native workbook](diagnostics/out/yahoo-rowcount/SMF-Mac-History-Public-Passed-20261108.xlsm)
- [Frozen native values and formulas](diagnostics/out/yahoo-rowcount/history-public-readback.json)
- [Native runner specification](diagnostics/out/yahoo-rowcount/history-public-runner-manifest.json)
- [Nine reusable worksheet regression specifications](tests/full_function_suite/history_rowcount_cases.json)

The nine reusable specifications have not been run through the full-suite harness in this checkpoint.
They are separate from the 46 recorded native controls. The full provider/function suite was not rerun.

| Evidence | Bytes | SHA-256 |
| --- | ---: | --- |
| `SMF-Mac-History-Public-Passed-20261108.xlsm` | 699850 | `6a8fb395877762327969b8b4bc82f2dd57347674ee65139fdb5fa3ca4d97372e` |
| `history-public-readback.json` | 17999 | `1f0c719e6813868aa2d59e23444490418bf6e45ed5ec9decc403b52b3a22ea9b` |

## Packaged add-in and restart checks

The versioned `out/SMF-Mac-0.5-dev.xlam` and standard `out/SMF-Mac.xlam` files are byte-identical.
Each contains 673,519 bytes with SHA-256 `3b9423659992348920ee08bd32898a9b20187073116271b20e7d090a33514331`.
The binary audit passed, with no extra runner/test streams in the clean add-in.
All 14 adjacent element files match `docs/Elements/` byte for byte. The `0.5-dev` fix does not change the AppleScript helper.

The first cold restart found a stale registration pointing to `/Applications/Microsoft Excel.app/SMF-Mac.xlam`.
The agent removed that missing entry and browsed to the full absolute path of the repository's `mac/out/SMF-Mac.xlam`.
After a clean menu quit/restart and workbook reopen, Excel showed no missing-add-in warning or link prompt.
The agent repasted and recalculated every check formula. All eight passed, and the live five-row, two-column result appeared.
The version formula reported `mac.0.5-dev` and the correct output directory.

The saved [history check workbook](out/SMF-Mac-History-Checks-20261108.xlsx) contains 11,974 bytes.
Its SHA-256 is `fbf5e56f9e4968a8d4bb300b5a99ce0df35ac8a3c546dd2dbd10be317e99e2de`.
The saved XML retains eight PASS results, the version/path result and the full `5x2` result.
The eight worksheet checks are follow-up evidence and are excluded from any new unique-case total.

After restart, the add-in remained checked and calculation was **Automatic**.
The final security readback confirmed VBA project access OFF; other macro-security settings were unchanged.
The passing `Double` controls and final restart completed without an Excel crash.
The earlier unsafe `Long` probe remains a separate local record, excluded from the 46 passing controls.
These focused results do not establish support for every history argument, provider, workbook or Mac Excel version.
