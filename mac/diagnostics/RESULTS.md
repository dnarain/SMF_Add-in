# Native diagnostic results: November 7, 2026

Workbook: `SMF-Mac-Minimal-Probe-1d7b121e.xlsm`.
Excel process ID: `67044`.
Excel version: `16.114.3` on Apple Silicon.
VBA project access was verified OFF before the test.

## Completed

- Excel saved the empty `.xlsm` file. SaveAs exceeded the 15-second AppleEvent
  timeout, but later readback confirmed the expected path and `saved=True`.
- The empty workbook passed ZIP integrity checks. It contained no VBA project
  before the manual import.
- `SMFMacProbe.bas` was imported through the native VBA editor.
- The workbook-qualified `SMFMacProbeBasic` call returned
  `PASS: basic VBA (3 checks)` immediately.
- The worksheet formula `=SMFMacProbeBasic()` returned the same text in A1.
- The unchanged `code/modMacCode.bas` was manually imported into the probe.
- The workbook-qualified `SMFMacTransportSelfTest` returned `6` immediately.
- The workbook-qualified `RCHGetURLData1Mac` fetched
  `https://finviz.com/quote.ashx?t=AAPL` through `AppleScriptTask` from real Excel:

  | Check | Result |
  | --- | --- |
  | Response length | 308,819 characters |
  | HTTP status | 200 |
  | Earnings marker | Present |
  | Market Cap marker | Present |
  | `SMFMacLastError()` | Empty |

The basic function ran before any core SMF module was loaded in the probe.
The basic function checked arithmetic, string replacement, and array access.
The transport function checked six fixed HTTP-response cases without a network request.
Excel did not restart during these completed steps.

The HTTP check retained the page in an AppleScript variable and returned only
the metadata above. No response body was logged.
The probe contained only `SMFMacProbe` and the unchanged `modMacCode`.

## Saved workbook

Cmd+S did not confirm persistence: Excel still reported `saved=False`, and the
file remained 8,489 bytes. One normal AppleScript Save then succeeded.
Excel reported `saved=True`. The saved workbook has this repository-relative path:

```text
mac/diagnostics/out/SMF-Mac-Minimal-Probe-1d7b121e.xlsm
```

The saved file is 25,745 bytes. Independent ZIP inspection passed, confirmed
`xl/vbaProject.bin`, and confirmed the A1 formula `SMFMacProbeBasic()`.
Excel stayed in process `67044`. No security setting was changed.
All native automation stopped after this readback.

The full 49-module project remains out of this diagnostic workbook.
These results establish basic VBA execution, one worksheet UDF, the response
decoder, and one native HTTP request. They do not establish full-project
compilation, Finviz field extraction, or a working add-in.
