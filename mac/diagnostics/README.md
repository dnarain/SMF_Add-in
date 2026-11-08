# Minimal native Excel diagnostics

The full 49-module import completed, but Excel exited at the first fixture call.
The log does not distinguish compilation from execution.

Keep **Trust access to the VBA project object model** OFF for these manual imports.
Use one dedicated disposable workbook with a saved, unique `.xlsm` filename.
Close the earlier disposable SMF test workbooks before running these checks.

1. Import `SMFMacProbe.bas` through the VBA editor's **File > Import File**.
2. Run the `SMFMacProbeBasicToSheet` macro. Cell A1 must show
   `PASS: basic VBA (3 checks)`. An automated direct call to
   `SMFMacProbeBasic` must return the same text.
3. Confirm a worksheet formula `=SMFMacProbeBasic()` returns the same text.
4. Import the unchanged `code/modMacCode.bas` into this same workbook.
5. Run `SMFMacTransportSelfTest`. Expect `6`.
6. If those checks pass, call `RCHGetURLData1Mac` once for
   `https://finviz.com/quote.ashx?t=AAPL`. Record only the response length,
   the Earnings and Market Cap markers, `SMFMacHTTPStatus()`, and
   `SMFMacLastError()`. Allow 45 seconds for this call.

The first function uses arithmetic, strings, and an array. The transport test
decodes fixed response strings. Neither function requests a web page.
The final check calls `AppleScriptTask` from Excel to retrieve one public page.

Stop if Excel exits unexpectedly or displays an error. Record the last completed
step, workbook name, process ID, and exact error before trying another test.
Do not load the full SMF project as part of these checks.

These checks do not establish that the full SMF project compiles or that the
add-in can be saved and loaded. A successful page response also does not prove
that an SMF field-extraction recipe returns the correct value.

Disposable workbook files in `diagnostics/out` are excluded from the source ZIP.
