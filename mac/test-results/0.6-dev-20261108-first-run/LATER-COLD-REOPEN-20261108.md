# Later cold-reopen failures after the first 0.6 run

This note preserves the immediate post-failure checkpoint. The [final candidate report](../0.6-dev-20261108-final/README.md) records subsequent checks.

The 0.6 add-in is not released. Two later cold-reopen checks of the data-and-picture example restarted Excel.
The installed 0.5 add-in remains unchanged and is unloaded during this investigation.

## Completed first run

The first 0.6 full run completed 292 cases: 238 PASS and 54 CHECK_FAILED.
It recorded no native phase errors or unfinished cases. All 53 added cases and all 32 worksheet-entry cases passed.
All 14 picture-group cases passed within that batch.
The saved results, assertions, workbook hashes and source fingerprints remain unchanged.
The later cold-reopen checks are separate lifecycle checks. They do not change the recorded 292-case outcomes.

The manual review and two small repairs are complete.
Mac history calls use supported default fields. Mac `smfGetYahooOptionQuote1` uses the existing Yahoo JSON function.
The option comparisons accepted numeric zero. A fresh usable positive option quote remains unestablished.

## Later native evidence

The operator observed two Excel restarts after cold reopens of the data-and-picture example.
Excel diagnostic logs record ungraceful exits and failed chart-insert operations at the terminal chart-error records.
The new PID2326 log gives these exact records:

| UTC time | Log line | Byte offset | Recorded evidence |
|---|---:|---:|---|
| 06: 46: 05.834 | 10 | 1717 | Previous process exited unexpectedly. `ExitWasGraceful=false`, `ExceptionEnum=1`, stack hash `1e53dc4383007179_1_main`. |
| 06: 46: 06.219 | 3193 | 1443958 | `Office.System.SystemHealthUngracefulAppExitMacAndiOS` repeats the previous-exit metadata. |
| 06: 51: 25.820 | 25825 | 18542166 | `Office.Charting.InsertChartE2o`, `Success=false`, HRESULT `-2147221504`, `ChartType=51`. |
| 06: 51: 25.820 | 25826 | 18542985 | The second telemetry record repeats the failed chart insertion. |
| 06: 51: 25.820 | 25827 | 18543709 | Final `Chart Error` record, tag `bpnoy`, error identifier `EAFBB740-D7AE-1748-8E25-A3B2EC57F0C5`. |

The log ends after that chart-error record, followed by NUL padding.
The log does not decode `ExceptionEnum=1` or provide a call stack that identifies the VBA line.
Cache-recovery messages report empty fault stores. They do not identify the workbook or chart failure.
The private log SHA256 is `f2c2c250cea37aa1f3c1330ac116a3c785d5ee211f60bb868b5b55e77661c1c7`.
Raw logs remain private because they contain account and file metadata.

## Current investigation

The operator is isolating the numeric-chart formula in cell B16.
The current Mac comment code removes its owned object, then creates a chart for `pChoice=98`.
The code calls `ChartObjects.Add` before it sets `ChartType=xlLine`.
The chart-insert errors support investigation of that lifecycle. They do not establish the underlying Excel defect.

A private candidate will reuse an existing owned numeric chart during recalculation.
The candidate still requires native recalculation, cold reopen and regression checks.
The first 292-case report remains a separate record. A later candidate needs its own source audit and native results.
The 0.6 bundle remains pending until the later chart failures are resolved and checked.
