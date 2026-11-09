# SMF for Excel on Mac — 0.5-dev

This experimental build includes the add-in and three example workbooks. It was checked in native Excel on an Apple Silicon Mac.

Download the [working branch](https://github.com/dnarain/SMF_Add-in/tree/dnarain/mac-native-support): select **Code > Download ZIP**, unzip it, then open `mac/built/0.5-dev`.

## Install

1. Keep every file in this folder together, with its original name.
2. Quit Excel. Double-click `install_helper.command`. It installs the web-data helper for Excel and preserves an existing helper as a backup.
3. Open Excel. Select **Tools > Excel Add-ins > Browse**, then select `SMF-Mac.xlam` from this folder. Enable its checkbox.
4. Allow this add-in's macros when Excel asks. Keep **Trust access to the VBA project object model** OFF. The installer does not change Excel security settings.
5. Keep calculation set to **Automatic**. Quit and reopen Excel, then open an example below.

The version result includes `mac.0.5-dev` and this folder's location:

```excel
=RCHGetElementNumber("Version")
```

If Excel reports an old missing add-in, remove that entry with **Delete from list**. Use **Browse** to select the full path to this folder's `SMF-Mac.xlam` again.

## Examples

| Workbook | What to try |
| --- | --- |
| `SMF-Mac-Checks-20261108.xlsx` | Version, HTML parsing and Finviz earnings/price checks. |
| `SMF-Mac-Data-Picture-Checks-20261108.xlsx` | Provider fields, a StockCharts picture and a native numeric chart. |
| `SMF-Mac-History-Checks-20261108.xlsx` | Yahoo history with explicit row counts and a five-row, two-column result. |

The examples were refreshed with this build on November 8, 2026. Recalculation can change live values. Fixed expected prices and dates may need updating.

Seven legacy provider checks still fail. Other SMF fields and example workbooks still require checks.

The [Mac documentation](https://github.com/dnarain/SMF_Add-in/blob/dnarain/mac-native-support/mac/README.md) records the checks and remaining limits. `SHA256SUMS` lists this bundle's file checksums.

## Build record

`BUILD-INFO.json` records the source commit and native Excel checks.
The VBA project matches the tested build byte for byte.
Excel retains the build Mac Forms reference and compiler-cache paths inside that project.
The active VBA source contains no hard-coded home-folder path.
The bundle passed the example checks from a separate folder on the same Mac. A fresh Mac installation remains untested.
