# Install SMF 0.7-dev on macOS

This is an experimental port for desktop Excel on Mac. It runs natively on Apple silicon without Windows.
It requires Excel with VBA support and internet access for provider data.
All of my changes in this fork were vibe-coded with AI agents.
The native checks used one Mac. Real-world use and other Mac configurations still need testing.

Use the [`dnarain/mac-native-support` branch](https://github.com/dnarain/SMF_Add-in/tree/dnarain/mac-native-support).
The 0.7 bundle is in [`mac/built/0.7-dev`](mac/built/0.7-dev/README.md).

## 1. Get the files

On the branch page, select **Code > Download ZIP**, then unzip the download.

Alternatively, clone that branch:

```sh
git clone --branch dnarain/mac-native-support https://github.com/dnarain/SMF_Add-in.git
```

Open [`mac/built/0.7-dev`](mac/built/0.7-dev/README.md) inside the downloaded or cloned repository.
Copy that entire folder to a stable location, for example `~/Documents/SMF-Mac/0.7-dev`.
Keep its files together and keep that folder in place after registering the add-in.
Excel remembers the full add-in path. A temporary folder or later folder move can break that registration.

The 24-file bundle contains:

- `SMF-Mac.xlam`, the Excel add-in
- `SMF-Mac-Checks-20261109.xlsx`, basic checks
- `SMF-Mac-Data-Picture-Checks-20261109.xlsx`, data and picture checks
- `SMF-Mac-History-Checks-20261109.xlsx`, history checks
- Fourteen `smf-elements-*.txt` definition files
- `SMFMac.applescript` and `install_helper.command`
- `NotSupported.csv`, `README.md`, `BUILD-INFO.json` and `SHA256SUMS`

Do not install from the ZIP preview. Unzip and copy the complete folder first.

## 2. Install the Mac helper

1. Quit Excel.
2. In the copied 0.7-dev folder, double-click `install_helper.command`.
3. Wait for Terminal to show **Installed the Mac HTTP helper**.

If the installer does not run, report its exact error message.

The installer compiles `SMFMac.applescript` and writes:

```text
~/Library/Application Scripts/com.microsoft.Excel/SMFMac.scpt
```

Excel's `AppleScriptTask` calls the helper from that application-specific folder.
The installer backs up an existing `SMFMac.scpt` before replacing it. It does not require `sudo` or change Excel security settings.
The add-in, examples and definition files stay in the copied bundle folder.

## 3. Register the add-in in Excel

1. Open Excel.
2. Open a blank workbook and leave it visible. The Add-ins command can be disabled without a visible workbook.
3. Select **Tools > Excel Add-ins > Browse**.
4. Select `SMF-Mac.xlam` from your stable 0.7-dev folder.
5. Enable its checkbox in the Add-ins list.
6. Allow this add-in's macros when Excel asks.
7. Set Excel calculation to **Automatic**.

Keep **Trust access to the VBA project object model** OFF.
Normal installation uses the built add-in. It does not require importing modules, compiling VBA or permitting programs to edit VBA projects.
Keep your other macro-security settings unchanged.

## 4. Check the installation

Open `SMF-Mac-Checks-20261109.xlsx` from the same folder.
It checks the version, local HTML parsing, Finviz earnings text and a positive Finviz price.

You can also enter these formulas in a blank workbook:

```excel
=RCHGetElementNumber("Version")
=smfStripHTML("<b>42</b>")
=RCHGetElementNumber("AAPL",4074)
```

The version should contain `3.0.2024.07.30-mac.0.7-dev`.
The basic example abbreviates the version. The full `RCHGetElementNumber("Version")` formula also reports the loaded add-in folder in parentheses.
The HTML formula should return `42`. The price formula needs a successful provider response and can change with the market.

Next, open the data/picture and history examples from that same folder.
Keep the add-in loaded while using them. If Excel asks for `SMF-Mac.xlam`, select the file beside the examples.

Quit and reopen Excel once, then check the version again.
If Excel cannot find the add-in or does not load it, use the registration steps in the troubleshooting section.
Unattended startup registration for the final 0.7 package remains under verification.

## Optional: FRED API key

Only the FRED functions need a FRED API key. Other functions do not use that key.
Obtain your own key from [FRED](https://fredaccount.stlouisfed.org/apikeys).
Save it as a private plain-text file named `smf-FRED-API-Key.txt` in the same folder as the loaded `SMF-Mac.xlam`.
The file must contain only the 32-character key on one line.

The key file is excluded from Git and the bundle. Never commit or publish it, or put the key in a worksheet formula or report.
Excel may ask for access to that specific file. Select your local key file when granting access.
Without a readable valid key, FRED calls return their configured error value.

## Troubleshooting

**Excel says it could not find SMF-Mac.xlam, especially after a restart.**
Open a blank workbook, then select **Tools > Excel Add-ins**. Uncheck the stale SMF entry, then use **Browse** to select `SMF-Mac.xlam` in your stable folder.
Enable the corrected entry and check the version formula again.
An entry pointing to a temporary build folder will fail after that folder moves or disappears.

**Formulas return #NAME? or the version formula does not work.**
Check that the add-in is enabled and its macros are permitted for this workbook/session.
Register it through **Browse** again if needed. VBA project object-model access can remain OFF.

**The version works, but downloads fail.**
Run `install_helper.command` from the same bundle again, then quit and reopen Excel.
Check that `SMFMac.scpt` exists in the application-specific folder shown earlier.
If macOS or Excel requests access to the bundle or a local source file, select the corresponding file/folder you installed.

**A financial formula returns #N/A.**
Confirmed retired modes and explicitly obsolete default definitions return Excel `#N/A`.
Open `NotSupported.csv`, or right-click a worksheet cell and select **SMF: Source Status** for the loaded status table.
The table records six retired source groups and 288 obsolete default fields. A valid custom definition can replace an obsolete default.

**Only one provider or field fails.**
Provider pages can change or deny automated access. The installation can pass its basic checks while an individual provider formula fails.
See [PROVIDER-CHANGES.md](PROVIDER-CHANGES.md) for the configured URLs and remaining limits.
The Google-history URL is encoded, but its endpoint currently returns Finance HTML. Historical CSV data remains unavailable through that function.

For a problem report, include the exact formula, expected and actual result, Excel version and macOS version.
Remove API keys, session values and private folder paths before sharing diagnostics.

## Current test scope

The full native run recorded 538 cases: 502 PASS and 36 CHECK_FAILED. All 51 worksheet-entry cases passed.
A separate 246-case focused run passed. Some cases check controlled fixtures or expected retirement errors.
These results cover selected calls across 121 public Functions. They do not establish every provider, field or future market value.
The examples' live price/yield/target checks establish numeric values or ranges; independent accuracy and freshness remain separate checks.
