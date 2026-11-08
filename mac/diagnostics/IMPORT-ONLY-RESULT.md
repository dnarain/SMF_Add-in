# Full-source import and native checks: November 7, 2026

The user explicitly approved temporary VBA project access for one automatic
49-module import, followed by restoring that setting to OFF before compilation.

## Native actions and result

The agent used Computer Use to verify the project-access checkbox was OFF,
enable it, and return to the saved builder owner:

```text
SMF-Mac-Compile-Owner-Manual-20261107-01.xlsm
```

The agent visually verified `SMFMacPrepareCompile` with the local repository path before pressing Return.
The command that follows substitutes a placeholder for that path:

```vb
? SMFMacPrepareCompile("<absolute repository folder>")
```

The wrapper calls the builder with `runTests=False` and `saveOutput=False`.
The import completed. The final persistent log lines are:

```text
Imports completed; target=Sheet1; project=SMFDiag_20261107_142858; imported=49; components=51
Completed after restoring flags; target=Sheet1; project=SMFDiag_20261107_142858; saved=False; runTests=False; saveOutput=False; DisplayAlerts=True; ScreenUpdating=True; EnableEvents=True; FlagsRestored=True
```

Computer Use then showed the selected VBA project
`SMFDiag_20261107_142858 (Sheet1)` in the Project Explorer.
The builder owner and Acrobat add-in remained separate projects.

The Computer Use request exceeded its initial 10-second deadline while Excel
was importing. The agent did not repeat the command. The import continued,
and the persistent log confirmed completion. The agent reconnected to Excel
afterward and observed the imported project.

## Security setting restored

After import completion, the agent reopened Excel's Security pane, verified
the project-access checkbox was ON, turned it OFF, and verified value `0`.
The other macro-security radio buttons remained at their prior values `0, 0, 1`.
No other security setting changed.

## Saved checkpoint and class repair

The builder returned an unsaved target. Later native Save actions created
`out/SMF-Mac-Full-Source-20261107-01.xlsm`.
Its VBA project remains `SMFDiag_20261107_142858`.

The first saved checkpoint contained 49 standard modules and no class modules.
Its `PROJECT` stream explicitly listed `Module=Dictionary`.
The source file was a class export, but the import classified it as a standard module.
The agent repaired Dictionary through a native class import using the CRLF source file.

The saved project now contains 48 standard modules, one `Dictionary` class,
and two document modules. Its 49 source components remain present.
Independent CFB inspection confirmed these counts in both saved checkpoints:
The hashes describe the post-class-repair saves, before the later HTML replacement.

| Checkpoint | Bytes | SHA256 |
| --- | ---: | --- |
| `out/SMF-Mac-Full-Source-20261107-01.xlsm` | 451,477 | `4411fbaa524f572f6dbb9d4dbf040da8818bb17c4b7cc6b8b014d6867f4b7e40` |
| `out/SMF-Mac-Full-Source-20261107-class-repaired.xlsm` | 451,814 | `58e77947a2a8d81382bf64dde105bab794b6313fe5791e4f7a7c280117d86bd4` |

Decompressed Dictionary source retains `Attribute Item.VB_UserMemId = 0`.
That attribute preserves its default `Item` property.

One keyboard action cut selected `modMenu` text. The agent immediately selected Undo Cut and saved.
Independent OVBA decompression confirmed that the current and backup `modMenu` source bytes match exactly.
Both contain 5,108 bytes, with SHA256 `cb35c9b22bff635b396486f0242e741a0e191b07b2588ba4d900981c0ba4ac8b`.
The raw module streams differ because the source offset changed from 6919 to 6935.

## Native checks after class repair

VBA project access remained OFF for these checks.
The full project's `SMFMacTransportSelfTest` returned `6`.
The agent selected the actual Debug > Compile command for the target project.
Excel returned without an error dialog.
These observations do not establish correct results for every SMF function.

The next HTML fixture run stopped at the `preformatted CRLF` assertion.
The dialog displayed these values:

```text
expected: [A<LF><CR>B]
actual:   [A<LF><CR><LF><CR>B]
```

The agent selected End. The later results are recorded below.
No final `mac/out/SMF-Mac.xlam` exists.

The HTML source now handles actual CR/LF input pairs in either order.
It contains 55 fixture cases after eight additions.

## Current native work

The Mac locked after the HTML fixture failure. The user subsequently unlocked it.
An Immediate-window probe returned these actual runtime values:

```text
AscW(vbCr) = 13
AscW(vbLf) = 10
Len(vbCrLf) = 2
AscW(Left$(vbCrLf, 1)) = 10
AscW(Right$(vbCrLf, 1)) = 13
```

The agent replaced the HTML module manually. Its native fixture returned `55`.

The next `SMFMacJSONSelfTest` stopped with runtime error 429, "ActiveX component can't create object".
Debug highlighted `CreateObject("MSScriptControl.ScriptControl")` in `InitScriptEngine`.
That statement was inside a `#If Not Mac Then` guard.
A native probe returned `Not 1 = -2` and `Not True = False`.
The static checker previously assumed the numeric Mac constant was `-1`.

The source now uses `#If Mac = 0 Then` at the three Windows-only guards.
These changes affect `modJSONExtract.bas` and `smfUtilities_.bas` only.
The checker covers both `Mac=1` and `Mac=-1`, with Win64 zero and nonzero.
Its 21 fixtures and 464 signature comparisons pass.
An offline regression confirms that restoring the old guard exposes ScriptControl under `Mac=1`.
The agent replaced those two modules manually. A direct `InitScriptEngine` call returned without error.
The native JSON fixture then returned `PASS: 84 JSON checks`.
The integration fixture returned `5`.
The full project has now passed transport `6`, HTML `55`, JSON `84`, and integration `5`.

The agent saved the current full-source workbook after these checks.
It contains 590,999 bytes, with SHA256
`78c4f79a12f21882327e19acacf1a5afb10dc322b52108708b30baac32b3f72e`.
Independent CFB inspection confirmed 48 standard modules, one Dictionary class, and two document modules.
All 49 expected source component names remain present.
Decompressed saved HTML source matches the current source after line-ending normalization.
Saved JSON and utility source differ only in VBE identifier casing; literal-preserving comparisons match.
All three explicit `Mac = 0` guards are present. Dictionary retains its hidden default-Item attribute.
Live numbered-field checks and a saved, loaded `.xlam` remain pending.
No 49-module rebuild or project-access change is required for these replacements.

An earlier lock also interrupted editor inspection after the initial import.
It did not establish the cause of the earlier editor-access problem.
