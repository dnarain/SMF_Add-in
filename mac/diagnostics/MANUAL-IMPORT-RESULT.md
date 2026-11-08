# Manual-copy inspection: November 7, 2026

The user asked the agent to inspect the open editor after the import instructions.
The VBA editor was open on
`SMF-Mac-Compile-Owner-Manual-20261107-01.xlsm`.

The Modules folder was collapsed. Expanding it through Computer Use showed:

- `build_SMFMac_noprompt`
- `modMacCode`
- `SMFMacProbe`

No new import was initiated during this inspection. The builder was already
present in the live project. Opening its code showed the expected entry point:

```vb
Public Function SMFMacPrepareCompile(ByVal projectFolder As String) As String
    SMFMacPrepareCompile = SMFMacBuildNoPrompt(projectFolder, False, False)
End Function
```

The agent clicked the editor's Save button for this workbook. Independent disk
checks then confirmed 37,366 bytes, ZIP integrity, and exactly these three standard
modules in the saved VBA PROJECT records. The saved project includes the
`SMFMacPrepareCompile` name.

Saved owner SHA256:

```text
385892ea5e0561d779d3b8dd36b4d5211cf8e35450aaff4e866e6f97f9ee11e5
```

The original passing probe remains separate at 25,745 bytes with its original
two modules. Its SHA256 is:

```text
234999698e4c728273e92033f759a7471633bd9075fb74bbbfc7adcaadb3dddf
```

No security setting changed. No builder function, full-project import, explicit
compilation, fixture, or HTTP request ran during the inspection.

The agent requested fresh approval to temporarily enable VBA project access
for the 49-module automatic import and restore it to OFF before compilation.
The user subsequently approved it. The automatic import completed and the
setting was restored to OFF. See `IMPORT-ONLY-RESULT.md`.
