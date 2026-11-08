#!/usr/bin/env python3
"""Check the alternative builder without starting Excel or AppleScript."""

from pathlib import Path
import re


def main() -> None:
    root = Path(__file__).resolve().parents[2]
    path = root / "mac" / "build_SMFMac_noprompt.bas"
    text = path.read_text()
    assert text.isascii(), "The builder must use ASCII source."
    assert max(map(len, text.splitlines())) < 1024, "A source line exceeds VBA's limit."
    code = "\n".join(line for line in text.splitlines() if not line.lstrip().startswith("'"))
    for token in ("GrantAccessToMultipleFiles", "FileDialog", "GetOpenFilename", "InputBox", "MsgBox", "MacScript"):
        assert token.casefold() not in code.casefold(), f"Unexpected native prompt entry point: {token}"
    assert not re.search(r"\bDir\$?\s*\(", code, re.IGNORECASE), "The builder must not enumerate exterior files."
    imports = re.findall(r"VBComponents\.Import\(([^)]+)\)", code, re.IGNORECASE)
    assert imports == ["stagedPath"], "Only a returned staged module path can be imported."
    imported_at = code.index("Set component = project.VBComponents.Import(stagedPath)")
    typed_at = code.index("actualType = CLng(component.Type)", imported_at)
    guarded_at = code.index("If actualType <> expectedType Then", typed_at)
    named_at = code.index("If StrComp(component.Name", guarded_at)
    counted_at = code.index("imported = imported + 1", named_at)
    assert imported_at < typed_at < guarded_at < named_at < counted_at, "Validate type and name before counting an import."
    assert re.search(r"expectedType\s*=\s*1\b", code[imported_at:typed_at]), "Standard modules must require type 1."
    assert re.search(r'If LCase\$\(Right\$\(CStr\(filename\), 4\)\) = "\.cls" Then expectedType = 2\b', code[imported_at:typed_at]), "Class files must require type 2."
    mismatch = code[guarded_at:named_at]
    assert "Err.Raise vbObjectError + 2466" in mismatch, "A wrong component type must stop the build."
    assert "CStr(filename)" in mismatch and "CStr(actualType)" in mismatch and "CStr(expectedType)" in mismatch, "The mismatch error must identify the file and both types."
    failed = code.split("BuildFailed:", 1)[1].split("Private Sub", 1)[0]
    assert '"; description=" & errorText' in failed, "The failure log must retain the component-type error details."
    wrapper = code.split("Public Function SMFMacPrepareCompile", 1)[1].split("End Function", 1)[0]
    assert "SMFMacBuildNoPrompt(projectFolder, False, False)" in wrapper, "The preparation wrapper must disable tests and saving."
    assert "Application.Run" not in wrapper, "The preparation wrapper must not execute a target function."
    assert 'AppleScriptTask("SMFMac.scpt", "stageModule", _' in code
    assert "projectFolder & vbLf & CStr(filename)" in code
    assert ":Library:Containers:com.microsoft.Excel:Data:Library:Caches:SMF-Mac-Build:v0.2:" in code
    assert "SMF-Mac-Build:v0.1:" not in code, "The alternative builder must use the v0.2 staging cache."
    assert "Optional ByVal runTests As Boolean = True" in code
    assert "Optional ByVal saveOutput As Boolean = False" in code
    assert 'AppleScriptTask("SMFMac.scpt", "logBuildState", projectFolder & vbLf & stateText)' in code
    logger = text.split("Private Sub smfMacNoPromptLog", 1)[1].split("End Sub", 1)[0]
    assert "On Error Resume Next" in logger, "Diagnostic logging must not block a build."
    for stage in ("Create a new build workbook", "Access the new workbook's VBA project", "Stage module ", "Import module ",
                  "Run the transport fixtures", "Run the HTML fixtures", "Run the JSON fixtures", "Run the integration fixtures"):
        pattern = r'stage = "' + re.escape(stage) + r'[^\n]*\n\s+smfMacNoPromptLog'
        assert re.search(pattern, code), f"Missing pre-operation progress log: {stage}"
    assert "No add-in file was created." in text
    assert "saved=" in text and "Target workbook:" in text
    assert not re.search(r"\btarget\.Close\b", code, re.IGNORECASE), "Keep the owned target available for inspection."

    block = text.split("BEGIN GENERATED SOURCE FILES", 1)[1].split("END GENERATED SOURCE FILES", 1)[0]
    names = "".join(re.findall(r'"([^"]+)"', block)).split("|")
    expected = {
        p.name for p in (root / "code").iterdir()
        if p.suffix.lower() in {".bas", ".cls"} and p.name.lower() != "exportmods.bas"
    }
    assert len(names) == len(set(names)) == 49
    assert set(names) == expected, "The explicit module list differs from code/."
    assert names[0] == "Dictionary.cls", "Import Dictionary before the JSON parser."
    assert all(name.endswith(".bas") for name in names[1:]), "Class modules must precede standard modules."
    for name in names:
        source = (root / "code" / name).read_text()
        match = re.search(r'^Attribute VB_Name = "([^"]+)"', source, re.MULTILINE)
        assert match and match.group(1).casefold() == name[:-4].casefold(), f"Component-name mismatch: {name}"

    starts = re.findall(r"(?mi)^(?:Public |Private )?(Sub|Function) \w+", text)
    ends = re.findall(r"(?mi)^End (Sub|Function)$", text)
    assert sorted(starts) == sorted(ends), "Procedure boundaries do not match."
    for name in ("SMFMacTransportSelfTest", "smfMacHTMLSelfTest", "SMFMacJSONSelfTest", "SMFMacIntegrationSelfTest"):
        assert name in code, f"Missing fixture group: {name}"
    print("PASS: staged imports with component-type and name guards, no source-file prompts, 49 exact modules, four fixture groups, import-only wrapper, optional output, and VBA structure.")
    print("Static checks do not establish native VBA compilation or runtime behavior.")


if __name__ == "__main__":
    main()
