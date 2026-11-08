-- Import the reviewed source and return without executing any target function.
-- Arguments: absolute project folder, then dedicated builder-owner workbook name.
-- This script does not change macro security or VBA project access.
on run argv
    if (count of argv) is not 2 then error "Expected project folder and builder-owner workbook name."
    set projectFolder to item 1 of argv
    set ownerName to item 2 of argv
    if projectFolder does not start with "/" then error "Expected an absolute project-folder path."
    if ownerName does not start with "SMF-Mac-Compile-Owner-" or ownerName does not end with ".xlsm" then error "Use the dedicated compile-diagnostic owner."
    if ownerName contains "'" then error "The owner filename cannot contain an apostrophe."
    with timeout of 300 seconds
        tell application "Microsoft Excel"
            set prepareResult to run VB Macro ("'" & ownerName & "'!SMFMacPrepareCompile") arg1 projectFolder
        end tell
    end timeout
    if prepareResult starts with "FAILED:" then error prepareResult
    set resultLines to paragraphs of prepareResult
    if resultLines does not contain "Imported modules: 49" then error "The builder did not confirm exactly 49 source imports."
    repeat with fixtureName in {"Transport", "HTML", "JSON", "Integration"}
        if resultLines does not contain ((fixtureName as text) & " fixtures: not run") then error "The builder did not confirm that every fixture group stayed unrun."
    end repeat
    if resultLines does not contain "Built in memory. No add-in file was created." then error "The builder did not confirm an unsaved target."
    if resultLines does not contain "saved=False" then error "The builder did not report saved=False."
    if resultLines does not contain "FlagsRestored=True" then error "Excel's display and event settings were not restored."
    set hasTargetWorkbook to false
    set hasTargetProject to false
    repeat with resultLine in resultLines
        set lineText to resultLine as text
        if lineText starts with "Target workbook: " and (length of lineText) > (length of "Target workbook: ") then set hasTargetWorkbook to true
        if lineText starts with "Target VBA project: SMFDiag_" and (length of lineText) > (length of "Target VBA project: SMFDiag_") then set hasTargetProject to true
    end repeat
    if not hasTargetWorkbook or not hasTargetProject then error "The builder did not identify the target workbook and diagnostic VBA project."
    return prepareResult
end run
