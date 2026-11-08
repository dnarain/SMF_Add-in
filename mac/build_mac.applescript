-- Run after importing build_SMFMac.bas into a disposable workbook.
-- Example: osascript mac/build_mac.applescript /absolute/path/to/smf-mac Book3
-- This script does not change Excel macro security or VBA project access.

on run argv
    if (count of argv) is not 2 then error "Expected project folder and builder workbook name."
    set projectFolder to item 1 of argv
    set builderName to item 2 of argv
    if builderName contains "'" then error "Use a builder workbook name without an apostrophe."
    with timeout of 45 seconds
        tell application "Microsoft Excel"
            activate
            set buildResult to run VB Macro ("'" & builderName & "'!SMFMacBuild") arg1 projectFolder arg2 true
        end tell
        if buildResult starts with "FAILED:" then error buildResult
        return buildResult
    end timeout
end run
