-- Import reviewed source files through the native VBA editor.
-- Run in batches of five files after selecting the dedicated build project.
-- Arguments: build workbook name, then absolute source-file paths.
-- This script does not use VBProject access or change macro security.

on run argv
    if (count of argv) < 2 then error "Expected build workbook name and source-file paths."
    set bookName to item 1 of argv
    if bookName does not start with "SMF-Mac-" then error "Use a dedicated SMF-Mac build workbook."
    set importedNames to {}
    with timeout of 45 seconds
        tell application "Microsoft Excel"
            activate
            activate object workbook bookName
        end tell
        tell application "System Events"
            tell process "Microsoft Excel"
                set hasBuildWindow to exists window ("Microsoft Visual Basic - " & bookName)
                if not hasBuildWindow and (exists menu item "Macro" of menu 1 of menu bar item "Tools" of menu bar 1) then
                    click menu item "Visual Basic Editor" of menu 1 of menu item "Macro" of menu 1 of menu bar item "Tools" of menu bar 1
                end if
            end tell
        end tell
        my waitForEditor(bookName)
        repeat with fileArgument in items 2 thru -1 of argv
            set sourcePath to fileArgument as text
            if sourcePath does not start with "/" then error "Use an absolute source-file path."
            if sourcePath does not end with ".bas" and sourcePath does not end with ".cls" then error "Expected a VBA module or class file."
            tell application "System Events"
                tell process "Microsoft Excel"
                    if exists window "Import File" then error "Resolve the existing import dialog before another batch."
                    if not (exists menu item "Import File..." of menu 1 of menu bar item "File" of menu bar 1) then error "Open the VBA editor and select the build project first."
                    if not (exists window ("Microsoft Visual Basic - " & bookName)) then error "The dedicated build project is not active."
                    click menu item "Import File..." of menu 1 of menu bar item "File" of menu bar 1
                end tell
            end tell
            my waitForImportDialog(true)
            tell application "System Events" to tell process "Microsoft Excel" to keystroke "g" using {command down, shift down}
            my waitForPathSheet(true)
            tell application "System Events"
                tell process "Microsoft Excel"
                    set value of text field 1 of sheet 1 of window "Import File" to sourcePath
                    keystroke return
                end tell
            end tell
            my waitForPathSheet(false)
            tell application "System Events"
                tell process "Microsoft Excel"
                    if not (enabled of button "Open" of splitter group 1 of window "Import File") then error "The selected source file cannot be opened."
                    click button "Open" of splitter group 1 of window "Import File"
                end tell
            end tell
            my waitForImportDialog(false)
            set end of importedNames to sourcePath
            log "Imported: " & sourcePath
        end repeat
    end timeout
    return importedNames
end run

on waitForEditor(bookName)
    repeat with attemptNo from 1 to 100
        tell application "System Events"
            tell process "Microsoft Excel"
                set hasImportMenu to exists menu item "Import File..." of menu 1 of menu bar item "File" of menu bar 1
                set hasBuildWindow to exists window ("Microsoft Visual Basic - " & bookName)
                set hasMacroMenu to exists menu item "Macro" of menu 1 of menu bar item "Tools" of menu bar 1
                if attemptNo is 1 or attemptNo is 100 then
                    set windowNames to name of every window
                    log {"Editor state", hasImportMenu, hasBuildWindow, hasMacroMenu, windowNames}
                end if
                if not hasBuildWindow and hasMacroMenu then
                    click menu item "Visual Basic Editor" of menu 1 of menu item "Macro" of menu 1 of menu bar item "Tools" of menu bar 1
                end if
            end tell
        end tell
        if hasImportMenu and hasBuildWindow then return
        if hasBuildWindow then my focusBuildWindow(bookName)
        delay 0.03
    end repeat
    error "The dedicated build project did not become active in the VBA editor."
end waitForEditor

on focusBuildWindow(bookName)
    tell application "System Events"
        tell process "Microsoft Excel"
            set w to window ("Microsoft Visual Basic - " & bookName)
            perform action "AXRaise" of w
            set p to position of w
            set s to size of w
            click at {(item 1 of p) + ((item 1 of s) div 2), (item 2 of p) + 12}
        end tell
    end tell
end focusBuildWindow

on waitForImportDialog(shouldExist)
    repeat 100 times
        tell application "System Events" to tell process "Microsoft Excel" to set isPresent to exists window "Import File"
        if isPresent is shouldExist then return
        delay 0.03
    end repeat
    error "The Import File dialog did not reach the expected state."
end waitForImportDialog

on waitForPathSheet(shouldExist)
    repeat 100 times
        tell application "System Events" to tell process "Microsoft Excel" to set isPresent to exists sheet 1 of window "Import File"
        if isPresent is shouldExist then return
        delay 0.03
    end repeat
    error "The source-path sheet did not reach the expected state."
end waitForPathSheet
