#!/bin/zsh
set -euo pipefail
scriptDirectory="${0:A:h}"
targetDirectory="$HOME/Library/Application Scripts/com.microsoft.Excel"
targetFile="$targetDirectory/SMFMac.scpt"
/bin/mkdir -p "$targetDirectory"
if [[ -e "$targetFile" ]]; then
    backupFile="$targetFile.backup.$(/bin/date +%Y%m%d-%H%M%S).$$"
    /bin/cp -p "$targetFile" "$backupFile"
    print -r -- "Saved the previous helper: $backupFile"
fi
temporaryFile="$targetDirectory/.SMFMac.$$.scpt"
trap '/bin/rm -f "$temporaryFile"' EXIT
/usr/bin/osacompile -o "$temporaryFile" "$scriptDirectory/SMFMac.applescript"
/bin/chmod 600 "$temporaryFile"
/bin/mv "$temporaryFile" "$targetFile"
print -r -- "Installed the Mac HTTP helper: $targetFile"
print -r -- "This script does not change Excel security settings."
