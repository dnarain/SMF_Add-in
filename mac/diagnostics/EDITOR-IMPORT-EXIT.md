# Unexpected editor-import exit: November 7, 2026

## What happened

Excel exited while the VBA editor imported only
`mac/build_SMFMac_noprompt.bas` into the new blank owner
`SMF-Mac-Compile-Owner-20261107-01.xlsm`.
The final UI action clicked Open in the Import File dialog.
The script then returned an empty window list. A bounded read-only check
reported that Excel was no longer running.

The imported source SHA256 was:

```text
e71d4ace774e8211ad64ac3f363d78b483179ba967b9cf8da41ca7fff4200ba7
```

No builder function ran. No full-project import, explicit compilation, fixture,
or new HTTP request occurred. VBA project access remained OFF; no Security
option was changed. No owned command session remains pending.

## Disk evidence

The blank owner reported the intended FullName and `saved=True`, but its actual
`.xlsm` file was absent. Only its 165-byte lock file existed.
The owner therefore cannot be treated as a saved build artifact.

The passing two-module probe remains on disk at 25,745 bytes. It had been
verified saved and closed before the failed editor import.
No original probe file was overwritten.

## System evidence

macOS launchservicesd recorded PID `67044` exiting at **13:47:54 PST** with
`LSExitStatus=0`, `exitStatus=0`, and `QUITTING`.
The entries identify neither a quit requester nor a fault signal.
No matching ReportCrash or runningboardd reason entry appeared in the focused query.
Exit status 0 does not identify why Excel quit or exclude an internal error handler.

No Excel report appeared in the readable DiagnosticReports folders.
The operating system denied access to Microsoft's error-reporting container;
that location was not inspected further.
A focused updater/installer query for 13:46:30-13:48:30 returned no matching
Excel update, quit-request, or relaunch evidence.

The installed app's Info.plist now reports Excel **16.114.4**,
build **16.114.26110224**. Earlier 16.114.3 notes are outdated.
The main executable's modification time is November 6 at 20:55:57 local time.
These facts do not establish that an update caused the exit.

## Next proposed diagnostic

All native automation stopped. No full-project compilation result exists.
Before another automatic build, compare one manual import with this automated
editor path. Use a disk copy of the already saved passing probe, which removes
the new blank workbook's failed SaveAs from that check.

Prepared copy:

```text
mac/diagnostics/out/SMF-Mac-Compile-Owner-Manual-20261107-01.xlsm
```

The copy passed ZIP inspection and is byte-for-byte identical to the passing
probe. It has not been opened in Excel. The original remains separate.
The proposed user action is only File > Import File for the builder module,
with VBA project access OFF and no macro execution.
