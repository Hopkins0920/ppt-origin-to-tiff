# Windows portable preview validation

The preview packages the unchanged 1.1.0 conversion engine and GUI. Packaging and native runtime changes are reviewed separately from the stable source release.

## Passed locally — 2026-10-28

- Nine synthetic regression tests and dependency consistency checks.
- Pinned Python package advisory audit: no known vulnerabilities reported at the time checked.
- Fresh conda-forge Python 3.12.16 runtime; lxml 6.1.0 with dynamic libxml2 2.15.4 and libxslt 1.1.45.
- DLL/PYD source mapping, SHA-256 verification and package notice collection for all 67 native binaries.
- libiconv 1.18 corresponding upstream source archive, exact recipe and patches included; source checksum verified.
- Extracted ZIP GUI smoke check and synthetic presentation scan with Python/conda removed from PATH, in a path containing Chinese characters and spaces.
- Five packaged native regression cases with Windows 11, OriginPro 2021 and desktop PowerPoint: successful conversion; failed graph retained followed by a successful conversion; strict failure; all graphs skipped; cooperative cancellation.
- Original file checksums, geometry/groups, skipped OLE objects and PowerPoint before/after rendering passed in successful native runs.

The native fixtures, figures, full reports and personal paths are not included in the repository or preview archive.

## Remaining scope

The independent GitHub Windows runner performs build/startup/scan checks without licensed Origin or PowerPoint. Its result is recorded by the Windows portable checks workflow; it does not validate native exports.

Manual startup on another Windows desktop is pending. That computer does not have Origin, so its test covers opening the GUI and file chooser only. Other Origin/Office versions remain unverified.

The preview is unsigned and requires the Microsoft Visual C++ v14 x64 runtime as a separate prerequisite. Microsoft system/runtime DLLs are excluded. No executable release or source-branch merge is automatic.
