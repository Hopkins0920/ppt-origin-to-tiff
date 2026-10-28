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

## Automated Windows validation

The [Windows portable checks](https://github.com/Hopkins0920/ppt-origin-to-tiff/actions/runs/40697244469) passed build, GUI startup, extracted package integrity, Unicode/space paths and synthetic graph inventory on a GitHub Windows runner with neither Origin nor PowerPoint registered. Native export evidence comes from the separate local regression above.

The preview is unsigned and requires the Microsoft Visual C++ v14 x64 runtime as a separate prerequisite. Microsoft system/runtime DLLs are excluded. Release files include component notices, corresponding LGPL source and SHA-256 checksums.
