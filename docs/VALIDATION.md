# Publication validation — 2026-10-28

## Passed locally

- Existing private v1.1.0 baseline: 9 synthetic regression tests passed before copying.
- Isolated candidate: 9 synthetic regression tests and Tk GUI smoke checks passed on Windows, 64-bit Python 3.12.3.
- Fresh Windows folder EXE builds and its embedded GUI smoke check passes. Exact bundled dependencies and notices are recorded in LICENSES/BUILD_INVENTORY.json.
- Source preparation leaves conversion functions and GUI logic unchanged. The engine's Python requirement docstring, test f-string compatibility and dependency requirements are the existing-file changes.
- After Origin was opened normally, five licensed-desktop regression cases passed with OriginPro 2021 and the updated dependencies: source success; packaged EXE skipping a failed first graph and converting a later graph; packaged strict failure without output; packaged all-skipped run without output; packaged cooperative cancellation without output. Successful presentations also passed native PowerPoint before/after rendering and original-input checksum checks. Skipped objects, groups, geometry and unrelated contents passed package verification.
- Dependency advisory audit of the updated pinned build requirements reports no known vulnerabilities at the time checked. This is an advisory-database check, not proof of absence of vulnerabilities in Python/native DLLs or applications.
- Personal documents, private native fixtures, image files, work reports, old binaries and virtual environments are excluded by the source allowlist. Public tests are synthetic. Original files and installed working tools are untouched.
- Authorized private GitHub upload completed: the remote preparation tree matches all 29 reviewed source files exactly. Windows Python 3.11 and 3.12 each passed unit tests, GUI smoke checks and source auditing in [Windows source checks](https://github.com/Hopkins0920/ppt-origin-to-tiff/actions/runs/40688158344). The pinned dependency advisory check passed in [Dependency advisory check](https://github.com/Hopkins0920/ppt-origin-to-tiff/actions/runs/40688158333). These runs validate the initial uploaded candidate; subsequent commits are checked independently by CI.

## Release blockers / unverified

- The initial unattended native call timed out after 180 seconds. After the user opened Origin, isolated hidden and visible instances both exported 600-DPI TIFFs successfully, and the five regression cases above passed. No conversion-code workaround was required. The initial timeout's precise cause was not established; cold-start behavior on other machines remains to be tested. Do not infer a licensing fault or universal compatibility from this recovery.
- Hosted CI covers Python 3.11/3.12 on Windows, but has no licensed Origin/PowerPoint. Native compatibility evidence remains the separate local regression above. Updated Pillow requires Python 3.11+.
- Second-machine EXE startup/native export, native DLL redistribution review and code signing are not complete. Candidate archives are for local review, not public release.
- The owner created the private `Hopkins0920/ppt-origin-to-tiff` repository, and authorized preparation-branch upload and CI execution. The preparation branch is uploaded and [draft PR #1](https://github.com/Hopkins0920/ppt-origin-to-tiff/pull/1) is open. Project license, attribution, merge and public release remain unapproved. Main contains only the initial review notice; branch protection and release settings are not configured by this upload.

Earlier local v1.1.0 testing recorded 20 conversions plus one explicitly retained failed graph in a 21-object presentation, including PowerPoint previews. This historical result does not validate the updated dependency environment.
