# Validation — 1.1.0

## Automated checks

Windows CI runs the nine synthetic regression tests, a Tkinter GUI smoke check and a source privacy audit on Python 3.11 and 3.12. A separate workflow checks the pinned build requirements against published dependency advisories. Each source release also requires these checks before packaging.

GitHub-hosted runners do not have licensed Origin or PowerPoint. Synthetic tests validate package handling and simulated failure behavior, not native graph export.

## Licensed desktop regression

Local environment: Windows 11 x64, CPython 3.12.3, OriginPro 2021 and desktop PowerPoint. The tested dependency snapshot includes NumPy 1.26.4, Pillow 13.0.0, lxml 6.1.0 and pywin32 310.

Five native cases passed:

- Successful source conversion to 600-DPI TIFF, with PowerPoint before/after rendering.
- Packaged execution that retained a failed first graph and converted a later graph.
- Strict failure without publishing an output presentation.
- An all-skipped run without publishing an output presentation.
- Cancellation without publishing an output presentation.

Successful runs passed input checksum, slide geometry, grouping, unrelated contents and skipped-object preservation checks. The native fixtures and resulting research images are kept outside this repository.

## Compatibility limits

An initial native call timed out after 180 seconds. After Origin was opened normally, isolated hidden and visible instances exported successfully. The cause of the initial timeout was not established. First-start, recovery and licensing dialogs may require attention before using the converter.

Other Origin/Office versions and application bitness combinations have not been validated. Source code, the nine synthetic tests and the GUI are included in this release; a Windows executable is not distributed. Independent-machine executable startup, native export and bundled native-library redistribution checks are incomplete. The local executable is unsigned.

Earlier local testing converted 20 graphs and retained one failed graph in a 21-object presentation. That historical result used a different dependency environment and is separate from the five cases above.
