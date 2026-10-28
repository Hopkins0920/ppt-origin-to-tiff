# Changelog

## 1.1.0 — first public release candidate (not published)

- Native Origin TIFF export with a graphical inventory and progress display.
- Default GUI skip mode retains failed editable graphs and continues; CLI remains strict by default.
- Geometry, package relationships, unaffected contents and native TIFF bytes are checked.
- Safe cancellation, per-object diagnostics and optional PowerPoint previews.
- Bilingual documentation, synthetic regression tests, manual Windows build workflow and publication audit.
- Publication preparation preserves conversion logic, updates the Python requirement text and test syntax, and adds release tooling/documentation.
- Isolated candidate dependencies use Pillow 13.0.0 and lxml 6.1.0 following advisory checks. Source installations now require Python 3.11+; existing private installations are untouched.

## Local history

1.0 introduced the local converter. 1.1 added per-graph skipping. These were private local builds, not prior GitHub releases.
