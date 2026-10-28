# ppt-origin-to-tiff

A GUI tool for converting embedded Origin graphs in PowerPoint to TIFF images while preserving their original placement

[English](#english) · [中文使用说明](README.zh-CN.md)

## English

A Windows desktop utility for researchers who need to replace embedded Origin graphs in PowerPoint with high-resolution TIFF images. It extracts the embedded native graph, asks the locally installed Origin application to export it through `expGraph`, and replaces only successfully verified objects in a new `.pptx` file.

**Private preparation branch:** the owner has approved uploading this branch for CI review. Merge and public release have not been approved. No public download is available yet. The proposed project license is awaiting approval; see [license decision](docs/LICENSE-DECISION.md).

The Windows executable is still named `OriginTifTool.exe`; existing command-line entry points remain unchanged.

## Features

- Chinese graphical interface with file selection, object inventory, slide numbers, display sizes, per-graph progress and cancellation. This release has English documentation, but no English UI yet.
- Native Origin TIFF export, 600 DPI by default; checks TIFF resolution and effective display density.
- Preserves object size, position, rotation, flips, grouping and stacking order; verifies unrelated slide contents.
- Skip failed exports while preserving their original editable Origin objects. The GUI enables this by default; the CLI requires `--skip-failed`.
- Original files and existing output files are never overwritten. Cooperative cancellation stops at a safe boundary.
- Per-graph comparison images, optional PowerPoint before/after slide renders and a structured report. Similarity checks are heuristics; inspect scientific content before submission.

## Requirements

| Component | Requirement |
| --- | --- |
| OS | 64-bit Windows with an interactive desktop; Windows 11 is the locally tested environment |
| Origin | Installed, licensed and usable through `Origin.Application` COM; Origin 2021 was tested locally |
| PowerPoint | Installed desktop Microsoft PowerPoint COM application for the default slide preview check; optional when previews are disabled |
| Source users | CPython 3.11–3.12, 64-bit, with Tkinter; the local build uses 3.12 (Pillow 13 requires Python 3.11+) |
| EXE users | No separate Python installation; Origin and optional PowerPoint are still required |
| Input | Standard `.pptx` with embedded Origin graphs and cached previews |

Other Office/Origin versions, mixed application bitness and other Windows versions need validation. A browser edition of Office, WPS, Origin Viewer, macOS or Linux is not a supported substitute for this local automation workflow. Font availability can affect output.

## Quick start for non-programmers

After an owner-approved Windows release becomes available, download the Windows ZIP from the repository's Releases page. Extract the **entire folder**, open `OriginTifTool.exe`, and keep its `_internal` and `LICENSES` directories alongside it.

1. Click **选择文件…** (choose file), then review the recognized graphs.
2. Choose a new output filename and DPI. The default is 600 DPI.
3. Leave **跳过失败图表，保留原 Origin 对象并继续** checked to keep failed graphs as editable Origin objects and continue.
4. Click **开始转换** (convert). Red rows were skipped; amber rows were converted but need visual review.
5. Double-click a row for its comparison, or its failure reason if no comparison exists. Open the result and the **检查报告** (report).

If every graph fails, no duplicate presentation is created. Document parsing, application startup, timeout and final validation errors still stop the task.

## Run from source

In a PowerShell window opened in this repository:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe origin_tif_gui.py
```

Command-line examples:

```powershell
python origin_pptx_to_tif.py "C:\Figures\example.pptx" --dry-run
python origin_pptx_to_tif.py "C:\Figures\example.pptx" --dpi 600 --skip-failed
python origin_pptx_to_tif.py "C:\Figures\example.pptx" --output "C:\Figures\result.pptx"
```

`--no-render` disables PowerPoint slide rendering; `--work-dir` chooses the parent of the work directory; `--help` lists options. CLI defaults to strict failure handling. Never lower matching thresholds simply to force acceptance.

## Results, privacy and disk use

Each run creates a `.origin_tif_*` directory beside the input, or beneath `--work-dir`. It contains embedded graph data, exported images, comparison images, logs and `report.json`. Allow several times the presentation size in free space. These files can reveal unpublished research and absolute paths: **do not attach the whole directory to a public issue**. See [privacy guidance](docs/PRIVACY.md).

`completed_with_skips` means that some original Origin objects remain intentionally. Check `converted_count`, `skipped_count` and each `skipped` entry. `all_origin_replaced=false` is expected in a partial conversion; `all_selected_origin_replaced=true` and `skipped_origin_preserved=true` describe the verified successful replacements and retained objects.

## FAQ and limitations

- **Origin cannot start:** open Origin normally and resolve licensing/first-start dialogs. Run `python scripts/check_environment.py` for a registration check. That check does not prove licensing or successful activation.
- **PowerPoint is unavailable:** disable the PowerPoint preview checkbox, or use `--no-render`. Per-graph checks still run.
- **Appearance mismatch or clipping:** inspect the comparison. Keep the skip option enabled to retain that graph. Different fonts, outdated OLE previews and export bounds may explain differences.
- **Does the EXE include Origin or Office?** No. Users install and license them separately. This is a local desktop tool, not a hosted conversion service.
- **Old `.ppt`, linked objects or signed/encrypted files?** Unsupported. Some damaged objects, special cropping/effects and mixed opaque/transparent white areas also require manual processing. Failures during document inspection can stop a whole run even with skipping enabled.
- **Can I interrupt it immediately?** Cancellation is cooperative; it waits for the current native call or file check. An application timeout can leave its temporary application instance open. The tool does not force-kill user applications.
- **Is every output visually identical?** No. Original frame geometry is preserved, but the cached preview and a fresh Origin rendering may differ. White transparency is restored through a PowerPoint color-key effect, which can affect white plot contents. Inspect flagged graphs.
- **Is the EXE signed?** The candidate build is unsigned. Code signing and validation on a second Windows computer are release checklist items. Do not disable security protections to run it.

## Development and maintenance

```powershell
python -m unittest discover -s tests -v
python origin_tif_gui.py --smoke-test
python scripts/audit_source.py
```

Automated tests use synthetic package structures and mock Origin calls; GitHub-hosted CI does not validate real licensed Origin/PowerPoint exports. Native release testing is performed separately on a licensed Windows desktop with private fixtures that are never committed.

See [contributing](CONTRIBUTING.md), [maintenance](docs/MAINTENANCE.md), [release checklist](docs/RELEASING.md), [third-party notices](THIRD_PARTY_NOTICES.md), and [changelog](CHANGELOG.md). Windows builds are manual workflow artifacts, not automatically published releases.

## Upstream interfaces

- [Origin expGraph documentation](https://docs.originlab.com/x-function/ref/expgraph/)
- [PowerPoint Presentations.Open](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)
- [OriginLab current licensing terms](https://www.originlab.com/Index.aspx?go=PURCHASE&pid=1775) — consult the agreement applicable to your installed version; this project grants no Origin license.
