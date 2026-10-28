# OriginTifTool 1.1.0 — publication candidate

Private preparation/CI review is authorized; no public release or merge is approved. Project license and owner attribution are pending.

Native Origin-to-TIFF conversion, Chinese GUI, per-object skipping, original-file protection, comparison reports and optional PowerPoint slide previews. Failed graphs remain editable Origin objects; partial conversion is explicitly reported.

Requirements: local Windows desktop, licensed Origin COM server, optional licensed desktop PowerPoint; source requires Python and listed dependencies. EXE does not include Origin/Office. No English GUI in this release.

Current local validation: nine synthetic tests and source/packaged GUI smoke tests pass. Five native regression cases pass: successful conversion with PowerPoint renders; skip then later success with preserved OLE; strict failure; all skipped without output; cancellation without output. This is separate from historical private regressions. See docs/VALIDATION.md for the initial cold-start timeout and recovery.

Build environment: Windows 11 x64, Python 3.12.3, OriginPro 2021, locally installed desktop PowerPoint; NumPy 1.26.4, Pillow 13.0.0, lxml 6.1.0, pywin32 310, PyInstaller 6.13.0. Fixed Python build dependencies passed the advisory check on 2026-10-28. Python/native DLL review, independent-machine validation, signature arrangements, project license and owner publication approval remain pending. Other Origin/Office versions are unverified.

Known limitations: cached previews versus fresh render differences; font availability; color-key transparency; special OLE effects/cropping; unsupported linked/signed/encrypted/legacy input; cooperative cancellation and possible leftover temporary native instances after timeouts.

中文：这是待审核候选，不是已公开的正式发布。正式发布记录须附实际测试环境、跳过数量说明、校验结果、依赖审查和签名状态；不要使用科研原图作为公开示例。
