# Contributing / 贡献指南

Use the bug and feature templates. Never submit confidential presentations, embedded data or complete unsanitized reports. A synthetic minimal example is preferred.

Start from `main` on `fix/<description>` or `feat/<description>`. Keep each change focused. Preserve strict and skip modes, cancellation, geometry, original-file protection, unrelated OLE objects and native TIFF export. Avoid changing export thresholds or upgrading dependencies together with unrelated changes.

Run the unit tests, GUI smoke test and source audit described in README. For engine/export or dependency changes, record a separate licensed desktop integration test, including skip and cancel behavior. Tests without Origin are not evidence of native compatibility.

Open a pull request describing the trigger, before/after behavior, risks and validation. The project owner must explicitly approve merging and public release. Do not enable automatic merge or directly push to `main`. Changes to licenses and dependency distribution require review.

中文：先使用独立分支，保留原有保护和转换行为。通过自动检查后再做必要的 Origin/PowerPoint 人工验证，写明结果，请项目所有者确认合并。真实数据只能留在本地，不随 PR 上传。
