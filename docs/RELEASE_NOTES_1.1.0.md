# ppt-origin-to-tiff 1.1.0

Convert embedded Origin graphs in PowerPoint to native TIFF images, with a Chinese graphical interface and a default resolution of 600 DPI.

## Features

- Inventory embedded graphs and show conversion progress.
- Preserve original placement, dimensions, rotation, grouping and stacking order.
- Keep failed graphs as editable Origin objects and continue with the remaining graphs.
- Write a new presentation, comparison images and a structured report.
- Support cancellation, command-line operation and optional PowerPoint slide previews.

## Download and installation

This is a **source release** under the MIT License. Download `ppt-origin-to-tiff-1.1.0-source.zip`, extract it and follow the English or Chinese README. `SHA256SUMS` contains the archive checksum.

Requirements: Windows desktop, 64-bit CPython 3.11–3.12 with Tkinter, the listed Python packages and a licensed local Origin installation. Desktop PowerPoint is required for the optional slide preview step. Origin and Office are not included.

Windows users can download the separate [portable package](https://github.com/Hopkins0920/ppt-origin-to-tiff/releases/tag/v1.1.0-windows-preview1).

## Validation and limitations

Windows CI covers Python 3.11/3.12, nine synthetic tests, GUI startup, source auditing and dependency advisories. Local native regression used Windows 11 x64, Python 3.12.3, OriginPro 2021 and desktop PowerPoint; five success/failure/cancellation cases passed. Other Origin/Office versions remain unverified. See [validation](https://github.com/Hopkins0920/ppt-origin-to-tiff/blob/main/docs/VALIDATION.md).

Font differences, stale cached previews, cropping, special effects and white color-key transparency can change the appearance. Inspect converted figures before submission. Linked objects, legacy `.ppt`, encrypted and digitally signed files are unsupported. Cancellation waits for a safe boundary; native application timeouts may leave temporary instances open.

## 中文说明

本版采用 MIT 许可，以源码形式发布。默认通过 Origin 原生导出 600 DPI TIF，保留图框位置和尺寸；失败图表可保留原对象并继续。请下载源码 ZIP，按照 `README.zh-CN.md` 安装依赖并打开窗口。

源码运行需要 Windows、本机授权 Origin 和 Python 3.11–3.12；页面预览另需桌面版 PowerPoint。另提供独立的 Windows 便携版下载。自动测试与本机原生回归已通过，其他 Origin/Office 版本尚未验证。投稿前请复核对照图；报告和中间文件可能含未发表数据，不应直接上传至公开 Issue。
