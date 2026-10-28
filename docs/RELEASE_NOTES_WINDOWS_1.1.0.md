# Windows portable preview1

Download `ppt-origin-to-tiff-1.1.0-Windows-x64-preview1.zip`, extract the **entire folder**, then double-click `OriginTifTool.exe`. No separate Python installation is needed. The interface is Chinese; English and Chinese instructions are included.

## Requirements

- 64-bit Windows 10 or 11.
- Licensed local Origin. Desktop PowerPoint is additionally needed for optional slide previews.
- Microsoft Visual C++ v14 x64 runtime; install or update it from [Microsoft](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist) if Windows reports missing VCRUNTIME DLLs.
- Windows PowerShell 5.1 and GDI+ for cached EMF/WMF previews, normally included with Windows desktop.

Default TIFF export is 600 DPI. Failed graphs can remain editable Origin objects while other graphs are converted. The input file and existing output files are preserved.

## Validation and contents

The conversion engine and GUI match version 1.1.0. Nine unit tests, extracted GUI startup, synthetic graph inventory, Unicode/space path checks, component notices and archive integrity checks passed. Five packaged native regression cases passed on Windows 11 x64 with OriginPro 2021 and desktop PowerPoint, including successful export, skip-then-success, strict failure, all-skipped and cancellation.

The package contains component notices, an inventory of 67 native binaries and corresponding source/recipe/patches for libiconv. Origin and Office are not bundled. The executable is unsigned. `SHA256SUMS` contains the ZIP checksum.

## 中文使用说明

下载名称含 **Windows-x64-preview1.zip** 的文件，完整解压后双击 `OriginTifTool.exe`，无需安装 Python。不要只复制 EXE；保留 `_internal`、`LICENSES` 和 `SOURCES` 文件夹。

转换仍需本机授权 Origin，页面预览另需桌面版 PowerPoint。程序默认导出 600 DPI TIF，支持跳过失败图表并保留原对象。若提示缺少 VCRUNTIME DLL，请从微软官方页面安装 x64 运行库。详细说明在包内《使用说明.txt》中。

本版已通过自动检查及 Windows 11、OriginPro 2021、桌面版 PowerPoint 环境的原生转换回归。请在转换后检查对照图和报告，再用于论文投稿。
