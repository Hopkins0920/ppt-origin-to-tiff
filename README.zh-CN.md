# ppt-origin-to-tiff：将 PPT 中的 Origin 图表替换为高清 TIF

[English README](README.md)

面向科研人员的 Windows 小工具。它读取 PowerPoint 中可双击打开的 Origin 嵌入图表，通过本机 Origin 的 `expGraph`（导出图形）生成 TIF，再将通过检查的图表替换到新 PPT 中。默认 600 DPI，原文件保留。

**1.1.0 版 · MIT 许可证 · Windows · 中文界面**

提供 [Windows 便携版 ZIP](https://github.com/Hopkins0920/ppt-origin-to-tiff/releases/tag/v1.1.0-windows-preview1) 和[源码包](https://github.com/Hopkins0920/ppt-origin-to-tiff/releases/tag/v1.1.0)。便携包已包含 Python、组件许可和 SHA-256 校验文件。详见[便携版验证记录](docs/PORTABLE_VALIDATION.md)。

项目名称为 `ppt-origin-to-tiff`；为兼容现有使用方式，Windows 程序仍叫 `OriginTifTool.exe`，命令行入口文件名也不变。

## 主要功能

- 中文窗口选择文件，显示页码、对象 ID、尺寸、组合层数和逐图进度。
- 原生 Origin 导出，核对 TIF 的 DPI 和放入 PPT 后的实际像素密度。
- 保留原图框尺寸、位置、旋转、翻转、组合和叠放顺序，并检查其他页面内容。
- 默认跳过单图导出或外观校验失败的对象：保留原来的可编辑 Origin 图，继续处理其余图表。
- 显示跳过原因、前后对比和检查报告；支持取消和命令行批处理。
- 英文文档已提供；本版窗口界面仍为中文。

## 环境与安装

| 依赖 | 要求 |
| --- | --- |
| Windows | 64 位、可交互桌面；本地验证环境为 Windows 11 |
| Origin | 自行安装、合法授权，可通过 COM 调用；本地已验证 Origin 2021 |
| PowerPoint | 默认的页面前后预览需要桌面版 PowerPoint；可取消勾选以跳过此步骤 |
| 缓存预览 | Windows PowerShell 5.1 和 GDI+，用于 EMF／WMF 预览；标准 Windows 桌面通常已包含 |
| Python | EXE 用户无需另装；源码用户使用带 Tkinter 的 64 位 CPython 3.11–3.12，本地打包使用 3.12；Pillow 13 要求 3.11+ |
| EXE 运行库 | Microsoft Visual C++ v14 x64；出现缺少 VCRUNTIME DLL 的提示时，按便携版说明安装微软官方运行库 |
| 文件 | 标准 `.pptx`，含原生数据和缓存预览的 Origin 嵌入对象 |

浏览器版 Office、WPS 和 Origin Viewer 不能替代这里调用的应用。其他 Origin/Office 版本、应用位数搭配和其他 Windows 版本需要实测。字体也会影响图形外观。工具不捆绑、不提供 Origin 或 Office 授权。

## 窗口使用教程

下载 [Windows 便携版 ZIP](https://github.com/Hopkins0920/ppt-origin-to-tiff/releases/tag/v1.1.0-windows-preview1)，**完整解压**后双击 `OriginTifTool.exe`，无需另装 Python。请保留旁边的 `_internal`、`LICENSES` 和 `SOURCES` 文件夹。运行要求见[便携版说明](docs/WINDOWS_PORTABLE.md)。

1. 点“选择文件”，选择 `.pptx`，查看识别列表。
2. 选择新的保存文件名，分辨率默认 600 DPI。
3. 建议保留勾选“跳过失败图表，保留原 Origin 对象并继续”。取消勾选则遇错停止。
4. 点“开始转换”。红色表示已跳过并保留原图，橙色表示已转换但建议复核。
5. 双击行查看对照图或失败原因。完成后打开结果与检查报告，复核投稿用图。

所有对象均失败时，不生成重复的 PPT。取消会等当前原生调用或文件检查结束；进度百分比不是剩余时间估算。原文件和已有结果不会被覆盖。

## 从源码运行

在本项目文件夹打开 PowerShell：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe origin_tif_gui.py
```

命令行示例（命令行默认遇错停止，需要显式开启跳过）：

```powershell
python origin_pptx_to_tif.py "C:\Figures\example.pptx" --dry-run
python origin_pptx_to_tif.py "C:\Figures\example.pptx" --dpi 600 --skip-failed
```

`--output` 指定新文件名，`--work-dir` 指定中间文件父目录，`--no-render` 跳过 PowerPoint 页面预览。`python scripts/check_environment.py` 只检查本机模块与 COM 注册，不会启动应用，也不能证明授权有效。

## 报告与隐私

每次运行生成 `.origin_tif_*` 工作目录，包含 Origin 数据、TIF、对照图、日志和 `report.json`；大文件应预留数倍空间。中间材料可能含未发表数据及个人路径，**不要直接上传到公开 Issue**，见[隐私说明](docs/PRIVACY.md)。

`completed_with_skips` 表示部分转换；`converted_count`、`skipped_count` 分别记录转换与跳过数量。此时 `all_origin_replaced=false` 是预期结果；另外两项 `all_selected_origin_replaced=true`、`skipped_origin_preserved=true` 表示目标替换及保留项通过检查。

## 常见问题和已知限制

- **Origin 启动失败：**先自行打开应用，解决首次启动、恢复和授权对话框；确认 COM 已注册。
- **没有 PowerPoint：**取消页面预览选项，仍保留逐图检查。其他 Office 替代软件未验证。
- **外观不匹配、边缘裁切：**查看对比图，保留跳过选项。不要仅降低相似度阈值强行通过。
- **旧 `.ppt`、外部链接、加密或数字签名文件：**不支持。损坏对象、特殊裁切/效果、透明背景混合不透明白色内容可能需要人工处理。读取结构时失败仍可能停止整次转换。
- **为何大小一致但字体不同：**PPT 缓存预览与当前 Origin 图形设置、版本、字体可能不同。几何检查与相似度不能证明科研内容完全一致。
- **白色透明效果：**为恢复原 OLE 的透明背景，PPT 使用白色透明色键，也可能使图中白色内容透明；检查橙色提示。
- **在哪里下载 EXE：**从 [Windows 便携版发布页](https://github.com/Hopkins0920/ppt-origin-to-tiff/releases/tag/v1.1.0-windows-preview1)下载 ZIP 并完整解压。程序未进行代码签名；若缺少 VCRUNTIME DLL，请按便携版说明安装微软官方 x64 运行库。
- **卡住或超时：**原生调用结束前取消可能需要等待；超时可能遗留临时应用实例，工具不会强杀用户应用。

## 测试与长期维护

自动测试只使用合成结构与模拟接口；公开 CI 不包含授权 Origin/PowerPoint，也不含真实科研 PPT。正式版本还需维护者在授权 Windows 桌面上做原生导出测试。

请通过 Bug／功能建议模板描述问题；提问前删除个人路径和未发表内容。维护者按独立分支修改、自动测试、人工回归、所有者批准合并、更新版本记录的顺序工作。

[贡献指南](CONTRIBUTING.md) · [维护流程](docs/MAINTENANCE.md) · [发布清单](docs/RELEASING.md) · [第三方说明](THIRD_PARTY_NOTICES.md) · [版本记录](CHANGELOG.md)

## 开源许可

项目代码和文档采用 [MIT 许可证](LICENSE)，版权署名为 © 2026 Hopkins0920。允许使用、修改和再分发，包括商业用途，需保留版权及许可声明。Origin、PowerPoint 和第三方依赖分别遵循其自身许可；本项目的许可不涉及用户处理的科研数据。
