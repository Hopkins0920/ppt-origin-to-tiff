# Windows 便携版使用说明

此包为 **1.1.0 Windows x64 preview1 测试版**，未进行代码签名。请完整解压后使用，不要只复制 EXE。

## 启动

1. 把整个 `OriginTifTool` 文件夹解压到可写目录，例如桌面或文档目录。
2. 双击 `OriginTifTool.exe`。打开中文窗口即表示启动成功，无需安装 Python。
3. 选择 PPTX 文件，设置新的保存文件名，分辨率默认 600 DPI。
4. 保留“跳过失败图表”选项可在遇到单图错误时保留原对象并继续。
5. 转换后复核对照图、跳过数量和检查报告。原文件不会被覆盖。

## 所需环境

- 64 位 Windows 10／11；本机完整转换测试使用 Windows 11。
- Windows PowerShell 5.1 和系统 GDI+，用于读取 PPT 中的 EMF／WMF 缓存预览；标准 Windows 桌面通常已包含。
- Microsoft Visual C++ v14 x64 运行库。此包不捆绑微软系统 DLL；若提示缺少 `VCRUNTIME140.dll` 或 `VCRUNTIME140_1.dll`，请从 [Microsoft 官方页面](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist)安装或更新 x64 运行库。
- 完整转换需要本机已授权、能正常启动的 Origin。
- 页面前后预览另需桌面版 PowerPoint；没有 PowerPoint 时取消预览选项。
- 未安装 Origin 的电脑仍可做窗口启动测试，不能验证实际图形导出。

测试版没有数字签名，不能用签名确认发布者身份。请从项目认可的下载位置获取，使用 `SHA256SUMS` 核对文件；不要关闭安全防护。

## 在另一台电脑上试运行

1. 复制 ZIP 并完整解压，双击 EXE。
2. 确认窗口正常显示，能打开选择文件对话框，然后关闭窗口。
3. 反馈 Windows 版本、是否装有 Python、是否启动成功。失败时记录错误原文即可，不必发送科研 PPT。
4. 没有 Origin 时，到这里即可；无需为启动测试安装 Origin。

另一台电脑的手动启动测试结果尚待确认。本机原生转换测试与异机启动测试分别记录，不代表已覆盖所有 Origin/Office 版本。

## 许可和组件

项目代码采用 MIT 许可，第三方组件遵循 `LICENSES` 中的各自条款。`BUILD_INVENTORY.json` 记录组件版本、原生 DLL 来源与文件校验值，不包含个人安装路径。

This product includes software developed by the FreeType Project (https://www.freetype.org/).

The dynamically linked libiconv library is covered by LGPL-2.1. Its corresponding upstream source archive, exact conda-forge build recipe and patches are included in `SOURCES/libiconv`. You may replace the shared library with a compatible modified version and reverse engineer this application as needed to debug modifications to that LGPL-covered library. The original notices and LGPL text are included in `LICENSES/conda/libiconv`.

Python has not been modified by this project; its runtime and notices are packaged from conda-forge. Origin and PowerPoint are installed separately and are not included in this package.
