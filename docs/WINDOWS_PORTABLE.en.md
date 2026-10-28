# Windows portable preview

This is the **1.1.0 Windows x64 preview1** package. It is unsigned. Extract the complete folder and keep the executable, `_internal`, `LICENSES` and `SOURCES` together.

## Start

Double-click `OriginTifTool.exe`. No separate Python installation is needed. The interface is Chinese; see `README.zh-CN.md` for the button labels. Select a PPTX, choose a new output filename and keep the skip-failed option enabled to retain failed editable graphs.

## Requirements

- 64-bit Windows 10 or 11; native local regression used Windows 11.
- Microsoft Visual C++ v14 x64 runtime. Microsoft system/runtime DLLs are not bundled. If Windows reports missing VCRUNTIME DLLs, install or update the x64 runtime from [Microsoft](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist).
- Windows PowerShell 5.1 and GDI+ for EMF/WMF cached previews; these are normally included in standard Windows desktop installations.
- Licensed local Origin for conversion. Desktop PowerPoint is additionally required for optional slide previews.

On a computer without Origin, opening the window and file dialog is still a useful startup test. It does not validate native graph export. Independent computer startup confirmation is pending; other Origin/Office versions remain unverified.

Use the archive checksum supplied alongside the package. No code signature is provided; do not disable security protections. Keep unpublished presentations and full work reports out of public issues.

## Components and licenses

Project code is MIT-licensed. Third-party terms are included in `LICENSES`; `BUILD_INVENTORY.json` records component versions, native binary origins and hashes without personal installation paths.

This product includes software developed by the FreeType Project (https://www.freetype.org/).

The dynamically linked libiconv library is LGPL-2.1. Its corresponding upstream source archive, exact conda-forge build recipe and patches are included in `SOURCES/libiconv`. You may replace the library with a compatible modified version and reverse engineer this application as needed to debug modifications to that LGPL-covered library. Notices and the LGPL text are in `LICENSES/conda/libiconv`.

Python is packaged from conda-forge without changes by this project. Origin and PowerPoint are installed and licensed separately; neither is included in this archive.
