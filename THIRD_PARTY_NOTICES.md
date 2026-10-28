# Third-party components and distribution boundary

This repository contains project Python source, synthetic tests and documentation. It does not include Origin, Office, installers, activation material, proprietary sample data, fonts, vendor documentation copies or the old `_internal` build folders.

Runtime dependencies are installed separately: pywin32 (PSF-style license), lxml (BSD-style project license plus bundled libxml2/libxslt and related notices), Pillow (HPND-style project license plus bundled codecs), and NumPy (BSD-3-Clause plus bundled numerical library notices). Treat these labels as an inventory, not a substitute for the exact files accompanying the installed versions.

A Windows bundle also includes CPython, Tcl/Tk, runtime DLLs and native libraries. `scripts/build_windows.py` records the build environment and copies installed distribution license/notice files, CPython's license and notices already supplied with Tcl/Tk. Review the generated `LICENSES` directory and native DLL inventory before release. A collection script cannot certify redistribution rights for every transitive component. Do not publish a candidate with unresolved notices or license issues.

PyInstaller's executable exception allows distribution subject to dependency terms: [official license](https://pyinstaller.org/en/stable/license.html). This does not grant rights to redistribute Origin or Office. Users install and license those applications themselves.

Use only local desktop COM automation in this project. The [current OriginLab EULA](https://www.originlab.com/Index.aspx?go=PURCHASE&pid=1775) places additional restrictions on automation-server access, including public network/Internet access. It is version-specific: review the agreement applying to the installed Origin version. No hosted service is provided or authorized by this project.
