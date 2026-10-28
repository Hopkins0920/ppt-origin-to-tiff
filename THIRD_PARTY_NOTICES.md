# Third-party components and distribution boundary

This repository contains project Python source, synthetic tests and documentation. It does not include Origin, Office, installers, activation material, proprietary sample data, fonts, vendor documentation copies or the old `_internal` build folders.

Source runtime dependencies are installed separately: pywin32 (PSF-style license), lxml (BSD-style project license plus XML library notices), Pillow (MIT-CMU project license plus bundled codec notices), and NumPy (BSD-3-Clause plus numerical library notices). Exact installed notice files take precedence over these inventory labels.

A Windows bundle also includes CPython, Tcl/Tk, runtime DLLs and native libraries. `scripts/build_windows.py` records the build environment and copies installed distribution license/notice files, CPython's license and notices already supplied with Tcl/Tk. Review the generated `LICENSES` directory and native DLL inventory before release. A collection script cannot certify redistribution rights for every transitive component. Do not publish a candidate with unresolved notices or license issues.

The portable preview uses `scripts/build_portable.py` and the exact conda-forge runtime in `requirements-portable-conda.txt`. It maps every bundled DLL/PYD to an installed package, preserves that package's notice files and records file hashes. Microsoft runtime/system DLLs are excluded; the Microsoft Visual C++ v14 x64 runtime is a separate prerequisite. Unrecognized native binaries fail packaging.

The dynamically linked libiconv component is LGPL-2.1. The preview includes the corresponding upstream source tarball, exact conda-forge recipe and patches in `SOURCES/libiconv`, in addition to its license. Users may replace the shared library and reverse engineer modifications as needed to debug it. This permission does not change the licenses of other components.

This product includes software developed by the FreeType Project (https://www.freetype.org/). Its notices and other Pillow codec terms accompany the Pillow distribution license.

PyInstaller's executable exception allows distribution subject to dependency terms: [official license](https://pyinstaller.org/en/stable/license.html). This does not grant rights to redistribute Origin or Office. Users install and license those applications themselves.

Use only local desktop COM automation in this project. The [current OriginLab EULA](https://www.originlab.com/Index.aspx?go=PURCHASE&pid=1775) places additional restrictions on automation-server access, including public network/Internet access. It is version-specific: review the agreement applying to the installed Origin version. No hosted service is provided or authorized by this project.
