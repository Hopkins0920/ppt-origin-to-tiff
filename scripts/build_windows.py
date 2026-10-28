"""Build an isolated folder bundle and license inventory; never uploads it."""
from pathlib import Path
import hashlib
import importlib.metadata as meta
import json
import os
import shutil
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def run(*args):
    subprocess.run([sys.executable, *args], cwd=ROOT, check=True)


def main():
    if sys.platform != 'win32':
        raise SystemExit('Windows desktop Python is required to build this candidate.')
    run('origin_tif_gui.py', '--smoke-test')
    run('-m', 'PyInstaller', '--noconfirm', '--onedir', '--windowed', '--noupx',
        '--name', 'OriginTifTool', '--distpath', 'dist', '--workpath', 'build', 'origin_tif_gui.py')
    bundle = ROOT / 'dist' / 'OriginTifTool'
    subprocess.run([str(bundle / 'OriginTifTool.exe'), '--smoke-test'], check=True,
                   creationflags=subprocess.CREATE_NO_WINDOW)
    licenses = bundle / 'LICENSES'; licenses.mkdir(exist_ok=True)
    inventory = {}
    for name in ('numpy', 'Pillow', 'lxml', 'pywin32', 'pyinstaller'):
        dist = meta.distribution(name)
        copied = []
        for file in dist.files or ():
            if any(word in Path(str(file)).name.lower() for word in ('license', 'copying', 'notice')):
                source = Path(dist.locate_file(file))
                if source.is_file():
                    # Keep full package-relative path so nested third-party notices cannot collide.
                    safe_parts = [p for p in Path(str(file)).parts if p not in ('..', '.', os.sep)]
                    dest = licenses / name / Path(*safe_parts)
                    dest.parent.mkdir(parents=True, exist_ok=True); shutil.copyfile(source, dest)
                    copied.append(str(dest.relative_to(bundle)).replace('\\', '/'))
        inventory[name] = {'version': dist.version, 'notice_files': copied}
        if not copied:
            raise RuntimeError(f'No installed license files found for {name}; inspect distribution before publishing.')
    python_license = next((p for p in (Path(sys.base_prefix)/'LICENSE.txt', Path(sys.base_prefix)/'Lib'/'LICENSE.txt', Path(sys.base_prefix)/'LICENSE_PYTHON.txt') if p.is_file()), None)
    if not python_license:
        raise RuntimeError('CPython license not located; do not distribute an incomplete bundle.')
    shutil.copyfile(python_license, licenses/'CPython-LICENSE.txt')
    # Some distributions omit notices from the PyInstaller hook data collection.
    import tkinter
    tcl = tkinter.Tcl()
    tcl_dir = Path(tcl.eval('info library'))
    for runtime in (tcl_dir, *tcl_dir.parent.glob('tk*')):
        if runtime.is_dir():
            for notice in runtime.glob('*license*'):
                if notice.is_file():
                    shutil.copyfile(notice, licenses/(runtime.name + '-' + notice.name))
    tcl_notices = [str(p.relative_to(bundle)).replace('\\', '/') for p in bundle.rglob('*')
                   if p.is_file() and any(t in str(p).lower() for t in ('_tcl_data', '_tk_data', 'licenses/tcl', 'licenses/tk', 'licenses\\tcl', 'licenses\\tk'))
                   and 'license' in p.name.lower()]
    if not tcl_notices:
        raise RuntimeError('Tcl/Tk license notices missing; review runtime before distribution.')
    inventory['CPython'] = {'version': sys.version.split()[0]}
    inventory['TclTk_notices'] = tcl_notices
    inventory['native_binaries'] = sorted(str(p.relative_to(bundle)).replace('\\', '/') for p in bundle.rglob('*') if p.suffix.lower() in ('.dll', '.pyd'))
    inventory['status'] = 'candidate: owner approval, dependency security review, native DLL license review and independent-machine validation required'
    (licenses/'BUILD_INVENTORY.json').write_text(json.dumps(inventory, indent=2), encoding='utf-8')
    for name in ('README.md', 'README.zh-CN.md', 'THIRD_PARTY_NOTICES.md'):
        shutil.copyfile(ROOT/name, bundle/name)
    if (ROOT/'LICENSE').is_file():
        shutil.copyfile(ROOT/'LICENSE', bundle/'LICENSE')
    else:
        shutil.copyfile(ROOT/'docs'/'LICENSE-DECISION.md', bundle/'LICENSE-PENDING.md')
    artifacts = ROOT/'artifacts'; artifacts.mkdir(exist_ok=True)
    archive = artifacts/'OriginTifTool-1.1.0-Windows-x64-candidate.zip'
    with zipfile.ZipFile(archive, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        for path in sorted(bundle.rglob('*')):
            if path.is_file(): z.write(path, 'OriginTifTool/' + path.relative_to(bundle).as_posix())
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    (artifacts/'SHA256SUMS').write_text(digest + '  ' + archive.name + '\n', encoding='utf-8')
    print('Created local candidate:', archive.name)


if __name__ == '__main__':
    main()
