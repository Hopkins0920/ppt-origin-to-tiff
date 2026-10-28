"""Build a traceable Windows preview without modifying previous build folders."""
from pathlib import Path
import argparse
from datetime import datetime, timezone
import hashlib
import importlib.metadata as metadata
import json
import os
import shutil
import subprocess
import sys
import tomllib
import urllib.request
import uuid
import zipfile

ROOT = Path(__file__).resolve().parents[1]
MS_RUNTIME = ('vcruntime', 'msvcp', 'mfc', 'concrt', 'vcomp', 'api-ms-', 'ucrtbase')


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def clean_environment():
    env = os.environ.copy()
    for key in ('PYTHONPATH', 'PYTHONHOME', 'PYTHONSTARTUP'):
        env.pop(key, None)
    env['PYTHONNOUSERSITE'] = '1'
    prefix = Path(sys.prefix)
    env['PATH'] = os.pathsep.join(str(prefix / p) for p in ('', 'Library/bin', 'Scripts')) + os.pathsep + env.get('PATH', '')
    return env


def run(*args, env):
    subprocess.run([sys.executable, '-E', '-s', *map(str, args)], cwd=ROOT, env=env, check=True)


def conda_owners():
    owners, records = {}, {}
    for file in (Path(sys.prefix) / 'conda-meta').glob('*.json'):
        record = json.loads(file.read_text(encoding='utf-8'))
        name = record['name']
        records[name] = record
        for relative in record.get('files', []):
            owners[str((Path(sys.prefix) / relative).resolve()).casefold()] = ('conda', name)
    if not records:
        raise RuntimeError('Use the documented conda-forge runtime so native library notices can be traced.')
    return owners, records


def distribution_owners():
    owners = {}
    for name in ('numpy', 'Pillow', 'lxml', 'pywin32', 'pyinstaller'):
        dist = metadata.distribution(name)
        for file in dist.files or ():
            owners[str(Path(dist.locate_file(file)).resolve()).casefold()] = ('python', name)
    return owners


def copy_notices(owner, records, licenses):
    kind, name = owner
    copied = []
    dest = licenses / kind / name
    if kind == 'conda':
        record = records[name]
        cache = Path(record.get('link', {}).get('source', ''))
        source = cache / 'info' / 'licenses'
        if not source.is_dir():
            raise RuntimeError(f'Conda notice cache missing for {name}; retain package caches until packaging completes.')
        for file in source.rglob('*'):
            if file.is_file():
                target = dest / file.relative_to(source)
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(file, target)
                copied.append(target.relative_to(licenses.parent).as_posix())
        info = {key: record.get(key) for key in ('version', 'build', 'license', 'url', 'sha256')}
    else:
        dist = metadata.distribution(name)
        for file in dist.files or ():
            if not any(word in Path(str(file)).name.lower() for word in ('license', 'copying', 'notice')):
                continue
            source = Path(dist.locate_file(file))
            if source.is_file():
                parts = [p for p in Path(str(file)).parts if p not in ('..', '.', os.sep)]
                target = dest / Path(*parts)
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(source, target)
                copied.append(target.relative_to(licenses.parent).as_posix())
        info = {'version': dist.version, 'license': dist.metadata.get('License-Expression') or dist.metadata.get('License')}
    if not copied:
        raise RuntimeError(f'No notices found for {kind}/{name}.')
    return {**info, 'notice_files': sorted(copied)}


def copy_corresponding_sources(components, records, bundle):
    config = tomllib.loads((ROOT / 'third_party_sources.toml').read_text(encoding='utf-8'))
    result = {}
    for kind, name in sorted(components):
        if kind != 'conda' or 'LGPL' not in records[name].get('license', ''):
            continue
        record, source = records[name], config.get(name)
        if not source or source['version'] != record['version']:
            raise RuntimeError(f'Corresponding source configuration missing for {name}.')
        target = bundle / 'SOURCES' / name
        target.mkdir(parents=True)
        archive = target / source['url'].rsplit('/', 1)[-1]
        cache = ROOT / 'artifacts' / 'source-cache' / archive.name
        cache.parent.mkdir(parents=True, exist_ok=True)
        if cache.is_file() and digest(cache) == source['sha256']:
            shutil.copyfile(cache, archive)
        else:
            with urllib.request.urlopen(source['url'], timeout=30) as response, archive.open('wb') as stream:
                shutil.copyfileobj(response, stream)
            if digest(archive) != source['sha256']:
                raise RuntimeError(f'Corresponding source checksum mismatch for {name}.')
            if not cache.exists():
                shutil.copyfile(archive, cache)
        recipe = Path(record['link']['source']) / 'info' / 'recipe'
        if not recipe.is_dir():
            raise RuntimeError(f'Exact build recipe and patches missing for {name}.')
        shutil.copytree(recipe, target / 'recipe')
        result[name] = {**source, 'archive': archive.relative_to(bundle).as_posix(),
                        'recipe': (target / 'recipe').relative_to(bundle).as_posix()}
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-root', type=Path)
    args = parser.parse_args()
    if sys.platform != 'win32' or sys.version_info[:2] != (3, 12):
        raise SystemExit('Build with the documented 64-bit Windows Python 3.12 runtime.')
    if sys.version_info < (3, 12, 16):
        raise SystemExit('Python 3.12.16 or newer security patch is required for this preview.')
    stamp = datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:6]
    output = (args.output_root or ROOT / 'artifacts' / ('portable-' + stamp)).resolve()
    output.mkdir(parents=True, exist_ok=False)
    env = clean_environment()
    owners, records = conda_owners()
    # Conda owns statically linked lxml builds and their native dependencies.
    for path, owner in distribution_owners().items():
        owners.setdefault(path, owner)
    run('origin_tif_gui.py', '--smoke-test', env=env)
    spec = output / 'portable.spec'
    sources = output / 'binary-sources.json'
    spec.write_text(f'''import json
from pathlib import Path
a = Analysis([{str(ROOT / 'origin_tif_gui.py')!r}], pathex=[{str(ROOT)!r}],
             excludes=['win32ui', 'win32trace', 'win32com.shell', 'numpy.testing', 'lxml.isoschematron'])
a.binaries = [row for row in a.binaries if not Path(row[0]).name.lower().startswith({MS_RUNTIME!r})]
Path({str(sources)!r}).write_text(json.dumps(list(a.binaries) + list(a.datas)), encoding='utf-8')
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='OriginTifTool', console=False, upx=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='OriginTifTool')
''', encoding='utf-8')
    run('-m', 'PyInstaller', '--distpath', output / 'dist', '--workpath', output / 'work', spec, env=env)
    bundle = output / 'dist' / 'OriginTifTool'
    source_map = {name.replace('\\', '/').casefold(): Path(source) for name, source, _ in json.loads(sources.read_text(encoding='utf-8'))}
    licenses = bundle / 'LICENSES'
    licenses.mkdir()
    components = {('python', 'pyinstaller'), ('conda', 'python')}
    native = []
    for file in sorted(bundle.rglob('*')):
        if file.suffix.lower() not in ('.dll', '.pyd'):
            continue
        if file.name.lower().startswith(MS_RUNTIME):
            raise RuntimeError(f'Microsoft system/runtime binary unexpectedly bundled: {file.name}')
        relative = file.relative_to(bundle).as_posix()
        key = relative.removeprefix('_internal/').casefold()
        source = source_map.get(key)
        owner = owners.get(str(source.resolve()).casefold()) if source else None
        if not owner or digest(source) != digest(file):
            raise RuntimeError(f'Untraced or changed native binary: {relative}')
        components.add(owner)
        native.append({'path': relative, 'sha256': digest(file), 'component': '/'.join(owner)})
    notices = {'/'.join(owner): copy_notices(owner, records, licenses) for owner in sorted(components)}
    corresponding_sources = copy_corresponding_sources(components, records, bundle)
    inventory = {
        'format_version': 1,
        'application_version': '1.1.0',
        'build_label': 'windows-x64-preview1',
        'created_utc': datetime.now(timezone.utc).isoformat(),
        'python_version': sys.version.split()[0],
        'components': notices,
        'corresponding_sources': corresponding_sources,
        'native_binaries': native,
        'excluded_microsoft_runtime': list(MS_RUNTIME),
        'signing': 'unsigned',
        'source_sha256': {name: digest(ROOT / name) for name in ('origin_pptx_to_tif.py', 'origin_tif_gui.py')},
        'validation_scope': 'Automated Windows build/startup/scan checks; native regression: Windows 11, OriginPro 2021 and desktop PowerPoint.',
    }
    (licenses / 'BUILD_INVENTORY.json').write_text(json.dumps(inventory, indent=2), encoding='utf-8')
    for name in ('LICENSE', 'THIRD_PARTY_NOTICES.md', 'CHANGELOG.md'):
        shutil.copyfile(ROOT / name, bundle / name)
    shutil.copytree(ROOT / 'docs', bundle / 'docs')
    shutil.copyfile(ROOT / 'docs' / 'WINDOWS_PORTABLE.en.md', bundle / 'README.md')
    shutil.copyfile(ROOT / 'docs' / 'WINDOWS_PORTABLE.md', bundle / 'README.zh-CN.md')
    shutil.copyfile(ROOT / 'docs' / 'WINDOWS_PORTABLE.md', bundle / '使用说明.md')
    shutil.copyfile(ROOT / 'docs' / 'WINDOWS_PORTABLE.md', bundle / '使用说明.txt')
    subprocess.run([str(bundle / 'OriginTifTool.exe'), '--smoke-test'], check=True, timeout=90,
                   cwd=output, env=env, creationflags=subprocess.CREATE_NO_WINDOW)
    archive = output / 'ppt-origin-to-tiff-1.1.0-Windows-x64-preview1.zip'
    with zipfile.ZipFile(archive, 'w', compression=zipfile.ZIP_DEFLATED) as zipped:
        for file in sorted(bundle.rglob('*')):
            if file.is_file():
                zipped.write(file, 'OriginTifTool/' + file.relative_to(bundle).as_posix())
    (output / 'SHA256SUMS').write_text(digest(archive) + '  ' + archive.name + '\n', encoding='utf-8')
    (output / 'build-result.json').write_text(json.dumps({'bundle': str(bundle), 'archive': str(archive)}, indent=2), encoding='utf-8')
    print(f'Built preview: {archive.name}; {len(native)} traced native binaries.')
    print('Build directory:', output)


if __name__ == '__main__':
    main()
