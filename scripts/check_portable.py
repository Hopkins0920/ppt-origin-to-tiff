"""Verify and smoke-test an extracted preview with Python paths removed."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import subprocess
import tempfile
import zipfile
import winreg


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def registered(prog_id):
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id + '\\CLSID'):
            return True
    except FileNotFoundError:
        return False


def synthetic_presentation(path):
    p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    ns = f'xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}"'
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        'ppt/presentation.xml': f'<p:presentation {ns}><p:sldIdLst><p:sldId id="256" r:id="s"/></p:sldIdLst></p:presentation>',
        'ppt/_rels/presentation.xml.rels': f'<Relationships xmlns="{rel}"><Relationship Id="s" Target="slides/slide1.xml"/></Relationships>',
        'ppt/slides/slide1.xml': f'<p:sld {ns}><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Synthetic graph"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></p:xfrm><a:graphic><a:graphicData><p:oleObj r:id="ole" progId="Origin95.Graph" imgW="914400" imgH="457200"><p:embed/><p:pic><p:blipFill><a:blip r:embed="preview"/></p:blipFill></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>',
        'ppt/slides/_rels/slide1.xml.rels': f'<Relationships xmlns="{rel}"><Relationship Id="ole" Type="{r}/oleObject" Target="../embeddings/synthetic.bin"/><Relationship Id="preview" Type="{r}/image" Target="../media/synthetic.emf"/></Relationships>',
        'ppt/embeddings/synthetic.bin': b'Synthetic inventory fixture; no native graph data.',
        'ppt/media/synthetic.emf': b'Synthetic cached preview placeholder; inspection only.',
    }
    with zipfile.ZipFile(path, 'w') as zipped:
        for name, value in parts.items():
            zipped.writestr(name, value)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('archive', type=Path)
    parser.add_argument('--work-parent', type=Path)
    args = parser.parse_args()
    work = Path(tempfile.mkdtemp(prefix='portable check 中文 ', dir=args.work_parent))
    with zipfile.ZipFile(args.archive) as zipped:
        for name in zipped.namelist():
            if not name.startswith('OriginTifTool/') or '..' in Path(name).parts:
                raise RuntimeError('Unexpected archive path.')
        if zipped.testzip():
            raise RuntimeError('Archive CRC check failed.')
        zipped.extractall(work)
    bundle = work / 'OriginTifTool'
    inventory = json.loads((bundle / 'LICENSES/BUILD_INVENTORY.json').read_text(encoding='utf-8'))
    assert (bundle / 'LICENSE').read_text().startswith('MIT License')
    traced = inventory['native_binaries']
    expected = {row['path'] for row in traced}
    actual = {p.relative_to(bundle).as_posix() for p in bundle.rglob('*') if p.suffix.lower() in ('.dll', '.pyd')}
    assert actual == expected, 'Native inventory does not match bundle.'
    for row in traced:
        assert sha(bundle / row['path']) == row['sha256'], row['path']
        component = inventory['components'][row['component']]
        assert component['notice_files'] and all((bundle / n).is_file() for n in component['notice_files'])
    for source in inventory['corresponding_sources'].values():
        assert sha(bundle / source['archive']) == source['sha256']
        assert (bundle / source['recipe']).is_dir()
    assert not any(p.suffix.lower() in ('.pptx', '.ppt', '.oggu', '.ogg') for p in bundle.rglob('*'))
    # The child executable must not find Python, conda or build-time DLLs on PATH.
    env = os.environ.copy()
    for key in list(env):
        if key.upper().startswith(('PYTHON', 'CONDA', '_CE_')):
            env.pop(key)
    windows = Path(env.get('SystemRoot') or env['SYSTEMROOT'])
    env['PATH'] = os.pathsep.join(str(p) for p in (windows / 'System32', windows, windows / 'System32/Wbem', windows / 'System32/WindowsPowerShell/v1.0'))
    exe = bundle / 'OriginTifTool.exe'
    flags = subprocess.CREATE_NO_WINDOW
    subprocess.run([str(exe), '--smoke-test'], cwd=work, env=env, check=True, timeout=90, creationflags=flags)
    source, events = work / 'synthetic.pptx', work / 'events.jsonl'
    synthetic_presentation(source)
    checksum = sha(source)
    subprocess.run([str(exe), '--engine', str(source), '--dry-run', '--work-dir', str(work), '--event-file', str(events)],
                   cwd=work, env=env, check=True, timeout=90, creationflags=flags)
    records = [json.loads(line) for line in events.read_text(encoding='utf-8').splitlines()]
    assert sha(source) == checksum
    assert any(row['event'] == 'inventory' and len(row['rows']) == 1 for row in records)
    result = {'archive_sha256': sha(args.archive), 'native_binaries': len(traced),
              'license_inventory': 'passed', 'corresponding_sources': 'passed',
              'gui_smoke_without_python_path': 'passed', 'synthetic_scan_without_origin': 'passed',
              'unicode_and_space_path': 'passed', 'independent_machine_test': 'not performed by this script'}
    result['origin_registered'] = registered('Origin.Application')
    result['powerpoint_registered'] = registered('PowerPoint.Application')
    (work / 'summary.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result, indent=2))
    print('Private check directory:', str(work).encode('ascii', 'backslashreplace').decode('ascii'))


if __name__ == '__main__':
    main()
