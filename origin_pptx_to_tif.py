#!/usr/bin/env python3
"""Convert embedded Origin graphs in .pptx files using Origin's expGraph.

Windows + licensed Origin + Python 3.11+. Dependencies: pywin32 lxml Pillow numpy.
PowerPoint is used for before/after previews unless --no-render is specified.
Never overwrites an input or an existing output. See README.md for limitations.
"""
from __future__ import annotations

import argparse
from copy import copy, deepcopy
from datetime import datetime
import hashlib
import io
import json
import math
import os
from pathlib import Path
import posixpath
import shutil
import subprocess
import sys
import time
import traceback
import uuid
import zipfile

from lxml import etree as E
import numpy as np
from PIL import Image, ImageFilter

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
EMU = 914400
XML_PARSER = E.XMLParser(resolve_entities=False, no_network=True)
VERSION = '1.1.0'
EVENT_FILE = None


def emit(event, message='', **data):
    record = {'event': event, 'message': message, **data}
    if EVENT_FILE:
        with Path(EVENT_FILE).open('a', encoding='utf-8') as f:
            f.write(json.dumps(record, ensure_ascii=False) + '\n')
    if sys.stdout is not None:
        print(message or event, flush=True)


class Cancelled(RuntimeError):
    pass


def check_cancel(job):
    if job.get('cancel_file') and Path(job['cancel_file']).exists():
        raise Cancelled('Cancelled at a safe boundary; the original file is unchanged.')


def engine_command(*args):
    if getattr(sys, 'frozen', False):
        return [sys.executable, '--engine', *args]
    return [sys.executable, str(Path(__file__).resolve()), *args]


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def xml(data):
    return E.fromstring(data, XML_PARSER)


def encode(root):
    return E.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def dump(path, value):
    path = Path(path)
    tmp = path.with_suffix(path.suffix + '.tmp')
    tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(tmp, path)


def sha(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda: f.read(8 * 1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def relpart(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def resolve(part, target):
    from urllib.parse import unquote
    target = unquote(target.split('#')[0])
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def relationships(z, part):
    name = relpart(part)
    return xml(z.read(name)) if name in z.namelist() else E.Element('{' + REL + '}Relationships')


def origin_frames(root):
    found = []
    for gf in root.findall('.//p:graphicFrame', NS):
        oles = gf.findall('.//p:oleObj', NS)
        origins = [o for o in oles if o.get('progId', '').lower().startswith('origin')]
        if origins:
            identities = {(o.get('{' + NS['r'] + '}id'), o.get('progId')) for o in origins}
            require(len(identities) == 1, 'Conflicting Origin objects in AlternateContent.')
            found.append((gf, origins[0]))
    all_oles = [o for o in root.findall('.//p:oleObj', NS) if o.get('progId', '').lower().startswith('origin')]
    require(sum(len(g.findall('.//p:oleObj', NS)) for g, _ in found) >= len(all_oles),
            'Unsupported Origin OLE placement outside a graphicFrame.')
    return found


def frame_id(gf):
    cn = gf.find('p:nvGraphicFramePr/p:cNvPr', NS)
    require(cn is not None, 'Origin frame has no shape ID.')
    return cn.get('id')


def matrix(xf, group=False):
    """Map child/local coordinates into parent coordinates, including rotations."""
    off, ext = xf.find('a:off', NS), xf.find('a:ext', NS)
    require(off is not None and ext is not None, 'Missing shape transform.')
    x, y = float(off.get('x')), float(off.get('y'))
    w, h = float(ext.get('cx')), float(ext.get('cy'))
    require(w > 0 and h > 0, 'Zero/negative graph or group extent.')
    angle = math.radians(float(xf.get('rot', 0)) / 60000)
    rot = np.array([[math.cos(angle), -math.sin(angle)], [math.sin(angle), math.cos(angle)]])
    flip = np.diag([-1 if xf.get('flipH') in ('1', 'true') else 1,
                    -1 if xf.get('flipV') in ('1', 'true') else 1])
    if group:
        co, ce = xf.find('a:chOff', NS), xf.find('a:chExt', NS)
        require(co is not None and ce is not None, 'Group child transform is missing.')
        cw, ch = float(ce.get('cx')), float(ce.get('cy'))
        require(cw > 0 and ch > 0, 'Invalid group scale.')
        center = np.array([float(co.get('x')) + cw / 2, float(co.get('y')) + ch / 2])
        linear = rot @ flip @ np.diag([w / cw, h / ch])
    else:
        center = np.array([.5, .5])
        linear = rot @ flip @ np.diag([w, h])
    result = np.eye(3)
    result[:2, :2] = linear
    result[:2, 2] = np.array([x + w / 2, y + h / 2]) - linear @ center
    return result


def geometry(gf):
    xf = gf.find('p:xfrm', NS)
    require(xf is not None, 'Origin frame transform is missing.')
    mat = matrix(xf)
    depth = 0
    for ancestor in gf.iterancestors('{' + NS['p'] + '}grpSp'):
        gx = ancestor.find('p:grpSpPr/a:xfrm', NS)
        require(gx is not None, 'Group transform is missing.')
        mat = matrix(gx, True) @ mat
        depth += 1
    points = (mat @ np.array([[0, 1, 1, 0], [0, 0, 1, 1], [1, 1, 1, 1]]))[:2]
    return {'display_inches': [float(np.linalg.norm(mat[:2, j]) / EMU) for j in (0, 1)],
            'bbox_emu': [float(points[0].min()), float(points[1].min()), float(points[0].max()), float(points[1].max())],
            'group_depth': depth}


def inspect(source, work, extract=True, cancel_file=None):
    rows, slides = [], []
    with zipfile.ZipFile(source) as z:
        require(len(z.namelist()) == len(set(z.namelist())), 'Duplicate ZIP part names.')
        require(not any(n.startswith('_xmlsignatures/') for n in z.namelist()), 'Digitally signed presentations are not supported.')
        pres = xml(z.read('ppt/presentation.xml'))
        ids = pres.find('p:sldIdLst', NS)
        require(ids is not None, 'Unsupported or empty PowerPoint presentation format.')
        mapping = {r.get('Id'): r for r in relationships(z, 'ppt/presentation.xml')}
        for number, item in enumerate(ids, 1):
            check_cancel({'cancel_file': cancel_file})
            link = mapping[item.get('{' + NS['r'] + '}id')]
            part = resolve('ppt/presentation.xml', link.get('Target'))
            slides.append(part)
            root = xml(z.read(part)); rels = {r.get('Id'): r for r in relationships(z, part)}
            seen = set()
            for gf, ole in origin_frames(root):
                ident = frame_id(gf)
                require(ident not in seen, f'Slide {number}: duplicate Origin frame IDs require manual review.')
                seen.add(ident)
                rid = ole.get('{' + NS['r'] + '}id')
                require(rid in rels, f'Slide {number}, object {ident}: missing OLE relationship.')
                rel = rels[rid]
                require(rel.get('TargetMode') != 'External' and ole.find('p:embed', NS) is not None,
                        f'Slide {number}, object {ident}: linked Origin objects are not supported.')
                embedded = resolve(part, rel.get('Target'))
                pic = gf.find('.//p:pic', NS)
                require(pic is not None, f'Slide {number}, object {ident}: no cached preview for verification.')
                require(pic.find('p:blipFill/a:srcRect', NS) is None, 'Cropped OLE preview requires manual review.')
                blip = pic.find('p:blipFill/a:blip', NS)
                require(blip is not None and not len(blip), 'OLE preview image effects require manual review.')
                preview_rid = blip.get('{' + NS['r'] + '}embed')
                require(preview_rid in rels, 'Cached preview relationship missing.')
                preview_part = resolve(part, rels[preview_rid].get('Target'))
                nw, nh = int(ole.get('imgW', 0)), int(ole.get('imgH', 0))
                require(nw > 0 and nh > 0, 'Origin native image dimensions are missing.')
                name = f'slide_{number:03d}_origin_{len(rows)+1:03d}'
                row = {'slide': number, 'part': part, 'id': ident, 'rid': rid, 'embedded_part': embedded,
                       'prog_id': ole.get('progId'), 'name': name, 'native_size': [nw, nh],
                       'reference': str(work / (name + '_reference.png')), 'tif': str(work / (name + '.tif')),
                       **geometry(gf)}
                if extract:
                    import pythoncom
                    import win32timezone  # pywin32 loads this dynamically for IStorage.Stat timestamps.
                    binary = work / (name + '.bin'); binary.write_bytes(z.read(embedded))
                    st = pythoncom.StgOpenStorage(str(binary), None, 0x20, None, 0)
                    stream = st.OpenStream('Contents', None, 0x10, 0)
                    data = stream.Read(stream.Stat()[2])
                    require(data.startswith((b'CPYU', b'CPYA')), f'Unsupported Origin native stream: {name}')
                    native = work / (name + ('.oggu' if data.startswith(b'CPYU') else '.ogg'))
                    native.write_bytes(data)
                    row['native'] = str(native)
                    preview = work / (name + '_preview' + Path(preview_part).suffix)
                    preview.write_bytes(z.read(preview_part)); row['preview'] = str(preview)
                    stream = st = None
                rows.append(row)
    return rows, slides


RENDER_REFERENCE_PS = r'''param([string]$Job)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$taskRows = Get-Content -LiteralPath $Job -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($taskRow in $taskRows) {
    $taskImage = [System.Drawing.Image]::FromFile($taskRow.preview)
    $taskScale = 1200.0 / [Math]::Max($taskRow.native_size[0], $taskRow.native_size[1])
    $taskW = [Math]::Max(32, [int][Math]::Round($taskRow.native_size[0] * $taskScale))
    $taskH = [Math]::Max(32, [int][Math]::Round($taskRow.native_size[1] * $taskScale))
    $taskBitmap = [System.Drawing.Bitmap]::new($taskW, $taskH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $taskGraphics = [System.Drawing.Graphics]::FromImage($taskBitmap)
    try {
        $taskGraphics.Clear([System.Drawing.Color]::Transparent)
        $taskGraphics.DrawImage($taskImage, [System.Drawing.Rectangle]::new(0,0,$taskW,$taskH))
        $taskBitmap.Save($taskRow.reference, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally { $taskGraphics.Dispose(); $taskBitmap.Dispose(); $taskImage.Dispose() }
}
'''


def render_references(rows, work, timeout):
    job, script = work / 'references.json', work / 'render_references.ps1'
    dump(job, rows); script.write_text(RENDER_REFERENCE_PS, encoding='utf-8-sig')
    # Fixed local script and an argument list; document text never becomes shell code.
    ps = shutil.which('powershell.exe')
    require(ps, 'Windows PowerShell is required to render EMF/WMF previews.')
    result = subprocess.run([ps, '-NoProfile', '-NonInteractive', '-File', str(script), '-Job', str(job)],
                            capture_output=True, timeout=timeout, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    require(result.returncode == 0, 'Reference rendering failed: ' + result.stderr.decode(errors='replace')[-1500:])


def white_image(path):
    with Image.open(path) as im:
        rgba = im.convert('RGBA')
        return Image.alpha_composite(Image.new('RGBA', rgba.size, 'white'), rgba).convert('RGB')


def content_bbox(im):
    pixels = np.asarray(im)
    ys, xs = np.where(pixels.min(axis=2) < 220)
    require(len(xs) > 10, 'Blank/near-blank graph cannot be matched safely.')
    return [int(xs.min()), int(ys.min()), int(xs.max()+1), int(ys.max()+1)]


def fit_padding(reference, candidate):
    a, b = content_bbox(reference), content_bbox(candidate)
    u = [a[0]/reference.width, a[1]/reference.height, (a[2]-a[0])/reference.width, (a[3]-a[1])/reference.height]
    v = [b[0]/candidate.width, b[1]/candidate.height, (b[2]-b[0])/candidate.width, (b[3]-b[1])/candidate.height]
    sx, sy = u[2]/v[2], u[3]/v[3]
    tx, ty = u[0]-sx*v[0], u[1]-sy*v[1]
    return {'l': round(-tx/sx*100000), 't': round(-ty/sy*100000),
            'r': round((1-(1-tx)/sx)*100000), 'b': round((1-(1-ty)/sy)*100000)}


def crop_transform(crop):
    l, t, r, b = (crop.get(k, 0)/100000 for k in ('l', 't', 'r', 'b'))
    require(1-l-r > 0 and 1-t-b > 0, 'Invalid image bounds.')
    sx, sy = 1/(1-l-r), 1/(1-t-b)
    return sx, sy, -l*sx, -t*sy


def normalized(im, crop):
    sx, sy, tx, ty = crop_transform(crop)
    return im.transform((512, 512), Image.Transform.AFFINE,
                        (im.width/(512*sx), 0, -tx*im.width/sx,
                         0, im.height/(512*sy), -ty*im.height/sy),
                        Image.Resampling.BICUBIC, fillcolor='white')


def match_score(reference, candidate, crop):
    a, b = normalized(reference, {}), normalized(candidate, crop)
    aa = np.asarray(a.filter(ImageFilter.GaussianBlur(1.5)), dtype=float)
    bb = np.asarray(b.filter(ImageFilter.GaussianBlur(1.5)), dtype=float)
    area = (aa.min(axis=2) < 245) | (bb.min(axis=2) < 245)
    require(area.any(), 'Blank graph preview.')
    similarity = float(1 - np.abs(aa-bb).mean(axis=2)[area].mean()/255)
    ink = lambda im: np.asarray(im.convert('L').point(lambda p: 255 if p < 210 else 0).filter(ImageFilter.MaxFilter(5))) > 0
    am, bm = ink(a), ink(b)
    overlap = float((am & bm).sum()/max(1, (am | bm).sum()))
    return {'similarity': similarity, 'ink_overlap': overlap}


def background_policy(row, mode):
    if mode != 'auto':
        return mode == 'white-transparent', []
    with Image.open(row['reference']) as im:
        px = np.asarray(im.convert('RGBA'))
    transparent = px[:, :, 3] < 128
    opaque_white = (px[:, :, :3].min(axis=2) > 245) & ~transparent
    if transparent.mean() > .05:
        require(opaque_white.mean() < .02,
                f"{row['name']}: mixed opaque white and transparency cannot be preserved safely with white color-keying. Review manually.")
        warnings = ['Small opaque white areas also become transparent in PowerPoint; inspect the comparison.'] if opaque_white.mean() > .002 else []
        return True, warnings
    return False, []


def progress(job, **fields):
    dump(job['progress'], {'time': time.time(), **fields})


def labtalk_path(path):
    value = str(path)
    require(not any(c in value for c in ('"', '\n', '\r', '%', '$')), 'Path contains characters with LabTalk substitution meaning. Use --work-dir with a simple path.')
    return value


def export_one(app, job, row):
    require(app.Execute('doc -s; doc -n;'), 'Cannot create temporary Origin project.')
    require(app.Load(row['native']), f"Origin cannot load {row['name']}")
    reference = white_image(row['reference'])
    transparent, warnings = background_policy(row, job['background'])
    w, h = row['display_inches']; ratio = row['native_size'][0]/row['native_size'][1]
    width = max(w, h*ratio)*1.03
    selected = None
    attempts = []
    for margin in (2, 0):
        for attempt in range(3):
            check_cancel(job)
            require((width*job['dpi'])**2/max(.01, ratio) <= job['max_pixels'], 'Graph export exceeds the configured pixel limit.')
            name = row['name'] + f'_m{margin}_{attempt}'
            path = Path(row['tif']).with_name(name + '.tif')
            command = (f'expGraph type:=tif export:=page filename:="{name}" path:="{labtalk_path(path.parent)}" '
                       f'overwrite:=replace theme:=<Original> sysopts:=0 tr.Margin:={margin} tr.BorderWid:=0 '
                       f'tr1.unit:=0 tr1.width:={width:.10f} tr2.TIF.dotsperinch:={job["dpi"]} '
                       'tr2.TIF.ColorSpace:="RGB" tr2.TIF.bitsperpixel:="24-bit Color" tr2.TIF.Compression:="LZW";')
            require(app.Execute(command) and path.is_file(), f"Origin expGraph failed: {row['name']}")
            with Image.open(path) as im:
                dpi = [float(v) for v in im.info.get('dpi', ())]
                require(im.format == 'TIFF' and len(dpi) == 2 and all(abs(v-job['dpi']) < .01 for v in dpi), 'TIFF DPI verification failed.')
                require(im.width*im.height <= job['max_pixels'], 'Exported TIFF exceeds the pixel limit.')
                candidate = im.convert('RGB')
            crop = fit_padding(reference, candidate) if margin == 0 else {}
            sx, sy, _, _ = crop_transform(crop)
            effective = [candidate.width/w/sx, candidate.height/h/sy]
            if min(effective) < job['dpi']:
                width *= job['dpi']/min(effective)*1.04
                continue
            score = match_score(reference, candidate, crop)
            details = {'command': command, 'margin': margin, 'pixels': list(candidate.size), 'dpi': dpi,
                       'effective_dpi': effective, 'srcRect': crop, 'white_transparent': transparent,
                       'score': score, 'warnings': warnings, 'candidate_path': str(path)}
            attempts.append(details)
            if score['ink_overlap'] >= job['min_overlap'] and score['similarity'] >= job['min_similarity']:
                shutil.copy2(path, row['tif']); selected = details
            break
        if selected is not None:
            break
    # Also retain a comparison for failed matches; the candidate is never used to replace that OLE.
    best = selected or (max(attempts, key=lambda a: a['score']['ink_overlap']) if attempts else None)
    if best:
        comparison = Image.new('RGB', (1024, 512), 'white')
        comparison.paste(normalized(reference, {}), (0, 0))
        comparison.paste(normalized(white_image(best['candidate_path']), best['srcRect']), (512, 0))
        comparison.save(Path(row['tif']).with_name(row['name'] + '_comparison.png'))
    require(selected is not None,
            f"{row['name']}: appearance mismatch or clipping; graph not converted. Attempts: {json.dumps(attempts)}")
    if selected['margin'] == 0 or selected['score']['ink_overlap'] < .75:
        selected['warnings'].append('Export bounds were adjusted or matching is borderline; inspect this graph comparison.')
    return {**row, **selected}


def export_worker(job):
    import win32com.client
    app = None; results = []; skipped = []
    skipped_file = job.get('skipped_results', str(Path(job['results']).with_name('skipped.json')))
    dump(job['results'], results); dump(skipped_file, skipped)
    try:
        progress(job, stage='starting_origin', skipped=skipped)
        app = win32com.client.DispatchEx('Origin.Application')
        require(app.Execute('sec -poc 30;'), 'Origin startup failed.')
        for index, row in enumerate(job['rows'], 1):
            check_cancel(job)
            progress(job, stage='export', index=index, total=len(job['rows']), name=row['name'], skipped=skipped)
            try:
                result = export_one(app, job, row)
            except Cancelled:
                raise
            except Exception as error:
                if not job.get('skip_failed', False):
                    raise
                check_cancel(job)
                skipped.append({**row, 'status': 'skipped', 'preserved_as': 'original_origin_ole', 'error': str(error)})
                dump(skipped_file, skipped)
                progress(job, stage='skipped', index=index, total=len(job['rows']), name=row['name'], skipped=skipped)
                continue
            results.append(result)
            dump(job['results'], results)
            progress(job, stage='exported', index=index, total=len(job['rows']), name=row['name'], skipped=skipped)
    finally:
        if app is not None:
            progress(job, stage='closing_origin', skipped=skipped)
            app.Exit()


def render_worker(job):
    import win32com.client
    app = win32com.client.DispatchEx('PowerPoint.Application')
    pres = None
    try:
        for which, path in (('before', job['source']), ('after', job['candidate'])):
            progress(job, stage='opening_powerpoint', name=which)
            pres = app.Presentations.Open(path, True, False, False)
            require(pres.Slides.Count == job['slide_count'], 'PowerPoint slide count changed.')
            width = 1600
            height = round(width*pres.PageSetup.SlideHeight/pres.PageSetup.SlideWidth)
            for n in job['slide_numbers']:
                check_cancel(job)
                progress(job, stage='render', name=f'{which} slide {n}')
                pres.Slides(n).Export(str(Path(job['work']) / f'{which}_slide_{n:03d}.png'), 'PNG', width, height)
            pres.Close(); pres = None
    finally:
        if pres is not None: pres.Close()
        # PowerPoint may share its application instance; never quit another open deck.
        if app.Presentations.Count == 0: app.Quit()


def run_worker(kind, job, work, timeout):
    job['progress'] = str(work / (kind + '_progress.json'))
    jobfile = work / (kind + '_job.json'); dump(jobfile, job)
    log = work / (kind + '.log')
    last_state = None; last_change = time.monotonic()
    with log.open('wb') as output:
        env = dict(os.environ, PYTHONUTF8='1')
        process = subprocess.Popen(engine_command('--_worker', kind, str(jobfile)),
                                   stdout=output, stderr=subprocess.STDOUT, env=env,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        try:
            while process.poll() is None:
                try:
                    state = json.loads(Path(job['progress']).read_text(encoding='utf-8'))
                    if state != last_state:
                        last_state = state; last_change = time.monotonic()
                        emit('progress', f"  {state.get('stage')} {state.get('index','')}/{state.get('total','')} {state.get('name','')}", **state)
                except (FileNotFoundError, json.JSONDecodeError):
                    pass
                if time.monotonic()-last_change > timeout:
                    raise TimeoutError(f'{kind} timed out after {timeout}s. Worker stopped; its temporary application may need closing. Existing user applications were not force-terminated.')
                time.sleep(.25)
        except BaseException:
            if process.poll() is None:
                process.terminate()
                try: process.wait(5)
                except subprocess.TimeoutExpired: process.kill(); process.wait()
            raise
    check_cancel(job)
    detail = log.read_text(encoding='utf-8', errors='replace')
    error_file = Path(job['progress']).with_suffix('.error.txt')
    if error_file.exists():
        detail += '\n' + error_file.read_text(encoding='utf-8', errors='replace')
    require(process.returncode == 0, f'{kind} failed. See {log}\n' + detail[-4000:])


def signature(e):
    return e.tag, tuple(sorted(e.attrib.items())), e.text, e.tail, tuple(signature(c) for c in e)


def indexpath(e):
    path = []
    while e.getparent() is not None:
        p = e.getparent(); path.append((p.tag, p.index(e))); e = p
    return tuple(reversed(path))


def find_pic(root, ident):
    found = [p for p in root.findall('.//p:pic', NS)
             if p.find('p:nvPicPr/p:cNvPr', NS) is not None and p.find('p:nvPicPr/p:cNvPr', NS).get('id') == ident]
    require(len(found) == 1, f'Replacement image ID {ident} is not unique.')
    return found[0]


def make_picture(gf, row, rid):
    pic = deepcopy(gf.find('.//p:pic', NS))
    nv = pic.find('p:nvPicPr', NS)
    nv.replace(nv.find('p:cNvPr', NS), deepcopy(gf.find('p:nvGraphicFramePr/p:cNvPr', NS)))
    if gf.find('p:nvGraphicFramePr/p:nvPr', NS) is not None:
        nv.replace(nv.find('p:nvPr', NS), deepcopy(gf.find('p:nvGraphicFramePr/p:nvPr', NS)))
    sp = pic.find('p:spPr', NS)
    xf = deepcopy(gf.find('p:xfrm', NS)); xf.tag = '{' + NS['a'] + '}xfrm'
    prev = sp.find('a:xfrm', NS)
    if prev is not None: sp.replace(prev, xf)
    else: sp.insert(0, xf)
    fill = pic.find('p:blipFill', NS); blip = fill.find('a:blip', NS)
    blip.set('{' + NS['r'] + '}embed', rid)
    if row['white_transparent']:
        cc = E.SubElement(blip, '{' + NS['a'] + '}clrChange', useA='1')
        for tag, alpha in (('clrFrom', None), ('clrTo', '0')):
            color = E.SubElement(E.SubElement(cc, '{' + NS['a'] + '}' + tag), '{' + NS['a'] + '}srgbClr', val='FFFFFF')
            if alpha is not None: E.SubElement(color, '{' + NS['a'] + '}alpha', val=alpha)
    if row['srcRect']:
        crop = E.Element('{' + NS['a'] + '}srcRect', **{k: str(v) for k, v in row['srcRect'].items()})
        fill.insert(list(fill).index(blip)+1, crop)
    return pic


def build(source, candidate, results):
    changed, added, possible_orphans = {}, {}, set()
    with zipfile.ZipFile(source) as z:
        for part in sorted({r['part'] for r in results}):
            root = xml(z.read(part)); rels = relationships(z, part)
            byid = {r.get('Id'): r for r in rels}
            frames = {frame_id(g): g for g, _ in origin_frames(root)}
            selected = [r for r in results if r['part'] == part]
            for row in selected:
                gf = frames[row['id']]
                rid = 'rIdOriginTif_' + row['id']
                while rid in byid: rid += '_'
                media = 'ppt/media/codex_origin_' + row['name'] + '.tif'
                while media in z.namelist() or media in added: media = media[:-4] + '_.tif'
                added[media] = Path(row['tif']); row['media_part'] = media
                pic = make_picture(gf, row, rid); gf.getparent().replace(gf, pic)
                E.SubElement(rels, '{' + REL + '}Relationship', Id=rid, Type=NS['r']+'/image', Target=posixpath.relpath(media, posixpath.dirname(part)))
                byid[rid] = rels[-1]
                possible_orphans.add(row['embedded_part'])
            used = {value for node in root.iter() for key, value in node.attrib.items() if key.startswith('{' + NS['r'] + '}')}
            for row in selected:
                if row['rid'] not in used and row['rid'] in byid:
                    oldrel = byid.pop(row['rid']); rels.remove(oldrel)
            changed[part] = encode(root); changed[relpart(part)] = encode(rels)
        # Keep shared objects and all non-Origin OLE payloads, e.g. Excel equations.
        referenced = set()
        for name in z.namelist():
            if name.endswith('.rels'):
                rr = xml(changed.get(name, z.read(name)))
                part = '' if name == '_rels/.rels' else posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
                referenced.update(resolve(part, r.get('Target')) for r in rr if r.get('TargetMode') != 'External')
        removed = possible_orphans - referenced
        ct = xml(z.read('[Content_Types].xml'))
        for child in list(ct):
            if child.get('PartName', '').lstrip('/') in removed: ct.remove(child)
        tif = [c for c in ct if c.get('Extension', '').lower() == 'tif']
        if not tif: E.SubElement(ct, '{' + CT + '}Default', Extension='tif', ContentType='image/tiff')
        else: require(all(c.get('ContentType') == 'image/tiff' for c in tif), 'Conflicting TIFF content type.')
        if signature(ct) != signature(xml(z.read('[Content_Types].xml'))): changed['[Content_Types].xml'] = encode(ct)
        with zipfile.ZipFile(candidate, 'w', allowZip64=True) as out:
            out.comment = z.comment
            for info in z.infolist():
                if info.filename in removed: continue
                if info.filename in changed:
                    out.writestr(info.filename, changed[info.filename], compress_type=zipfile.ZIP_DEFLATED)
                else:
                    ni = copy(info)
                    with z.open(info) as src, out.open(ni, 'w', force_zip64=info.file_size > 2**31) as dst:
                        shutil.copyfileobj(src, dst, 8*1024*1024)
            for name, path in added.items(): out.write(path, name, compress_type=zipfile.ZIP_DEFLATED)
    return set(changed), removed


def verify(source, candidate, rows, slides, changed, removed, skipped=()):
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(candidate) as z:
        require(z.testzip() is None, 'ZIP CRC validation failed.')
        require(len(z.namelist()) == len(set(z.namelist())), 'Duplicate output parts.')
        for info in a.infolist():
            if info.filename not in changed and info.filename not in removed:
                ni = z.getinfo(info.filename)
                require((info.CRC, info.file_size) == (ni.CRC, ni.file_size), 'Unrelated package part changed: '+info.filename)
        for part in slides:
            old, new = xml(a.read(part)), xml(z.read(part))
            selected = [r for r in rows if r['part'] == part]
            retained = {r['id'] for r in skipped if r['part'] == part}
            frames = {frame_id(g): g for g, _ in origin_frames(old)}
            require(set(frames) == {r['id'] for r in selected} | retained, 'An Origin object was neither converted nor explicitly skipped.')
            require({frame_id(g) for g, _ in origin_frames(new)} == retained, 'Unexpected remaining or missing Origin OLE object.')
            for row in selected:
                gf, pic = frames[row['id']], find_pic(new, row['id'])
                require(indexpath(gf) == indexpath(pic), 'Group or stacking order changed.')
                ox, nx = gf.find('p:xfrm', NS), pic.find('p:spPr/a:xfrm', NS)
                require(ox.attrib == nx.attrib and [signature(c) for c in ox] == [signature(c) for c in nx], 'Graph size or position changed.')
                require(hashlib.sha256(z.read(row['media_part'])).hexdigest() == sha(row['tif']), 'TIFF bytes changed during packaging.')
            for row in selected:
                gf, pic = frames[row['id']], find_pic(new, row['id'])
                gf.getparent().remove(gf); pic.getparent().remove(pic)
            require(signature(old) == signature(new), 'Non-target slide content changed: '+part)
        names = set(z.namelist())
        for name in names:
            if name.endswith('.rels'):
                part = '' if name == '_rels/.rels' else posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
                for r in xml(z.read(name)):
                    if r.get('TargetMode') != 'External':
                        require(resolve(part, r.get('Target')) in names, 'Broken internal relationship: '+name)
    return {'zip_crc': True, 'all_origin_replaced': not bool(skipped), 'all_selected_origin_replaced': True,
            'skipped_origin_preserved': True, 'geometry_and_groups_exact': True,
            'non_target_content_unchanged': True, 'relationships_resolve': True, 'tif_bytes_preserved': True}


def process_file(source, args):
    source = source.resolve()
    require(source.suffix.lower() == '.pptx' and source.is_file(), f'Input must be an existing .pptx: {source}')
    output = Path(args.output).resolve() if args.output else source.with_name(source.stem + f'_Origin_TIF_{args.dpi}dpi.pptx')
    require(output.suffix.lower() == '.pptx', 'Output must use .pptx.')
    require(args.dry_run or (source != output and not output.exists()), f'Output already exists or equals input: {output}')
    work = (Path(args.work_dir).resolve() if args.work_dir else source.parent) / ('.origin_tif_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:8])
    work.mkdir(parents=True, exist_ok=False)
    report = {'version': VERSION, 'source': str(source), 'output': str(output), 'work_dir': str(work),
              'status': 'started', 'dpi': args.dpi, 'skip_failed': args.skip_failed,
              'visual_check_note': 'Similarity metrics are heuristics, not proof of scientific-data identity. Inspect before/after previews.'}
    try:
        original_hash = sha(source)
        check_cancel(vars(args))
        rows, slides = inspect(source, work, extract=not args.dry_run, cancel_file=args.cancel_file)
        report.update(objects=len(rows), slides=len(slides), inventory=rows)
        dump(work/'report.json', report)
        emit('inventory', f'{source.name}: {len(slides)} slides, {len(rows)} embedded Origin objects', rows=rows, slides=len(slides), work_dir=str(work))
        if args.dry_run or not rows:
            report['status'] = 'inspected' if args.dry_run else 'no_origin_objects'
            return report
        require(os.name == 'nt', 'Origin conversion requires Windows.')
        check_cancel(vars(args))
        emit('progress', 'Rendering cached previews', stage='references')
        render_references(rows, work, args.timeout)
        job = {'rows': rows, 'results': str(work/'exports.json'), 'dpi': args.dpi, 'background': args.background,
               'min_overlap': args.min_overlap, 'min_similarity': args.min_similarity, 'max_pixels': args.max_megapixels*1_000_000,
               'cancel_file': args.cancel_file, 'skip_failed': args.skip_failed, 'skipped_results': str(work/'skipped.json')}
        run_worker('export', job, work, args.timeout)
        results = json.loads(Path(job['results']).read_text(encoding='utf-8'))
        skipped = json.loads(Path(job['skipped_results']).read_text(encoding='utf-8'))
        require(len(results) + len(skipped) == len(rows) and
                {r['name'] for r in results + skipped} == {r['name'] for r in rows},
                'Not all Origin graphs have a conversion or skip result.')
        require(args.skip_failed or not skipped, 'Skipped objects found in strict mode.')
        report.update(results=results, skipped=skipped, converted_count=len(results), skipped_count=len(skipped))
        emit('skipped_results', skipped=skipped)
        check_cancel(vars(args))
        if not results:
            report['status'] = 'no_success'
            emit('no_success', 'All graphs were skipped; no new PPT was created.', skipped=len(skipped))
            return report
        emit('progress', 'Replacing objects and checking the package', stage='packaging')
        candidate = work/'candidate.pptx'
        changed, removed = build(source, candidate, results)
        report['verification'] = verify(source, candidate, results, slides, changed, removed, skipped)
        report['results'] = results
        emit('export_results', results=results)
        if not args.no_render:
            run_worker('render', {'source': str(source), 'candidate': str(candidate), 'work': str(work),
                                 'slide_count': len(slides), 'slide_numbers': sorted({r['slide'] for r in rows}),
                                 'cancel_file': args.cancel_file}, work, args.timeout)
            report['powerpoint_preview'] = 'passed: before/after affected slides rendered'
        else:
            report['powerpoint_preview'] = 'not performed (--no-render)'
        check_cancel(vars(args))
        require(sha(source) == original_hash, 'Input changed while conversion was running; output not published.')
        output.parent.mkdir(parents=True, exist_ok=True)
        # Atomic no-clobber publication on the same filesystem; original is never modified.
        staging = output.with_name('.' + output.name + '.' + uuid.uuid4().hex + '.tmp')
        try:
            shutil.copyfile(candidate, staging)
            require(sha(staging) == sha(candidate), 'Final copy checksum mismatch.')
            check_cancel(vars(args))
            if os.name == 'nt': os.rename(staging, output)  # Windows rename refuses an existing destination.
            else: os.link(staging, output); staging.unlink()
        finally:
            if staging.exists(): staging.unlink()
        report.update(status='completed_with_skips' if skipped else 'completed', source_sha256=original_hash, output_sha256=sha(output))
        emit('saved', 'SAVED ' + str(output), output=str(output), warnings=sum(bool(r['warnings']) for r in results),
             converted=len(results), skipped=len(skipped))
        return report
    except BaseException as error:
        report.update(status='cancelled' if isinstance(error, Cancelled) else 'failed', error=str(error))
        raise
    finally:
        dump(work/'report.json', report)
        emit('report', 'REPORT ' + str(work/'report.json'), report=str(work/'report.json'), status=report['status'])


def main(argv=None):
    global EVENT_FILE
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == '--_worker':
        job = json.loads(Path(argv[2]).read_text(encoding='utf-8'))
        try:
            (export_worker if argv[1] == 'export' else render_worker)(job)
        except BaseException:
            Path(job['progress']).with_suffix('.error.txt').write_text(traceback.format_exc(), encoding='utf-8')
            raise
        return 0
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('files', nargs='+', type=Path, help='One or more .pptx files; PowerShell wildcards may be expanded explicitly.')
    parser.add_argument('-o', '--output', help='Output filename; only with one input; existing files are never overwritten.')
    parser.add_argument('--dpi', type=int, default=600)
    parser.add_argument('--dry-run', action='store_true', help='Inspect only; do not start Origin or modify PPTX.')
    parser.add_argument('--no-render', action='store_true', help='Skip native PowerPoint previews; per-graph visual and structural checks still run.')
    parser.add_argument('--skip-failed', action='store_true', help='Preserve failed graph OLE objects and continue exporting the other graphs.')
    parser.add_argument('--work-dir', help='Parent directory for unique work folders, extracted objects, TIFFs and reports.')
    parser.add_argument('--timeout', type=int, default=180, help='Timeout for each native application phase (seconds).')
    parser.add_argument('--background', choices=['auto', 'opaque', 'white-transparent'], default='auto')
    parser.add_argument('--min-overlap', type=float, default=.68, help='Foreground matching threshold. Lowering requires manual visual review.')
    parser.add_argument('--min-similarity', type=float, default=.80)
    parser.add_argument('--max-megapixels', type=int, default=100)
    parser.add_argument('--event-file', help=argparse.SUPPRESS)
    parser.add_argument('--cancel-file', help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    EVENT_FILE = args.event_file
    require(not args.output or len(args.files) == 1, '--output requires a single input.')
    require(72 <= args.dpi <= 2400 and args.timeout >= 10 and args.max_megapixels > 0, 'Invalid resolution, timeout or pixel limit.')
    require(0 <= args.min_overlap <= 1 and 0 <= args.min_similarity <= 1, 'Invalid comparison thresholds.')
    failed = False
    for source in args.files:
        try: process_file(source, args)
        except Exception as e:
            failed = True; emit('cancelled' if isinstance(e, Cancelled) else 'error', str(e))
    return 1 if failed else 0


if __name__ == '__main__':
    try: sys.exit(main())
    except Exception:
        traceback.print_exc(); sys.exit(1)
