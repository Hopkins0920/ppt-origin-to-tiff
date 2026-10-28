"""Regression tests for document preservation and conversion failure boundaries."""
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile
import json
from unittest.mock import patch, Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import origin_pptx_to_tif as c
from PIL import Image, ImageDraw


def graph(ident, prog='Origin95.Graph', linked=False):
    return f'''<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{ident}" name="Graph"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm rot="1800000" flipH="1"><a:off x="500" y="600"/><a:ext cx="914400" cy="457200"/></p:xfrm>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
    <p:oleObj r:id="ole{ident}" progId="{prog}" imgW="914400" imgH="457200"><p:{'link' if linked else 'embed'}/>
    <p:pic><p:nvPicPr><p:cNvPr id="999" name="cached"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="preview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>
    </p:oleObj></a:graphicData></a:graphic></p:graphicFrame>'''


class PackageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.work = Path(self.temp.name)
        self.source = self.work/'sample.pptx'

    def tearDown(self):
        self.temp.cleanup()

    def fixture(self, linked=False, second_origin=False):
        ns = ' '.join(f'xmlns:{k}="{v}"' for k,v in c.NS.items())
        group = '''<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>
        <a:xfrm rot="5400000" flipV="1"><a:off x="200" y="300"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/></a:xfrm></p:grpSpPr>'''
        second_type = 'Origin95.Graph' if second_origin else 'Excel.Sheet.12'
        slide = f'<p:sld {ns}><p:cSld><p:spTree>{group}{graph(7, linked=linked)}</p:grpSp>{graph(8,second_type)}</p:spTree></p:cSld></p:sld>'
        external_attr = 'TargetMode="External"' if linked else ''
        rels = f'<Relationships xmlns="{c.REL}"><Relationship Id="ole7" Type="{c.NS["r"]}/oleObject" Target="../embeddings/shared.bin" {external_attr}/><Relationship Id="ole8" Type="{c.NS["r"]}/oleObject" Target="../embeddings/shared.bin"/><Relationship Id="preview" Type="{c.NS["r"]}/image" Target="../media/preview.emf"/></Relationships>'
        parts = {
            '[Content_Types].xml':f'<Types xmlns="{c.CT}"/>',
            '_rels/.rels':f'<Relationships xmlns="{c.REL}"><Relationship Id="p" Type="{c.NS["r"]}/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
            'ppt/presentation.xml':f'<p:presentation {ns}><p:sldIdLst><p:sldId id="256" r:id="late"/><p:sldId id="257" r:id="early"/></p:sldIdLst></p:presentation>',
            'ppt/_rels/presentation.xml.rels':f'<Relationships xmlns="{c.REL}"><Relationship Id="early" Target="slides/slide1.xml"/><Relationship Id="late" Target="slides/slide9.xml"/></Relationships>',
            'ppt/slides/slide9.xml':slide,
            'ppt/slides/slide1.xml':f'<p:sld {ns}><p:cSld><p:spTree/></p:cSld></p:sld>',
            'ppt/slides/_rels/slide9.xml.rels':rels,
            'ppt/embeddings/shared.bin':b'shared payload must survive',
            'ppt/media/preview.emf':b'untouched cached preview',
        }
        with zipfile.ZipFile(self.source,'w') as z:
            for name,data in parts.items():z.writestr(name,data)

    def test_display_order_and_group_scale(self):
        self.fixture()
        rows,slides=c.inspect(self.source,self.work,extract=False)
        self.assertEqual(slides[0],'ppt/slides/slide9.xml')
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]['slide'],1)
        self.assertEqual(rows[0]['group_depth'],1)
        for actual,expected in zip(rows[0]['display_inches'],[2,1]):self.assertAlmostEqual(actual,expected)

    def test_shared_embedding_non_origin_and_exact_geometry_survive(self):
        self.fixture()
        rows,slides=c.inspect(self.source,self.work,extract=False)
        tif=self.work/'native.tif';Image.new('RGB',(200,100),'white').save(tif,dpi=(600,600))
        rows[0].update(tif=str(tif),white_transparent=True,srcRect={'l':-500,'t':0,'r':0,'b':-900})
        candidate=self.work/'result.pptx'
        changed,removed=c.build(self.source,candidate,rows)
        self.assertNotIn('ppt/embeddings/shared.bin',removed)
        self.assertTrue(all(c.verify(self.source,candidate,rows,slides,changed,removed).values()))
        with zipfile.ZipFile(candidate) as z:
            root=c.xml(z.read(slides[0]))
            pic=c.find_pic(root,'7')
            self.assertIsNotNone(pic.find('.//a:clrChange',c.NS))
            self.assertEqual(z.read('ppt/embeddings/shared.bin'),b'shared payload must survive')
            self.assertEqual(root.find('.//p:oleObj',c.NS).get('progId'),'Excel.Sheet.12')

    def test_external_link_is_rejected(self):
        self.fixture(linked=True)
        with self.assertRaisesRegex(RuntimeError,'linked Origin'):
            c.inspect(self.source,self.work,extract=False)

    def test_skipped_origin_and_shared_payload_are_preserved(self):
        self.fixture(second_origin=True)
        rows,slides=c.inspect(self.source,self.work,extract=False)
        success,skipped=rows[:1],rows[1:]
        tif=self.work/'native.tif';Image.new('RGB',(200,100),'white').save(tif,dpi=(600,600))
        success[0].update(tif=str(tif),white_transparent=False,srcRect={})
        candidate=self.work/'partial.pptx'
        changed,removed=c.build(self.source,candidate,success)
        checks=c.verify(self.source,candidate,success,slides,changed,removed,skipped)
        self.assertFalse(checks['all_origin_replaced'])
        self.assertTrue(checks['skipped_origin_preserved'])
        with zipfile.ZipFile(self.source) as a,zipfile.ZipFile(candidate) as b:
            before=c.xml(a.read(slides[0]));after=c.xml(b.read(slides[0]))
            old=next(g for g,_ in c.origin_frames(before) if c.frame_id(g)=='8')
            new=next(g for g,_ in c.origin_frames(after) if c.frame_id(g)=='8')
            self.assertEqual(c.signature(old),c.signature(new))
            self.assertEqual(a.read('ppt/embeddings/shared.bin'),b.read('ppt/embeddings/shared.bin'))
            new.find('p:xfrm/a:off',c.NS).set('x','999999')
            tampered=self.work/'tampered.pptx'
            with zipfile.ZipFile(tampered,'w') as out:
                for info in b.infolist():
                    out.writestr(info,b.read(info.filename) if info.filename!=slides[0] else c.encode(after))
        with self.assertRaisesRegex(RuntimeError,'Non-target slide content changed'):
            c.verify(self.source,tampered,success,slides,changed,removed,skipped)

    def export_job(self, skip_failed=True):
        return {'rows':[{'name':'one','slide':1,'id':'1'},{'name':'two','slide':1,'id':'2'},{'name':'three','slide':2,'id':'3'}],
                'progress':str(self.work/'progress.json'),'results':str(self.work/'exports.json'),
                'skipped_results':str(self.work/'skipped.json'),'skip_failed':skip_failed}

    def test_worker_continues_past_failed_graph(self):
        job=self.export_job();app=Mock();app.Execute.return_value=True
        with patch('win32com.client.DispatchEx',return_value=app),patch.object(c,'export_one',side_effect=[{'name':'one'},RuntimeError('bad match'),{'name':'three'}]) as export:
            c.export_worker(job)
            self.assertEqual(export.call_count,3)
        self.assertEqual([r['name'] for r in json.loads(Path(job['results']).read_text())],['one','three'])
        skipped=json.loads(Path(job['skipped_results']).read_text())
        self.assertEqual([(r['name'],r['preserved_as']) for r in skipped],[('two','original_origin_ole')])
        app.Exit.assert_called_once()

    def test_strict_mode_and_cancel_never_continue(self):
        for skip,error in [(False,RuntimeError('bad match')),(True,c.Cancelled('cancel'))]:
            with self.subTest(skip=skip,error=type(error).__name__):
                app=Mock();app.Execute.return_value=True;job=self.export_job(skip)
                with patch('win32com.client.DispatchEx',return_value=app),patch.object(c,'export_one',side_effect=[{'name':'one'},error,{'name':'three'}]) as export:
                    with self.assertRaises(type(error)):c.export_worker(job)
                    self.assertEqual(export.call_count,2)
                self.assertEqual(json.loads(Path(job['skipped_results']).read_text()),[])

    def test_all_skipped_does_not_publish_a_presentation(self):
        self.fixture();rows,slides=c.inspect(self.source,self.work,extract=False)
        output=self.work/'must_not_exist.pptx'
        def worker(kind,job,work,timeout):
            self.assertEqual(kind,'export')
            c.dump(job['results'],[])
            c.dump(job['skipped_results'],[{**r,'status':'skipped','error':'bad match'} for r in rows])
        with patch.object(c,'inspect',return_value=(rows,slides)),patch.object(c,'render_references'),patch.object(c,'run_worker',side_effect=worker),patch.object(c,'emit') as events:
            code=c.main([str(self.source),'--skip-failed','--output',str(output),'--work-dir',str(self.work)])
        self.assertEqual(code,0);self.assertFalse(output.exists())
        self.assertTrue(any(call.args[0]=='no_success' for call in events.call_args_list))
        self.assertFalse(any(call.args[0]=='saved' for call in events.call_args_list))

    def test_scan_cancel_is_honored(self):
        self.fixture();flag=self.work/'cancel';flag.touch()
        with self.assertRaises(c.Cancelled):c.inspect(self.source,self.work,extract=False,cancel_file=str(flag))

    def test_clipped_graph_fails_similarity(self):
        ref=Image.new('RGB',(512,512),'white');draw=ImageDraw.Draw(ref)
        draw.rectangle((40,40,470,470),outline='black',width=5)
        draw.line((40,440,180,300,300,380,470,80),fill='blue',width=7)
        identical=c.match_score(ref,ref,{})
        clipped=ref.copy();ImageDraw.Draw(clipped).rectangle((250,0,512,512),fill='white')
        bad=c.match_score(ref,clipped,{})
        self.assertGreater(identical['ink_overlap'],.99)
        self.assertLess(bad['ink_overlap'],.68)


if __name__=='__main__':unittest.main()
