import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from extract import ITEMS
from report import build_report
from image_evidence import verify_manifest, settings, safe_file


def fixture(n=120):
    checks=[]
    for ident,(name,sheet,headers) in ITEMS.items():
        rows=[]
        if ident=='20.1':
            rows=[['https://example.com/page'+str(i%3),'https://cdn.example.com/中文'+str(i%2)+'.png','=literal'] for i in range(n)]
        checks.append(dict(id=ident,item_name=name,sheet=sheet,headers=headers,result='X' if rows else '√',
                           finding=f'{len(rows)} issue rows recorded.' if rows else 'No issues.',action='Review',
                           rows=rows,row_count=len(rows),sources=[],gaps=[]))
    return dict(checks=checks,site_name='Example',site_url='https://example.com',audit_date='2026-11-19',run_id='test-images')


class ImageEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.root=Path(self.tmp.name)
    def make(self,options=None,n=120):
        paths=build_report(fixture(n),self.root,options or {'row_limit':5,'preview_rows':3,'target_bytes':8000}, {'export_sha256':'synthetic'})
        return paths[-1],paths[-1].with_suffix('.evidence.json')
    def test_preview_complete_order_duplicates_unicode_formula_and_parts(self):
        w,p=self.make();m=verify_manifest(p,w);c=m['checks']['20.1']
        self.assertEqual((c['reference_rows'],c['displayed_rows'],c['unique_assets'],c['affected_pages']),(120,3,2,3))
        self.assertGreater(len(c['parts']),1)
        self.assertEqual(c['preview'],fixture()['checks'][-2]['rows'][:3])
        import openpyxl
        book=openpyxl.load_workbook(w);self.assertEqual(book['20.1 Image - Alt Text Missing'].max_row,4)
        self.assertEqual(book['20.1 Image - Alt Text Missing']['C2'].data_type,'s');book.close()
    def test_small_legacy_output(self):
        paths=build_report(fixture(),self.root)
        self.assertFalse(paths[-1].with_suffix('.evidence.json').exists())
    def test_default_row_boundary(self):
        w,p=self.make({'row_limit':50000},50001)
        c=verify_manifest(p,w)['checks']['20.1'];self.assertEqual(c['displayed_rows'],1000)
        self.assertEqual(c['reference_rows'],50001)
    def test_real_byte_threshold_switch(self):
        # Compressible long evidence makes full workbook bigger than its 1-row preview.
        data=fixture(600)
        from report import render_workbook
        def sized(expected):
            raw=render_workbook(expected)
            return raw
        with patch('report.render_workbook',side_effect=sized):
            paths=build_report(data,self.root,{'row_limit':10000,'preview_rows':1,'target_bytes':8000})
        p=paths[-1].with_suffix('.evidence.json');m=verify_manifest(p,paths[-1])
        self.assertEqual(m['checks']['20.1']['reason'],'workbook_size')
    def test_missing_corrupt_and_changed_preview(self):
        w,p=self.make();m=json.loads(p.read_text());part=self.root/m['checks']['20.1']['parts'][0]['file']
        original=part.read_bytes();part.write_bytes(original+b'x')
        with self.assertRaises(ValueError):verify_manifest(p,w)
        part.write_bytes(original);m['checks']['20.1']['preview'][0][0]='changed';p.write_text(json.dumps(m))
        with self.assertRaises(ValueError):verify_manifest(p,w)
        part.unlink()
        with self.assertRaises(FileNotFoundError):verify_manifest(p,w)
    def test_no_overwrite_and_failed_write_cleanup(self):
        self.make();before={p.name:p.read_bytes() for p in self.root.iterdir()}
        with self.assertRaises(FileExistsError):self.make()
        self.assertEqual(before,{p.name:p.read_bytes() for p in self.root.iterdir()})
    def test_bounds_and_paths(self):
        for v in ({'preview_rows':1001},{'target_bytes':1},{'row_limit':True}):
            with self.assertRaises(ValueError):settings(v)
        for name in ('../x','/tmp/x','a/b','a\\b'):
            with self.assertRaises(ValueError):safe_file(self.root,name)

if __name__=='__main__':unittest.main()
