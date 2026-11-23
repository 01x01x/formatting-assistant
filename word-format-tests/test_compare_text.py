"""Black-box content-audit regressions using small synthetic OOXML packages."""
from __future__ import annotations

import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import unittest
import uuid
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZIP_DEFLATED, ZipFile
import zlib

ROOT = Path(__file__).resolve().parent
FIX = ROOT / "fixtures"
FIX.mkdir(exist_ok=True)
SCRIPT = ROOT.parent / "word-format-from-content-template" / "scripts" / "compare_text.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = f'xmlns:w="{W}" xmlns:r="{R}" xmlns:m="{M}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
RESULTS = []


def run(text, props=""):
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(text):
    return '<w:p>' + run(text) + '</w:p>'


def table(rows):
    return '<w:tbl>' + ''.join('<w:tr>' + ''.join('<w:tc>' + para(cell) + '</w:tc>' for cell in row) + '</w:tr>' for row in rows) + '</w:tbl>'


def png(red):
    def chunk(kind, value):
        return struct.pack('>I', len(value)) + kind + value + struct.pack('>I', zlib.crc32(kind + value) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\x00' + bytes((red, 20, 30)))) + chunk(b'IEND', b'')


def picture(rid="rPic"):
    return f'<w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="demo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'


def field(instruction=" PAGE ", cached="1"):
    return '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">' + escape(instruction) + '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run(cached) + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'


def docx(name, body, *, header=None, footer=None, footnotes=None, endnotes=None, comments=None, rels=(), media=None, extra=None, styles=None):
    path = FIX / (name + '.docx')
    parts, relationships, sections = {}, list(rels), []
    for kind, text in (("header", header), ("footer", footer)):
        if text is not None:
            parts[f'word/{kind}1.xml'] = f'<w:{"hdr" if kind == "header" else "ftr"} {NS}>{para(text)}</w:{"hdr" if kind == "header" else "ftr"}>'
            relationships.append((kind, kind, kind + '1.xml', False))
            sections.append(f'<w:{kind}Reference w:type="default" r:id="{kind}"/>')
    for kind, mapping in (("footnote", footnotes), ("endnote", endnotes), ("comment", comments)):
        if mapping is None:
            continue
        plural = kind + 's'
        content = ''.join(f'<w:{kind} w:id="{key}"' + (' w:author="Tester" w:initials="T"' if kind == 'comment' else '') + '>' + para(value) + f'</w:{kind}>' for key, value in mapping.items())
        parts['word/' + plural + '.xml'] = f'<w:{plural} {NS}>{content}</w:{plural}>'
        relationships.append((plural, plural, plural + '.xml', False))
    if styles is not None:
        parts['word/styles.xml'] = f'<w:styles {NS}>{styles}</w:styles>'
        relationships.append(('styles', 'styles', 'styles.xml', False))
    parts['word/document.xml'] = f'<w:document {NS}><w:body>{body}<w:sectPr>{"".join(sections)}</w:sectPr></w:body></w:document>'
    parts['_rels/.rels'] = f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="root" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>'
    parts['word/_rels/document.xml.rels'] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(f'<Relationship Id={quoteattr(rid)} Type={quoteattr(R + "/" + kind)} Target={quoteattr(target)}' + (' TargetMode="External"' if external else '') + '/>' for rid, kind, target, external in relationships) + '</Relationships>'
    parts['[Content_Types].xml'] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    parts.update(media or {})
    parts.update(extra or {})
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for member, value in parts.items():
            archive.writestr(member, value)
    return path


class ContentAuditTests(unittest.TestCase):
    def pair(self, label, left, right, expected, *flags, warning=None, strict=None):
        report = FIX / (label + '-report.json')
        result = subprocess.run([sys.executable, '-X', 'utf8', str(SCRIPT), str(left), str(right), '--report', str(report), *flags], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, expected, label + '\n' + result.stdout + '\n' + result.stderr)
        data = json.loads(report.read_text(encoding='utf-8'))
        self.assertEqual(data['exit_code'], expected)
        self.assertEqual(data['status'], {0:'PASS', 1:'FAIL', 2:'ERROR', 3:'REVIEW'}[expected])
        if warning:
            self.assertIn(warning, {item['code'] for item in data['warnings']})
        if strict is not None:
            self.assertEqual(data['strict_equal'], strict)
        RESULTS.append({'case': label, 'expected_exit': expected, 'actual_exit': result.returncode, 'status': data['status']})
        return data

    def texts(self, label, left, right, expected, *flags):
        a, b = FIX / (label + '-a.txt'), FIX / (label + '-b.txt')
        a.write_bytes(left.encode('utf-8'))
        b.write_bytes(right.encode('utf-8'))
        return self.pair(label, a, b, expected, *flags)

    def test_text_contract(self):
        cases = [
            ('identical', '原稿。\n第二段。', '原稿。\n第二段。', 0, ()),
            ('bom_crlf', '\ufeff甲\r\n乙\r', '甲\n乙\n', 0, ()),
            ('interior_bom', '甲\ufeff乙', '甲乙', 1, ()),
            ('old_phrase_falsepass', '操作要求：请在 Word 中更新目录。', '操作要求：。', 1, ()),
            ('old_field_tutorial_falsepass', '示例：TOC \\o "1-3" \\h \\z \\u。', '示例：。', 1, ()),
            ('indentation', 'if ok:\n    run()', 'if ok:\nrun()', 1, ()),
            ('tab_boundary', '甲\t乙', '甲 乙', 1, ()),
            ('punctuation', '甲，乙。', '甲,乙。', 1, ()),
            ('paragraph_order', '甲\n乙', '乙\n甲', 1, ('--show-diff',)),
            ('loose_only_review', 'a bc', 'ab c', 3, ('--loose',)),
            ('loose_identical_review', '甲', '甲', 3, ('--loose',)),
            ('loose_still_difference', '甲', '乙', 1, ('--loose',)),
        ]
        for label, a, b, expected, flags in cases:
            with self.subTest(label=label):
                self.texts(label, a, b, expected, *flags)

    def test_bad_inputs(self):
        good = FIX / 'good.txt'; good.write_text('原稿', encoding='utf-8')
        empty = FIX / 'empty.txt'; empty.write_bytes(b'')
        white = FIX / 'whitespace.txt'; white.write_text(' \t\n',encoding='utf-8')
        gbk = FIX / 'gbk.txt'; gbk.write_bytes('原稿'.encode('gbk'))
        badzip = FIX / 'bad.docx'; badzip.write_bytes(b'not a zip')
        validzip = FIX / 'notdocx.docx'
        with ZipFile(validzip, 'w') as archive: archive.writestr('hello.txt', 'hello')
        for label, left, right in (
            ('empty_error', empty, empty), ('whitespace_only_error',white,white), ('encoding_error', gbk, gbk), ('missing_error', FIX/'missing.txt', good),
            ('bad_zip_error', badzip, badzip), ('zip_not_docx_error', validzip, validzip),
            ('mixed_type_error', good, docx('mixed', para('原稿'))),
        ):
            with self.subTest(label=label): self.pair(label, left, right, 2)
        badxml = docx('badxml', para('原稿'), extra={'word/document.xml':'<unclosed'})
        self.pair('bad_xml_error', badxml, badxml, 2)

    def test_report_must_not_modify_inputs(self):
        for linked_input in ('source', 'final'):
            for alias_type in ('same_path', 'hard_link'):
                with self.subTest(input=linked_input, alias=alias_type):
                    token = uuid.uuid4().hex
                    source = FIX / (token + '-source.txt')
                    final = FIX / (token + '-final.txt')
                    original = '原稿与结果都应保留。'.encode('utf-8')
                    source.write_bytes(original)
                    final.write_bytes(original)
                    target = source if linked_input == 'source' else final
                    report = target
                    if alias_type == 'hard_link':
                        report = FIX / (token + '-linked-report.json')
                        os.link(target, report)
                    result = subprocess.run([sys.executable, '-X', 'utf8', str(SCRIPT), str(source), str(final), '--report', str(report)], capture_output=True, text=True, encoding='utf-8')
                    self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
                    self.assertNotIn('Traceback', result.stderr)
                    self.assertEqual(source.read_bytes(), original)
                    self.assertEqual(final.read_bytes(), original)
                    self.assertEqual(report.read_bytes(), original)
                    RESULTS.append({'case': f'report_{alias_type}_{linked_input}_protected', 'expected_exit': 2, 'actual_exit': result.returncode, 'status': 'ERROR'})

    def test_console_cannot_encode_glyphs(self):
        source = FIX / 'console-emoji-source.txt'
        final = FIX / 'console-emoji-final.txt'
        report = FIX / 'console-emoji-report.json'
        source.write_text('项目🚀', encoding='utf-8')
        final.write_text('项目✅', encoding='utf-8')
        env = {**os.environ, 'PYTHONIOENCODING': 'gbk:strict', 'PYTHONUTF8': '0'}
        result = subprocess.run([sys.executable, str(SCRIPT), str(source), str(final), '--show-diff', '--report', str(report)], capture_output=True, env=env)
        self.assertEqual(result.returncode, 1)
        self.assertEqual(result.stderr, b'')
        self.assertIn(b'\\U0001f680', result.stdout)
        data = json.loads(report.read_text(encoding='utf-8'))
        self.assertEqual(data['status'], 'FAIL')
        self.assertIn('🚀', '\n'.join(data['diff_preview']))
        RESULTS.append({'case': 'gbk_console_diff_no_crash_utf8_report', 'expected_exit': 1, 'actual_exit': result.returncode, 'status': 'FAIL'})

    def test_docx_text_order_and_runs(self):
        base = docx('basic', para('甲乙') + table([['一', '二']]) + para('丙'))
        self.pair('docx_identical', base, base, 0)
        split = docx('split', '<w:p>' + run('甲', '<w:rPr><w:b/></w:rPr>') + run('乙') + '</w:p>' + table([['一','二']]) + para('丙'))
        self.pair('formatting_run_split', base, split, 0)
        reordered = docx('reordered', table([['一','二']]) + para('甲乙') + para('丙'))
        self.pair('body_table_order', base, reordered, 1)
        onecell = docx('onecell', para('甲乙') + '<w:tbl><w:tr><w:tc>' + para('一') + para('二') + '</w:tc></w:tr></w:tbl>' + para('丙'))
        self.pair('table_cell_boundaries', base, onecell, 1)
        self.pair('docx_literal_toc_phrase', docx('literal_a', para('操作要求：请在 Word 中更新目录。')), docx('literal_b', para('操作要求：。')), 1)
        self.pair('docx_whitespace_loose', docx('loose_doc_a', para('a bc')), docx('loose_doc_b', para('ab c')), 3, '--loose')

    def test_notes_and_furniture_scope(self):
        body = '<w:p>' + run('甲') + '<w:r><w:footnoteReference w:id="1"/><w:endnoteReference w:id="2"/></w:r>' + run('乙') + '</w:p>'
        a = docx('notes_a', body, footnotes={'1':'原脚注'}, endnotes={'2':'原尾注'}, header='原页眉', footer='原页脚')
        self.pair('footnote_change', a, docx('notes_b', body, footnotes={'1':'改脚注'}, endnotes={'2':'原尾注'}, header='原页眉',footer='原页脚'), 1)
        self.pair('endnote_change', a, docx('notes_c', body, footnotes={'1':'原脚注'}, endnotes={'2':'改尾注'}, header='原页眉',footer='原页脚'), 1)
        anchor = '<w:p><w:r><w:footnoteReference w:id="1"/><w:endnoteReference w:id="2"/></w:r>' + run('甲乙') + '</w:p>'
        self.pair('footnote_anchor_change', a, docx('notes_anchor', anchor, footnotes={'1':'原脚注'}, endnotes={'2':'原尾注'}, header='原页眉',footer='原页脚'), 1)
        b = docx('headers_b', body, footnotes={'1':'原脚注'}, endnotes={'2':'原尾注'}, header='改页眉', footer='原页脚')
        self.pair('header_all', a, b, 1)
        self.pair('header_content_excluded', a, b, 0, '--scope','content')
        c = docx('footers_b', body, footnotes={'1':'原脚注'}, endnotes={'2':'原尾注'}, header='原页眉', footer='改页脚')
        self.pair('footer_all', a, c, 1)
        self.pair('footer_content_excluded', a, c, 0, '--scope','content')
        self.pair('orphan_note_review', docx('orphan', para('正文'),footnotes={'7':'无锚点脚注'}), docx('orphan2',para('正文'),footnotes={'7':'无锚点脚注'}), 3, warning='unreferenced_note')

    def test_comments_and_anchors(self):
        body = '<w:p><w:commentRangeStart w:id="4"/>' + run('甲') + '<w:commentRangeEnd w:id="4"/>' + run('乙') + '<w:r><w:commentReference w:id="4"/></w:r></w:p>'
        moved = '<w:p>' + run('甲') + '<w:commentRangeStart w:id="4"/>' + run('乙') + '<w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r></w:p>'
        a, b = docx('comment_a',body,comments={'4':'原批注'}),docx('comment_b',body,comments={'4':'改批注'})
        self.pair('comment_text_change',a,b,1)
        self.pair('comment_anchor_change',a,docx('comment_moved',moved,comments={'4':'原批注'}),1)
        self.pair('comment_content_excluded',a,b,0,'--scope','content')

    def test_hyperlinks(self):
        body = '<w:p><w:hyperlink r:id="url">' + run('链接') + '</w:hyperlink></w:p>'
        a = docx('link_a',body,rels=[('url','hyperlink','https://example.com/a',True)])
        b = docx('link_b',body,rels=[('url','hyperlink','https://example.com/b',True)])
        self.pair('hyperlink_target_change',a,b,1)
        renamed = docx('link_renamed',body.replace('r:id="url"','r:id="newid"'),rels=[('newid','hyperlink','https://example.com/a',True)])
        self.pair('hyperlink_rid_renamed',a,renamed,0)
        missing = docx('link_missing',body)
        self.pair('missing_relationship_review',missing,missing,3,warning='missing_relationship')
        anchor = docx('anchor', '<w:p><w:hyperlink w:anchor="target">' + run('链接') + '</w:hyperlink></w:p>')
        self.pair('internal_anchor_review',anchor,anchor,3,warning='internal_anchor_review')

    def test_images_and_equations(self):
        a = docx('image_a',picture(),rels=[('rPic','image','media/a.png',False)],media={'word/media/a.png':png(10)})
        b = docx('image_b',picture(),rels=[('rPic','image','media/a.png',False)],media={'word/media/a.png':png(200)})
        self.pair('image_bytes_change',a,b,1)
        renamed = docx('image_rename',picture('newpic'),rels=[('newpic','image','media/renamed.png',False)],media={'word/media/renamed.png':png(10)})
        self.pair('image_internal_names_changed',a,renamed,0)
        missing = docx('image_missing',picture(),rels=[('rPic','image','media/a.png',False)])
        self.pair('missing_media_review',missing,missing,3,warning='missing_target_part')
        extra_action = docx('image_action',picture().replace('<wp:docPr id="1" name="demo"/>','<wp:docPr id="1" name="demo"><a:hlinkClick r:id="action"/></wp:docPr>'),rels=[('rPic','image','media/a.png',False),('action','hyperlink','https://example.com/',True)],media={'word/media/a.png':png(10)})
        self.pair('image_action_requires_review',extra_action,extra_action,3,warning='unverified_drawing_relationship')
        def equation(content): return '<w:p><m:oMath>' + content + '</m:oMath></w:p>'
        math_a = docx('math_a',equation('<m:r><m:t>xy</m:t></m:r>'))
        self.pair('equation_text_change',math_a,docx('math_b',equation('<m:r><m:t>xz</m:t></m:r>')),1)
        self.pair('equation_structure_change',math_a,docx('math_fraction',equation('<m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>y</m:t></m:r></m:den></m:f>')),1)
        self.pair('equation_run_format_split',math_a,docx('math_split',equation('<m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r><m:r><m:t>y</m:t></m:r>')),0)

    def test_review_conditions_and_fields(self):
        numbering = docx('numbering','<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>' + run('项目') + '</w:p>')
        self.pair('automatic_numbering_review',numbering,numbering,3,warning='automatic_numbering')
        revision = docx('revision','<w:p><w:ins w:id="1">' + run('待确认') + '</w:ins></w:p>')
        self.pair('revision_review',revision,revision,3,warning='tracked_revision')
        a,b = docx('field_a',para('正文')+field()),docx('field_cache',para('正文')+field(cached='2'))
        self.pair('field_identical_review',a,a,3,warning='field_requires_review',strict=True)
        report = self.pair('field_cache_changed_review',a,b,3,warning='derived_field_results_changed',strict=True)
        self.assertTrue(report['derived_field_records_changed'])
        self.pair('field_instruction_change',a,docx('field_changed',para('正文')+field(' NUMPAGES ')),1,strict=False)
        self.pair('field_removed',a,docx('field_removed',para('正文')),1,strict=False)
        spanning = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>' + para('缓存目录') + '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + para('正文')
        multi = docx('multi_field',spanning)
        self.pair('spanning_field_instruction_change',multi,docx('multi_field_changed',spanning.replace(' TOC ',' REF target ')),1,strict=False)
        self.pair('spanning_field_removed',multi,docx('multi_field_removed',para('正文')),1,strict=False)
        chart = docx('complex',para('正文')+'<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="chart"/></a:graphic></w:drawing></w:r></w:p>')
        self.pair('complex_object_review',chart,chart,3,warning='complex_drawing_object')
        empty = docx('empty_docx','')
        self.pair('empty_docx_review',empty,empty,3,warning='empty_checked_content')


if __name__ == '__main__':
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(ContentAuditTests)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    (ROOT/'regression-results.json').write_text(json.dumps({'interpreter':sys.executable,'case_count':len(RESULTS),'cases':RESULTS,'successful':result.wasSuccessful()},ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'CLI cases checked: {len(RESULTS)}')
    raise SystemExit(0 if result.wasSuccessful() else 1)
