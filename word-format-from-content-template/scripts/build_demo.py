#!/usr/bin/env python3
"""Build a synthetic, reproducible formatting example, not a general formatter.

Requires python-docx. Inputs and output are generated under a fresh --output-dir.
Existing files are never overwritten. The sample deliberately has no external
links or media relationships so copying its body XML is safe for this fixture.
Do not reuse this copying shortcut for arbitrary user DOCX documents.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    dest = args.output_dir.resolve()
    names = ["source.docx", "template.docx", "formatted.docx", "format-rules.md",
             "REQUEST.txt", "structural-checks.json"]
    if any((dest / name).exists() for name in names):
        parser.error("Choose a fresh output directory; demonstration files already exist.")
    try:
        from docx import Document
        from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Mm, Pt, RGBColor
    except ImportError:
        parser.error("python-docx is required; use the document runtime in this environment.")
    dest.mkdir(parents=True, exist_ok=True)

    def set_font(style, size: float, east_asia="宋体", latin="Times New Roman"):
        style.font.name = latin
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        props = style.element.get_or_add_rPr()
        fonts = props.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            props.insert(0, fonts)
        for key in ("ascii", "hAnsi", "cs"):
            fonts.set(qn("w:" + key), latin)
        fonts.set(qn("w:eastAsia"), east_asia)
        for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn("w:" + key), None)

    def page_setup(doc):
        section = doc.sections[0]
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.top_margin = section.bottom_margin = Mm(25.4)
        section.left_margin = section.right_margin = Mm(28)
        section.header_distance = section.footer_distance = Mm(12.7)
        section.header.paragraphs[0].text = "排版验证报告"
        section.footer.paragraphs[0].text = "公开演示材料"
        for p in (section.header.paragraphs[0], section.footer.paragraphs[0]):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    source = Document()
    page_setup(source)
    set_font(source.styles["Normal"], 10.5, "等线", "Calibri")
    source.styles["Normal"].paragraph_format.line_spacing = 1.0
    title = source.add_paragraph("课程报告排版验证示例", "Title")
    title.runs[0].font.color.rgb = RGBColor(190, 40, 40)
    source.add_paragraph("一 任务说明", "Heading 1")
    source.add_paragraph("以下为公开演示用的虚构记录。排版任务只改变外观，不改写正文、表格数值或公式。")
    source.add_paragraph("操作要求：请在 Word 中更新目录。")
    source.add_paragraph("数据字段 data base 与 database 不应混淆；金额 1 250.50 元须原样保留。")
    source.add_paragraph("二 数据记录", "Heading 1")
    table = source.add_table(rows=1, cols=3)
    rows = [
        ("项目", "数值", "说明"),
        ("样例 A", "1 250.50", "数值中的空格属于原稿内容。"),
        ("样例 B", "0.25", "不得自行改写为百分比。"),
        ("样例 C", "12", "不新增单位或推断业务成效。"),
    ]
    for cell, text in zip(table.rows[0].cells, rows[0]):
        cell.text = text
    for row in rows[1:]:
        for cell, text in zip(table.add_row().cells, row):
            cell.text = text
    source.add_paragraph("三 公式与文字", "Heading 1")
    p = source.add_paragraph("原生公式：")
    math = OxmlElement("m:oMath")
    for text in ("x", "=", "a", "+", "b"):
        run = OxmlElement("m:r")
        token = OxmlElement("m:t")
        token.text = text
        run.append(token)
        math.append(run)
    p._p.append(math)
    source.add_paragraph("核对时同时检查公式结构、表格顺序、页眉页脚，以及正文中的空格和标点。")
    source_path = dest / "source.docx"
    source.save(source_path)

    template = Document()
    page_setup(template)
    # The runtime's default template may carry a blue rule in Title/Subtitle.
    # This sample's rules request typography only, so remove inherited rules.
    for style in template.styles:
        ppr = style.element.find(qn("w:pPr"))
        if ppr is not None:
            border = ppr.find(qn("w:pBdr"))
            if border is not None:
                ppr.remove(border)
    body = template.styles["Normal"]
    set_font(body, 12)
    body.paragraph_format.line_spacing = 1.5
    body.paragraph_format.first_line_indent = Pt(24)
    body.paragraph_format.space_after = Pt(6)
    for name, size in (("Title", 18), ("Heading 1", 14)):
        st = template.styles[name]
        set_font(st, size, "黑体")
        st.font.bold = True
        st.paragraph_format.first_line_indent = Pt(0)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.space_before = Pt(12 if name == "Heading 1" else 0)
        st.paragraph_format.space_after = Pt(8 if name == "Heading 1" else 14)
    template.styles["Title"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for name in ("Header", "Footer"):
        set_font(template.styles[name], 9)
        template.styles[name].paragraph_format.first_line_indent = Pt(0)
    template.add_paragraph("模板旧题目", "Title")
    template.add_paragraph("模板旧作者：示例人员")
    template.add_paragraph("此处是需要替换的模板示例正文。")
    template_path = dest / "template.docx"
    template.save(template_path)

    rules = """# 演示格式要求

本例是合成测试资料，不是学校或公司的真实规范。

- A4 纵向；上下边距 25.4 mm，左右边距 28 mm。
- 中文正文宋体 12 pt，西文 Times New Roman；1.5 倍行距，首行缩进 24 pt，段后 6 pt。
- 文档标题黑体 18 pt、黑色加粗居中；一级标题黑体 14 pt、黑色加粗，左对齐且不缩进。
- 固定页眉为“排版验证报告”，固定页脚为“公开演示材料”，均居中、9 pt。
- 表格宽度 154 mm，三列分别为 36、28、90 mm；表内 10.5 pt，无首行缩进，表头加粗且重复。
- 表格边框为浅灰色，表头底色为浅灰色；不设置截断内容的固定行高。
- 保留所有原稿文字、手动编号、空格、数字、表格顺序及原生公式；不新增目录、题注或页码。
- 清除模板旧题目、旧作者及示例正文，保留本例固定页眉页脚。
"""
    (dest / "format-rules.md").write_text(rules, encoding="utf-8")
    (dest / "REQUEST.txt").write_text(
        "请用 $word-format-from-content-template 将 source.docx 按 format-rules.md 排入 template.docx，"
        "输出到你自己的新文件名。只调整格式，保留原稿所有文字、表格和原生公式。"
        "不要覆盖输入；请做内容核对和逐页视觉检查。\n",
        encoding="utf-8",
    )
    input_hashes = {p.name: digest(p) for p in (source_path, template_path, dest / "format-rules.md")}

    # This synthetic source has only body/table text and inline OMML. Arbitrary
    # DOCX bodies may have relationship-backed images/links: never copy them by
    # this shortcut without also preserving and verifying their relationships.
    final = Document(template_path)
    for child in list(final.element.body):
        if child.tag != qn("w:sectPr"):
            final.element.body.remove(child)
    for child in source.element.body:
        if child.tag != qn("w:sectPr"):
            final.element.body.insert(len(final.element.body) - 1, deepcopy(child))
    for p in final.paragraphs:
        for run in p.runs:
            if run._r.rPr is not None:
                run._r.remove(run._r.rPr)
    for tbl in final.tables:
        tbl.autofit = False
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        props = tbl._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement("w:" + edge)
            for key, value in (("val", "single"), ("sz", "4"), ("color", "D9D9D9")):
                border.set(qn("w:" + key), value)
            borders.append(border)
        props.append(borders)
        widths = (36, 28, 90)
        for col, width in zip(tbl.columns, widths):
            col.width = Mm(width)
        for row_index, row in enumerate(tbl.rows):
            for col_index, (cell, width) in enumerate(zip(row.cells, widths)):
                cell.width = Mm(width)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                tcpr = cell._tc.get_or_add_tcPr()
                margins = OxmlElement("w:tcMar")
                for edge in ("top", "left", "bottom", "right"):
                    margin = OxmlElement("w:" + edge)
                    margin.set(qn("w:w"), "100")
                    margin.set(qn("w:type"), "dxa")
                    margins.append(margin)
                tcpr.append(margins)
                if row_index == 0:
                    shading = OxmlElement("w:shd")
                    shading.set(qn("w:fill"), "EAEAEA")
                    tcpr.append(shading)
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.first_line_indent = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1.2
                    paragraph.paragraph_format.space_after = Pt(2)
                    paragraph.paragraph_format.space_before = Pt(2)
                    paragraph.alignment = (WD_ALIGN_PARAGRAPH.CENTER if row_index == 0
                        else WD_ALIGN_PARAGRAPH.RIGHT if col_index == 1 else WD_ALIGN_PARAGRAPH.LEFT)
                    for run in paragraph.runs:
                        run.font.size = Pt(10.5)
                        run.bold = row_index == 0
            if row_index == 0:
                repeat = OxmlElement("w:tblHeader")
                row._tr.get_or_add_trPr().append(repeat)
    output_path = dest / "formatted.docx"
    final.save(output_path)

    after = Document(output_path)
    content = "\n".join(p.text for p in after.paragraphs)
    checks = {
        "fixture_only": True,
        "input_hashes": input_hashes,
        "inputs_unchanged": all(digest(dest / name) == value for name, value in input_hashes.items()),
        "template_residue_absent": all(s not in content for s in ("模板旧题目", "模板旧作者", "模板示例正文")),
        "page_width_mm": round(after.sections[0].page_width / Mm(1), 2),
        "page_height_mm": round(after.sections[0].page_height / Mm(1), 2),
        "table_count": len(after.tables),
        "equation_count": len(after.element.body.findall(".//" + qn("m:oMath"))),
        "text_comparison": "Run scripts/compare_text.py source.docx formatted.docx --report comparison.json separately.",
        "visual_review": "Not asserted by this builder; render and inspect separately.",
    }
    assert checks["inputs_unchanged"] and checks["template_residue_absent"]
    assert checks["table_count"] == 1 and checks["equation_count"] == 1
    (dest / "structural-checks.json").write_text(json.dumps(checks, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_dir": str(dest), "files": names}, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
