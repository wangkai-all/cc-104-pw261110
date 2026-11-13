#!/usr/bin/env python3
"""Generate synthetic, privacy-free DOCX survey samples for the importer demo.

Requires python-docx (`pip install python-docx`). Output goes to samples/.

Samples:
  sample_survey.docx          survey with a merged title, a genuine two-row header,
                              vertical region names, a repeated header mid-table,
                              a subtotal row, in-table notes, blank/无/0 values,
                              and a nested table in a later sheet.
  sample_survey_revised.docx  structurally changed version (extra column, renamed
                              header, swapped table order) to demo structure diffs.
  sample_nested.docx          just the nested-table sheet.
"""
import os
import sys

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:  # pragma: no cover
    sys.exit("python-docx is required: pip install python-docx")

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "samples")


def put(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text or "")
    run.bold = bold


def build_main(path, dataset=None):
    """Build the main survey sample. `dataset` lets us emit a structurally
    identical file with different, still fictional, data."""
    data_rows = dataset or [
        ("甲镇", "东河村", "1200", "每月", "地下水", "达标", "12.5", "万m³", "无"),
        ("甲镇", "西河村", "", "每季", "地表水", "", "8.0", "万m³", ""),
        ("甲镇", "小计", "1200", "", "", "", "20.5", "", ""),
        ("甲镇", "南坡村", "0", "每年", "地下水", "达标", "0", "万m³", "未通水"),
        ("乙乡", "北沟村", "", "每月", "泉水", "不达标", "3.2", "万m³", "已上报"),
        ("乙乡", "林场村", "860", "每月", "地下水", "达标", "6.4", "万m³", "无"),
    ]
    second_block = dataset is None or len(dataset) > 6
    doc = Document()
    doc.add_heading("乡村供水情况调查表", level=1)
    p = doc.add_paragraph("填报单位：示例县水务局（示例数据，非真实信息）")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    t = doc.add_table(rows=1, cols=9)
    t.style = "Table Grid"

    # row 0: merged title
    title = t.cell(0, 0).merge(t.cell(0, 8))
    put(title, "表1  各村供水基本情况（2026年度）", bold=True)
    title.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # rows 1-2: two-row header. 地区/村名/水源类型/水质/备注 span both rows;
    # 供水情况 covers 供水人口+检测频次, 年供水量 covers 数值+计量单位.
    t.add_row()  # row 1
    t.add_row()  # row 2
    for i, v in [(0, "地区"), (1, "村名"), (4, "水源类型"), (5, "水质"), (8, "备注")]:
        put(t.cell(1, i), v, bold=True)
    put(t.cell(1, 2).merge(t.cell(1, 3)), "供水情况", bold=True)
    put(t.cell(1, 6).merge(t.cell(1, 7)), "年供水量", bold=True)
    for i, v in [(2, "供水人口"), (3, "检测频次"), (6, "数值"), (7, "计量单位")]:
        put(t.cell(2, i), v, bold=True)
    for col, label in [(0, "地区"), (1, "村名"), (4, "水源类型"), (5, "水质"), (8, "备注")]:
        put(t.cell(1, col).merge(t.cell(2, col)), label, bold=True)

    # data starts at row 3
    rows = data_rows
    start = 3
    for vals in rows:
        r = t.add_row()
        for i, v in enumerate(vals):
            put(r.cells[i], v)
    # vertical merges: 甲镇 rows 3..6 (subtotal inside), 乙乡 rows 7..8
    put(t.cell(start, 0).merge(t.cell(start + 3, 0)), "甲镇")
    put(t.cell(start + 4, 0).merge(t.cell(start + 5, 0)), "乙乡")

    note = t.add_row()
    put(note.cells[0].merge(note.cells[8]),
        "注：①空白表示未填报；②“无”表示确无此项；③“0”表示实测为零；④小计行不参与汇总。")
    doc.add_paragraph("填表人：示例　联系电话：0000-00000000")

    # Table 2: repeated header inside one table + nested table
    doc.add_page_break()
    doc.add_paragraph("表2  水源地明细（含嵌套表）")
    t2 = doc.add_table(rows=1, cols=4)
    t2.style = "Table Grid"
    for i, v in enumerate(["地区", "水源地名称", "类型", "明细"]):
        put(t2.cell(0, i), v, bold=True)
    r1 = t2.add_row()
    for i, v in enumerate(["丙镇", "一号井", "地下水", "见单元格内嵌套表"]):
        put(r1.cells[i], v)
    nested = r1.cells[3].add_table(rows=3, cols=3)
    nested.style = "Table Grid"
    for i, v in enumerate(["检测项", "结果", "单位"]):
        put(nested.cell(0, i), v, bold=True)
    for r, vals in enumerate([["浊度", "1.2", "NTU"], ["pH", "7.4", ""]], start=1):
        for c, v in enumerate(vals):
            put(nested.cell(r, c), v)
    # repeated header row then a second block
    r2 = t2.add_row()
    for i, v in enumerate(["地区", "水源地名称", "类型", "明细"]):
        put(r2.cells[i], v, bold=True)
    r3 = t2.add_row()
    for i, v in enumerate(["丁乡", "三号泉", "泉水", ""]):
        put(r3.cells[i], v)

    doc.save(path)


def build_revised(path):
    """Deliberately changed structure: table order swapped, one extra column,
    one renamed header, and no two-row header."""
    doc = Document()
    doc.add_heading("乡村供水情况调查表（修订版）", level=1)

    t2 = doc.add_table(rows=1, cols=3)
    t2.style = "Table Grid"
    for i, v in enumerate(["水源", "水量", "状态"]):
        put(t2.cell(0, i), v, bold=True)
    r = t2.add_row()
    for i, v in enumerate(["一号井", "120", "在用"]):
        put(r.cells[i], v)

    doc.add_paragraph("以下为修订后的表1：")
    t = doc.add_table(rows=1, cols=8)
    t.style = "Table Grid"
    put(t.cell(0, 0).merge(t.cell(0, 7)), "表1  各村供水基本情况（2026年度修订）", bold=True)
    # single header row; 供水方式 is new; 水质 column renamed
    t.add_row()  # row 1
    heads = ["地区", "村名", "供水人口", "水源类型", "供水方式", "水质评价", "年供水量", "备注"]
    for i, v in enumerate(heads):
        put(t.cell(1, i), v, bold=True)
    data = [
        ("甲镇", "东河村", "1200", "地下水", "集中", "达标", "12.5", "无"),
        ("甲镇", "西河村", "", "地表水", "分散", "", "8.0", ""),
        ("乙乡", "北沟村", "", "泉水", "集中", "不达标", "3.2", "已上报"),
        ("乙乡", "林场村", "860", "地下水", "集中", "达标", "6.4", "无"),
    ]
    for vals in data:
        rr = t.add_row()
        for i, v in enumerate(vals):
            put(rr.cells[i], v)
    # region vertical merges: data rows are 2..5
    put(t.cell(2, 0).merge(t.cell(3, 0)), "甲镇")
    put(t.cell(4, 0).merge(t.cell(5, 0)), "乙乡")
    note = t.add_row()
    put(note.cells[0].merge(note.cells[7]), "注：空白表示未填报；“无”表示确无此项。")
    doc.save(path)


def build_nested(path):
    doc = Document()
    doc.add_heading("嵌套表示例", level=1)
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    for i, v in enumerate(["名称", "状态", "检测明细"]):
        put(t.cell(0, i), v, bold=True)
    r = t.add_row()
    put(r.cells[0], "一号井")
    put(r.cells[1], "在用")
    nested = r.cells[2].add_table(rows=2, cols=2)
    nested.style = "Table Grid"
    for i, v in enumerate(["检测项", "结果"]):
        put(nested.cell(0, i), v, bold=True)
    put(nested.cell(1, 0), "浊度")
    put(nested.cell(1, 1), "1.2")
    doc.save(path)


def main():
    os.makedirs(OUT, exist_ok=True)
    build_main(os.path.join(OUT, "sample_survey.docx"))
    build_main(os.path.join(OUT, "sample_survey_second.docx"), dataset=[
        ("河东镇", "柳湾村", "640", "每月", "地表水", "达标", "5.1", "万m³", ""),
        ("河东镇", "沙岭村", "无", "每季", "地下水", "", "无", "", "无"),
        ("河东镇", "小计", "640", "", "", "", "5.1", "", ""),
        ("河东镇", "石桥村", "0", "每年", "泉水", "达标", "0", "万m³", "未通水"),
        ("河西乡", "苇塘村", "1250", "每月", "地下水", "达标", "9.8", "万m³", ""),
        ("河西乡", "杨树村", "730", "每季", "地表水", "不达标", "4.3", "万m³", "已上报"),
    ])
    build_revised(os.path.join(OUT, "sample_survey_revised.docx"))
    build_nested(os.path.join(OUT, "sample_nested.docx"))
    with open(os.path.join(OUT, "index.json"), "w", encoding="utf-8") as f:
        import json
        json.dump({"samples": [
            "sample_survey.docx",
            "sample_survey_second.docx",
            "sample_survey_revised.docx",
            "sample_nested.docx",
        ]}, f, ensure_ascii=False, indent=2)
    print("written to", OUT)


if __name__ == "__main__":
    main()
