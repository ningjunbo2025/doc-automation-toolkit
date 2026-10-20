# -*- coding: utf-8 -*-
"""Convert Markdown to a formatted Word document (Chinese typography, tables, code blocks)."""
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

CN_FONT = "微软雅黑"
CODE_FONT = "Consolas"
CODE_BG = "F2F2F2"
BODY_SIZE = Pt(10.5)


def set_run_font(run, name=CN_FONT, size=None, bold=None, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def shade_paragraph(p, fill):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def bottom_border(p):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "BFBFBF")
    pBdr.append(bottom)
    pPr.append(pBdr)


def setup_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = CN_FONT
    normal.font.size = BODY_SIZE
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CN_FONT)
    pf = normal.paragraph_format
    pf.space_before = Pt(3)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.35

    sizes = {"Heading 1": Pt(20), "Heading 2": Pt(15), "Heading 3": Pt(12.5), "Heading 4": Pt(11)}
    for name, size in sizes.items():
        st = doc.styles[name]
        st.font.name = CN_FONT
        st.element.rPr.rFonts.set(qn("w:eastAsia"), CN_FONT)
        st.font.size = size
        st.font.bold = True
        st.font.color.rgb = RGBColor(0x1F, 0x2A, 0x44)
        st.paragraph_format.space_before = Pt(14 if name == "Heading 1" else 10)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True

    for name in ("List Bullet", "List Bullet 2", "List Number"):
        st = doc.styles[name]
        st.font.name = CN_FONT
        st.element.rPr.rFonts.set(qn("w:eastAsia"), CN_FONT)
        st.font.size = BODY_SIZE


INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")


def add_inline(p, text, size=None, bold=False, color=None):
    """处理 **加粗** 和 `行内代码`。"""
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = p.add_run(part[2:-2])
            set_run_font(r, size=size, bold=True, color=color)
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            r = p.add_run(part[1:-1])
            set_run_font(r, name=CODE_FONT, size=Pt(9.5), bold=False, color=RGBColor(0xC7, 0x25, 0x4E))
        else:
            r = p.add_run(part)
            set_run_font(r, size=size, bold=bold, color=color)


def add_code_block(doc, lines):
    for i, line in enumerate(lines):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.15
        pf.left_indent = Cm(0.3)
        pf.right_indent = Cm(0.1)
        shade_paragraph(p, CODE_BG)
        # 代码块里保留原样，但把 **加粗** 转成真正的加粗（文案类内容会用到）
        text = line if line.strip() else " "
        for part in re.split(r"(\*\*[^*]+\*\*)", text):
            if not part:
                continue
            is_bold = part.startswith("**") and part.endswith("**") and len(part) > 4
            r = p.add_run(part[2:-2] if is_bold else part)
            set_run_font(r, name=CODE_FONT, size=Pt(9), bold=True if is_bold else None)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_table(doc, rows):
    header, body = rows[0], rows[1:]
    table = doc.add_table(rows=len(rows), cols=len(header))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, cell_text in enumerate(header):
        cell = table.cell(0, j)
        cell.text = ""
        p = cell.paragraphs[0]
        add_inline(p, cell_text, size=Pt(9.5), bold=True)
        for run in p.runs:
            run.font.bold = True
    for i, row in enumerate(body, start=1):
        for j in range(len(header)):
            cell = table.cell(i, j)
            cell.text = ""
            add_inline(cell.paragraphs[0], row[j] if j < len(row) else "", size=Pt(9.5))
    # 表头跨页重复
    trPr = table.rows[0]._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def parse_pipe_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def is_separator(line):
    return bool(re.fullmatch(r"\s*\|?[\s:\-|]+\|?\s*", line)) and "-" in line


def convert(md_path, docx_path):
    doc = Document()
    setup_styles(doc)
    sec = doc.sections[0]
    sec.top_margin = Cm(2.4)
    sec.bottom_margin = Cm(2.4)
    sec.left_margin = Cm(2.6)
    sec.right_margin = Cm(2.6)

    lines = Path(md_path).read_text(encoding="utf-8").splitlines()
    i = 0
    first_h1 = True
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 代码块
        if stripped.startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i].rstrip())
                i += 1
            i += 1
            add_code_block(doc, buf)
            continue

        # 表格
        if stripped.startswith("|") and i + 1 < len(lines) and is_separator(lines[i + 1]):
            rows = [parse_pipe_row(stripped)]
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(parse_pipe_row(lines[i]))
                i += 1
            add_table(doc, rows)
            continue

        # 空行
        if not stripped:
            i += 1
            continue

        # 水平线
        if re.fullmatch(r"-{3,}|\*{3,}", stripped):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(8)
            bottom_border(p)
            i += 1
            continue

        # 标题
        m = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            text = text.replace("**", "")
            h = doc.add_heading(level=level)
            if level == 1 and first_h1:
                h.alignment = WD_ALIGN_PARAGRAPH.CENTER
                h.paragraph_format.space_before = Pt(0)
                h.paragraph_format.space_after = Pt(18)
                first_h1 = False
            r = h.add_run(text)
            set_run_font(r, size=None, bold=True)
            i += 1
            continue

        # 引用
        if stripped.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip()[1:].strip())
                i += 1
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.4)
            bottom_border(p) if False else None
            for k, t in enumerate(buf):
                if k:
                    p.add_run().add_break()
                add_inline(p, t, size=Pt(10), color=RGBColor(0x55, 0x55, 0x55))
            # 左侧竖线
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            left = OxmlElement("w:left")
            left.set(qn("w:val"), "single")
            left.set(qn("w:sz"), "12")
            left.set(qn("w:space"), "6")
            left.set(qn("w:color"), "8EA3C4")
            pBdr.append(left)
            pPr.append(pBdr)
            continue

        # 无序列表（含缩进层级）
        m = re.match(r"^(\s*)[-·•]\s+(.*)$", line)
        if m:
            level = min(len(m.group(1)) // 2, 1)
            style = "List Bullet" if level == 0 else "List Bullet 2"
            p = doc.add_paragraph(style=style)
            if level:
                p.paragraph_format.left_indent = Cm(1.5)
            add_inline(p, m.group(2).strip())
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            i += 1
            continue

        # 有序列表
        m = re.match(r"^(\s*)(\d+)\.\s+(.*)$", line)
        if m:
            level = min(len(m.group(1)) // 2, 1)
            style = "List Number" if level == 0 else "List Number 2"
            p = doc.add_paragraph(style=style)
            if level:
                p.paragraph_format.left_indent = Cm(1.5)
            add_inline(p, m.group(3).strip())
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            i += 1
            continue

        # 普通段落（保留手动换行的行合并）
        p = doc.add_paragraph()
        add_inline(p, stripped)
        i += 1

    doc.save(docx_path)
    return docx_path


def main():
    if len(sys.argv) >= 2:
        md = Path(sys.argv[1])
        docx_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else md.with_suffix(".docx")
        docx_path.parent.mkdir(parents=True, exist_ok=True)
        convert(md, docx_path)
        print(f"OK  {md} -> {docx_path}")
        return

    base = Path(__file__).resolve().parent
    out_dir = base / "Word版"
    out_dir.mkdir(exist_ok=True)
    md_files = sorted(base.rglob("*.md"))
    for md in md_files:
        docx_path = out_dir / (md.stem + ".docx")
        convert(md, docx_path)
        print(f"OK  {md.relative_to(base)} -> {docx_path.name}")


if __name__ == "__main__":
    main()
