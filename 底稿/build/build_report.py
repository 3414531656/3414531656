# -*- coding: utf-8 -*-
"""把 6 个任务底稿 + 总表合并为一份 Word 汇总和一份 Excel 总表。"""
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(r"C:\Users\341453\Downloads\ln_marine_data")
BUILD = BASE / "build"
TASKS = [BASE / f"task{i}.md" for i in range(1, 7)]
SEP = re.compile(r"^:?-{2,}:?$")

# ---------- 读入并拆分 ----------
header_md = (BUILD / "总表.md").read_text(encoding="utf-8").strip()

bodies, n4s, n5s, titles = {}, {}, {}, {}
for n in range(1, 7):
    lines = TASKS[n - 1].read_text(encoding="utf-8").splitlines()
    i4 = next(i for i, l in enumerate(lines) if re.match(rf"^## {n}-4", l))
    i5 = next(i for i, l in enumerate(lines) if re.match(rf"^## {n}-5", l))
    body = lines[:i4]
    titles[n] = body[0].lstrip("# ").strip()
    dem = []
    for l in body[1:]:
        if l.startswith("### "):
            dem.append("#### " + l[4:])
        elif l.startswith("## "):
            dem.append("### " + l[3:])
        elif l.startswith("# "):
            dem.append("## " + l[2:])
        else:
            dem.append(l)
    bodies[n] = dem
    n4s[n] = lines[i4 + 1:i5]
    n5s[n] = lines[i5 + 1:]

# ---------- 合并 markdown ----------
out = [header_md, "", "## 二、六个任务逐任务明细"]
for n in range(1, 7):
    out += ["", "<!-- PAGEBREAK -->", "", "## " + titles[n], ""] + bodies[n]
out += ["", "<!-- PAGEBREAK -->", "", "## 三、无法获取清单"]
for n in range(1, 7):
    out += ["", f"### 任务{n}", ""] + n4s[n]
out += ["", "<!-- PAGEBREAK -->", "", "## 四、口径冲突与提醒"]
for n in range(1, 7):
    out += ["", f"### 任务{n}", ""] + n5s[n]
out += [
    "", "## 附录：检索与取证说明",
    "",
    "- 检索工具：政府网站检索 + web_fetch 打开 HTML 原始页；PDF 一律下载后抽取文本或逐页渲染核对，页码按 PDF 实际页序与页脚印刷页码双记。",
    "- URL 核验：本汇总所有 URL 均于 2026-10-31 实测可打开，逐条结果见各任务“URL 核验记录”表。",
    "- 本地存档：下载的 PDF/网页存档在「原件存档」文件夹，6 个任务的独立底稿在「底稿」文件夹。",
    "",
]
merged = "\n".join(out)
(BUILD / "汇总全文.md").write_text(merged, encoding="utf-8")
print("merged md chars:", len(merged))

# ---------- docx ----------
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = Cm(29.7), Cm(21)
for attr in ("left_margin", "right_margin"):
    setattr(sec, attr, Cm(1.2))
sec.top_margin, sec.bottom_margin = Cm(1.2), Cm(1.2)

def set_style_font(name, ascii_font, ea_font, size=None, bold=None):
    try:
        st = doc.styles[name]
    except KeyError:
        return
    if size:
        st.font.size = Pt(size)
    if bold is not None:
        st.font.bold = bold
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea_font)

set_style_font("Normal", "Times New Roman", "宋体", 10)
for lvl, sz in ((0, 20), (1, 16), (2, 13), (3, 11), (4, 10)):
    set_style_font(f"Title" if lvl == 0 else f"Heading {lvl}", "Arial", "黑体", sz, True)
for s in ("List Bullet", "List Number", "List Bullet 2"):
    set_style_font(s, "Times New Roman", "宋体", 9.5)

def add_runs(p, text, size=None, bold=False):
    text = text.replace("`", "").replace("<br>", " ")
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1（\2）", text)
    for k, seg in enumerate(re.split(r"\*\*", text)):
        if not seg:
            continue
        r = p.add_run(seg)
        r.bold = bold or (k % 2 == 1)
        if size:
            r.font.size = Pt(size)
        r.font.name = "Times New Roman"
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")

def add_md_table(rows):
    ncols = max(len(r) for r in rows)
    rows = [r + [""] * (ncols - len(r)) for r in rows]
    t = doc.add_table(rows=len(rows), cols=ncols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    lens = [min(max(max(len(r[j]) for r in rows), 8), 110) for j in range(ncols)]
    usable = 27.0
    widths = [max(1.25, usable * l / sum(lens)) for l in lens]
    s = sum(widths)
    if s > usable:
        widths = [w * usable / s for w in widths]
    for j, w in enumerate(widths):
        t.columns[j].width = Cm(w)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.width = Cm(widths[j])
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            add_runs(p, val, size=6.5, bold=(i == 0))
            if i == 0:
                shd = OxmlElement("w:shd")
                shd.set(qn("w:fill"), "DCE6F1")
                cell._tc.get_or_add_tcPr().append(shd)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)
    doc.add_paragraph()

lines = merged.splitlines()
i = 0
first_h1 = True
while i < len(lines):
    line = lines[i]
    s = line.strip()
    if s == "<!-- PAGEBREAK -->":
        doc.add_page_break()
        i += 1
        continue
    if s.startswith("|"):
        block = []
        while i < len(lines) and lines[i].strip().startswith("|"):
            cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
            if not all(SEP.match(c) for c in cells if c):
                block.append(cells)
            i += 1
        if block:
            add_md_table(block)
        continue
    m = re.match(r"^(#{1,4})\s+(.*)$", s)
    if m:
        lvl = len(m.group(1))
        text = m.group(2).strip()
        if lvl == 1 and first_h1:
            doc.add_heading(text, level=0)
            first_h1 = False
        else:
            doc.add_heading(text, level=lvl)
        i += 1
        continue
    if s.startswith(">"):
        buf = []
        while i < len(lines) and lines[i].strip().startswith(">"):
            buf.append(lines[i].strip().lstrip(">").strip())
            i += 1
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_before = Pt(0)
        add_runs(p, " ".join(x for x in buf if x), size=9)
        for r in p.runs:
            r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        continue
    m = re.match(r"^(\s*)[-*]\s+(.*)$", line)
    if m:
        indent = len(m.group(1))
        style = "List Bullet 2" if indent >= 2 else "List Bullet"
        p = doc.add_paragraph(style=style)
        add_runs(p, m.group(2), size=9.5)
        i += 1
        continue
    m = re.match(r"^\s*(\d+)\.\s+(.*)$", line)
    if m:
        # 用原文字面序号，避免跨任务自动编号串号
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.3)
        add_runs(p, f"{m.group(1)}. {m.group(2)}", size=10)
        i += 1
        continue
    if s in ("", "---", "***"):
        i += 1
        continue
    p = doc.add_paragraph()
    add_runs(p, s, size=10)
    i += 1

# 页脚页码
def add_page_field(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run("第 ")
    r = paragraph.add_run()
    for t, txt in (("begin", None), ("instr", "PAGE"), ("separate", None), ("text", "1"), ("end", None)):
        if t == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
            r._r.append(el)
            r = paragraph.add_run()
        elif t == "text":
            r.text = txt
            r = paragraph.add_run()
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), t)
            r._r.append(el)
            if t != "end":
                r = paragraph.add_run()
    paragraph.add_run(" 页")

add_page_field(sec.footer.paragraphs[0])
doc.core_properties.title = "辽宁海洋数据 · 六任务公开信息核验汇总"
doc.core_properties.author = "政府公开信息检索"
OUT_DOCX = BUILD / "辽宁海洋数据_六任务汇总.docx"
doc.save(OUT_DOCX)

chk = Document(OUT_DOCX)
print("docx paragraphs:", len(chk.paragraphs), "tables:", len(chk.tables))

# ---------- xlsx ----------
def parse_first_table(lines):
    rows = []
    started = False
    for l in lines:
        if l.strip().startswith("|"):
            cells = [c.strip() for c in l.strip().strip("|").split("|")]
            if all(SEP.match(c) for c in cells if c):
                started = True
                continue
            rows.append(cells)
            started = True
        elif started and rows:
            break
    return rows

total_tbl = parse_first_table(header_md.splitlines())
wb = Workbook()
ws = wb.active
ws.title = "总表"
head_fill = PatternFill("solid", fgColor="DCE6F1")
thin = Side(style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
for row in total_tbl:
    ws.append(row)
for c in ws[1]:
    c.font = Font(bold=True, size=10)
    c.fill = head_fill
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border
for r in ws.iter_rows(min_row=2):
    for c in r:
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.border = border
widths = [7, 34, 12, 30, 10, 10, 42, 20]
for j, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = "A2"

ws2 = wb.create_sheet("无法获取")
ws2.append(["任务号", "指标", "试过的检索式（至少3条）", "建议去哪个网站哪个栏目"])
for n in range(1, 7):
    tbl = parse_first_table(n4s[n])
    for row in tbl[1:]:
        row = (row + ["", "", ""])[:3]
        ws2.append([f"任务{n}"] + row)
ws3 = wb.create_sheet("口径冲突")
ws3.append(["任务号", "口径冲突与提醒"])
for n in range(1, 7):
    for l in n5s[n]:
        s = l.strip()
        if s.startswith("- "):
            ws3.append([f"任务{n}", s[2:].strip()])
for sheet in (ws2, ws3):
    for c in sheet[1]:
        c.font = Font(bold=True, size=10)
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    for r in sheet.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = border
for j, w in zip(range(1, 5), (9, 36, 80, 55)):
    ws2.column_dimensions[get_column_letter(j)].width = w
ws2.freeze_panes = "A2"
ws3.column_dimensions["A"].width = 9
ws3.column_dimensions["B"].width = 150
ws3.freeze_panes = "A2"
OUT_XLSX = BUILD / "辽宁海洋数据_总表.xlsx"
wb.save(OUT_XLSX)

wb2 = Workbook()
print("xlsx saved; sheets:", wb.sheetnames, "rows(总表):", ws.max_row, "rows(无法获取):", ws2.max_row, "rows(口径):", ws3.max_row)
