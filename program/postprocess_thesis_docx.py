#!/usr/bin/env python3
"""Apply thesis page geometry, semantic breaks, headers, footers, and TOC to DOCX."""

from __future__ import annotations

import re
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.shared import RGBColor
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def normalized(text):
    return re.sub(r"\s+", "", text or "").upper()


def add_field(paragraph, instruction, placeholder=""):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instr, separate, text, end):
        run._r.append(node)
    return run


def set_page_number_format(section, fmt, start=None):
    sect = section._sectPr
    node = sect.find(qn("w:pgNumType"))
    if node is None:
        node = OxmlElement("w:pgNumType")
        sect.append(node)
    node.set(qn("w:fmt"), fmt)
    if start is None:
        node.attrib.pop(qn("w:start"), None)
    else:
        node.set(qn("w:start"), str(start))


def configure_geometry(section):
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)
    section.header_distance = Cm(1.50)
    section.footer_distance = Cm(1.75)


def clear_story(story):
    for p in story.paragraphs:
        p._element.getparent().remove(p._element)
    story.add_paragraph()


def configure_header_footer(section, mode, even_title=""):
    section.header.is_linked_to_previous = False
    section.even_page_header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    section.even_page_footer.is_linked_to_previous = False
    for story in (section.header, section.even_page_header, section.footer, section.even_page_footer):
        clear_story(story)

    if mode == "blank":
        return

    odd = section.header.paragraphs[0]
    odd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    odd.add_run("博士学位论文")

    even = section.even_page_header.paragraphs[0]
    even.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Read the heading on the current page, rather than the last heading
    # assigned to this whole section (which can be a later back-matter title).
    add_field(even, ' STYLEREF 1 \\* MERGEFORMAT ', even_title or "博士学位论文")
    for paragraph in (odd, even):
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.name = "KaiTi"
            run.font.size = Pt(10.5)
            run.font.bold = False
            run._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "KaiTi")
        borders = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        for key, value in {"val": "single", "sz": "6", "space": "1", "color": "000000"}.items():
            bottom.set(qn("w:" + key), value)
        borders.append(bottom)
        paragraph._p.get_or_add_pPr().append(borders)

    for footer in (section.footer, section.even_page_footer):
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_field(p, " PAGE \\* MERGEFORMAT ", "1")


def move_section_break(break_paragraph, target_paragraph):
    sect = break_paragraph._p.pPr.sectPr
    break_paragraph._p.pPr.remove(sect)
    ppr = target_paragraph._p.get_or_add_pPr()
    old = ppr.find(qn("w:sectPr"))
    if old is not None:
        ppr.remove(old)
    ppr.append(sect)
    parent = break_paragraph._p.getparent()
    parent.remove(break_paragraph._p)


def section_index_for(doc, target):
    count = 0
    for child in doc.element.body:
        if child is target._p:
            return count
        if child.tag == qn("w:p") and child.find("w:pPr/w:sectPr", child.nsmap) is not None:
            count += 1
    return count


def previous_paragraph(doc, target):
    idx = next(i for i, p in enumerate(doc.paragraphs) if p._p is target._p)
    if idx == 0:
        blank = OxmlElement("w:p")
        target._p.addprevious(blank)
        return doc.paragraphs[0]
    return doc.paragraphs[idx - 1]


def find_paragraph(doc, candidates, style_prefix=None):
    wanted = {normalized(x) for x in candidates}
    for p in doc.paragraphs:
        if style_prefix and not (p.style and p.style.name.startswith(style_prefix)):
            continue
        if normalized(p.text) in wanted:
            return p
    return None


def insert_toc_after_heading(heading):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    p.append(ppr)
    heading._p.addnext(p)
    from docx.text.paragraph import Paragraph
    toc = Paragraph(p, heading._parent)
    add_field(toc, ' TOC \\o "1-3" \\h \\z \\u ', "目录将在 Word 中自动更新")


def strip_pandoc_number(text):
    return re.sub(r"^Section\s+\d+(?:\.\d+)*\s*", "", text.strip(), flags=re.I)


def normalize_headings(doc):
    chinese = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    special = {
        "学位论文独创性声明": "学位论文独创性声明",
        "学位论文版权使用授权书": "学位论文版权使用授权书",
        "摘要": "摘  要",
        "ABSTRACT": "ABSTRACT",
        "攻读博士学位期间的科研成果": "攻读博士学位期间的科研成果",
        "致谢": "致  谢",
    }
    chapter = section = subsection = 0
    for p in doc.paragraphs:
        if not p.style or not p.style.name.startswith("Heading"):
            continue
        base = strip_pandoc_number(p.text)
        key = normalized(base)
        if p.style.name == "Heading 1":
            section = subsection = 0
            if key in {normalized(k) for k in special}:
                for source, replacement in special.items():
                    if key == normalized(source):
                        p.text = replacement
                        break
            elif key == normalized("补充材料"):
                p.text = "附录 A  补充材料"
            else:
                chapter += 1
                label = chinese[chapter - 1] if chapter <= len(chinese) else str(chapter)
                p.text = f"第{label}章  {base}"
        elif p.style.name == "Heading 2" and chapter:
            section += 1
            subsection = 0
            p.text = f"{chapter}.{section}  {base}"
        elif p.style.name == "Heading 3" and chapter:
            subsection += 1
            p.text = f"{chapter}.{section}.{subsection}  {base}"


def insert_front_lists(doc, after_paragraph):
    specs = [
        ("目  录", ' TOC \\o "1-3" \\h \\z \\u ', "目录将在 Word 中自动更新"),
    ]
    anchor = after_paragraph._p
    for title, instruction, placeholder in specs:
        heading = doc.add_paragraph(title, style="Heading 1")
        field = doc.add_paragraph()
        add_field(field, instruction, placeholder)
        heading_node = heading._p
        field_node = field._p
        heading_node.getparent().remove(heading_node)
        field_node.getparent().remove(field_node)
        anchor.addnext(heading_node)
        heading_node.addnext(field_node)
        anchor = field_node


def move_bibliography(doc):
    entries = [p for p in doc.paragraphs if p.style and p.style.name == "Bibliography"]
    target = find_paragraph(doc, ["攻读博士学位期间的科研成果"], "Heading")
    if not entries or target is None:
        return
    heading = doc.add_paragraph("参考文献", style="Heading 1")
    hnode = heading._p
    hnode.getparent().remove(hnode)
    target._p.addprevious(hnode)
    for entry in entries:
        node = entry._p
        node.getparent().remove(node)
        target._p.addprevious(node)


def remove_all_content_before(doc, target):
    body = doc.element.body
    for child in list(body):
        if child is target._p:
            break
        body.remove(child)


def set_run_font(run, size, bold=False, east_asia="SimSun"):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east_asia)


def set_paragraph_text(p, text, size=12, bold=False, alignment=WD_ALIGN_PARAGRAPH.CENTER, east_asia="SimSun"):
    p.clear()
    p.alignment = alignment
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_run_font(run, size, bold, east_asia)
    return p


def borderless(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "nil")


def move_before(nodes, target):
    for node in nodes:
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
        target._p.addprevious(node)


def rebuild_cover_pages(doc):
    declaration = find_paragraph(doc, ["学位论文独创性声明"], "Heading")
    if declaration is None:
        return
    paras = doc.paragraphs
    front = paras[:10]
    source_table = doc.tables[0] if doc.tables else None
    class_text = front[0].text if front else ""
    match1 = re.search(r"分类号：\s*(.*?)密级：\s*(.*)", class_text.splitlines()[0] if class_text.splitlines() else "")
    lines = class_text.splitlines()
    match2 = re.search(r"U\s*D\s*C：\s*(.*?)编号：\s*(.*)", lines[1] if len(lines) > 1 else "")
    classification, secret = (match1.group(1), match1.group(2)) if match1 else ("", "")
    udc, number = (match2.group(1), match2.group(2)) if match2 else ("", "")
    degree = front[1].text.strip() if len(front) > 1 else "博士学位论文"
    title_cn = front[2].text.strip() if len(front) > 2 else "论文题目"
    date_cn = front[4].text.strip() if len(front) > 4 else ""
    en_intro = front[5].text.splitlines() if len(front) > 5 else []
    school = en_intro[-1] if len(en_intro) > 1 else ""
    title_en = front[6].text.strip() if len(front) > 6 else ""
    by_lines = front[7].text.splitlines() if len(front) > 7 else []
    author = by_lines[-1] if len(by_lines) > 1 else ""
    sup_lines = front[8].text.splitlines() if len(front) > 8 else []
    supervisor = sup_lines[-1] if len(sup_lines) > 1 else ""
    date_en = front[9].text.strip() if len(front) > 9 else ""
    info = []
    if source_table is not None:
        info = [[cell.text.strip() for cell in row.cells] for row in source_table.rows]

    remove_all_content_before(doc, declaration)
    nodes = []
    top = doc.add_table(rows=2, cols=2)
    top.alignment = WD_TABLE_ALIGNMENT.CENTER
    top.autofit = False
    values = [(f"分类号：{classification}", f"密级：{secret}"), (f"U D C：{udc}", f"编号：{number}")]
    for r, row_values in enumerate(values):
        for c, value in enumerate(row_values):
            cell = top.cell(r, c)
            cell.width = Cm(7.2)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_paragraph_text(cell.paragraphs[0], value, 12, alignment=WD_ALIGN_PARAGRAPH.LEFT if c == 0 else WD_ALIGN_PARAGRAPH.RIGHT)
    borderless(top)
    nodes.append(top._tbl)

    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(72); set_paragraph_text(p, degree, 26, True, east_asia="SimHei"); nodes.append(p._p)
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(48); set_paragraph_text(p, title_cn, 26, True, east_asia="SimHei"); nodes.append(p._p)
    info_table = doc.add_table(rows=max(len(info), 1), cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False
    for r in range(max(len(info), 1)):
        row_values = info[r] if r < len(info) else ["", ""]
        for c in range(2):
            cell = info_table.cell(r, c)
            cell.width = Cm(5.2 if c == 0 else 7.4)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_paragraph_text(cell.paragraphs[0], row_values[c], 14, alignment=WD_ALIGN_PARAGRAPH.RIGHT if c == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            cell.paragraphs[0].paragraph_format.space_after = Pt(8)
    borderless(info_table)
    nodes.append(info_table._tbl)
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(70); set_paragraph_text(p, date_cn, 14); p.add_run().add_break(WD_BREAK.PAGE); nodes.append(p._p)

    for text_value, size, bold, before in (
        ("A Dissertation Submitted to", 16, False, 30),
        (school, 16, False, 8),
        (title_en, 20, True, 70),
        ("BY", 16, False, 90),
        (author, 16, False, 12),
        ("Supervised by", 16, False, 55),
        (supervisor, 16, False, 12),
        (date_en, 14, False, 80),
    ):
        p = doc.add_paragraph(); set_paragraph_text(p, text_value, size, bold); p.paragraph_format.space_before = Pt(before); nodes.append(p._p)
    p.add_run().add_break(WD_BREAK.PAGE)
    move_before(nodes, declaration)

    # Declarations are display headings but must not enter the main TOC.
    for title in ("学位论文独创性声明", "学位论文版权使用授权书"):
        heading = find_paragraph(doc, [title], "Heading")
        if heading is not None:
            heading.style = doc.styles["Normal"]
            set_paragraph_text(heading, title, 16, True, east_asia="SimHei")
            heading.paragraph_format.page_break_before = title != "学位论文独创性声明"
            heading.paragraph_format.space_after = Pt(24)


def replace_visible_numbers(doc):
    # Pandoc counts the four unnumbered front-matter headings as chapters.
    # Rewrite every chapter-number token, including math and hyperlink runs.
    for node in doc._element.body.xpath(".//w:t | .//m:t"):
        if node.text:
            node.text = re.sub(r"(?<![\d.])([5-8])\.(\d+)(?![\d.])", lambda m: ("A" if m[1] == "8" else str(int(m[1]) - 4)) + "-" + m[2], node.text)
            node.text = re.sub(r"Figure ([1-3A]-\d+):?", r"图 \1 ", node.text)
            node.text = re.sub(r"Table ([1-3A]-\d+):?", r"表 \1 ", node.text)
    return
    replacements = (
        ("Figure 5.1:", "图 1-1 "),
        ("Table 6.1:", "表 2-1 "),
        ("Figure 5.1", "图 1-1"),
        ("Table 6.1", "表 2-1"),
        ("5.1", "1-1"),
        ("6.1", "2-1"),
    )
    for p in doc.paragraphs:
        for node in p._p.xpath(".//w:t | .//m:t"):
            if node.text:
                for old, new in replacements:
                    node.text = node.text.replace(old, new)


def format_word_tables(doc):
    for index, table in enumerate(doc.tables):
        if index == 0:
            continue
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = [Cm(4.0), Cm(3.0), Cm(3.0)] if len(table.columns) == 3 else [Cm(5.0)] * len(table.columns)
        for i, width in enumerate(widths):
            table.columns[i].width = width
            for cell in table.columns[i].cells:
                cell.width = width
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.first_line_indent = Pt(0)


def section_heading_map(doc):
    mapping = {}
    current = ""
    for p in doc.paragraphs:
        if p.style and p.style.name == "Heading 1":
            current = p.text.strip()
        if current:
            mapping[section_index_for(doc, p)] = current
    # Carry the latest heading into sections without a new Heading 1.
    result, latest = {}, ""
    for index in range(len(doc.sections)):
        if index in mapping:
            latest = mapping[index]
        result[index] = latest
    return result


def apply(input_docx: Path, output_docx: Path):
    doc = Document(str(input_docx))

    doc.styles["Heading 1"].paragraph_format.page_break_before = True
    normalize_headings(doc)
    move_bibliography(doc)
    rebuild_cover_pages(doc)
    replace_visible_numbers(doc)
    format_word_tables(doc)

    abstract = find_paragraph(doc, ["摘要", "摘 要", "摘  要", "摘　要"], "Heading")
    abstract_en = find_paragraph(doc, ["ABSTRACT"], "Heading")
    first_chapter = next(
        (p for p in doc.paragraphs if p.style and p.style.name == "Heading 1" and p.text.strip().startswith("第一章")),
        None,
    )
    if abstract_en is not None and first_chapter is not None:
        en_idx = next(i for i, p in enumerate(doc.paragraphs) if p._p is abstract_en._p)
        main_idx = next(i for i, p in enumerate(doc.paragraphs) if p._p is first_chapter._p)
        between = doc.paragraphs[en_idx + 1:main_idx]
        anchor = between[-1] if between else abstract_en
        insert_front_lists(doc, anchor)

    if abstract is None or first_chapter is None:
        raise RuntimeError("Could not locate abstract or first chapter heading for section breaks")

    sec2_break = doc.add_section(WD_SECTION.NEW_PAGE)
    sec2_paragraph = doc.paragraphs[-1]
    sec3_break = doc.add_section(WD_SECTION.NEW_PAGE)
    sec3_paragraph = doc.paragraphs[-1]
    doc.settings.odd_and_even_pages_header_footer = True

    move_section_break(sec2_paragraph, previous_paragraph(doc, abstract))
    move_section_break(sec3_paragraph, previous_paragraph(doc, first_chapter))

    abstract_section = section_index_for(doc, abstract)
    main_section = section_index_for(doc, first_chapter)
    titles = section_heading_map(doc)
    for index, section in enumerate(doc.sections):
        configure_geometry(section)
        if index < abstract_section:
            configure_header_footer(section, "blank")
        elif index < main_section:
            configure_header_footer(section, "front", titles.get(index, ""))
            set_page_number_format(section, "lowerRoman", 1 if index == abstract_section else None)
        else:
            configure_header_footer(section, "main", titles.get(index, ""))
            set_page_number_format(section, "decimal", 1 if index == main_section else None)

    settings = doc.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")

    # Avoid stretched Latin words and over-justified placeholder messages.
    main_chapter = find_paragraph(doc, ["第1章", "第 1 章"], "Heading")
    if main_chapter is not None:
        # The main-text section break already starts a new page.  Keeping the
        # inherited Heading 1 page break creates an extra blank page in Word.
        main_chapter.paragraph_format.page_break_before = False

    chinese = find_paragraph(doc, ["摘  要", "摘要"], "Heading")
    english = find_paragraph(doc, ["ABSTRACT"], "Heading")
    if chinese is not None and english is not None:
        start = next(i for i, p in enumerate(doc.paragraphs) if p._p is chinese._p)
        end = next(i for i, p in enumerate(doc.paragraphs) if p._p is english._p)
        for paragraph in doc.paragraphs[start + 1 : end]:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if english is not None:
        start = next(i for i, p in enumerate(doc.paragraphs) if p._p is english._p)
        for p in doc.paragraphs[start + 1:]:
            if p.style and p.style.name == "Heading 1":
                break
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Pt(0)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: postprocess_thesis_docx.py INPUT.docx OUTPUT.docx")
    apply(Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve())
