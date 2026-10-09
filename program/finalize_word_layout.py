"""Final Word-only typography fixes applied after Pandoc conversion."""

from __future__ import annotations

import argparse
import os
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


def all_paragraphs(document):
    yield from document.paragraphs
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def force_black_hyperlink(hyperlink) -> None:
    for run_element in hyperlink.xpath(".//w:r"):
        r_pr = run_element.get_or_add_rPr()
        color = r_pr.find(qn("w:color"))
        if color is None:
            color = OxmlElement("w:color")
            r_pr.append(color)
        color.set(qn("w:val"), "000000")

        underline = r_pr.find(qn("w:u"))
        if underline is None:
            underline = OxmlElement("w:u")
            r_pr.append(underline)
        underline.set(qn("w:val"), "none")


def set_run_superscript(run_element, half_points: str = "24") -> None:
    """Format a citation run as a compact black superscript."""
    r_pr = run_element.get_or_add_rPr()

    vertical = r_pr.find(qn("w:vertAlign"))
    if vertical is None:
        vertical = OxmlElement("w:vertAlign")
        r_pr.append(vertical)
    vertical.set(qn("w:val"), "superscript")

    size = r_pr.find(qn("w:sz"))
    if size is None:
        size = OxmlElement("w:sz")
        r_pr.append(size)
    size.set(qn("w:val"), half_points)

    size_cs = r_pr.find(qn("w:szCs"))
    if size_cs is None:
        size_cs = OxmlElement("w:szCs")
        r_pr.append(size_cs)
    size_cs.set(qn("w:val"), half_points)

    color = r_pr.find(qn("w:color"))
    if color is None:
        color = OxmlElement("w:color")
        r_pr.append(color)
    color.set(qn("w:val"), "000000")


def fix_equation_reference(paragraph, hyperlink) -> None:
    text_nodes = hyperlink.xpath(".//w:t")
    link_text = "".join(node.text or "" for node in text_nodes).strip()
    full_text = "".join(
        node.text or "" for node in paragraph._p.xpath(".//w:t")
    )
    if not re.fullmatch(r"\(?\d+[\-‑]\d+\)?", link_text):
        return
    anchor = hyperlink.get(qn("w:anchor"), "").lower()
    if not anchor.startswith(("eq", "equation")):
        return
    number = link_text.strip("()").replace("‑", "-")
    if text_nodes:
        text_nodes[0].text = f"({number})"
        for extra_node in text_nodes[1:]:
            extra_node.text = ""


def set_border(container, edge: str, value: str, size: str = "8") -> None:
    """Set one OOXML border on a table or row property container."""
    borders_tag = "w:tblBorders" if container.tag == qn("w:tblPr") else "w:tcBorders"
    borders = container.find(qn(borders_tag))
    if borders is None:
        borders = OxmlElement(borders_tag)
        container.append(borders)
    border = borders.find(qn(f"w:{edge}"))
    if border is None:
        border = OxmlElement(f"w:{edge}")
        borders.append(border)
    border.set(qn("w:val"), value)
    border.set(qn("w:sz"), size)
    border.set(qn("w:space"), "0")
    border.set(qn("w:color"), "000000")


def format_three_line_table(table) -> None:
    """Apply a thesis-style three-line table without vertical rules."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for column in table.columns:
        column.width = Cm(14.66 / len(table.columns))
    table_pr = table._tbl.tblPr

    table_width = table_pr.find(qn("w:tblW"))
    if table_width is None:
        table_width = OxmlElement("w:tblW")
        table_pr.append(table_width)
    table_width.set(qn("w:type"), "dxa")
    table_width.set(qn("w:w"), "8311")

    # Remove every grid border first, then add the three required rules.
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        set_border(table_pr, edge, "nil", "0")
    set_border(table_pr, "top", "single", "12")
    set_border(table_pr, "bottom", "single", "12")

    column_count = max((len(row.cells) for row in table.rows), default=1)
    equal_width = max(1, 8311 // column_count)
    for row_index, row in enumerate(table.rows):
        row_pr = row._tr.get_or_add_trPr()
        row_pr.append(OxmlElement("w:cantSplit"))
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tc_pr = cell._tc.get_or_add_tcPr()
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                set_border(tc_pr, edge, "nil", "0")
            tc_width = tc_pr.find(qn("w:tcW"))
            if tc_width is None:
                tc_width = OxmlElement("w:tcW")
                tc_pr.append(tc_width)
            tc_width.set(qn("w:type"), "dxa")
            tc_width.set(qn("w:w"), str(equal_width))

            # Header cells provide the middle rule of the three-line table.
            if row_index == 0:
                set_border(tc_pr, "top", "single", "12")
                set_border(tc_pr, "bottom", "single", "4") # eighths of a point: 0.5 pt
            elif row_index == len(table.rows) - 1:
                set_border(tc_pr, "bottom", "single", "12")
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.first_line_indent = None
                paragraph.paragraph_format.line_spacing = 1.5
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.keep_with_next = True
                for run in paragraph.runs:
                    run.font.size = Pt(10.5)
                    if row_index == 0:
                        run.bold = True
    table.rows[-1].cells[-1].paragraphs[-1].paragraph_format.keep_with_next = False


def fix_document(path: Path) -> None:
    document = Document(path)
    # Pandoc may inject FirstParagraph with eastAsia=Times New Roman even
    # when Normal is SimSun. That makes mixed-script runs behave like long
    # Western words, causing premature wraps and huge justified spaces.
    from calibrate_reference_docx import set_east_asian_font
    for name in ("Body Text", "First Paragraph"):
        if name in document.styles:
            set_east_asian_font(document.styles[name], "宋体", "Times New Roman")
    source_tex = Path(os.environ.get("THESIS_TEX", str(path.parent / "thesis-template.tex")))
    bilingual = {}
    if source_tex.exists():
        for kind, chinese, english in re.findall(r"\\(Figure|Table)Caption\{([^{}]+)\}\{([^{}]+)\}", source_tex.read_text(encoding="utf-8-sig")):
            bilingual[chinese] = (kind, english)
    for paragraph in list(document.paragraphs):
        match = re.match(r"^([图表])\s*([1-3A][\-‑]\d+)\s*[:：]?\s*(.*)$", paragraph.text)
        if match and match[3].strip() in bilingual:
            kind, english = bilingual[match[3].strip()]
            number = match[2]
            paragraph.text = f"{match[1]} {number} {match[3].strip()}"
            paragraph.style = document.styles["Caption"]
            en = document.add_paragraph(style="Caption")
            en.add_run(("Fig." if kind == "Figure" else "Table ") + number + " " + english).bold = False
            en.paragraph_format.line_spacing = 1.5
            en.paragraph_format.space_before = Pt(0)
            en.paragraph_format.space_after = Pt(0)
            en.paragraph_format.first_line_indent = Pt(0)
            en.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.keep_with_next = True
            en.paragraph_format.keep_with_next = kind == "Table"
            paragraph._p.addnext(en._p)

    # Reassert A4 geometry after conversion.  This prevents Word from using a
    # Letter-sized section or a different right margin than the PDF template.
    for section in document.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.gutter = Cm(0)

    usable_width = min(
        section.page_width - section.left_margin - section.right_margin
        for section in document.sections
    )
    # Explicit upright control glyphs prevent the preceding italic variable
    # from leaking into scalable matrix delimiters during Word normalization.
    for control in document._element.xpath(".//m:ctrlPr"):
        rpr = control.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            control.append(rpr)
        upright = rpr.find(qn("w:i"))
        if upright is None:
            upright = OxmlElement("w:i")
            rpr.append(upright)
        upright.set(qn("w:val"), "0")
    # Normal-text math runs permit Times New Roman letters/digits in native
    # OMML. Structural symbols and extensible delimiters retain the math font.
    for math_run in document._element.xpath(".//m:r"):
        equation = next((node for node in math_run.iterancestors() if node.tag == qn("m:oMath")), None)
        # Word stretches matrix brackets incorrectly when their contents are
        # normal-text math. Keep matrix equations in the native math font.
        if equation is not None and next(equation.iter(qn("m:m")), None) is not None:
            continue
        text = "".join(node.text or "" for node in math_run.iter(qn("m:t")))
        # Operators/delimiters must remain math runs; normal-text conversion
        # of these glyphs can destroy Word's stretchy brackets and spacing.
        if not text or not all(c.isalnum() or c.isspace() for c in text):
            continue
        math_pr = math_run.find(qn("m:rPr"))
        if math_pr is None:
            math_pr = OxmlElement("m:rPr")
            math_run.insert(0, math_pr)
        was_normal = math_pr.find(qn("m:nor")) is not None
        style = math_pr.find(qn("m:sty"))
        style_value = style.get(qn("m:val")) if style is not None else "i"
        normal = math_pr.find(qn("m:nor"))
        if normal is None:
            normal = OxmlElement("m:nor")
            math_pr.append(normal)
        normal.set(qn("m:val"), "1")
        run_pr = math_run.find(qn("w:rPr"))
        if run_pr is None:
            run_pr = OxmlElement("w:rPr")
            math_run.insert(list(math_run).index(math_pr) + 1, run_pr)
        fonts = run_pr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            run_pr.insert(0, fonts)
        for attr in ("ascii", "hAnsi", "cs"):
            fonts.set(qn("w:" + attr), "Times New Roman")
        for attr in ("asciiTheme", "hAnsiTheme", "cstheme"):
            fonts.attrib.pop(qn("w:" + attr), None)
        for tag, enabled in (("i", not was_normal and style_value in {"i", "bi"} and any(c.isalpha() for c in text)), ("b", style_value in {"b", "bi"})):
            prop = run_pr.find(qn("w:" + tag))
            if prop is None:
                prop = OxmlElement("w:" + tag)
                run_pr.append(prop)
            prop.set(qn("w:val"), "1" if enabled else "0")
    for shape in document.inline_shapes:
        if shape.width > usable_width:
            ratio = usable_width / shape.width
            shape.width = int(shape.width * ratio)
            shape.height = int(shape.height * ratio)

    for style_name in ("Hyperlink", "FollowedHyperlink"):
        try:
            style = document.styles[style_name]
        except KeyError:
            continue
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.font.underline = False

    cover_markers = (
        "分类号",
        "密级",
        "博士生姓名",
        "导师姓名",
        "培养方向",
        "培养单位",
        "UDC",
        "Dissertation",
        "Candidate",
        "Supervisor",
    )
    for table in document.tables:
        table_text = "".join(cell.text for row in table.rows for cell in row.cells)
        if not any(marker in table_text for marker in cover_markers):
            format_three_line_table(table)

    in_chinese_abstract = False
    in_reference_list = False
    in_achievements = False
    for paragraph in all_paragraphs(document):
        normalized = re.sub(r"\s+", "", paragraph.text).upper()
        if normalized in {"攻读博士学位期间的科研成果", "攻读学位期间取得的研究成果"}:
            in_achievements = True
        elif paragraph.style and paragraph.style.name == "Heading 1":
            in_achievements = False
        if in_achievements and paragraph.style.name != "Heading 1":
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.left_indent = Pt(0)
            paragraph.paragraph_format.right_indent = Pt(0)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if normalized.startswith("第一章"):
            paragraph.paragraph_format.page_break_before = False
        if normalized in {"目录", "目錄"}:
            paragraph.style = document.styles["Title"]
            for run in paragraph.runs:
                run.font.name = "SimHei"
                run.font.size = Pt(16)
                run._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "SimHei")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            outline = OxmlElement("w:outlineLvl")
            outline.set(qn("w:val"), "9")
            paragraph._p.get_or_add_pPr().append(outline)
        if normalized == "摘要":
            in_chinese_abstract = True
        elif normalized == "ABSTRACT":
            in_chinese_abstract = False
        elif in_chinese_abstract:
            # Avoid Word stretching isolated English words such as
            # "Times New Roman" across a justified Chinese line.
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if normalized in {"参考文献", "REFERENCES"}:
            in_reference_list = True
        elif normalized in {
            "攻读博士学位期间的科研成果",
            "攻读学位期间取得的研究成果",
            "致谢",
            "ACKNOWLEDGEMENTS",
        }:
            in_reference_list = False

        if in_reference_list and re.match(r"^\[\d+\]", paragraph.text):
            # Original thesis bibliography: 12 pt, 1.5 lines, hanging numbers.
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.left_indent = Cm(0.8)
            paragraph.paragraph_format.first_line_indent = Cm(-0.8)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(0.8))
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for element in paragraph._p.xpath(".//w:r"):
                properties = element.get_or_add_rPr()
                for tag in ("sz", "szCs"):
                    value = properties.find(qn("w:" + tag))
                    if value is None:
                        value = OxmlElement("w:" + tag)
                        properties.append(value)
                    value.set(qn("w:val"), "24")

        if paragraph._p.xpath(".//w:drawing | .//w:pict"):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.space_before = Pt(18)
            paragraph.paragraph_format.space_after = Pt(6)
            paragraph.paragraph_format.keep_with_next = True

        if paragraph._p.xpath(".//m:oMathPara") or (
            paragraph._p.xpath(".//m:oMath") and paragraph._p.xpath(".//w:tab")
        ):
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.first_line_indent = Pt(0)
            # Do not snap tall matrices, fractions or integrals to the body grid.
            p_pr = paragraph._p.get_or_add_pPr()
            snap = p_pr.find(qn("w:snapToGrid"))
            if snap is None:
                snap = OxmlElement("w:snapToGrid")
                p_pr.append(snap)
            snap.set(qn("w:val"), "0")
            paragraph.paragraph_format.keep_together = True

        if re.match(r"^[图表]\s*[1-3A][\-‑]\d+", paragraph.text):
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.space_before = Pt(12 if paragraph.text.startswith("表") else 0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.keep_with_next = paragraph.text.startswith("表")

        for hyperlink in paragraph._p.xpath(".//w:hyperlink"):
            force_black_hyperlink(hyperlink)
            fix_equation_reference(paragraph, hyperlink)

        # CSL/Pandoc normally places each citation marker in its own run or
        # hyperlink.  Format only bracketed numeric markers, so the numbered
        # bibliography entries at the end of the thesis remain unchanged.
        if not in_reference_list:
            for run_element in paragraph._p.xpath(".//w:r"):
                run_text = "".join(
                    node.text or "" for node in run_element.xpath(".//w:t")
                ).strip()
                if re.fullmatch(r"\[[0-9,，;；\-–—\s]+\]", run_text):
                    set_run_superscript(run_element)

        # A chapter-number token must not split at the end of a Word line.
        for text_node in paragraph._p.xpath(".//w:t"):
            if text_node.text:
                text_node.text = re.sub(
                    r"(?<=\d)-(?=\d)", "‑", text_node.text
                )

    # Space around whole figure/table blocks, not between bilingual captions.
    for paragraph in document.paragraphs:
        if re.match(r"^Fig\.\s*\d+[\-‑]\d+", paragraph.text):
            paragraph.paragraph_format.space_after = Pt(12)
        previous = paragraph._p.getprevious()
        if previous is not None and previous.tag == qn("w:tbl"):
            if not paragraph._p.xpath(".//w:drawing | .//w:pict"):
                paragraph.paragraph_format.space_before = Pt(12)

    native_layout = Path(os.environ.get("THESIS_FRONT_LAYOUT", str(path.parent / "._thesis_build" / "front-layout.json")))
    if native_layout.exists():
        from native_frontmatter import rebuild
        rebuild(document, native_layout)
    cover_image = path.parent / "._thesis_build" / "cover-content.png"
    if not native_layout.exists() and cover_image.exists():
        english = next((p for p in document.paragraphs if p.text == "A Dissertation Submitted to"), None)
        if english is not None:
            body = document._element.body
            for node in list(body):
                if node is english._p:
                    break
                body.remove(node)
            cover = document.add_paragraph()
            cover.paragraph_format.first_line_indent = Pt(0)
            cover.paragraph_format.line_spacing = 1.0
            cover.paragraph_format.space_before = Pt(0)
            cover.paragraph_format.space_after = Pt(0)
            cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cover.add_run().add_picture(str(cover_image), width=Cm(14.66))
            english._p.addprevious(cover._p)
            english.paragraph_format.page_break_before = True
    document.save(path)


def format_cached_toc(path: Path) -> None:
    """Format real TOC field results after Word has regenerated them."""
    document = Document(path)
    for paragraph in document.paragraphs:
        if paragraph.style.style_id != "TOC1":
            continue
        paragraph.style.font.bold = True
        after_tab = False
        for run in paragraph._p.xpath(".//w:r"):
            if run.xpath("./w:tab"):
                bold_value = "1" # the tab's font also controls its leader
                after_tab = True
            else:
                bold_value = "0" if after_tab else "1"
            rpr = run.get_or_add_rPr()
            for tag in ("w:b", "w:bCs"):
                node = rpr.find(qn(tag))
                if node is None:
                    node = OxmlElement(tag)
                    rpr.append(node)
                node.set(qn("w:val"), bold_value)
    document.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", nargs="?", default="thesis-template.docx")
    args = parser.parse_args()
    path = Path(args.docx).resolve()
    fix_document(path)
    print(f"Finalized Word layout: {path}")


if __name__ == "__main__":
    main()
