"""Calibrate the Pandoc reference DOCX to the thesis PDF page geometry."""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


PAGE_WIDTH = Cm(21.0)
PAGE_HEIGHT = Cm(29.7)
LEFT_MARGIN = Cm(3.17)
RIGHT_MARGIN = Cm(3.17)
TOP_MARGIN = Cm(2.54)
BOTTOM_MARGIN = Cm(2.54)


def set_east_asian_font(style, chinese: str, western: str) -> None:
    style.font.name = western
    r_pr = style.element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), western)
    r_fonts.set(qn("w:hAnsi"), western)
    r_fonts.set(qn("w:eastAsia"), chinese)


def calibrate(path: Path) -> None:
    document = Document(path)

    for section in document.sections:
        section.page_width = PAGE_WIDTH
        section.page_height = PAGE_HEIGHT
        section.left_margin = LEFT_MARGIN
        section.right_margin = RIGHT_MARGIN
        section.top_margin = TOP_MARGIN
        section.bottom_margin = BOTTOM_MARGIN
        section.header_distance = Cm(1.5)
        section.footer_distance = Cm(1.5)
        section.gutter = Cm(0)

    normal = document.styles["Normal"]
    set_east_asian_font(normal, "宋体", "Times New Roman")
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.first_line_indent = Pt(24)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.widow_control = True

    # These styles override Normal in Pandoc-generated chapter/abstract text.
    for name in ("Body Text", "First Paragraph"):
        if name in document.styles:
            set_east_asian_font(document.styles[name], "宋体", "Times New Roman")

    heading_specs = {
        "Title": ("黑体", "Times New Roman", 22, WD_ALIGN_PARAGRAPH.CENTER),
        "Heading 1": ("黑体", "Times New Roman", 16, WD_ALIGN_PARAGRAPH.CENTER),
        "Heading 2": ("黑体", "Times New Roman", 14, WD_ALIGN_PARAGRAPH.LEFT),
        "Heading 3": ("黑体", "Times New Roman", 12, WD_ALIGN_PARAGRAPH.LEFT),
        "Caption": ("宋体", "Times New Roman", 10.5, WD_ALIGN_PARAGRAPH.CENTER),
    }
    for style_name, (cn_font, en_font, size, alignment) in heading_specs.items():
        try:
            style = document.styles[style_name]
        except KeyError:
            continue
        set_east_asian_font(style, cn_font, en_font)
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.alignment = alignment
        style.paragraph_format.keep_with_next = True

    for style_name in ("Hyperlink", "FollowedHyperlink"):
        try:
            style = document.styles[style_name]
        except KeyError:
            continue
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.font.underline = False

    document.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", nargs="?", default="thesis-reference.docx")
    args = parser.parse_args()
    path = Path(args.docx).resolve()
    calibrate(path)
    print(f"Calibrated reference DOCX: {path}")


if __name__ == "__main__":
    main()
