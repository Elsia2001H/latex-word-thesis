"""Compose vector PDF panels with centred labels for both PDF and Word."""
from pathlib import Path
import fitz

ROOT = Path.cwd()


def compose(files, output, columns=2):
    width, panel_height, gap, label_height = 416.0, 140.0, 12.0, 22.0
    cell_width = (width - gap * (columns - 1)) / columns
    rows = (len(files) + columns - 1) // columns
    doc = fitz.open()
    page = doc.new_page(width=width, height=rows * (panel_height + label_height + gap) - gap)
    for i, name in enumerate(files):
        x = (i % columns) * (cell_width + gap)
        y = (i // columns) * (panel_height + label_height + gap)
        with fitz.open(ROOT / 'figures' / name) as source:
            page.show_pdf_page(fitz.Rect(x, y, x + cell_width, y + panel_height), source, 0, keep_proportion=True)
        page.insert_textbox(fitz.Rect(x, y + panel_height + 2, x + cell_width, y + panel_height + label_height), '(' + chr(97 + i) + ')', fontname='tiro', fontsize=11, align=1)
    doc.save(ROOT / 'figures' / output)
    doc.close()


if __name__ == '__main__':
    compose(['测试多图.pdf'] * 4, 'test-new-panels-abcd.pdf')
