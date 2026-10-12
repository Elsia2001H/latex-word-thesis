"""Remove conversion-only spaces, preserving equation objects and references."""
import re
import sys
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'o': 'urn:schemas-microsoft-com:office:office'}
W = '{' + NS['w'] + '}'
CJK = re.compile(r'[\u3400-\u9fff\u3000-\u303f\uff00-\uffef]')


def clean_tree(root):
    removed = 0
    for paragraph in root.xpath('//w:tc//w:p[.//m:oMath or .//o:OLEObject]', namespaces=NS):
        # Pandoc's empty table-math tail becomes a visible box with Show All.
        # Limit removal to text-only zero-width-space runs beside equations.
        for node in paragraph.xpath('./w:r/w:t', namespaces=NS):
            if node.text and set(node.text) == {'\u200b'}:
                removed += len(node.text)
                node.getparent().remove(node)
    for link in root.xpath('//w:hyperlink[@w:anchor]', namespaces=NS):
        label = ''.join(link.xpath('.//w:t/text()', namespaces=NS))
        if not re.fullmatch(r'\(\d+[.\-\u2011]\d+\)|\([A-Z][\-\u2011]\d+\)', label):
            continue
        parent = link.getparent()
        siblings = list(parent); index = siblings.index(link)
        for direction in (-1, 1):
            j = index + direction
            spaces = []
            # Only the adjacent NBSP separator belongs to this reference.
            while 0 <= j < len(siblings):
                item = siblings[j]
                if item.tag != W + 'r' or item.find(W + 't') is None:
                    break
                node = item.find(W + 't')
                if node.text and set(node.text) == {'\u00a0'}:
                    spaces.append(node); j += direction
                    continue
                text = node.text or ''
                edge = text[-1:] if direction == -1 else text[:1]
                if spaces and CJK.fullmatch(edge):
                    for space in spaces:
                        removed += len(space.text)
                        space.getparent().remove(space)
                break
    return removed


def clean_file(path):
    with ZipFile(path) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    output = []; removed = 0
    for info, data in entries:
        if info.filename == 'word/document.xml':
            root = E.fromstring(data)
            removed = clean_tree(root)
            if removed:
                data = E.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        output.append((info, data))
    if removed:
        temporary = path.with_suffix('.markers.tmp.docx')
        with ZipFile(temporary, 'w', ZIP_DEFLATED) as archive:
            for info, data in output:
                archive.writestr(info, data)
        temporary.replace(path)
    print('Cleaned equation editing markers:', removed)
    return removed


if __name__ == '__main__':
    clean_file(Path(sys.argv[1]))
