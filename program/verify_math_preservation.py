"""Check that replacement changes only math XML and its package dependencies."""
import copy
import json
import sys
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
from mathtype_convert import NS, files_of
from mathtype_cfb import read_streams


def stripped(data, native):
    xml=E.fromstring(data)
    nodes=xml.xpath('//m:oMathPara | //m:oMath[not(ancestor::m:oMathPara)]',namespaces=NS) if native else xml.xpath('//w:r[w:object/o:OLEObject[@ProgID="Equation.DSMT4"]]',namespaces=NS)
    for node in nodes:
        node.getparent().remove(node)
    return E.tostring(xml,method='c14n')


def main(path):
    old=files_of(path.with_suffix('.native.docx')); new=files_of(path)
    allowed={'word/document.xml','word/_rels/document.xml.rels','[Content_Types].xml'}
    assert all(new[k]==v for k,v in old.items() if k not in allowed),'Non-equation package parts changed'
    assert stripped(old['word/document.xml'],True)==stripped(new['word/document.xml'],False),'Non-equation document XML changed'
    xml=E.fromstring(new['word/document.xml'])
    rels={r.get('Id'):r.get('Target') for r in E.fromstring(new['word/_rels/document.xml.rels'])}
    objects=xml.xpath('//o:OLEObject[@ProgID="Equation.DSMT4"]',namespaces=NS)
    for obj in objects:
        import posixpath,struct
        streams,_=read_streams(new[posixpath.normpath('word/'+rels[obj.get('{'+NS['r']+'}id')])])
        native=streams['Equation Native']; assert len(native)==28+struct.unpack_from('<I',native,8)[0]
    print('PASS unchanged layout/text/numbering/references/package parts; objects=%d'%len(objects))


if __name__=='__main__':
    main(Path(sys.argv[1]))
