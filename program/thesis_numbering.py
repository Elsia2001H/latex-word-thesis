"""Map Pandoc chapter counters to actual numbered TeX chapters and appendices."""
import os
import re
from pathlib import Path


def key(text):
    return re.sub(r'\s+', '', text)


def chapter_mapping(document):
    source=Path(os.environ.get('THESIS_TEX',''))
    if not source.is_file():
        raise RuntimeError('THESIS_TEX is required for chapter-number mapping')
    text=source.read_text(encoding='utf-8-sig')
    text=re.sub(r'(?<!\\)%[^\n]*','',text)
    chapters=[]; body=0; appendix=None
    pattern=r'\\appendix\b|\\chapter(\*)?(?:\[[^\]]*\])?\{([^{}]*)\}'
    for match in re.finditer(pattern,text):
        if match.group(0)==r'\appendix':
            appendix=0
        elif not match.group(1):
            if appendix is None:
                body+=1; number=str(body)
            else:
                appendix+=1
                if appendix>26:
                    raise ValueError('More than 26 appendices require custom numbering')
                number=chr(64+appendix)
            chapters.append((key(match.group(2)),number))
    mapping={}; cursor=0
    for paragraph in document.paragraphs:
        if paragraph.style.name!='Heading 1':
            continue
        match=re.match(r'^Section\s+(\d+)\s+(.*)$',paragraph.text)
        if match and cursor<len(chapters) and key(match.group(2))==chapters[cursor][0]:
            mapping[match.group(1)]=chapters[cursor][1]; cursor+=1
    if cursor!=len(chapters):
        raise ValueError('TeX/Word chapter mapping incomplete; expand custom chapter macros or input files first')
    return mapping


def renumber(text,mapping,count=0):
    return re.sub(r'(?<![\d.])(\d+)\.(\d+)(?![\d.])',lambda m:mapping.get(m[1],m[1])+('-' if m[1] in mapping else '.')+m[2],text,count=count)
