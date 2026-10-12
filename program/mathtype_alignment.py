"""Preserve a simple aligned environment as a MathType relation-aligned pile.

Record layout follows Wiris MTEF v5. Only a standalone, two-column aligned
environment is changed; matrices, cases and multiple alignment pairs stay intact.
"""
import re
import struct


class Reader:
    def __init__(self, data):
        self.data = data; self.pos = 0; self.matrices = []

    def byte(self):
        value = self.data[self.pos]; self.pos += 1; return value

    def skip(self, count):
        self.pos += count
        if self.pos > len(self.data): raise ValueError('Truncated MTEF')

    def integer(self):
        value = self.byte()
        if value == 255:
            value = struct.unpack_from('<H', self.data, self.pos)[0]; self.skip(2)
        return value

    def string(self):
        self.pos = self.data.index(0, self.pos) + 1

    def dimensions(self):
        count = self.byte(); terminated = 0
        while terminated < count:
            value = self.byte()
            terminated += (value >> 4 == 15) + (value & 15 == 15)

    def records(self, depth=0):
        if depth > 128: raise ValueError('MTEF nesting limit')
        nodes = []
        while self.data[self.pos] != 0:
            nodes.append(self.record(depth))
        self.skip(1)
        return nodes

    def record(self, depth):
        start = self.pos; tag = self.byte(); content_start = None; children = []
        if tag in (1,2,3,4,5,6):
            options = self.byte()
            if options & 8:
                a, b = self.byte(), self.byte()
                if (a,b) == (128,128): self.skip(4)
            if tag == 1:
                if options & 4: self.skip(2)
                if options & 2: self.record(depth+1)
                content_start = self.pos
                if not options & 1: children = self.records(depth+1)
            elif tag == 2:
                self.integer()
                if not options & 32: self.skip(2)
                if options & 4: self.skip(1)
                if options & 16: self.skip(2)
                if options & 1: self.records(depth+1)
            elif tag == 3:
                self.skip(1)
                if self.byte() & 128: self.skip(1)
                self.skip(1); children = self.records(depth+1)
            elif tag == 4:
                self.skip(2)
                if options & 2: self.record(depth+1)
                children = self.records(depth+1)
            elif tag == 5:
                self.skip(3); rows, cols = self.byte(), self.byte()
                self.skip((rows+4)//4 + (cols+4)//4)
                children = self.records(depth+1)
                self.matrices.append((start,self.pos,rows,cols,children,depth))
            else: self.skip(1)
        elif tag == 7:
            self.skip(self.byte()*3)
        elif tag == 8:
            self.integer(); self.skip(1)
        elif tag == 9:
            # General SIZE encodings vary; never guess a record boundary.
            raise ValueError('General MTEF SIZE inside aligned requires explicit support')
        elif 10 <= tag <= 14: pass
        elif tag == 15: self.integer()
        elif tag == 16:
            options = self.byte(); self.skip(8 if options & 1 else 6)
            if options & 4: self.string()
        elif tag == 17:
            self.integer(); self.string()
        elif tag == 18:
            self.skip(1); self.dimensions(); self.dimensions()
            for _ in range(self.byte()):
                if self.integer(): self.skip(1)
        elif tag == 19: self.string()
        elif tag >= 100: self.skip(self.integer())
        else: raise ValueError('Unsupported MTEF record %d' % tag)
        return (start,self.pos,tag,content_start,children)


def align_native(native, tex):
    # Do not guess alignment for explicit arrays or nested/multiple environments.
    source = re.sub(r'\{\\qquad\([^()]+\)\}\s*$', '', tex).strip()
    source = re.sub(r'\\(?:begin|end)\{equation\*?\}', '', source)
    source = re.sub(r'\\(?:label|tag)\{[^{}]*\}', '', source).strip()
    match = re.fullmatch(r'\\begin\{aligned\}(.*?)\\end\{aligned\}', source, re.S)
    if not match: return native
    body = match[1]
    rows = re.split(r'\\\\(?:\[[^]]*\])?', body)
    rows = [row for row in rows if row.strip()]
    if len(rows)<2 or any(row.count('&') != 1 for row in rows): return native
    size = struct.unpack_from('<I',native,8)[0]; mtef = native[28:28+size]
    if mtef[:2] != b'\x05\x01': raise ValueError('Expected MTEF v5')
    reader=Reader(mtef); reader.pos=5; reader.string(); reader.skip(1)
    reader.records()
    candidates=[matrix for matrix in reader.matrices if matrix[2:4] == (len(rows),2)]
    if not candidates: raise ValueError('Aligned MathType matrix not found')
    shallow=min(matrix[5] for matrix in candidates)
    candidates=[matrix for matrix in candidates if matrix[5]==shallow]
    if len(candidates)!=1: raise ValueError('Ambiguous aligned MathType matrix')
    start,end,_,_,cells,_=candidates[0]
    lines=[cell for cell in cells if cell[2]==1]
    if len(lines)!=len(rows)*2: raise ValueError('Aligned matrix cell count differs')
    # PILE halign=4 aligns each full row at its relation symbol, unlike array
    # column alignment, which MathType 7 cannot independently preserve.
    pile=bytearray(b'\x04\x00\x04\x04')
    for left,right in zip(lines[::2],lines[1::2]):
        pile += b'\x01\x00'
        for line in (left,right):
            if mtef[line[0]+1] & 1: continue
            # A matrix cell starts in full size. Merging cells removes that
            # scope boundary; restore it explicitly after any trailing script.
            pile += b'\x0a'
            pile += mtef[line[3]:line[1]-1]
        pile += b'\x00'
    pile += b'\x00'
    changed=mtef[:start]+pile+mtef[end:]
    header=bytearray(native[:28]); struct.pack_into('<I',header,8,len(changed))
    return bytes(header)+bytes(changed)
