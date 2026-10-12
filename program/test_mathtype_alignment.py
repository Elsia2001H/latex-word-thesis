import struct
import unittest
from mathtype_alignment import Reader, align_native


def fixture():
    def line(value):
        return b'\x01\x00' + b''.join(b'\x02\x00\x83'+struct.pack('<H',ord(c)) for c in value) + b'\x00'
    matrix=b'\x05\x00\x01\x01\x01\x02\x02\x00\x00'+line('a')+line('=b')+line('long')+line('=c')+b'\x00'
    mtef=b'\x05\x01\x00\x07\x08test\x00\x00\x0a\x01\x00'+matrix+b'\x00\x00'
    header=bytearray(28); struct.pack_into('<I',header,8,len(mtef))
    return bytes(header)+mtef


class AlignmentTests(unittest.TestCase):
    def test_equation_wrapper_and_label(self):
        native=fixture()
        changed=align_native(native,r'\begin{equation}\begin{aligned}a&=b\\long&=c\end{aligned}\label{eq:test}\end{equation}{\qquad(4-4)}')
        self.assertNotEqual(native,changed)
        self.assertIn(b'\x04\x00\x04\x04',changed)

    def test_cell_boundary_restores_full_size(self):
        native=fixture()
        # Model a left cell ending with a subscript-size state.
        marker=b'\x02\x00\x83g\x00\x00\x01\x00'
        native=native.replace(marker,b'\x02\x00\x83g\x00\x0b\x00\x01\x00')
        native=bytearray(native)
        struct.pack_into('<I',native,8,len(native)-28)
        changed=align_native(bytes(native),r'\begin{aligned}a&=b\\long&=c\end{aligned}')
        self.assertIn(b'\x0b\x0a\x02\x00\x83=\x00',changed)

    def test_relation_aligned_pile(self):
        native=fixture()
        changed=align_native(native,r'\begin{aligned}a&=b\\long&=c\end{aligned}')
        self.assertNotEqual(native,changed)
        self.assertIn(b'\x04\x00\x04\x04',changed[28:])
        self.assertEqual(struct.unpack_from('<I',changed,8)[0],len(changed)-28)
        reader=Reader(changed[28:]); reader.pos=5; reader.string(); reader.skip(1); reader.records()
        self.assertEqual(reader.matrices,[])
        # Every variable/relation character retains its original record bytes.
        for c in 'ablongc=':
            self.assertEqual(native.count(b'\x02\x00\x83'+struct.pack('<H',ord(c))),changed.count(b'\x02\x00\x83'+struct.pack('<H',ord(c))))

    def test_leave_matrices_and_multiple_pairs(self):
        native=fixture()
        self.assertEqual(align_native(native,r'\begin{bmatrix}a&b\\c&d\end{bmatrix}'),native)
        self.assertEqual(align_native(native,r'\begin{aligned}a&=b&c&=d\\e&=f&g&=h\end{aligned}'),native)


if __name__ == '__main__': unittest.main()
