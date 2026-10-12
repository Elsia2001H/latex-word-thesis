import struct
import unittest
from mathtype_convert import wmf_math


def fixture(chunks, total):
    data = struct.pack('<IHhhhhHIH', 0x9ac6cdd7, 0, 0, 0, 1000, 500, 1000, 0, 0) + bytes(18)
    for chunk in chunks:
        payload = b'AppsMFCC\x01\x00' + struct.pack('<II', total, len(chunk)) + b'test\0' + chunk
        if len(payload) % 2: payload += b'\0'
        data += struct.pack('<IHHH', (10+len(payload))//2, 0x626, 15, len(payload)) + payload
    return data + struct.pack('<IH', 3, 0)


class WmfTests(unittest.TestCase):
    def test_multipart(self):
        native = b'\x05\x01' + bytes(range(256))*100
        mtef, width, height = wmf_math(fixture([native[:10000], native[10000:]], len(native)))
        self.assertEqual(mtef, native)
        self.assertEqual((width,height), (72,36))

    def test_incomplete(self):
        with self.assertRaises(ValueError):
            wmf_math(fixture([b'\x05\x01abc'], 20))

    def test_unknown_tex(self):
        with self.assertRaisesRegex(ValueError, 'unsupported'):
            wmf_math(fixture([b'\x05\x01abc'], 5) + b'<mtext>&#x005C;unknown</mtext>')


if __name__ == '__main__': unittest.main()
