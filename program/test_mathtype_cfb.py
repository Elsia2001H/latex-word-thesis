import io
import sys
from pathlib import Path
from zipfile import ZipFile
from mathtype_cfb import read_streams, build_streams, replace_native


def main():
    path = Path(sys.argv[1]) if len(sys.argv)>1 else Path(__file__).parents[1]/'examples/thesis-mathtype.docx'
    with ZipFile(path) as z:
        bins = [z.read(n) for n in z.namelist() if n.endswith('.bin')]
    for original in bins:
        streams, clsid = read_streams(original)
        assert read_streams(build_streams(streams, clsid)) == (streams, clsid)
        for size in (0, 1, 63, 64, 65, 4095, 4096, 4097, 65536, 1000000):
            changed = dict(streams); changed['Equation Native'] = b'x'*size
            assert read_streams(build_streams(changed, clsid)) == (changed, clsid)
        native = streams['Equation Native']
        assert read_streams(replace_native(original, native[28:]))[0]['Equation Native'][28:] == native[28:]
    print('PASS: %d containers, small/large stream boundary and 1 MB growth' % len(bins))


if __name__ == '__main__':
    main()
