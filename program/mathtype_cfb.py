"""Read and rebuild flat CFB v3 equation containers, including growing streams.

No MathType binaries or third-party source are redistributed by this module.
"""
import math
import struct

FREE, END, FAT = 0xffffffff, 0xfffffffe, 0xfffffffd


def read_streams(data):
    if data[:8] != bytes.fromhex('d0cf11e0a1b11ae1'):
        raise ValueError('Invalid compound file')
    shift, mini_shift = struct.unpack_from('<HH', data, 30)
    size, mini_size = 1 << shift, 1 << mini_shift
    def sector(i):
        start = (i + 1) * size
        if start + size > len(data):
            raise ValueError('CFB sector outside file')
        return data[start:start + size]
    fat_ids = list(struct.unpack_from('<109I', data, 76))
    difat, count = struct.unpack_from('<II', data, 68)
    seen = set()
    for _ in range(count):
        if difat in seen:
            raise ValueError('Cyclic DIFAT')
        seen.add(difat)
        entries = struct.unpack('<%dI' % (size // 4), sector(difat))
        fat_ids.extend(entries[:-1]); difat = entries[-1]
    fat_count = struct.unpack_from('<I', data, 44)[0]
    ids = [i for i in fat_ids if i != FREE][:fat_count]
    if len(ids) != fat_count:
        raise ValueError('Incomplete FAT')
    fat = [v for i in ids for v in struct.unpack('<%dI' % (size // 4), sector(i))]
    def chain(start, table, getter):
        parts, visited = [], set()
        while start != END:
            if start >= len(table) or start in visited:
                raise ValueError('Invalid/cyclic CFB chain')
            visited.add(start); parts.append(getter(start)); start = table[start]
        return b''.join(parts)
    directory = chain(struct.unpack_from('<I', data, 48)[0], fat, sector)
    entries = [directory[p:p+128] for p in range(0, len(directory), 128)]
    root = entries[0]
    if root[66] != 5:
        raise ValueError('Missing CFB root')
    mini = chain(struct.unpack_from('<I', root, 116)[0], fat, sector)[:struct.unpack_from('<Q', root, 120)[0]]
    mini_start, mini_count = struct.unpack_from('<II', data, 60)
    mini_raw = chain(mini_start, fat, sector) if mini_count else b''
    mini_fat = struct.unpack('<%dI' % (len(mini_raw)//4), mini_raw)
    result = {}
    for entry in entries[1:]:
        if entry[66] == 1:
            raise ValueError('Nested storages are not supported by equation writer')
        if entry[66] != 2:
            continue
        length = struct.unpack_from('<H', entry, 64)[0]
        name = entry[:length-2].decode('utf-16le')
        start, length = struct.unpack_from('<IQ', entry, 116)
        raw = chain(start, mini_fat, lambda i: mini[i*mini_size:(i+1)*mini_size]) if length < 4096 else chain(start, fat, sector)
        if len(raw) < length:
            raise ValueError('Truncated CFB stream')
        result[name] = raw[:length]
    return result, root[80:96]


def build_streams(streams, root_clsid=b'\0'*16):
    """Build valid FAT/miniFAT chains, preserving all stream contents and root CLSID."""
    names = sorted(streams, key=lambda n: (len(n), n.upper()))
    sectors, links, mini, mini_links, locations = [], [], bytearray(), [], {}
    def allocate(raw):
        if not raw:
            return END
        first = len(sectors)
        for p in range(0, len(raw), 512):
            sectors.append(raw[p:p+512].ljust(512, b'\0'))
            links.append(len(sectors) if p+512 < len(raw) else END)
        return first
    for name in names:
        raw = streams[name]
        if len(name.encode('utf-16le')) > 62:
            raise ValueError('CFB stream name too long')
        if len(raw) < 4096:
            start = len(mini_links) if raw else END
            for p in range(0, len(raw), 64):
                mini.extend(raw[p:p+64].ljust(64, b'\0'))
                mini_links.append(len(mini_links)+1 if p+64 < len(raw) else END)
        else:
            start = allocate(raw)
        locations[name] = (start, len(raw))
    mini_start = allocate(bytes(mini))
    mini_fat_raw = struct.pack('<%dI' % len(mini_links), *mini_links) if mini_links else b''
    mini_fat_raw = mini_fat_raw.ljust(math.ceil(len(mini_fat_raw)/512)*512, b'\xff')
    mini_fat_start = allocate(mini_fat_raw)
    tree = {}
    def make_tree(lo, hi, depth=0):
        if lo >= hi:
            return FREE
        mid = (lo+hi)//2
        left, right = make_tree(lo, mid, depth+1), make_tree(mid+1, hi, depth+1)
        tree[mid+1] = (left, right, depth)
        return mid+1
    child = make_tree(0, len(names))
    deepest = max((v[2] for v in tree.values()), default=0)
    def entry(name, kind, start, length, left=FREE, right=FREE, child=FREE, color=1, clsid=b'\0'*16):
        raw = bytearray(128); label = (name+'\0').encode('utf-16le')
        raw[:len(label)] = label
        struct.pack_into('<HBBIII', raw, 64, len(label), kind, color, left, right, child)
        raw[80:96] = clsid; struct.pack_into('<IQ', raw, 116, start, length)
        return raw
    directory = entry('Root Entry', 5, mini_start, len(mini), child=child, clsid=root_clsid)
    for i, name in enumerate(names, 1):
        left, right, depth = tree[i]
        directory += entry(name, 2, *locations[name], left=left, right=right, color=0 if depth==deepest and depth else 1)
    dir_start = allocate(bytes(directory))
    fat_count = 1
    while math.ceil((len(sectors)+fat_count)/128) != fat_count:
        fat_count = math.ceil((len(sectors)+fat_count)/128)
    if fat_count > 109:
        raise ValueError('Equation container exceeds supported 7 MB capacity')
    fat_start = len(sectors)
    table = links + [FAT]*fat_count
    table += [FREE]*(128*fat_count-len(table))
    for offset in range(0, len(table), 128):
        sectors.append(struct.pack('<128I', *table[offset:offset+128]))
    header = bytearray(512); header[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<HHHHH', header, 24, 0x3e, 3, 0xfffe, 9, 6)
    struct.pack_into('<IIIIIIIII', header, 40, 0, fat_count, dir_start, 0, 4096, mini_fat_start, len(mini_fat_raw)//512, END, 0)
    struct.pack_into('<109I', header, 76, *([fat_start+i for i in range(fat_count)] + [FREE]*(109-fat_count)))
    return bytes(header)+b''.join(sectors)


def replace_native(container, mtef):
    streams, clsid = read_streams(container)
    native = streams['Equation Native']
    header = bytearray(native[:28])
    if len(header) != 28:
        raise ValueError('Invalid Equation Native header')
    struct.pack_into('<I', header, 8, len(mtef))
    streams['Equation Native'] = bytes(header)+mtef
    result = build_streams(streams, clsid)
    reread, reread_clsid = read_streams(result)
    if reread != streams or reread_clsid != clsid:
        raise ValueError('CFB rebuilding verification failed')
    return result
