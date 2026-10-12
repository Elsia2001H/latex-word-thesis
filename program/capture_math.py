"""Identity Pandoc filter recording the TeX of math nodes in document order."""
import json
import os
import sys


def main():
    tree = json.load(sys.stdin)
    found = []
    def walk(node):
        if isinstance(node, dict):
            if node.get('t') == 'Math':
                found.append({'display': node['c'][0]['t']=='DisplayMath', 'tex': node['c'][1]})
            else:
                for value in node.values():
                    walk(value)
        elif isinstance(node, list):
            for value in node:
                walk(value)
    walk(tree['blocks'])
    with open(os.environ['THESIS_MATH_MANIFEST'], 'w', encoding='utf-8') as f:
        json.dump(found, f, ensure_ascii=False, indent=2)
    json.dump(tree, sys.stdout, ensure_ascii=False)


if __name__ == '__main__':
    main()
