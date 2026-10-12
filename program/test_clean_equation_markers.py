import unittest
from lxml import etree as E
from clean_equation_markers import NS, clean_tree


class MarkerTests(unittest.TestCase):
    def tree(self, body):
        return E.fromstring(('<w:document xmlns:w="%s" xmlns:m="%s" xmlns:o="%s"><w:body>%s</w:body></w:document>' % (NS['w'], NS['m'], NS['o'], body)).encode())

    def test_table_tail_and_cjk_ref(self):
        root = self.tree('<w:tbl><w:tr><w:tc><w:p><w:r><w:object><o:OLEObject/></w:object></w:r><w:r><w:t>\u200b</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>式</w:t></w:r><w:r><w:t>\u00a0</w:t></w:r><w:hyperlink w:anchor="eq:test"><w:r><w:t>(4-1)</w:t></w:r></w:hyperlink></w:p>')
        before = E.tostring(root.find('.//o:OLEObject', NS))
        self.assertEqual(clean_tree(root), 2)
        self.assertEqual(E.tostring(root.find('.//o:OLEObject', NS)), before)
        self.assertEqual(root.find('.//w:hyperlink', NS).get('{'+NS['w']+'}anchor'), 'eq:test')
        self.assertEqual(clean_tree(root), 0)

    def test_legitimate_spaces_and_text(self):
        root = self.tree('<w:p><w:r><w:t>Eq.</w:t></w:r><w:r><w:t>\u00a0</w:t></w:r><w:hyperlink w:anchor="eq:test"><w:r><w:t>(4-1)</w:t></w:r></w:hyperlink><w:r><w:t> 10\u00a0kg\u200b</w:t></w:r></w:p>')
        before = E.tostring(root)
        self.assertEqual(clean_tree(root), 0)
        self.assertEqual(E.tostring(root), before)

    def test_figure_and_table_references(self):
        for prefix, anchor in [('图','fig:demo'),('表','tbl:demo')]:
            root=self.tree('<w:p><w:r><w:t>'+prefix+'</w:t></w:r><w:r><w:t>\u00a0</w:t></w:r><w:hyperlink w:anchor="'+anchor+'"><w:r><w:t>2-1</w:t></w:r></w:hyperlink></w:p>')
            self.assertEqual(clean_tree(root),1)
            self.assertEqual(clean_tree(root),0)
            self.assertEqual(root.find('.//w:hyperlink',NS).get('{'+NS['w']+'}anchor'),anchor)


if __name__ == '__main__': unittest.main()
