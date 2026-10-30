import unittest
import os
import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


class TestXmlHelper(unittest.TestCase):
    def test_border(self):
        doc = docx.Document()
        p = doc.add_paragraph("Test paragraph with bottom border")
        pPr = p._p.get_or_add_pPr()
        bdr_xml = f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>'
        pPr.append(parse_xml(bdr_xml))
        out_file = "test_bdr.docx"
        doc.save(out_file)
        self.assertTrue(os.path.exists(out_file))
        if os.path.exists(out_file):
            os.remove(out_file)


if __name__ == "__main__":
    unittest.main()
