import os
import unittest
import docx
from docx.shared import Pt, Mm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn


class TestDocxFeatures(unittest.TestCase):
    def test_page_number_and_borders(self):
        doc = docx.Document()
        section = doc.sections[0]
        section.different_first_page_header_footer = True
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)

        # Header with page number
        header = section.header
        p = header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13.5)
        fld = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        run._r.append(fld)

        # First page header should be empty
        first_header = section.first_page_header
        first_p = first_header.paragraphs[0]
        first_p.text = ""

        # Table with invisible borders
        table = doc.add_table(rows=1, cols=2)
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
            f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
            f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
            f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
            f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
            f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)

        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p_left.add_run("ỦY BAN NHÂN DÂN\nTHÀNH PHỐ HÀ NỘI")
        r1.font.name = "Times New Roman"
        r1.font.bold = True

        p_right = cell_right.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc")
        r2.font.name = "Times New Roman"
        r2.font.bold = True

        test_out = "test_output.docx"
        doc.save(test_out)
        self.assertTrue(os.path.exists(test_out))

        # Reopen to verify validity
        reopened = docx.Document(test_out)
        self.assertEqual(len(reopened.tables), 1)
        if os.path.exists(test_out):
            os.remove(test_out)


if __name__ == "__main__":
    unittest.main()
