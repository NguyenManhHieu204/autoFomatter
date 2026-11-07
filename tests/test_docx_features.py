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

    def test_image_formatting_scaling_and_preservation(self):
        """Kiểm tra tính năng căn chỉnh và bảo vệ hình ảnh trong văn bản:
        - Không được xóa hay loại bỏ hình ảnh
        - Hình ảnh được căn giữa trang in (WD_ALIGN_PARAGRAPH.CENTER)
        - Không bị thụt đầu dòng (first_line_indent = 0)
        - Tự động co giãn tỉ lệ nếu ảnh quá to (> 160mm)
        - Dòng chú thích hình ảnh (Hình 1:...) được căn giữa, in nghiêng 11pt, không thụt dòng
        """
        import io
        from PIL import Image
        from core.formatter_engine import FormatterEngine
        from core.doc_parser import paragraph_has_drawings_or_images

        # Tạo ảnh mẫu kích thước lớn (chiều rộng 200mm > 160mm lề A4)
        img = Image.new("RGB", (800, 400), color="blue")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)

        doc = docx.Document()
        p0 = doc.add_paragraph("ỦY BAN NHÂN DÂN TỈNH QUẢNG NINH")
        p1 = doc.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        p2 = doc.add_paragraph("BÁO CÁO")
        p3 = doc.add_paragraph("Nội dung trước hình ảnh.")
        
        # Thêm đoạn chứa ảnh quá khổ
        p_img = doc.add_paragraph()
        p_img.paragraph_format.first_line_indent = Mm(12.7) # Thụt lề sai trước khi chuẩn hóa
        r_img = p_img.add_run()
        pic = r_img.add_picture(buf, width=Mm(200)) # 200mm quá khổ

        # Thêm dòng chú thích ảnh
        p_cap = doc.add_paragraph("Hình 1: Sơ đồ kiến trúc phần mềm")
        p_cap.paragraph_format.first_line_indent = Mm(12.7)

        p_end = doc.add_paragraph("Nội dung sau hình ảnh.")

        test_in = "test_img_input.docx"
        test_out = "test_img_output.docx"
        doc.save(test_in)

        try:
            engine = FormatterEngine()
            report = engine.format_document(test_in, output_docx_path=test_out, standard_id="ND30")

            self.assertTrue(report.success)
            self.assertGreaterEqual(report.images_formatted, 1)

            # Kiểm tra tài liệu sau chuẩn hóa
            doc_res = docx.Document(test_out)
            
            # 1. Tìm đoạn văn chứa hình ảnh
            img_paras = [p for p in doc_res.paragraphs if paragraph_has_drawings_or_images(p)]
            self.assertEqual(len(img_paras), 1, "Hình ảnh phải được bảo toàn 100%, không được xóa hay loại bỏ!")

            p_res_img = img_paras[0]
            # Căn giữa
            self.assertEqual(p_res_img.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            # Thụt đầu dòng = 0
            self.assertEqual(p_res_img.paragraph_format.first_line_indent, 0)
            
            # Co giãn kích thước tối đa 160mm
            inline = p_res_img._element.xpath('.//*[local-name()="inline"]')[0]
            extent = inline.find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
            cx = int(extent.attrib.get('cx', 0))
            w_mm = cx / 36000
            self.assertLessEqual(w_mm, 160.01)

            # 2. Tìm đoạn chú thích ảnh (Hình 1: Sơ đồ kiến trúc phần mềm)
            cap_paras = [p for p in doc_res.paragraphs if "Hình 1" in p.text]
            self.assertEqual(len(cap_paras), 1)
            p_res_cap = cap_paras[0]
            self.assertEqual(p_res_cap.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(p_res_cap.paragraph_format.first_line_indent, 0)
            # Kiểm tra font in nghiêng
            has_italic = any(r.font.italic for r in p_res_cap.runs)
            self.assertTrue(has_italic)

        finally:
            if os.path.exists(test_in):
                os.remove(test_in)
            if os.path.exists(test_out):
                os.remove(test_out)


if __name__ == "__main__":
    unittest.main()
