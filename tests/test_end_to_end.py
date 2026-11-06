import os
import unittest
import docx
from docx.shared import Pt, Inches

from core.formatter_engine import FormatterEngine
from core.win32_helper import Win32WordHelper


class TestEndToEndFormatting(unittest.TestCase):
    def setUp(self):
        self.draft_path = "test_draft_sample.docx"
        self.output_path = "test_draft_sample_ChuanHoa.docx"

        # Tạo file word nháp giả lập thực tế với nhiều lỗi thể thức
        doc = docx.Document()

        # Tiêu đề nháp gõ tay bằng spaces
        doc.add_paragraph("UBND THÀNH PHỐ HÀ NỘI              CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        doc.add_paragraph("SỞ NỘI VỤ                                     Độc lập - Tự do - Hạnh phúc")
        doc.add_paragraph("Số: 15/QĐ-SNV                                  Hà Nội, ngày 5 tháng 03 năm 2026")

        # Tên văn bản và trích yếu
        p_title = doc.add_paragraph("QUYẾT ĐỊNH")
        p_subj = doc.add_paragraph("Về việc ban hành \"Quy chế quản lý văn bản\" tại cơ quan")

        # Căn cứ
        doc.add_paragraph("Căn cứ Luật Tổ chức chính quyền địa phương  ngày 19 tháng 6 năm 2015 ;")
        doc.add_paragraph("Căn cứ Nghị định số 30/2020/NĐ-CP  ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư .")

        # Điều khoản & nội dung có lỗi chính tả hành chính
        p_d1 = doc.add_paragraph("Điều 1 . Ban hành kèm theo Quyết định này  \"Quy chế quản lý văn bản\" .")
        p_d2 = doc.add_paragraph("Điều 2. Giao cho các Đồng chí tại Đơn vị trực thuộc phối hợp thực hiện theo chỉ đạo của chính phủ và ban chấp hành trung ương .")
        p_d3 = doc.add_paragraph("Điều 3. Quyết định này có hiệu lực kể từ  ngày 02 tháng 1 năm 2026 .")

        # Nơi nhận và chữ ký sai thể thức (dùng T/M thay vì TM., chữ ký thiếu chuẩn)
        doc.add_paragraph("Nơi nhận:")
        doc.add_paragraph("- Như Điều 3")
        doc.add_paragraph("- Giám đốc Sở")
        doc.add_paragraph("- Lưu: VT")

        doc.add_paragraph("T/M. GIÁM ĐỐC")
        doc.add_paragraph("PHÓ GIÁM ĐỐC")
        doc.add_paragraph("Phạm Văn Dũng")

        doc.save(self.draft_path)

    def tearDown(self):
        for f in [self.draft_path, self.output_path, f"{self.draft_path}.bak"]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass

    def test_full_pipeline_nd30(self):
        engine = FormatterEngine()
        report = engine.format_document(
            input_docx_path=self.draft_path,
            output_docx_path=self.output_path,
            standard_id="ND30"
        )

        self.assertTrue(report.success, f"Formatting failed: {report.error_message}")
        self.assertTrue(os.path.exists(self.output_path))
        self.assertGreater(report.paragraphs_formatted, 0)
        self.assertGreater(report.spaces_fixed, 0)
        self.assertGreater(report.quotes_fixed, 0)
        self.assertGreater(report.dates_fixed, 0)
        self.assertGreater(report.cap_warnings_count, 0)

        # Mở file kết quả kiểm tra cấu trúc
        doc_res = docx.Document(self.output_path)
        # Bảng đầu tiên phải là Header Table 2 cột
        self.assertGreaterEqual(len(doc_res.tables), 1)
        self.assertEqual(len(doc_res.tables[0].columns), 2)
        # Kiểm tra bảo toàn dữ liệu cơ quan ban hành và số ký hiệu từ văn bản gốc
        cell0_text = doc_res.tables[0].cell(0, 0).text
        self.assertIn("UBND THÀNH PHỐ HÀ NỘI", cell0_text)
        self.assertIn("SỞ NỘI VỤ", cell0_text)
        self.assertIn("15/QĐ-SNV", cell0_text)

        # Kiểm tra lề khổ A4
        section = doc_res.sections[0]
        self.assertAlmostEqual(section.page_width.mm, 210, delta=1)
        self.assertAlmostEqual(section.page_height.mm, 297, delta=1)

    def test_space_delimited_report_preservation(self):
        """Kiểm tra tài liệu BÁO CÁO gõ tay 2 cột bằng khoảng trắng/tab: không bị gộp dòng tiêu đề và bảo toàn tên cơ quan"""
        test_file = "test_report_spacing.docx"
        test_out = "test_report_spacing_out.docx"
        try:
            doc = docx.Document()
            doc.add_paragraph("UBND TỈNH QUẢNG NINH                                CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
            doc.add_paragraph("SỞ NỘI VỤ                                                            Độc lập - Tự do - Hạnh phúc")
            doc.add_paragraph("Số:  45 /SNV - VP                                  Quảng Ninh ,ngày 5 tháng 04 năm 2026")
            doc.add_paragraph("BÁO   CÁO")
            doc.add_paragraph("Về    việc rà  soát ,đánh  giá công tác cải cách hành chính và chuyển đổi số")
            doc.add_paragraph("Thực hiện ý kiến chỉ đạo của ủy ban nhân dân tỉnh tại Công văn số 123/UBND-TH, Sở Nội vụ đã chủ trì kiểm tra tình hình thực tế.")
            doc.add_paragraph("Nơi nhận :                                                          TM .UBND TỈNH")
            doc.add_paragraph("- Như trên ;                                                         KT .CHỦ TỊCH")
            doc.add_paragraph("- Lưu :VT ,SNV .                                                    PHÓ CHỦ TỊCH")
            doc.add_paragraph("                                                                     nguyễn văn an")
            doc.save(test_file)

            engine = FormatterEngine()
            rep = engine.format_document(test_file, test_out, standard_id="ND30")
            self.assertTrue(rep.success)

            res = docx.Document(test_out)
            # Kiểm tra bảo toàn thông tin Header
            h_c0 = res.tables[0].cell(0, 0).text
            h_c1 = res.tables[0].cell(0, 1).text
            self.assertIn("UBND TỈNH QUẢNG NINH", h_c0)
            self.assertIn("SỞ NỘI VỤ", h_c0)
            self.assertIn("45/SNV-VP", h_c0)
            self.assertIn("Quảng Ninh", h_c1)

            # Kiểm tra Tiêu đề, Trích yếu và Thân bài tách biệt độc lập
            self.assertEqual(res.paragraphs[0].text, "BÁO CÁO")
            self.assertTrue(res.paragraphs[1].text.startswith("Về việc rà soát"))
            self.assertTrue(res.paragraphs[2].text.startswith("Thực hiện ý kiến chỉ đạo"))

            # Kiểm tra Footer
            f_c0 = res.tables[1].cell(0, 0).text
            f_c1 = res.tables[1].cell(0, 1).text
            self.assertIn("Như trên;", f_c0)
            self.assertIn("Lưu: VT, SNV.", f_c0)
            self.assertIn("PHÓ CHỦ TỊCH", f_c1)
            self.assertIn("Nguyễn Văn An", f_c1)
        finally:
            for f in [test_file, test_out, f"{test_file}.bak"]:
                if os.path.exists(f):
                    try:
                        os.remove(f)
                    except Exception:
                        pass

    def test_full_pipeline_hd36(self):
        engine = FormatterEngine()
        out_party = "test_party_ChuanHoa.docx"
        report = engine.format_document(
            input_docx_path=self.draft_path,
            output_docx_path=out_party,
            standard_id="HD36"
        )
        self.assertTrue(report.success)
        self.assertTrue(os.path.exists(out_party))
        doc_res = docx.Document(out_party)
        self.assertGreaterEqual(len(doc_res.tables), 1)
        header_text = "\n".join([c.text for c in doc_res.tables[0].columns[1].cells])
        self.assertIn("ĐẢNG CỘNG SẢN VIỆT NAM", header_text)
        if os.path.exists(out_party):
            os.remove(out_party)



if __name__ == "__main__":
    unittest.main()
