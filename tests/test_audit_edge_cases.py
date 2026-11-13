"""
tests/test_audit_edge_cases.py: Bộ kiểm thử tự động rà soát các rủi ro biên
và điểm nghẽn kỹ thuật theo Báo cáo Đánh giá An toàn Thể thức (Audit Matrix).
Gồm 4 nhóm kiểm thử cốt lõi:
- Nhóm 1: Sai lệch lề in (< 0.5mm) & Đánh số trang từ trang 2
- Nhóm 2: Bảng biểu phức tạp (Merged Cells), lặp tiêu đề trang (tblHeader), chống gãy dòng (cantSplit)
- Nhóm 3: Thụt lề phân cấp văn bản Đảng & Nhà nước (I., 1., a), -, +)
- Nhóm 4: Nắn thẳng ảnh scan nghiêng > 5 độ & Phân tách mực dấu đỏ qua không gian màu HSV
- Nhóm 5: Timeout Watchdog cho MS Word COM
"""

import os
import sys
import unittest
import tempfile
import docx
from docx.shared import Mm, Pt
from docx.oxml.ns import qn
import numpy as np

# Thêm đường dẫn gốc dự án vào sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.formatter_engine import FormatterEngine
from core.ocr_engine import deskew_image, remove_red_seal_hsv
from core.signature_engine import ui_to_pdf_coordinates, pdf_to_ui_coordinates, VisualStampGenerator
from core.win32_helper import WordWatchdog
from core.batch_editor import BatchEditor, BatchEditOptions
from core.administrative_sanitizer import AdministrativeSanitizer


class TestAuditEdgeCases(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="af_test_audit_")
        self.engine = FormatterEngine()

    def tearDown(self):
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    # ==================== NHÓM 1: LỀ IN & ĐÁNH SỐ TRANG ====================
    def test_group1_margins_tolerance_and_page_numbering(self):
        """Kiểm tra dung sai lề in (< 0.5mm) và cấu hình ẩn số trang tại trang 1"""
        doc = docx.Document()
        doc.add_paragraph("Trang đầu tiên của tài liệu")

        rules = self.engine.get_rules("ND30")
        self.engine._apply_page_setup(doc, rules)

        section = doc.sections[0]
        # Quy chuẩn NĐ 30: Top=20mm, Bottom=20mm, Left=30mm, Right=15mm
        top_diff = abs(section.top_margin.mm - 20.0)
        bottom_diff = abs(section.bottom_margin.mm - 20.0)
        left_diff = abs(section.left_margin.mm - 30.0)
        right_diff = abs(section.right_margin.mm - 15.0)

        self.assertLess(top_diff, 0.5, f"Lề trên lệch {top_diff}mm, vượt ngưỡng 0.5mm")
        self.assertLess(bottom_diff, 0.5, f"Lề dưới lệch {bottom_diff}mm, vượt ngưỡng 0.5mm")
        self.assertLess(left_diff, 0.5, f"Lề trái lệch {left_diff}mm, vượt ngưỡng 0.5mm")
        self.assertLess(right_diff, 0.5, f"Lề phải lệch {right_diff}mm, vượt ngưỡng 0.5mm")

        # Đánh số trang: Ẩn trang 1, căn giữa lề trên từ trang 2
        self.assertTrue(section.different_first_page_header_footer)
        header_para = section.header.paragraphs[0]
        self.assertEqual(header_para.alignment, docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER)

    # ==================== NHÓM 2: BẢNG BIỂU PHỨC TẠP ====================
    def test_group2_complex_tables_multipage_protection(self):
        """Kiểm tra bảng biểu nhiều trang: có tblHeader lặp tiêu đề và cantSplit chống gãy dòng"""
        test_file = os.path.join(self.temp_dir, "test_table.docx")
        out_file = os.path.join(self.temp_dir, "out_table.docx")

        doc = docx.Document()
        # Tạo bảng 5 hàng x 3 cột có ô gộp (Merged Cells)
        t = doc.add_table(rows=5, cols=3)
        t.cell(0, 0).text = "STT"
        t.cell(0, 1).text = "Họ và tên"
        t.cell(0, 2).text = "Ghi chú"
        t.cell(1, 0).merge(t.cell(2, 0))  # Gộp 2 ô STT
        t.cell(1, 0).text = "1"
        t.cell(1, 1).text = "Nguyễn Văn A"
        t.cell(2, 1).text = "Trần Thị B"
        doc.save(test_file)

        report = self.engine.format_document(test_file, out_file, standard_id="ND30")
        self.assertTrue(report.success)

        doc_res = docx.Document(out_file)
        res_table = doc_res.tables[0]

        # Kiểm tra lặp dòng tiêu đề trên từng trang (tblHeader)
        trPr0 = res_table.rows[0]._tr.trPr
        self.assertIsNotNone(trPr0.find(qn('w:tblHeader')), "Thiếu thẻ w:tblHeader trên dòng tiêu đề bảng")

        # Kiểm tra chống gãy dòng (cantSplit) trên các dòng
        for row_idx, row in enumerate(res_table.rows):
            trPr = row._tr.trPr
            self.assertIsNotNone(trPr.find(qn('w:cantSplit')), f"Dòng {row_idx} thiếu thẻ w:cantSplit")

    # ==================== NHÓM 3: THỤT LỀ PHÂN CẤP ĐẢNG & NHÀ NƯỚC ====================
    def test_group3_multilevel_indentation(self):
        """Kiểm tra thụt lề phân cấp: I., 1., a), -, +"""
        test_file = os.path.join(self.temp_dir, "test_hierarchy.docx")
        out_file = os.path.join(self.temp_dir, "out_hierarchy.docx")

        doc = docx.Document()
        doc.add_paragraph("I. TỔ CHỨC BỘ MÁY")
        doc.add_paragraph("1. Cơ cấu phòng ban trực thuộc")
        doc.add_paragraph("a) Văn phòng cấp ủy")
        doc.add_paragraph("- Tiếp nhận và xử lý văn bản")
        doc.add_paragraph("+ Sắp xếp lịch họp ban thường vụ")
        doc.save(test_file)

        report = self.engine.format_document(test_file, out_file, standard_id="HD36")
        self.assertTrue(report.success)

        doc_res = docx.Document(out_file)
        paras = doc_res.paragraphs

        # Cấp La Mã: first line indent
        self.assertAlmostEqual(paras[0].paragraph_format.first_line_indent.mm, 12.7, delta=1.0)
        # Cấp Số: first line indent
        self.assertAlmostEqual(paras[1].paragraph_format.first_line_indent.mm, 12.7, delta=1.0)
        # Cấp chữ cái a): left indent 10mm
        self.assertAlmostEqual(paras[2].paragraph_format.left_indent.mm, 10.0, delta=1.0)
        # Cấp gạch đầu dòng -: left indent 14mm
        self.assertAlmostEqual(paras[3].paragraph_format.left_indent.mm, 14.0, delta=1.0)
        # Cấp cộng đầu dòng +: left indent 18mm
        self.assertAlmostEqual(paras[4].paragraph_format.left_indent.mm, 18.0, delta=1.0)

    # ==================== NHÓM 4: DESKEW & HSV RED SEAL FILTER ====================
    def test_group4_deskew_and_red_seal_hsv(self):
        """Kiểm tra thuật toán nắn thẳng xoay > 5 độ và bóc tách mực dấu đỏ bằng HSV"""
        import cv2

        # 1. Kiểm tra nắn thẳng góc xoay 7.5 độ
        img = np.ones((400, 600, 3), dtype=np.uint8) * 255
        cv2.putText(img, "DANG CONG SAN VIET NAM", (50, 100), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)
        cv2.putText(img, "QUYET DINH VE PHAN CONG", (80, 150), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)
        M = cv2.getRotationMatrix2D((300, 200), 7.5, 1.0)
        rotated = cv2.warpAffine(img, M, (600, 400), borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))

        straightened, corrected_angle = deskew_image(rotated)
        self.assertGreater(abs(corrected_angle), 6.5, "Không phát hiện được góc nghiêng > 6.5 độ")

        # Kiểm tra ảnh sau khi nắn thẳng có góc lệch còn lại < 0.5 độ
        _, residual_angle = deskew_image(straightened)
        self.assertEqual(residual_angle, 0.0, "Ảnh sau nắn thẳng vẫn còn độ nghiêng")

        # 2. Kiểm tra bộ lọc tách dấu đỏ HSV (Color Channel Separation)
        seal_img = np.ones((200, 400, 3), dtype=np.uint8) * 255
        cv2.putText(seal_img, "BI THU DANG UY", (50, 100), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 0), 2)
        # Đóng dấu đỏ đè lên chữ (BGR: 0, 0, 220)
        cv2.circle(seal_img, (150, 95), 45, (0, 0, 220), 4)

        cleaned_img, red_px_removed = remove_red_seal_hsv(seal_img)
        self.assertGreater(red_px_removed, 500, "Không nhận diện được pixel mực đỏ của con dấu")

        # Kiểm tra trên ảnh đã tẩy dấu: pixel đỏ đã được triệt tiêu
        hsv_clean = cv2.cvtColor(cleaned_img, cv2.COLOR_BGR2HSV)
        m1 = cv2.inRange(hsv_clean, np.array([0, 40, 40]), np.array([12, 255, 255]))
        m2 = cv2.inRange(hsv_clean, np.array([160, 40, 40]), np.array([180, 255, 255]))
        remaining_red = int(np.sum((m1 | m2) > 0))
        self.assertEqual(remaining_red, 0, f"Còn sót {remaining_red} pixel đỏ sau khi lọc")

    # ==================== NHÓM 5: TỌA ĐỘ KÝ SỐ & TIMEOUT WATCHDOG ====================
    def test_group5_coordinate_mapping_and_watchdog(self):
        """Kiểm tra ánh xạ tọa độ PostScript PDF và cơ chế Watchdog thời gian chờ"""
        # Kiểm tra chuyển đổi hệ tọa độ PostScript (gốc 0,0 góc dưới bên trái)
        page_w = 595.28
        page_h = 841.89
        box_w = 150.0
        box_h = 80.0

        pdf_x, pdf_y = ui_to_pdf_coordinates(0.55, 0.70, page_w, page_h, box_w, box_h)
        expected_x = round(0.55 * page_w, 2)
        expected_y = round(page_h - (0.70 * page_h) - box_h, 2)
        self.assertEqual(pdf_x, expected_x)
        self.assertEqual(pdf_y, expected_y)

        # Kiểm tra chuyển đổi ngược
        norm_x, norm_y = pdf_to_ui_coordinates(pdf_x, pdf_y, page_w, page_h, box_h)
        self.assertAlmostEqual(norm_x, 0.55, places=2)
        self.assertAlmostEqual(norm_y, 0.70, places=2)

        # Kiểm tra Timeout Watchdog kích hoạt khi timeout
        watchdog = WordWatchdog(timeout_sec=0.1, target_pid=999999)
        watchdog.start()
        import time
        time.sleep(0.2)
        self.assertTrue(watchdog.triggered, "Watchdog không kích hoạt khi quá thời gian chờ")

    # ==================== NHÓM 6: BẢO TOÀN ĐỊA DANH ĐỊA PHƯƠNG & XỬ LÝ NĐ 30 ====================
    def test_group6_preservation_of_local_location_thuan_hanh(self):
        """Kiểm tra tuyệt đối bảo toàn địa danh Thuận Hạnh (không bị chuyển thành Hà Nội) theo NĐ 30"""
        from core.administrative_sanitizer import AdministrativeSanitizer

        # 1. Kiểm tra OCR healing nhận dạng lỗi gõ 'thuận hnahj'
        healed, _ = AdministrativeSanitizer.heal_ocr_text("UBND xa Thuan Hnahj")
        self.assertIn("Thuận Hạnh", healed)

        # 2. Kiểm tra suy luận địa danh theo NĐ 30 (Phụ lục I Mục II khoản 4a)
        loc1 = AdministrativeSanitizer.infer_location_from_agencies("XÃ THUẬN HẠNH", "UỶ BAN NHÂN DÂN")
        self.assertEqual(loc1, "Thuận Hạnh")

        loc2 = AdministrativeSanitizer.infer_location_from_agencies("ỦY BAN NHÂN DÂN XÃ THUẬN HẠNH", "")
        self.assertEqual(loc2, "Thuận Hạnh")

        loc3 = AdministrativeSanitizer.infer_location_from_agencies("SỞ TÀI CHÍNH", "UBND TỈNH QUẢNG NINH")
        self.assertEqual(loc3, "Quảng Ninh")

        # 3. Kiểm tra tách địa danh và ngày tháng từ dự thảo có khoảng trống
        loc, raw_date = AdministrativeSanitizer.extract_location_and_date("Thuận Hạnh, ngày    tháng    năm 2026")
        self.assertEqual(loc, "Thuận Hạnh")
        self.assertEqual(raw_date, "ngày    tháng    năm 2026")

        # 4. Kiểm tra xử lý hoàn chỉnh qua FormatterEngine với bảng có sẵn từ file mẫu Thuận Hạnh
        test_file = os.path.join(self.temp_dir, "test_thuan_hanh.docx")
        doc = docx.Document()
        tbl = doc.add_table(rows=1, cols=2)
        tbl.cell(0, 0).text = "UỶ BAN NHÂN DÂN\nXÃ THUẬN HẠNH\n\nSố:      /TB-ĐKĐĐ"
        tbl.cell(0, 1).text = "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nThuận Hạnh, ngày    tháng    năm 2026"
        doc.add_paragraph("THÔNG BÁO KẾT QUẢ ĐĂNG KÝ ĐẤT ĐAI")
        doc.save(test_file)

        rep = self.engine.format_document(test_file)
        self.assertTrue(rep.success)

        res_doc = docx.Document(rep.output_file)
        res_cell1 = res_doc.tables[0].cell(0, 1).text
        # Tuyệt đối không được chứa Hà Nội và phải giữ nguyên Thuận Hạnh
        self.assertNotIn("Hà Nội", res_cell1)
        self.assertIn("Thuận Hạnh", res_cell1)
        self.assertIn("Thuận Hạnh, ngày    tháng    năm 2026", res_cell1)

    # ==================== NHÓM 7: CỠ CHỮ NỘI DUNG 13PT (ND30 & HD36) ====================
    def test_group7_body_font_size_13pt(self):
        """Kiểm tra cỡ chữ phần lời văn/nội dung đạt chuẩn 13pt theo yêu cầu và khung quy chuẩn NĐ 30"""
        test_file = os.path.join(self.temp_dir, "test_body_size.docx")
        out_file = os.path.join(self.temp_dir, "out_body_size.docx")

        doc = docx.Document()
        doc.add_paragraph("Căn cứ Luật Đất đai ngày 29 tháng 11 năm 2013;")
        doc.add_paragraph("Uỷ ban nhân dân xã Thuận Hạnh thông báo kết quả đăng ký đất đai như sau:")
        doc.add_paragraph("1. Thông tin người sử dụng đất và thửa đất đăng ký.")
        doc.add_paragraph("a) Người sử dụng đất chấp hành đầy đủ nghĩa vụ tài chính.")
        doc.save(test_file)

        # Định dạng theo chuẩn NĐ 30
        report = self.engine.format_document(test_file, out_file, standard_id="ND30")
        self.assertTrue(report.success)

        res_doc = docx.Document(out_file)
        paras = res_doc.paragraphs

        # Căn cứ pháp lý: 13.0 pt
        self.assertEqual(paras[0].runs[0].font.size.pt, 13.0)
        self.assertTrue(paras[0].runs[0].font.italic)

        # Lời văn nội dung: 13.0 pt
        self.assertEqual(paras[1].runs[0].font.size.pt, 13.0)

        # Tiêu đề mục số: 13.0 pt, in đậm
        self.assertEqual(paras[2].runs[0].font.size.pt, 13.0)
        self.assertTrue(paras[2].runs[0].font.bold)

        # Đoạn mục a): 13.0 pt
        self.assertEqual(paras[3].runs[0].font.size.pt, 13.0)

    # ==================== NHÓM 8: CHUẨN HÓA CƠ QUAN XÃ & CHỈNH SỬA HÀNG LOẠT ====================
    def test_group8_agency_split_and_batch_editor(self):
        """Kiểm tra chuẩn hóa ngắt dòng cơ quan NĐ 30 và bộ máy chỉnh sửa hàng loạt BatchEditor"""
        from core.administrative_sanitizer import AdministrativeSanitizer
        from core.batch_editor import BatchEditor, BatchEditOptions

        # 1. Kiểm tra thuật toán bóc tách và chuẩn hóa ngắt dòng tên cơ quan
        # A. Thuận Hạnh
        s1, i1 = AdministrativeSanitizer.split_and_normalize_agency_header("ỦY BAN NHÂN DÂN XÃ", "THUẬN HẠNH")
        self.assertEqual(s1, "ỦY BAN NHÂN DÂN")
        self.assertEqual(i1, "XÃ THUẬN HẠNH")

        # B. Đức An
        s2, i2 = AdministrativeSanitizer.split_and_normalize_agency_header("ỦY BAN NHÂN DÂN XÃ", "ĐỨC AN")
        self.assertEqual(s2, "ỦY BAN NHÂN DÂN")
        self.assertEqual(i2, "XÃ ĐỨC AN")

        # C. Lỗi chính tả / OCR typo: UBND XÃ + thuận hnahj
        s3, i3 = AdministrativeSanitizer.split_and_normalize_agency_header("UBND XÃ", "thuận hnahj")
        self.assertEqual(s3, "ỦY BAN NHÂN DÂN")
        self.assertEqual(i3, "XÃ THUẬN HẠNH")

        # D. Dồn chung 1 dòng: UBND XÃ ĐỨC AN
        s4, i4 = AdministrativeSanitizer.split_and_normalize_agency_header("", "ỦY BAN NHÂN DÂN XÃ ĐỨC AN")
        self.assertEqual(s4, "ỦY BAN NHÂN DÂN")
        self.assertEqual(i4, "XÃ ĐỨC AN")

        # E. Hội đồng nhân dân xã
        s5, i5 = AdministrativeSanitizer.split_and_normalize_agency_header("HỘI ĐỒNG NHÂN DÂN XÃ", "THUẬN HẠNH")
        self.assertEqual(s5, "HỘI ĐỒNG NHÂN DÂN")
        self.assertEqual(i4, "XÃ ĐỨC AN")

        # 2. Kiểm tra xuất file sang thư mục tùy chọn (output_dir)
        custom_out_dir = os.path.join(self.temp_dir, "custom_export")
        os.makedirs(custom_out_dir, exist_ok=True)

        test_file = os.path.join(self.temp_dir, "test_agency_format.docx")
        doc = docx.Document()
        tbl = doc.add_table(rows=1, cols=2)
        tbl.cell(0, 0).text = "ỦY BAN NHÂN DÂN XÃ\nTHUẬN HẠNH\n\nSố: 01/TB-UBND"
        tbl.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nThuận Hạnh, ngày 05 tháng 11 năm 2026"
        doc.add_paragraph("THÔNG BÁO")
        doc.save(test_file)

        rep = self.engine.format_document(test_file, output_dir=custom_out_dir)
        self.assertTrue(rep.success)
        self.assertTrue(os.path.exists(rep.output_file))
        self.assertEqual(os.path.dirname(os.path.abspath(rep.output_file)), os.path.abspath(custom_out_dir))

        res_doc = docx.Document(rep.output_file)
        cell_0_text = res_doc.tables[0].cell(0, 0).text
        # Dòng 1 phải là ỦY BAN NHÂN DÂN, Dòng 2 phải là XÃ THUẬN HẠNH
        self.assertIn("ỦY BAN NHÂN DÂN", cell_0_text)
        self.assertIn("XÃ THUẬN HẠNH", cell_0_text)
        self.assertNotIn("ỦY BAN NHÂN DÂN XÃ\nTHUẬN HẠNH", cell_0_text)

        # 3. Kiểm tra BatchEditor xử lý hàng loạt: Xóa ngày, đổi địa danh sang Đức An, tìm và thay thế
        batch_files = []
        for idx in range(3):
            bf = os.path.join(self.temp_dir, f"batch_in_{idx}.docx")
            d = docx.Document()
            t = d.add_table(rows=1, cols=2)
            t.cell(0, 0).text = "ỦY BAN NHÂN DÂN XÃ\nTHUẬN HẠNH\n\nSố: 05/TB-UBND"
            t.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nThuận Hạnh, ngày 10 tháng 10 năm 2026"
            d.add_paragraph(f"Nội dung văn bản số {idx}: áp dụng quy chế cũ.")
            d.save(bf)
            batch_files.append(bf)

        batch_out_dir = os.path.join(self.temp_dir, "batch_result")
        os.makedirs(batch_out_dir, exist_ok=True)

        opts = BatchEditOptions(
            enable_location=True,
            new_location="Đức An",
            sync_agency_location=True,
            enable_date=True,
            date_mode="draft_spaces",
            enable_find_replace=True,
            find_text="quy chế cũ",
            replace_text="quy chế mới 2026",
            output_dir=batch_out_dir
        )

        results = BatchEditor.process_batch(batch_files, opts)
        self.assertEqual(len(results), 3)
        for r in results:
            self.assertTrue(r.success)
            self.assertTrue(os.path.exists(r.output_path))
            out_d = docx.Document(r.output_path)

            # Cột trái: Đã đổi thành XÃ ĐỨC AN
            left_t = out_d.tables[0].cell(0, 0).text
            self.assertIn("ỦY BAN NHÂN DÂN", left_t)
            self.assertIn("XÃ ĐỨC AN", left_t)

            # Cột phải: Đã xóa ngày thành dự thảo Đức An
            right_t = out_d.tables[0].cell(0, 1).text
            self.assertIn("Đức An, ngày    tháng    năm 2026", right_t)

            # Thân văn bản: Đã thay thế "quy chế cũ" -> "quy chế mới 2026"
            body_t = " ".join(p.text for p in out_d.paragraphs)
            self.assertIn("quy chế mới 2026", body_t)
            self.assertNotIn("quy chế cũ", body_t)

    def test_group6_batch_doc_number_formatting(self):
        """Kiểm tra chức năng chỉnh sửa số hiệu văn bản hàng loạt (thay dấu chấm bằng khoảng trống dự thảo)"""
        # 1. Test hàm format_doc_number trực tiếp
        # Trường hợp 1: Dấu chấm sang khoảng trống
        num1 = BatchEditor.format_doc_number("Số: ............../TB-ĐKĐĐ", "dots_to_spaces")
        self.assertEqual(num1, "Số:               /TB-ĐKĐĐ")

        # Trường hợp 2: Số trống (chỉ có /) sang khoảng trống
        num2 = BatchEditor.format_doc_number("Số: /TB-ĐKĐĐ", "dots_to_spaces")
        self.assertEqual(num2, "Số:               /TB-ĐKĐĐ")

        # Trường hợp 3: Khoảng trống sang dấu chấm
        num3 = BatchEditor.format_doc_number("Số:               /TB-ĐKĐĐ", "spaces_to_dots")
        self.assertEqual(num3, "Số: ............../TB-ĐKĐĐ")

        # Trường hợp 4: Đổi ký hiệu mới sau dấu /
        num4 = BatchEditor.format_doc_number("Số: ............../TB-ĐKĐĐ", "dots_to_spaces", new_symbol="QĐ-UBND")
        self.assertEqual(num4, "Số:               /QĐ-UBND")

        # Trường hợp 5: Nhập số hiệu cụ thể
        num5 = BatchEditor.format_doc_number("", "custom", custom_text="125/QĐ-UBND")
        self.assertEqual(num5, "Số: 125/QĐ-UBND")

        # 2. Test thực tế trên tệp DOCX
        bf = os.path.join(self.temp_dir, "test_doc_num_file.docx")
        d = docx.Document()
        t = d.add_table(rows=1, cols=2)
        t.cell(0, 0).text = "ỦY BAN NHÂN DÂN XÃ\nTHUẬN HẠNH\n\nSố: ............../TB-ĐKĐĐ"
        t.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nThuận Hạnh, ngày 05 tháng 11 năm 2026"
        d.add_paragraph("Nội dung thông báo đăng ký đất đai")
        d.save(bf)

        opts = BatchEditOptions(
            enable_doc_number=True,
            doc_number_mode="dots_to_spaces",
            auto_fix_agency_split=True,
            enable_date=True,
            date_mode="draft_spaces"
        )
        res = BatchEditor.process_file(bf, opts)
        self.assertTrue(res.success)

        d_out = docx.Document(bf)
        left_text = d_out.tables[0].cell(0, 0).text
        # Kiểm tra ngắt dòng cơ quan: Dòng 1 ỦY BAN NHÂN DÂN, Dòng 2 XÃ THUẬN HẠNH
        self.assertIn("ỦY BAN NHÂN DÂN", left_text)
        self.assertIn("XÃ THUẬN HẠNH", left_text)
        self.assertNotIn("ỦY BAN NHÂN DÂN XÃ", left_text)
        # Kiểm tra số hiệu: đã thay thế dấu chấm bằng 15 khoảng trắng dự thảo
        self.assertIn("Số:               /TB-ĐKĐĐ", left_text)
        self.assertNotIn("..............", left_text)

    def test_group9_anti_spillover_footer_page_fit(self):
        """Kiểm tra cơ chế chống rớt chữ ký sang trang mới (Anti-Spillover & Page Fit)"""
        bf = os.path.join(self.temp_dir, "test_spill_doc.docx")
        out_f = os.path.join(self.temp_dir, "test_spill_doc_out.docx")

        d = docx.Document()
        t = d.add_table(rows=1, cols=2)
        t.cell(0, 0).text = "ỦY BAN NHÂN DÂN\nXÃ THUẬN HẠNH\n—————\nSố:      /TB-UBND"
        t.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n—————\nThuận Hạnh, ngày 05 tháng 11 năm 2026"

        d.add_paragraph("THÔNG BÁO")
        d.add_paragraph("Về việc niêm yết công khai kết quả kiểm tra hiện trạng sử dụng đất")

        for i in range(1, 9):
            p = d.add_paragraph(f"{i}. Thửa đất số {i*10}, tờ bản đồ số 5, diện tích {i*100} m2, mục đích sử dụng đất trồng cây lâu năm, địa chỉ tại thôn 2, xã Thuận Hạnh, huyện Đắk Song. Người sử dụng đất cam kết không có tranh chấp khiếu nại.")
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

        # Đoạn văn trắng thừa trước Nơi nhận
        d.add_paragraph("")
        d.add_paragraph("Nay Ủy ban nhân dân xã thông báo công khai để nhân dân được biết./.")
        d.add_paragraph("")
        d.add_paragraph("Nơi nhận:\n- Như trên;\n- Lưu: VT.\n\nTM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH\n\n\n\nNguyễn Văn A")
        d.save(bf)

        rep = self.engine.format_document(bf, out_f, standard_id="ND30")
        self.assertTrue(rep.success)

        # Kiểm tra qua Word COM nếu có Word, hoặc kiểm tra thuộc tính DOCX
        d_res = docx.Document(out_f)
        # Bảng Nơi nhận & Chữ ký phải có mặt
        tbl_footer = d_res.tables[-1]
        t_txt = " ".join(c.text for r in tbl_footer.rows for c in r.cells)
        self.assertIn("Nơi nhận:", t_txt)
        self.assertIn("CHỦ TỊCH", t_txt)
        self.assertIn("Nguyễn Văn A", t_txt)

        # Không còn đoạn văn trắng thừa ngay trước bảng Nơi nhận
        prev_node = tbl_footer._tbl.getprevious()
        if prev_node is not None and prev_node.tag.endswith('p'):
            p_prev = docx.text.paragraph.Paragraph(prev_node, d_res)
            self.assertTrue(len(p_prev.text.strip()) > 0)

    # ==================== NHÓM 11: ĐỔI TÊN TỆP HÀNG LOẠT (BATCH FILE RENAMING) ====================
    def test_group11_batch_file_renaming(self):
        """Kiểm tra chức năng đổi tên tệp hàng loạt: Xóa _ChuanHoa, Tìm & Thay thế, Va chạm tên tệp"""
        from core.batch_editor import BatchEditor, BatchRenameOptions

        # 1. Kiểm tra tính toán tên tệp mới: Xóa _ChuanHoa / _chuanhoa / chuẩn hóa
        opts_del = BatchRenameOptions(remove_chuanhoa=True)
        self.assertEqual(BatchEditor.compute_new_filename("Thong_bao_01_ChuanHoa.docx", opts_del), "Thong_bao_01.docx")
        self.assertEqual(BatchEditor.compute_new_filename("CV_123_chuanhoa.docx", opts_del), "CV_123.docx")
        self.assertEqual(BatchEditor.compute_new_filename("VanBan_chuẩn hóa.docx", opts_del), "VanBan.docx")
        self.assertEqual(BatchEditor.compute_new_filename("QD_2026_chuan hoa.docx", opts_del), "QD_2026.docx")
        self.assertEqual(BatchEditor.compute_new_filename("BaoCao_ChuanHoa_KySo.docx", opts_del), "BaoCao_KySo.docx")
        self.assertEqual(BatchEditor.compute_new_filename("VB - ChuanHoa - 2026.docx", opts_del), "VB - 2026.docx")

        # 2. Kiểm tra Tìm kiếm & Thay thế trong tên tệp (Đổi địa danh Thuận Hạnh -> Đức An)
        opts_fr = BatchRenameOptions(
            remove_chuanhoa=True,
            enable_find_replace=True,
            find_text="Thuận Hạnh",
            replace_text="Đức An"
        )
        self.assertEqual(
            BatchEditor.compute_new_filename("Thong_Bao_Thuận Hạnh_ChuanHoa.docx", opts_fr),
            "Thong_Bao_Đức An.docx"
        )

        # 3. Kiểm tra Tiền tố, Hậu tố & Đổi kiểu chữ
        opts_fmt = BatchRenameOptions(
            remove_chuanhoa=True,
            prefix="VB_",
            suffix="_final",
            case_transform="lower"
        )
        self.assertEqual(
            BatchEditor.compute_new_filename("Report_ChuanHoa.docx", opts_fmt),
            "VB_report_final.docx"
        )

        # 4. Kiểm tra Đánh số thứ tự tự động
        opts_num = BatchRenameOptions(
            remove_chuanhoa=True,
            enable_numbering=True,
            numbering_start=1,
            numbering_digits=2,
            numbering_pos="prefix"
        )
        self.assertEqual(
            BatchEditor.compute_new_filename("VanBan_A_ChuanHoa.docx", opts_num, index=0),
            "01_VanBan_A.docx"
        )
        self.assertEqual(
            BatchEditor.compute_new_filename("VanBan_B_ChuanHoa.docx", opts_num, index=1),
            "02_VanBan_B.docx"
        )

        # 5. Kiểm tra thực hiện đổi tên tệp thực tế trên đĩa (execute_rename)
        test_dir = os.path.join(self.temp_dir, "test_rename_dir")
        os.makedirs(test_dir, exist_ok=True)

        f1 = os.path.join(test_dir, "Doc1_ChuanHoa.docx")
        f2 = os.path.join(test_dir, "Doc2_chuanhoa.docx")
        with open(f1, "w", encoding="utf-8") as f: f.write("f1")
        with open(f2, "w", encoding="utf-8") as f: f.write("f2")

        res = BatchEditor.execute_rename([f1, f2], opts_del)
        self.assertEqual(len(res), 2)
        self.assertTrue(res[0].success)
        self.assertTrue(res[1].success)
        self.assertEqual(res[0].new_name, "Doc1.docx")
        self.assertEqual(res[1].new_name, "Doc2.docx")
        self.assertTrue(os.path.exists(os.path.join(test_dir, "Doc1.docx")))
        self.assertTrue(os.path.exists(os.path.join(test_dir, "Doc2.docx")))

        # 6. Kiểm tra giải quyết va chạm tên tệp (auto_index)
        f_dup = os.path.join(test_dir, "Doc1_ChuanHoa.docx")
        with open(f_dup, "w", encoding="utf-8") as f: f.write("dup")
        # Doc1.docx đã tồn tại sẵn, do đó Doc1_ChuanHoa.docx phải được đổi thành Doc1 (1).docx
        res_col = BatchEditor.execute_rename([f_dup], opts_del)
        self.assertTrue(res_col[0].success)
        self.assertEqual(res_col[0].new_name, "Doc1 (1).docx")
        self.assertTrue(os.path.exists(os.path.join(test_dir, "Doc1 (1).docx")))

    # ==================== NHÓM 12: XỬ LÝ ĐỒNG THỜI NỘI DUNG & ĐỔI TÊN TỆP ====================
    def test_group12_combined_batch_content_and_rename(self):
        """Kiểm tra xử lý đồng thời trong 1 lần chạy: vừa chỉnh sửa nội dung vừa đổi tên tệp (Single-pass Combined)"""
        from core.batch_editor import BatchEditor, BatchEditOptions, BatchRenameOptions

        test_dir = os.path.join(self.temp_dir, "test_combined_dir")
        os.makedirs(test_dir, exist_ok=True)

        f_in = os.path.join(test_dir, "ThongBao_ThuanHanh_ChuanHoa.docx")
        d = docx.Document()
        t = d.add_table(rows=1, cols=2)
        t.cell(0, 0).text = "ỦY BAN NHÂN DÂN XÃ\nTHUẬN HẠNH\n\nSố: ............../TB-ĐKĐĐ"
        t.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nThuận Hạnh, ngày 05 tháng 11 năm 2026"
        d.add_paragraph("Nội dung thông báo đất đai tại xã Thuận Hạnh.")
        d.save(f_in)

        edit_opts = BatchEditOptions(
            enable_location=True,
            new_location="Đức An",
            sync_agency_location=True,
            enable_doc_number=True,
            doc_number_mode="dots_to_spaces",
            enable_date=True,
            date_mode="draft_spaces",
            enable_find_replace=True,
            find_text="xã Thuận Hạnh",
            replace_text="xã Đức An"
        )

        rename_opts = BatchRenameOptions(
            remove_chuanhoa=True,
            enable_find_replace=True,
            find_text="ThuanHanh",
            replace_text="DucAn"
        )

        results = BatchEditor.process_combined_batch([f_in], edit_opts, rename_opts)
        self.assertEqual(len(results), 1)
        self.assertTrue(results[0].success)

        expected_new_path = os.path.join(test_dir, "ThongBao_DucAn.docx")
        self.assertEqual(results[0].output_path, expected_new_path)
        self.assertTrue(os.path.exists(expected_new_path))
        self.assertFalse(os.path.exists(f_in), "File gốc cũ phải được dọn dẹp sạch sẽ khi đổi tên cùng thư mục")

        # Kiểm tra nội dung bên trong tệp mới
        doc_out = docx.Document(expected_new_path)
        tbl_out = doc_out.tables[0]
        left_text = tbl_out.cell(0, 0).text
        right_text = tbl_out.cell(0, 1).text
        body_text = " ".join(p.text for p in doc_out.paragraphs)

        # 1. Cơ quan đã chuẩn hóa ngắt dòng: Dòng 1 ỦY BAN NHÂN DÂN, Dòng 2 XÃ ĐỨC AN
        self.assertIn("ỦY BAN NHÂN DÂN", left_text)
        self.assertIn("XÃ ĐỨC AN", left_text)
        self.assertNotIn("ỦY BAN NHÂN DÂN XÃ", left_text)

        # 2. Số hiệu chuyển dấu chấm sang khoảng trống
        self.assertIn("Số:               /TB-ĐKĐĐ", left_text)
        self.assertNotIn("..............", left_text)

        # 3. Địa danh & ngày tháng dự thảo chuẩn NĐ 30
        self.assertIn("Đức An, ngày    tháng    năm 2026", right_text)

        # 4. Tìm kiếm thay thế trong thân văn bản
        self.assertIn("xã Đức An", body_text)
        self.assertNotIn("xã Thuận Hạnh", body_text)


if __name__ == "__main__":
    unittest.main()
