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


if __name__ == "__main__":
    unittest.main()
