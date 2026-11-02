"""
tests/test_embedded_editor.py: Kiểm thử tự động Phân hệ Soạn thảo & Tạo nhanh Văn bản (Version 1.2)
Bao gồm:
- Khởi tạo khung mẫu OpenXML (OpenXmlBuilder)
- Quản trị bố cục bảng (TableLayoutManager)
- Giám sát ràng buộc pháp lý thời gian thực (LegalConstraintGuard)
- Bộ điều phối khuôn mẫu (TemplateWizardController)
"""

import os
import sys
import unittest
import tempfile
import docx
from docx.shared import Mm, Pt
from docx.oxml.ns import qn

# Thêm đường dẫn gốc dự án vào sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.editor_core.table_layout_manager import TableLayoutManager
from core.editor_core.openxml_builder import OpenXmlBuilder
from core.legal_constraint_guard import LegalConstraintGuard
from core.template_wizard import TemplateWizardController


class TestEmbeddedEditor(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="af_test_editor_")

    def tearDown(self):
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    # ==================== 1. KIỂM THỬ BỐ CỤC BẢNG (TABLE LAYOUT) ====================
    def test_table_layout_manager(self):
        doc = docx.Document()
        table = doc.add_table(rows=3, cols=2)
        
        # Kiểm tra bảng ẩn viền
        TableLayoutManager.make_table_borders_invisible(table)
        tblBorders = table._tbl.tblPr.find(qn('w:tblBorders'))
        self.assertIsNotNone(tblBorders)

        # Kiểm tra thiết lập độ rộng cột
        TableLayoutManager.set_column_widths(table, [68.0, 97.0])
        self.assertEqual(len(table.rows[0].cells), 2)

        # Kiểm tra bảo vệ nhiều trang
        TableLayoutManager.apply_multipage_safety(table)
        trPr0 = table.rows[0]._tr.trPr
        self.assertIsNotNone(trPr0.find(qn('w:tblHeader')))
        trPr1 = table.rows[1]._tr.trPr
        self.assertIsNotNone(trPr1.find(qn('w:cantSplit')))

    # ==================== 2. KIỂM THỬ OPENXML BUILDER (NĐ 30 & HD 36) ====================
    def test_openxml_builder_nd30(self):
        meta = {
            "standard_id": "ND30",
            "superior_agency": "ỦY BAN NHÂN DÂN THÀNH PHỐ HÀ NỘI",
            "issuing_agency": "SỞ TƯ PHÁP",
            "doc_number": "Số: 45/QĐ-STP",
            "location": "Hà Nội",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "Về việc ban hành quy chế làm việc",
            "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
            "signer_position": "GIÁM ĐỐC",
            "signer_name": "Trần Văn Cường"
        }
        doc = OpenXmlBuilder.create_scaffolded_document(meta, lock_scaffolding=True)
        test_path = os.path.join(self.temp_dir, "test_nd30.docx")
        doc.save(test_path)

        # Đọc lại và kiểm tra
        saved_doc = docx.Document(test_path)
        all_text_nodes = saved_doc._element.xpath('//w:t/text()')
        full_text = " ".join(str(t) for t in all_text_nodes)

        self.assertIn("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", full_text)
        self.assertIn("Độc lập - Tự do - Hạnh phúc", full_text)
        self.assertIn("QUYẾT ĐỊNH", full_text)
        self.assertIn("SỞ TƯ PHÁP", full_text)
        self.assertIn("Trần Văn Cường", full_text)

    def test_openxml_builder_hd36(self):
        meta = {
            "standard_id": "HD36",
            "superior_agency": "TỈNH ỦY ĐẮK NÔNG",
            "issuing_agency": "BAN THƯỜNG VỤ",
            "doc_number": "Số 16-QĐ/TU",
            "location": "Đắk Nông",
            "doc_type": "NGHỊ QUYẾT",
            "doc_subject": "Về công tác phát triển đảng viên",
            "authority_prefix": "T/M BAN THƯỜNG VỤ",
            "signer_position": "BÍ THƯ",
            "signer_name": "Nguyễn Thị Mai"
        }
        doc = OpenXmlBuilder.create_scaffolded_document(meta, lock_scaffolding=True)
        test_path = os.path.join(self.temp_dir, "test_hd36.docx")
        doc.save(test_path)

        saved_doc = docx.Document(test_path)
        all_text_nodes = saved_doc._element.xpath('//w:t/text()')
        full_text = " ".join(str(t) for t in all_text_nodes)

        self.assertIn("ĐẢNG CỘNG SẢN VIỆT NAM", full_text)
        self.assertNotIn("Độc lập - Tự do - Hạnh phúc", full_text)
        self.assertIn("NGHỊ QUYẾT", full_text)
        self.assertIn("T/M BAN THƯỜNG VỤ", full_text)

    # ==================== 3. KIỂM THỬ RÀ SOÁT PHÁP LÝ (LEGAL GUARD) ====================
    def test_legal_constraint_guard(self):
        # 1. Văn bản chuẩn 100%
        meta = {
            "standard_id": "ND30",
            "superior_agency": "ỦY BAN NHÂN DÂN",
            "issuing_agency": "THÀNH PHỐ HÀ NỘI",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "Về việc chuyển đổi số"
        }
        doc_compliant = OpenXmlBuilder.create_scaffolded_document(meta)
        res = LegalConstraintGuard.validate_document(doc_compliant, standard_id="ND30")
        self.assertTrue(res["is_compliant"])
        self.assertGreaterEqual(res["score"], 95.0)

        # 2. Văn bản vi phạm lề trái
        doc_bad_margin = docx.Document()
        doc_bad_margin.add_paragraph("Nội dung thử nghiệm")
        doc_bad_margin.sections[0].left_margin = Mm(10.0) # Lề trái chỉ 10mm (yêu cầu 30mm)
        res_bad = LegalConstraintGuard.validate_document(doc_bad_margin, standard_id="ND30")
        self.assertFalse(res_bad["is_compliant"])
        self.assertGreater(res_bad["error_count"], 0)

    # ==================== 4. KIỂM THỬ TEMPLATE WIZARD CONTROLLER ====================
    def test_template_wizard_controller(self):
        controller = TemplateWizardController()
        profiles = controller.get_profiles()
        self.assertGreater(len(profiles), 0, "Không nạp được danh sách hồ sơ cơ quan mẫu")

        types_nd30 = controller.get_available_document_types("ND30")
        self.assertIn("Quyết định (Cá biệt)", types_nd30)
        self.assertIn("Công văn hành chính", types_nd30)

        types_hd36 = controller.get_available_document_types("HD36")
        self.assertIn("Nghị quyết của Đảng", types_hd36)

        # Thử khởi tạo văn bản từ Wizard
        p = profiles[0]
        out_file = os.path.join(self.temp_dir, "wizard_gen.docx")
        doc_gen = controller.generate_document(
            profile_data=p,
            doc_type_display="Quyết định (Cá biệt)",
            doc_subject="Về công tác cán bộ năm 2026",
            output_docx_path=out_file
        )
        self.assertTrue(os.path.exists(out_file))

    # ==================== 5. KIỂM THỬ CANVAS & THAO TÁC BẢNG BIỂU ====================
    def test_canvas_table_and_metadata_editing(self):
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        
        from ui.editor.document_canvas import DocumentCanvas
        from ui.editor.metadata_editor_dlg import MetadataEditorDialog
        from ui.editor.table_properties_dlg import TablePropertiesDialog

        canvas = DocumentCanvas()
        
        # 1. Kiểm tra cập nhật Thể thức
        new_meta = {
            "standard_id": "HD36",
            "superior_agency": "TỈNH ỦY ĐẮK NÔNG",
            "issuing_agency": "ĐẢNG BỘ XÃ THUẬN HẠNH",
            "doc_number": "Số 16-QĐ/ĐU",
            "location": "Thuận Hạnh",
            "date_str": "Thuận Hạnh, ngày 02 tháng 11 năm 2026",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "Phân công các đồng chí Ủy viên BCH Đảng bộ xã",
            "recipients": ["- Như Điều 3;", "- BTV Tỉnh ủy (b/c);", "- Lưu: VP."],
            "authority_prefix": "T/M BAN THƯỜNG VỤ",
            "signer_position": "BÍ THƯ",
            "signer_name": "K'Thanh"
        }
        canvas.set_metadata(new_meta)
        self.assertIn("ĐẢNG BỘ XÃ THUẬN HẠNH", canvas.lbl_header_left.text())
        self.assertIn("ĐẢNG CỘNG SẢN VIỆT NAM", canvas.lbl_header_right.text())
        self.assertIn("Thuận Hạnh, ngày 02 tháng 11 năm 2026", canvas.lbl_header_right.text())
        self.assertIn("K'Thanh", canvas.lbl_footer_right.text())

        # 2. Kiểm tra thao tác Bảng biểu (Thêm/Xóa Hàng/Cột)
        canvas.insert_table(rows=3, cols=3, border_type="grid")
        cur = canvas.body_editor.textCursor()
        tbl = cur.currentTable()
        self.assertIsNotNone(tbl)
        self.assertEqual(tbl.rows(), 3)
        self.assertEqual(tbl.columns(), 3)

        # Thêm hàng dưới
        canvas.insert_table_row_below()
        self.assertEqual(tbl.rows(), 4)

        # Xóa hàng
        canvas.delete_table_row()
        self.assertEqual(tbl.rows(), 3)

        # Thêm cột phải
        canvas.insert_table_col_right()
        self.assertEqual(tbl.columns(), 4)

        # Xóa cột
        canvas.delete_table_col()
        self.assertEqual(tbl.columns(), 3)

        # 3. Kiểm tra danh sách & ngắt trang
        canvas.insert_bullet_list()
        canvas.insert_numbered_list()
        canvas.insert_page_break()
        self.assertIn("NGẮT TRANG", canvas.body_editor.toHtml())

        # 4. Kiểm tra hộp thoại cấu hình bảng
        dlg_tbl = TablePropertiesDialog()
        cfg = dlg_tbl.get_table_config()
        self.assertIn("rows", cfg)
        self.assertIn("cols", cfg)
        self.assertIn("border_type", cfg)
        dlg_tbl.close()


if __name__ == "__main__":
    unittest.main()
