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
        # 1. Văn bản chuẩn 100% đầy đủ số hiệu và ngày tháng
        meta = {
            "standard_id": "ND30",
            "superior_agency": "ỦY BAN NHÂN DÂN",
            "issuing_agency": "THÀNH PHỐ HÀ NỘI",
            "doc_number": "Số: 10/QĐ-UBND",
            "date_str": "Hà Nội, ngày 15 tháng 10 năm 2026",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "Về việc chuyển đổi số"
        }
        doc_compliant = OpenXmlBuilder.create_scaffolded_document(meta)
        res = LegalConstraintGuard.validate_document(doc_compliant, standard_id="ND30")
        self.assertTrue(res["is_compliant"])
        self.assertGreaterEqual(res["score"], 95.0)

        # 2. Văn bản vi phạm xung đột loại văn bản (TỜ TRÌNH chứa lệnh QUYẾT ĐỊNH và Điều 1)
        doc_mismatch = docx.Document()
        sec = doc_mismatch.sections[0]
        sec.top_margin = Mm(20.0)
        sec.bottom_margin = Mm(20.0)
        sec.left_margin = Mm(30.0)
        sec.right_margin = Mm(15.0)
        doc_mismatch.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        doc_mismatch.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        doc_mismatch.add_paragraph("..., ngày ... tháng ... năm 2026")
        doc_mismatch.add_paragraph("QĐ-UBND")
        doc_mismatch.add_paragraph("TỜ TRÌNH")
        doc_mismatch.add_paragraph("Về việc 0")
        doc_mismatch.add_paragraph("QUYẾT ĐỊNH:")
        doc_mismatch.add_paragraph("Điều 1. Ban hành kế hoạch...")
        res_mismatch = LegalConstraintGuard.validate_document(doc_mismatch, standard_id="ND30")
        self.assertFalse(res_mismatch["is_compliant"])
        self.assertLess(res_mismatch["score"], 60.0)
        rule_ids = [v["rule_id"] for v in res_mismatch["violations"]]
        self.assertIn("DOC_TYPE_STRUCTURE_MISMATCH", rule_ids)
        self.assertIn("PLACEHOLDER_DATE", rule_ids)
        self.assertIn("PLACEHOLDER_SUBJECT", rule_ids)

        # 3. Văn bản vi phạm lề trái
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
        self.assertIn("Nghị quyết Đảng ủy", types_hd36)

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
        from PyQt6.QtWidgets import QApplication, QLabel
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

        # 3. Kiểm tra danh sách & ngắt trang (phân tách trang rời rạc thực tế)
        canvas.insert_bullet_list()
        canvas.insert_numbered_list()
        canvas.insert_page_break()
        self.assertEqual(len(canvas.pages), 2)

        # 4. Kiểm tra hộp thoại cấu hình bảng
        dlg_tbl = TablePropertiesDialog()
        cfg = dlg_tbl.get_table_config()
        self.assertIn("rows", cfg)
        self.assertIn("cols", cfg)
        self.assertIn("border_type", cfg)
        dlg_tbl.close()

    def test_multipage_canvas_and_landscape_sections(self):
        """Kiểm tra phân trang rời rạc và xoay ngang trang tiếp theo độc lập"""
        from PyQt6.QtWidgets import QApplication, QLabel
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas

        canvas = DocumentCanvas()
        self.assertEqual(len(canvas.pages), 1)
        self.assertEqual(canvas.pages[0].orientation, "portrait")

        # Thêm trang 2 và xoay sang khổ ngang (Landscape)
        p2 = canvas.add_page(orientation="landscape")
        self.assertEqual(len(canvas.pages), 2)
        self.assertEqual(canvas.pages[0].orientation, "portrait")
        self.assertEqual(canvas.pages[1].orientation, "landscape")

        p2.editor.setText("Nội dung Phụ lục bảng biểu khổ ngang")
        paras = canvas.get_body_paragraphs()
        self.assertIn("[PAGE_BREAK_LANDSCAPE]", paras)

        # Xuất OpenXML docx và kiểm tra 2 Section độc lập
        meta = dict(canvas.metadata)
        meta["body_paragraphs"] = paras
        doc = OpenXmlBuilder.create_scaffolded_document(meta)
        self.assertEqual(len(doc.sections), 2)
        self.assertEqual(doc.sections[0].orientation, docx.enum.section.WD_ORIENT.PORTRAIT)
        self.assertEqual(doc.sections[1].orientation, docx.enum.section.WD_ORIENT.LANDSCAPE)

        # Xóa trang 2
        canvas.remove_page(1)
        self.assertEqual(len(canvas.pages), 1)

    def test_paragraph_spacing_and_indent_controls(self):
        """Kiểm tra chức năng giãn đoạn (xóa khoảng thừa trên/dưới) và thụt lề chuẩn Word"""
        from PyQt6.QtWidgets import QApplication, QLabel
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas

        canvas = DocumentCanvas()
        ed = canvas.active_editor
        ed.setPlainText("Đoạn văn kiểm thử giãn dòng và khoảng cách đoạn.")

        # 1. Thêm khoảng cách trên & dưới
        canvas.add_space_before(6.0)
        canvas.add_space_after(6.0)
        fmt = ed.textCursor().blockFormat()
        self.assertAlmostEqual(fmt.topMargin(), 6.0 * 1.333, delta=0.5)
        self.assertAlmostEqual(fmt.bottomMargin(), 6.0 * 1.333, delta=0.5)

        # 2. Xóa khoảng thừa trên
        canvas.remove_space_before()
        fmt = ed.textCursor().blockFormat()
        self.assertEqual(fmt.topMargin(), 0.0)
        self.assertAlmostEqual(fmt.bottomMargin(), 6.0 * 1.333, delta=0.5)

        # 3. Xóa sạch khoảng thừa trên & dưới
        canvas.remove_all_spacing()
        fmt = ed.textCursor().blockFormat()
        self.assertEqual(fmt.topMargin(), 0.0)
        self.assertEqual(fmt.bottomMargin(), 0.0)

        # 4. Tăng & giảm thụt lề (Increase/Decrease Indent)
        canvas.increase_indent(10.0)
        fmt = ed.textCursor().blockFormat()
        self.assertGreater(fmt.leftMargin(), 30)

        canvas.decrease_indent(10.0)
        fmt = ed.textCursor().blockFormat()
        self.assertEqual(fmt.leftMargin(), 0.0)

        # 5. Thụt dòng đầu tiên (First Line Indent)
        canvas.set_first_line_indent(10.0)
        fmt = ed.textCursor().blockFormat()
        self.assertAlmostEqual(fmt.textIndent(), 37.795, delta=1.0)

        # 6. Chuẩn hóa Nghị định 30
        canvas.apply_nd30_spacing()
        fmt = ed.textCursor().blockFormat()
        self.assertEqual(fmt.topMargin(), 0.0)
        self.assertAlmostEqual(fmt.bottomMargin(), 6.0 * 1.333, delta=0.5)
        self.assertAlmostEqual(fmt.textIndent(), 37.795, delta=1.0)

    def test_word_editor_window_and_ribbon_initialization(self):
        """Đảm bảo khởi tạo WordEditorWindow và AdministrativeRibbonBar thành công không lỗi NameError QMenu"""
        from PyQt6.QtWidgets import QApplication, QLabel
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.administrative_ribbon import AdministrativeRibbonBar
        from ui.editor.word_editor_window import WordEditorWindow

        ribbon = AdministrativeRibbonBar()
        self.assertIsNotNone(ribbon.btn_first_line.menu())
        self.assertIsNotNone(ribbon.btn_paragraph_spacing.menu())

        editor_win = WordEditorWindow(None)
        self.assertIsNotNone(editor_win.ribbon)
        self.assertIsNotNone(editor_win.canvas)
        editor_win.close()

    def test_canvas_apply_nd30_alignment_and_leading_spaces_removal(self):
        """Kiểm tra chức năng căn chuẩn NĐ 30: xóa bỏ khoảng trắng thụt lề thủ công và căn đều 2 bên"""
        from PyQt6.QtWidgets import QApplication, QLabel
        from PyQt6.QtCore import Qt
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas

        canvas = DocumentCanvas()
        ed = canvas.active_editor
        # Thiết lập văn bản bị lệch dòng và có khoảng trắng gõ tay ở đầu dòng
        ed.setPlainText("   Căn cứ Luật Tổ chức chính quyền địa phương;\n   QUYẾT ĐỊNH:\n   Điều 1. Ban hành kế hoạch.")

        res = canvas.apply_nd30_standards_and_alignment()
        self.assertGreater(res["fixed_paragraphs"], 0)
        self.assertGreater(res["stripped_spaces"], 0)

        # Kiểm tra văn bản sau khi xóa khoảng trắng đầu dòng
        plain_text = ed.toPlainText()
        lines = plain_text.splitlines()
        self.assertFalse(lines[0].startswith(" "))
        self.assertFalse(lines[1].startswith(" "))
        self.assertFalse(lines[2].startswith(" "))

        # Kiểm tra căn chỉnh
        doc = ed.document()
        b0 = doc.findBlockByLineNumber(0)
        b1 = doc.findBlockByLineNumber(1)
        b2 = doc.findBlockByLineNumber(2)
        # b1 là "QUYẾT ĐỊNH:" nên phải căn giữa
        self.assertEqual(b1.blockFormat().alignment(), Qt.AlignmentFlag.AlignCenter)
        # b2 là "Điều 1..." nên phải căn đều 2 bên và có textIndent
        self.assertEqual(b2.blockFormat().alignment(), Qt.AlignmentFlag.AlignJustify)
        self.assertGreater(b2.blockFormat().textIndent(), 0.0)

    def test_native_numbered_list_generation(self):
        """Kiểm tra danh sách đánh số tự động chuẩn Word (1, 2, 3...) sử dụng QTextListFormat"""
        from PyQt6.QtWidgets import QApplication, QLabel
        from PyQt6.QtGui import QTextListFormat
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas

        canvas = DocumentCanvas()
        ed = canvas.active_editor
        ed.setPlainText("Mục thử nghiệm")

        # 1. Kích hoạt đánh số
        canvas.insert_numbered_list()
        qlist = ed.textCursor().currentList()
        self.assertIsNotNone(qlist, "Không tạo được đối tượng QTextList")
        self.assertEqual(qlist.format().style(), QTextListFormat.Style.ListDecimal)

        # 2. Bấm lại nút để toggle off
        canvas.insert_numbered_list()
        self.assertIsNone(ed.textCursor().currentList(), "Không thoát được danh sách khi bấm toggle")

    def test_metadata_dialog_dynamic_switching(self):
        """Kiểm tra chuyển đổi động NĐ 30 vs HD 36 trên form Chỉnh sửa Thể thức"""
        from PyQt6.QtWidgets import QApplication, QLabel
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.metadata_editor_dlg import MetadataEditorDialog

        dlg = MetadataEditorDialog(None, metadata={
            "standard_id": "ND30",
            "superior_agency": "ỦY BAN NHÂN DÂN HUYỆN",
            "issuing_agency": "ỦY BAN NHÂN DÂN XÃ"
        })
        self.assertTrue(dlg.rb_nd30.isChecked())

        # Chuyển sang Đảng
        dlg.rb_hd36.setChecked(True)
        self.assertEqual(dlg.txt_sup.text(), "HUYỆN ỦY")
        self.assertEqual(dlg.txt_iss.text(), "ĐẢNG ỦY XÃ")
        self.assertEqual(dlg.txt_auth.text(), "T/M ĐẢNG ỦY")
        self.assertEqual(dlg.txt_pos.text(), "BÍ THƯ")
        self.assertIn("NGHỊ QUYẾT ĐẢNG", [dlg.cb_type.itemText(i) for i in range(dlg.cb_type.count())])

        # Chuyển lại Nhà nước
        dlg.rb_nd30.setChecked(True)
        self.assertEqual(dlg.txt_sup.text(), "ỦY BAN NHÂN DÂN HUYỆN")
        self.assertEqual(dlg.txt_iss.text(), "ỦY BAN NHÂN DÂN XÃ")
        self.assertEqual(dlg.txt_auth.text(), "TM. ỦY BAN NHÂN DÂN")
        self.assertEqual(dlg.txt_pos.text(), "CHỦ TỊCH")

    def test_compliance_audit_dialog_and_quick_fix(self):
        """Kiểm tra hộp thoại rà soát thể thức với nút sửa nhanh ngày tháng và điểm số nguyên"""
        from PyQt6.QtWidgets import QApplication, QLabel
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas
        from ui.editor.compliance_audit_dlg import ComplianceAuditDialog

        canvas = DocumentCanvas()
        canvas.metadata["date_str"] = "..., ngày ... tháng ... năm 2026"

        audit_res = {
            "score": 55.4,
            "error_count": 2,
            "warning_count": 1,
            "is_compliant": False,
            "checklist": [{"name": "Lề in", "passed": True, "detail": "Đạt"}],
            "violations": [
                {
                    "rule_id": "PLACEHOLDER_DATE",
                    "severity": "error",
                    "message": "Ngày tháng còn chứa ...",
                    "suggestion": "Điền ngày thực tế"
                }
            ]
        }

        dlg = ComplianceAuditDialog(None, audit_result=audit_res, canvas=canvas)
        # Sửa nhanh ngày tháng
        dlg._on_quick_fix(audit_res["violations"][0])
        self.assertNotIn("...", canvas.metadata["date_str"])
        # Kiểm tra không có emoji trong nút và checklist dùng [ĐẠT]
        self.assertNotIn("⚡", dlg.btn_auto_all.text())
        self.assertNotIn("✓", dlg.scroll_widget.findChild(QLabel).text())
        dlg.close()

    def test_signer_name_boundary_check(self):
        """Kiểm tra không báo lỗi người ký mẫu nếu tên người ký là 'Nguyễn Văn Anf', 'Nguyễn Văn An'"""
        import docx
        from docx.shared import Mm
        from core.legal_constraint_guard import LegalConstraintGuard

        # 1. Tên người ký là 'Nguyễn Văn Anf' -> Không được báo lỗi SAMPLE_SIGNER_NAME
        doc1 = docx.Document()
        doc1.sections[0].top_margin = Mm(20)
        doc1.sections[0].bottom_margin = Mm(20)
        doc1.sections[0].left_margin = Mm(30)
        doc1.sections[0].right_margin = Mm(20)
        doc1.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        doc1.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        doc1.add_paragraph("Hà Nội, ngày 04 tháng 11 năm 2026")
        doc1.add_paragraph("QUYẾT ĐỊNH")
        doc1.add_paragraph("Về việc khen thưởng cán bộ xuất sắc")
        doc1.add_paragraph("QUYẾT ĐỊNH:")
        doc1.add_paragraph("Điều 1. Khen thưởng đồng chí Nguyễn Văn Anf.")
        doc1.add_paragraph("CHỦ TỊCH")
        doc1.add_paragraph("Nguyễn Văn Anf")

        res1 = LegalConstraintGuard.validate_document(doc1, standard_id="ND30")
        rule_ids1 = [v["rule_id"] for v in res1["violations"]]
        self.assertNotIn("SAMPLE_SIGNER_NAME", rule_ids1, "Bắt nhầm 'Nguyễn Văn Anf' thành người ký mặc định 'Nguyễn Văn A'")

        # 2. Tên người ký là 'Nguyễn Văn An' -> Không được báo lỗi SAMPLE_SIGNER_NAME
        doc2 = docx.Document()
        doc2.sections[0].top_margin = Mm(20)
        doc2.sections[0].bottom_margin = Mm(20)
        doc2.sections[0].left_margin = Mm(30)
        doc2.sections[0].right_margin = Mm(20)
        doc2.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        doc2.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        doc2.add_paragraph("Hà Nội, ngày 04 tháng 11 năm 2026")
        doc2.add_paragraph("QUYẾT ĐỊNH")
        doc2.add_paragraph("Về việc phân công công tác")
        doc2.add_paragraph("QUYẾT ĐỊNH:")
        doc2.add_paragraph("Điều 1. Phân công công tác đồng chí.")
        doc2.add_paragraph("CHỦ TỊCH")
        doc2.add_paragraph("Nguyễn Văn An")

        res2 = LegalConstraintGuard.validate_document(doc2, standard_id="ND30")
        rule_ids2 = [v["rule_id"] for v in res2["violations"]]
        self.assertNotIn("SAMPLE_SIGNER_NAME", rule_ids2, "Bắt nhầm 'Nguyễn Văn An' thành người ký mặc định 'Nguyễn Văn A'")

        # 3. Tên người ký chính xác là 'Nguyễn Văn A' -> Phải cảnh báo SAMPLE_SIGNER_NAME
        doc3 = docx.Document()
        doc3.sections[0].top_margin = Mm(20)
        doc3.sections[0].bottom_margin = Mm(20)
        doc3.sections[0].left_margin = Mm(30)
        doc3.sections[0].right_margin = Mm(20)
        doc3.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        doc3.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        doc3.add_paragraph("Hà Nội, ngày 04 tháng 11 năm 2026")
        doc3.add_paragraph("QUYẾT ĐỊNH")
        doc3.add_paragraph("Về việc kiện toàn tổ chức")
        doc3.add_paragraph("QUYẾT ĐỊNH:")
        doc3.add_paragraph("Điều 1. Kiện toàn tổ chức.")
        doc3.add_paragraph("CHỦ TỊCH")
        doc3.add_paragraph("Nguyễn Văn A")

        res3 = LegalConstraintGuard.validate_document(doc3, standard_id="ND30")
        rule_ids3 = [v["rule_id"] for v in res3["violations"]]
        self.assertIn("SAMPLE_SIGNER_NAME", rule_ids3, "Không phát hiện tên người ký mẫu mặc định 'Nguyễn Văn A'")

    def test_metadata_dialog_scroll_and_no_emojis(self):
        """Kiểm tra MetadataEditorDialog có vùng cuộn và hoàn toàn không chứa emoji rối mắt"""
        from PyQt6.QtWidgets import QApplication, QLabel, QScrollArea
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.metadata_editor_dlg import MetadataEditorDialog

        dlg = MetadataEditorDialog(None, metadata={"standard_id": "ND30"})
        # Đảm bảo có QScrollArea
        self.assertIsInstance(dlg.scroll_area, QScrollArea)
        # Đảm bảo không có emoji trong các nút radio chọn quy chuẩn
        self.assertNotIn("🏛️", dlg.rb_nd30.text())
        self.assertNotIn("⭐", dlg.rb_hd36.text())
        # Đảm bảo không có emoji trong gợi ý
        self.assertNotIn("🏛️", dlg.lbl_std_hint.text())
        self.assertNotIn("⭐", dlg.lbl_std_hint.text())
        dlg.close()

    def test_metadata_dialog_smart_datepicker_and_combobox(self):
        """Kiểm tra Bộ chọn Lịch ngày tháng thông minh và QComboBox chống nền đen"""
        from PyQt6.QtWidgets import QApplication, QLabel, QListView
        from PyQt6.QtCore import QDate
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.metadata_editor_dlg import MetadataEditorDialog

        # 1. Khởi tạo với ngày cụ thể
        dlg1 = MetadataEditorDialog(None, metadata={
            "location": "Thuận Hạnh",
            "date_str": "Thuận Hạnh, ngày 04 tháng 11 năm 2026"
        })
        self.assertTrue(dlg1.rb_date_specific.isChecked())
        self.assertEqual(dlg1.date_picker.date(), QDate(2026, 11, 4))
        self.assertIn("ngày 04 tháng 11 năm 2026", dlg1._get_current_formatted_date())
        # QComboBox view phải là QListView
        self.assertIsInstance(dlg1.cb_type.view(), QListView)
        dlg1.close()

        # 2. Khởi tạo với ngày dự thảo chờ ký (...)
        dlg2 = MetadataEditorDialog(None, metadata={
            "location": "Xã",
            "date_str": "..., ngày ... tháng ... năm 2026"
        })
        self.assertTrue(dlg2.rb_date_draft.isChecked())
        self.assertFalse(dlg2.date_picker.isEnabled())
        self.assertIn("ngày ... tháng ... năm 2026", dlg2._get_current_formatted_date())

        # 3. Chuyển sang ngày cụ thể và chọn ngày 5 tháng 2 năm 2026 (yêu cầu số 0 phía trước per NĐ 30)
        dlg2.rb_date_specific.setChecked(True)
        dlg2.txt_loc.setText("Hà Nội")
        dlg2.date_picker.setDate(QDate(2026, 2, 5))
        formatted = dlg2._get_current_formatted_date()
        self.assertEqual(formatted, "Hà Nội, ngày 05 tháng 02 năm 2026")
        dlg2.close()

    def test_template_switching_and_quick_fix_structure(self):
        """Kiểm tra chuyển đổi mẫu văn bản và chức năng sửa nhanh cấu trúc tương ứng"""
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas
        from ui.editor.compliance_audit_dlg import ComplianceAuditDialog

        canvas = DocumentCanvas()
        # Chuyển mẫu sang BÁO CÁO
        tpl_bc = canvas.get_template_for_type("BÁO CÁO")
        canvas.metadata.update(tpl_bc)
        canvas.pages[0].editor.setHtml(tpl_bc["body_html"])
        self.assertEqual(canvas.metadata["doc_type"], "BÁO CÁO")
        self.assertIn("KẾT QUẢ ĐẠT ĐƯỢC", canvas.pages[0].editor.toPlainText())

        # Tạo tình huống lỗi: Loại văn bản là BÁO CÁO nhưng body bị lẫn QUYẾT ĐỊNH
        canvas.pages[0].editor.setPlainText("QUYẾT ĐỊNH: Điều 1. Thi hành.")
        dlg = ComplianceAuditDialog(None, audit_result={"violations": [{"rule_id": "DOC_TYPE_STRUCTURE_MISMATCH"}]}, canvas=canvas)
        # Sửa nhanh
        dlg._on_quick_fix({"rule_id": "DOC_TYPE_STRUCTURE_MISMATCH"})
        # Phải nạp đúng mẫu BÁO CÁO
        self.assertIn("KẾT QUẢ ĐẠT ĐƯỢC", canvas.pages[0].editor.toPlainText())
        dlg.close()

    def test_metadata_date_mode_switching_and_today(self):
        """Kiểm tra chuyển đổi linh hoạt giữa Ngày cụ thể và Chờ ký trong MetadataEditorDialog"""
        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtCore import QDate
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.metadata_editor_dlg import MetadataEditorDialog

        # Khởi tạo ở chế độ dự thảo chờ ký
        dlg = MetadataEditorDialog(None, metadata={
            "location": "Xã",
            "date_str": "..., ngày ... tháng ... năm 2026"
        })
        self.assertTrue(dlg.rb_date_draft.isChecked())
        self.assertFalse(dlg.date_picker.isEnabled())
        self.assertTrue(dlg.btn_today.isEnabled())  # Nút Hôm nay luôn sẵn sàng

        # Nhấp nút Hôm nay -> phải tự động bật Ngày cụ thể và kích hoạt date_picker
        dlg.btn_today.click()
        self.assertTrue(dlg.rb_date_specific.isChecked())
        self.assertFalse(dlg.rb_date_draft.isChecked())
        self.assertTrue(dlg.date_picker.isEnabled())
        self.assertEqual(dlg.date_picker.date(), QDate.currentDate())
        self.assertNotIn("...", dlg._get_current_formatted_date())

        # Chuyển lại sang Chờ ký
        dlg.rb_date_draft.setChecked(True)
        self.assertTrue(dlg.rb_date_draft.isChecked())
        self.assertFalse(dlg.date_picker.isEnabled())

        # Tương tác với DatePicker (click / focus / dateChanged) -> phải tự động quay lại Ngày cụ thể
        dlg._on_date_picker_interacted()
        self.assertTrue(dlg.rb_date_specific.isChecked())
        self.assertTrue(dlg.date_picker.isEnabled())
        dlg.close()

    def test_party_resolution_punctuation_and_quick_fix(self):
        """Kiểm tra dấu câu chuẩn mực cho Nghị quyết Đảng và cơ chế Sửa nhanh Căn cứ pháp lý"""
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas
        from ui.editor.compliance_audit_dlg import ComplianceAuditDialog

        canvas = DocumentCanvas()
        # 1. Kiểm tra mẫu NGHỊ QUYẾT ĐẢNG
        tpl = canvas.get_template_for_type("NGHỊ QUYẾT ĐẢNG")
        self.assertIn("Căn cứ Điều lệ Đảng Cộng sản Việt Nam;", tpl["body_html"])
        self.assertIn("Căn cứ Quy chế làm việc của Ban Chấp hành Đảng bộ xã;", tpl["body_html"])
        self.assertNotIn("Căn cứ Quy chế làm việc của Ban Chấp hành Đảng bộ xã,", tpl["body_html"])

        # 2. Giả lập văn bản có dòng Căn cứ kết thúc sai bằng dấu phẩy
        canvas.pages[0].editor.setHtml(
            "<p style='font-style: italic;'>Căn cứ Luật Ban hành văn bản quy phạm pháp luật,</p>"
            "<p style='font-style: italic;'>Theo đề nghị của Công chức Văn phòng;</p>"
            "<p>QUYẾT ĐỊNH:</p>"
        )

        # 3. Kích hoạt chuẩn hóa Nghị định 30
        res = canvas.apply_nd30_standards_and_alignment()
        body_text = canvas.pages[0].editor.toPlainText()

        # Dòng Căn cứ phải được tự động sửa thành dấu chấm phẩy ';'
        self.assertIn("Căn cứ Luật Ban hành văn bản quy phạm pháp luật;", body_text)
        # Dòng Theo đề nghị phải được tự động sửa thành dấu phẩy ','
        self.assertIn("Theo đề nghị của Công chức Văn phòng,", body_text)

        # 4. Kiểm tra Sửa nhanh thông qua ComplianceAuditDialog
        dlg = ComplianceAuditDialog(None, audit_result={
            "violations": [{"rule_id": "LEGAL_GROUNDS_SYNTAX"}]
        }, canvas=canvas)
        refresh_called = []
        dlg.audit_refresh_requested.connect(lambda: refresh_called.append(True))
        dlg._on_quick_fix({"rule_id": "LEGAL_GROUNDS_SYNTAX"})
        self.assertTrue(refresh_called)
        dlg.close()

    def test_editor_action_buttons(self):
        """Kiểm tra đầy đủ 3 nút hành động tinh gọn trên thanh tiêu đề: Lưu văn bản, Xuất PDF, Ký số"""
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.word_editor_window import WordEditorWindow

        win = WordEditorWindow(None)
        self.assertIsNotNone(win.btn_save)
        self.assertEqual(win.btn_save.text(), "Lưu văn bản")

        self.assertIsNotNone(win.btn_export_pdf)
        self.assertEqual(win.btn_export_pdf.text(), "Xuất PDF")

        self.assertIsNotNone(win.btn_sign)
        self.assertEqual(win.btn_sign.text(), "Ký số")
        win.close()

    def test_load_document_preserves_metadata_and_avoids_duplicate_headers(self):
        """Kiểm tra việc mở file DOCX có sẵn trong trình soạn thảo:
        - Tự động bóc tách và nạp đúng metadata (UBND TỈNH QUẢNG NINH, SỞ NỘI VỤ, BÁO CÁO)
        - Không bị chèn đè khung mẫu mặc định (QUYẾT ĐỊNH / ỦY BAN NHÂN DÂN XÃ)
        - Thân bài chỉ chứa nội dung văn bản thực tế, không bị nhân đôi tiêu đề / trích yếu
        """
        sample_docx = "Tai_Lieu_Mau_Test/Van_Ban_Hanh_Chinh_Mau_Loi_ChuanHoa.docx"
        if not os.path.exists(sample_docx):
            self.skipTest("Không tìm thấy tệp DOCX mẫu kiểm tra.")

        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.word_editor_window import WordEditorWindow

        win = WordEditorWindow(None, current_file_path=sample_docx)

        # 1. Kiểm tra metadata nạp chuẩn
        self.assertEqual(win.canvas.metadata.get("doc_type"), "BÁO CÁO")
        self.assertEqual(win.canvas.metadata.get("superior_agency"), "UBND TỈNH QUẢNG NINH")
        self.assertEqual(win.canvas.metadata.get("issuing_agency"), "SỞ NỘI VỤ")
        self.assertEqual(win.canvas.metadata.get("doc_number"), "Số: 45/SNV-VP")
        self.assertEqual(win.canvas.metadata.get("signer_name"), "Nguyễn Văn An")
        self.assertEqual(win.canvas.metadata.get("signer_position"), "PHÓ CHỦ TỊCH")

        # 2. Kiểm tra hiển thị Scaffolding
        self.assertEqual(win.canvas.lbl_doc_type.text(), "BÁO CÁO")
        self.assertIn("SỞ NỘI VỤ", win.canvas.lbl_header_left.text())
        self.assertNotIn("ỦY BAN NHÂN DÂN XÃ", win.canvas.lbl_header_left.text())

        # 3. Kiểm tra thân bài không chứa tiêu đề mẫu hoặc nhân đôi
        body_text = win.canvas.body_editor.toPlainText()
        self.assertNotIn("QUYẾT ĐỊNH", body_text)
        self.assertNotIn("Kế hoạch phát triển kinh tế", body_text)
        self.assertTrue(body_text.strip().startswith("Thực hiện ý kiến chỉ đạo"))
        win.close()

    def test_table_properties_dialog_layout_and_config(self):
        """Kiểm tra hộp thoại cấu hình bảng (TablePropertiesDialog):
        - Kích thước rộng rãi, không bị đè/squished (min width >= 460, min height >= 400)
        - Bố cục lưới 2 cột hàng/cột thoáng đãng
        - Lấy đúng cấu hình get_table_config()
        """
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.table_properties_dlg import TablePropertiesDialog

        dlg = TablePropertiesDialog()
        self.assertGreaterEqual(dlg.minimumWidth(), 460)
        self.assertGreaterEqual(dlg.minimumHeight(), 400)
        self.assertIsNotNone(dlg.sp_rows)
        self.assertIsNotNone(dlg.sp_cols)
        self.assertEqual(dlg.sp_rows.value(), 4)
        self.assertEqual(dlg.sp_cols.value(), 3)

        cfg = dlg.get_table_config()
        self.assertEqual(cfg["rows"], 4)
        self.assertEqual(cfg["cols"], 3)
        self.assertEqual(cfg["border_type"], "grid")
        self.assertTrue(cfg["repeat_header"])
        self.assertTrue(cfg["cant_split"])
        dlg.close()

    def test_top_action_bar_undo_redo_and_rename(self):
        """Kiểm tra thanh tiêu đề phía trên (Top Action Bar):
        - Có nút Quay lại (Undo) và Làm lại (Redo) giống Microsoft Word
        - Có phím tắt Ctrl+Z, Ctrl+Y
        - Nhấp vào tiêu đề hoặc nút Đổi tên kích hoạt chức năng đổi tên tệp
        """
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.word_editor_window import WordEditorWindow

        win = WordEditorWindow(None)
        self.assertIsNotNone(win.btn_undo)
        self.assertEqual(win.btn_undo.text(), "Quay lại")
        self.assertIsNotNone(win.btn_redo)
        self.assertEqual(win.btn_redo.text(), "Làm lại")
        self.assertIsNotNone(win.btn_rename_file)
        self.assertEqual(win.btn_rename_file.text(), "Đổi tên")

        # Kiểm tra phím tắt
        self.assertIsNotNone(win.shortcut_undo)
        self.assertIsNotNone(win.shortcut_redo)

        # Kiểm tra Undo / Redo hoạt động không lỗi
        win._on_undo()
        win._on_redo()

        win.close()

    def test_numbering_continuation_and_multilevel_list(self):
        """Kiểm tra cơ chế đánh số thông minh và phân cấp đa cấp:
        - Bấm từng đoạn riêng lẻ sẽ TỰ ĐỘNG NỐI TIẾP số thứ tự (1, 2, 3...) thay vì toàn ra số 1
        - Hỗ trợ số La Mã hoa (I., II., III...)
        - Hỗ trợ điểm quy chuẩn NĐ 30 (a), b), c)...)
        - Phím Tab tăng cấp phân cấp (I. -> 1. -> a) -> -)
        - Phím Shift+Tab giảm cấp phân cấp và thoát danh sách
        - Phím Enter trên dòng rỗng tự động thoát danh sách
        """
        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtGui import QTextCursor, QTextListFormat, QKeyEvent
        from PyQt6.QtCore import Qt
        app = QApplication.instance() or QApplication(sys.argv)
        from ui.editor.document_canvas import DocumentCanvas

        canvas = DocumentCanvas()
        ed = canvas.active_editor
        ed.setPlainText("Đoạn thứ nhất\nĐoạn thứ hai\nĐoạn thứ ba\nĐoạn thứ tư")

        # 1. Bấm số đoạn 0: Thành số 1
        cur0 = QTextCursor(ed.document().findBlockByNumber(0))
        ed.setTextCursor(cur0)
        canvas.insert_numbered_list()
        b0 = ed.document().findBlockByNumber(0)
        self.assertIsNotNone(b0.textList())
        self.assertEqual(b0.textList().itemNumber(b0), 0) # 0-indexed = Số 1

        # 2. Bấm số đoạn 1: Phải tự động nối thành số 2 (KHÔNG BỊ LỖI TOÀN RA SỐ 1)
        cur1 = QTextCursor(ed.document().findBlockByNumber(1))
        ed.setTextCursor(cur1)
        canvas.insert_numbered_list()
        b1 = ed.document().findBlockByNumber(1)
        self.assertIsNotNone(b1.textList())
        self.assertEqual(b1.textList().itemNumber(b1), 1) # Số 2

        # 3. Bấm số đoạn 2: Phải tự động nối thành số 3
        cur2 = QTextCursor(ed.document().findBlockByNumber(2))
        ed.setTextCursor(cur2)
        canvas.insert_numbered_list()
        b2 = ed.document().findBlockByNumber(2)
        self.assertIsNotNone(b2.textList())
        self.assertEqual(b2.textList().itemNumber(b2), 2) # Số 3

        # 4. Kiểm tra phân cấp bằng Tab (Demote): chuyển đoạn 2 sang Cấp 2
        ed._demote_current_list_item(cur2)
        self.assertEqual(b2.textList().format().indent(), 2)

        # 5. Kiểm tra Shift+Tab (Promote): chuyển đoạn 2 trở lại Cấp 1
        ed._promote_current_list_item(cur2)
        self.assertEqual(b2.textList().format().indent(), 1)

        # 6. Kiểm tra Shift+Tab ở Cấp 1: Thoát khỏi danh sách
        ed._promote_current_list_item(cur2)
        self.assertIsNone(b2.textList())

        # 7. Kiểm tra kiểu số La Mã hoa (I., II., III...)
        cur3 = QTextCursor(ed.document().findBlockByNumber(3))
        ed.setTextCursor(cur3)
        canvas.apply_list_format({"style": "roman_upper", "suffix": "."})
        b3 = ed.document().findBlockByNumber(3)
        self.assertIsNotNone(b3.textList())
        self.assertEqual(b3.textList().format().style(), QTextListFormat.Style.ListUpperRoman)

        canvas.close()


if __name__ == "__main__":
    unittest.main()
