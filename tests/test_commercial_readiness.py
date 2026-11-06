"""
tests/test_commercial_readiness.py: Bộ kiểm thử đánh giá tính sẵn sàng thương mại
Kiểm tra an toàn bảo mật, khả năng chống lỗi bất thường và độ tin cậy của ứng dụng:
1. Phòng vệ tệp độc hại / không hợp lệ (Security Guard: .docm, zip-slip, corrupt file)
2. Chuẩn hóa đường dẫn dài (Extended path normalization)
3. Chống chịu tài liệu rỗng / bất thường (Empty doc, tables-only doc)
4. Chuyển đổi chuẩn hóa bảng mã tiếng Việt (Unicode Tổ hợp NFD -> Dựng sẵn NFC)
5. Tính điểm tuân thủ thể thức pháp lý (Legal Compliance Score)
"""

import os
import sys
import unittest
import tempfile
import shutil
import docx

from core.security_guard import SecurityGuard
from core.charset_converter import CharsetConverter
from core.formatter_engine import FormatterEngine
from core.legal_constraint_guard import LegalConstraintGuard
from core.editor_core.openxml_builder import OpenXmlBuilder


class TestCommercialReadiness(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="af_test_comm_")
        self.engine = FormatterEngine()

    def tearDown(self):
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_security_forbidden_extensions(self):
        """Kiểm tra chặn tệp chứa Macro (.docm) hoặc tệp thực thi"""
        for ext in [".docm", ".dotm", ".exe", ".bat", ".ps1"]:
            fake_file = os.path.join(self.temp_dir, f"malicious{ext}")
            with open(fake_file, "w") as f:
                f.write("content")
            ok, msg = SecurityGuard.validate_input_file(fake_file)
            self.assertFalse(ok, f"Hệ thống không chặn tệp nguy hiểm: {ext}")
            self.assertIn("Từ chối tệp tiềm ẩn nguy cơ bảo mật", msg)

    def test_security_corrupt_docx(self):
        """Kiểm tra xử lý tệp .docx giả mạo hoặc bị hỏng không gây crash hệ thống"""
        corrupt_file = os.path.join(self.temp_dir, "corrupt.docx")
        with open(corrupt_file, "w") as f:
            f.write("This is not a zip file at all!")
        
        ok, msg = SecurityGuard.validate_docx_package(corrupt_file)
        self.assertFalse(ok)
        self.assertIn("không đúng định dạng", msg)

        # FormatterEngine phải xử lý an toàn, trả về report.success = False, không văng ngoại lệ
        out_file = os.path.join(self.temp_dir, "corrupt_out.docx")
        report = self.engine.format_document(corrupt_file, out_file)
        self.assertFalse(report.success)
        self.assertIsNotNone(report.error_message)

    def test_security_long_path_handling(self):
        """Kiểm tra chuẩn hóa đường dẫn siêu dài (trên 240 ký tự) trên Windows"""
        long_dir = os.path.join(self.temp_dir, "a" * 120, "b" * 120)
        norm = SecurityGuard.normalize_long_path(long_dir)
        self.assertIsInstance(norm, str)
        if sys.platform == "win32" and len(os.path.abspath(long_dir)) >= 240:
            self.assertTrue(norm.startswith("\\\\?\\"))

    def test_charset_nfd_to_nfc_conversion(self):
        """Kiểm tra chuyển đổi triệt để lỗi gõ tiếng Việt Unicode Tổ hợp (NFD) sang Dựng sẵn (NFC)"""
        # "Đảng" dạng NFD (D với dấu gạch ngang + a + ̉ + ng)
        text_nfd = "Đa\u0309ng Bô\u0323 Xa\u0303"
        text_nfc = CharsetConverter.normalize_to_nfc(text_nfd)
        # Ký tự đã được dựng sẵn gọn gàng
        self.assertEqual(text_nfc, "Đảng Bộ Xã")
        self.assertEqual(len(text_nfc), 10)

    def test_empty_document_handling(self):
        """Kiểm tra tài liệu hoàn toàn rỗng không gây crash hệ thống"""
        empty_doc_path = os.path.join(self.temp_dir, "empty.docx")
        doc = docx.Document()
        doc.save(empty_doc_path)

        out_path = os.path.join(self.temp_dir, "empty_out.docx")
        report = self.engine.format_document(empty_doc_path, out_path, standard_id="ND30")
        self.assertTrue(report.success)
        self.assertTrue(os.path.exists(out_path))

    def test_tables_only_document_handling(self):
        """Kiểm tra tài liệu chỉ chứa bảng biểu (không có đoạn văn thuần)"""
        table_doc_path = os.path.join(self.temp_dir, "tables_only.docx")
        doc = docx.Document()
        t = doc.add_table(rows=2, cols=2)
        t.cell(0, 0).text = "Tiêu đề 1"
        t.cell(0, 1).text = "Tiêu đề 2"
        t.cell(1, 0).text = "Dữ liệu A"
        t.cell(1, 1).text = "Dữ liệu B"
        doc.save(table_doc_path)

        out_path = os.path.join(self.temp_dir, "tables_only_out.docx")
        report = self.engine.format_document(table_doc_path, out_path, standard_id="ND30")
        self.assertTrue(report.success)
        self.assertTrue(os.path.exists(out_path))

    def test_legal_constraint_guard_scoring(self):
        """Kiểm tra bộ máy chấm điểm thể thức pháp lý (0 - 100 điểm)"""
        meta = {
            "standard_id": "ND30",
            "superior_agency": "ỦY BAN NHÂN DÂN HUYỆN",
            "issuing_agency": "ỦY BAN NHÂN DÂN XÃ",
            "doc_number": "Số: 01/QĐ-UBND",
            "location": "Xã",
            "date_str": "ngày 06 tháng 11 năm 2026",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "Về việc ban hành Kế hoạch phát triển kinh tế",
            "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
            "signer_position": "CHỦ TỊCH",
            "signer_name": "Nguyễn Văn An",
            "recipients": ["- Thường trực HĐND xã;", "- Lưu: VT."]
        }
        doc = OpenXmlBuilder.create_scaffolded_document(meta)
        
        # Thêm căn cứ và nội dung hợp chuẩn
        doc.add_paragraph("Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;")
        doc.add_paragraph("QUYẾT ĐỊNH:")
        doc.add_paragraph("Điều 1. Ban hành kèm theo Quyết định này kế hoạch phát triển kinh tế.")

        audit = LegalConstraintGuard.validate_document(doc, standard_id="ND30")
        self.assertIsInstance(audit, dict)
        self.assertIn("score", audit)
        self.assertIn("checklist", audit)
        self.assertIn("violations", audit)
        # Văn bản dựng chuẩn mẫu phải đạt điểm xuất sắc (>= 80 điểm)
        self.assertGreaterEqual(audit["score"], 80.0)


if __name__ == "__main__":
    unittest.main()
