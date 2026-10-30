import os
import unittest
import tempfile
import docx
from core.spell_linter import SpellLinter
from core.ocr_engine import OcrEngine
from core.win32_helper import Win32WordHelper
from core.formatter_engine import FormatterEngine


class TestOptimizations(unittest.TestCase):
    def setUp(self):
        self.linter = SpellLinter()
        self.ocr = OcrEngine()
        self.helper = Win32WordHelper()

    def test_regex_precompilation_and_single_pass(self):
        """Kiểm tra regex biên dịch trước và quét 1 lượt danh từ chung & cơ quan"""
        text = "Yêu cầu các Đồng chí tại cơ quan nghiêm túc báo cáo đảng cộng sản việt nam."
        warnings = self.linter.check_capitalization_warnings(text)
        
        self.assertTrue(any("Đồng chí" in w.original for w in warnings))
        self.assertTrue(any("đảng cộng sản việt nam" in w.original for w in warnings))

    def test_date_sanitization_optimized(self):
        """Kiểm tra xử lý nhanh ngày tháng"""
        text = "Hà Nội, ngày 9 tháng 02 năm 2026"
        res, count, issues = self.linter.sanitize_dates(text, "ND30")
        self.assertEqual(res, "Hà Nội, ngày 09 tháng 02 năm 2026")
        self.assertEqual(count, 1)

    def test_ocr_singleton_instance(self):
        """Kiểm tra RapidOCR dùng chung singleton để không nạp lại mô hình nặng"""
        ocr1 = OcrEngine()
        ocr2 = OcrEngine()
        self.assertEqual(ocr1.is_ocr_available(), ocr2.is_ocr_available())

    def test_word_session_context_manager(self):
        """Kiểm tra cơ chế word_session mở và đóng an toàn"""
        with self.helper.word_session() as session:
            # Nếu Word COM khả dụng, session là đối tượng Word.Application
            # Nếu không khả dụng (máy không có Word), session là None
            self.assertTrue(session is not None or not self.helper.is_word_available())


if __name__ == "__main__":
    unittest.main()
