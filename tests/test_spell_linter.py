import unittest
from core.spell_linter import SpellLinter


class TestSpellLinter(unittest.TestCase):
    def setUp(self):
        self.linter = SpellLinter()

    def test_clean_whitespace(self):
        text = "Đây  là   văn bản    chuẩn hóa."
        cleaned, count = self.linter.clean_whitespace(text)
        self.assertEqual(cleaned, "Đây là văn bản chuẩn hóa.")
        self.assertGreater(count, 0)

    def test_clean_punctuation(self):
        text = "Kính gửi ,đồng chí ;cụ thể :sau đây ."
        cleaned, count, issues = self.linter.clean_punctuation_spacing(text)
        self.assertEqual(cleaned, "Kính gửi, đồng chí; cụ thể: sau đây.")
        self.assertGreater(count, 0)

    def test_standardize_quotes(self):
        text = 'Dự thảo "Quy chế làm việc" được ban hành.'
        cleaned, count, issues = self.linter.standardize_quotes(text)
        self.assertEqual(cleaned, "Dự thảo “Quy chế làm việc” được ban hành.")
        self.assertEqual(count, 2)

    def test_sanitize_dates_nd30(self):
        # Ngày 1 chữ số -> thêm 0; tháng 3 -> bỏ 0
        text = "Hà Nội, ngày 5 tháng 03 năm 2026"
        cleaned, count, issues = self.linter.sanitize_dates(text, "ND30")
        self.assertEqual(cleaned, "Hà Nội, ngày 05 tháng 3 năm 2026")
        self.assertEqual(count, 1)

        # Tháng 1 -> thêm 0
        text2 = "ngày 2 tháng 1 năm 2026"
        cleaned2, count2, issues2 = self.linter.sanitize_dates(text2, "ND30")
        self.assertEqual(cleaned2, "ngày 02 tháng 01 năm 2026")
        self.assertEqual(count2, 1)

    def test_capitalization_warnings(self):
        text = "Yêu cầu các Đồng chí tại cơ quan nghiêm túc thực hiện."
        warnings = self.linter.check_capitalization_warnings(text)
        self.assertTrue(any(w.issue_type == "capitalization" for w in warnings))

        text2 = "Nghị quyết của đảng cộng sản việt nam"
        warnings2 = self.linter.check_capitalization_warnings(text2)
        self.assertTrue(any("đảng cộng sản việt nam" in w.original for w in warnings2))


if __name__ == "__main__":
    unittest.main()
