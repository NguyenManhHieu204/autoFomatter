import os
import unittest
import docx
from PIL import Image, ImageDraw

from core.formatter_engine import FormatterEngine
from core.doc_parser import DocParser


class TestScannedDocHandling(unittest.TestCase):
    def setUp(self):
        self.img_path = "temp_scan_page.png"
        self.doc_path = "test_scanned_image_only.docx"
        self.out_safe = "test_scanned_safe_output.docx"
        self.out_ocr = "test_scanned_ocr_output.docx"

        # Tạo 1 ảnh giả lập trang scan có chữ rõ nét
        img = Image.new("RGB", (1200, 600), color=(255, 255, 255))
        d = ImageDraw.Draw(img)
        try:
            from PIL import ImageFont
            font = ImageFont.truetype("arial.ttf", 36)
        except Exception:
            font = None

        d.text((50, 50), "UBND THANH PHO HA NOI", font=font, fill=(0, 0, 0))
        d.text((50, 150), "QUYET DINH", font=font, fill=(0, 0, 0))
        d.text((50, 250), "Ve viec ban hanh quy che lam viec", font=font, fill=(0, 0, 0))
        d.text((50, 350), "Dieu 1. Ban hanh quy che lam viec moi tai co quan don vi.", font=font, fill=(0, 0, 0))
        d.text((50, 450), "Dieu 2. Quyet dinh nay co hieu luc ke tu ngay ky ban hanh.", font=font, fill=(0, 0, 0))
        img.save(self.img_path)

        # Tạo file word nhúng bức ảnh này (giả lập file word chuyển từ PDF scan)
        doc = docx.Document()
        p = doc.add_paragraph()
        p.add_run().add_picture(self.img_path)
        doc.save(self.doc_path)

    def tearDown(self):
        for f in [self.img_path, self.doc_path, self.out_safe, self.out_ocr, f"{self.doc_path}.bak"]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass

    def test_safe_mode_prevents_image_deletion(self):
        """Kiểm tra Safe Mode: Tuyệt đối không xóa ảnh scan của người dùng"""
        engine = FormatterEngine()
        report = engine.format_document(
            input_docx_path=self.doc_path,
            output_docx_path=self.out_safe,
            standard_id="ND30"
        )

        self.assertTrue(report.success)
        self.assertTrue(report.is_scanned_image, "Phải nhận diện được đây là file dạng ảnh scan")
        self.assertTrue(report.safe_mode_activated)
        self.assertTrue(os.path.exists(self.out_safe))

        # Mở file kiểm tra xem ảnh scan còn nguyên vẹn không
        doc_res = docx.Document(self.out_safe)
        drawings = doc_res.element.xpath('.//*[local-name()="drawing" or local-name()="pict"]')
        self.assertGreater(len(drawings), 0, "Ảnh scan phải được bảo vệ nguyên vẹn 100%, không bị xóa!")

    def test_ocr_flow_reconstructs_editable_document(self):
        """Kiểm tra luồng OCR: Nhận dạng chữ từ ảnh scan và dựng lại văn bản chuẩn"""
        engine = FormatterEngine()
        report = engine.format_scanned_document_with_ocr(
            input_file=self.doc_path,
            output_docx_path=self.out_ocr,
            standard_id="ND30"
        )

        self.assertTrue(report.success, f"OCR formatting failed: {report.error_message}")
        self.assertTrue(os.path.exists(self.out_ocr))

        # Kiểm tra văn bản sinh ra có chữ thực tế trong paragraphs hoặc tables
        doc_res = docx.Document(self.out_ocr)
        text_content = " ".join([p.text for p in doc_res.paragraphs])
        for tbl in doc_res.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    text_content += " " + cell.text

        self.assertTrue(len(text_content.strip()) > 0, "Văn bản phải chứa chữ có thể chỉnh sửa được sau OCR")


if __name__ == "__main__":
    unittest.main()
