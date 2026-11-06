"""
tests/test_signature_engine.py: Kiểm tra tính đúng đắn của Bộ máy Ký số Hybrid
- Quét kho chứng thư số Windows
- Cầu nối Ban Cơ yếu Chính phủ (VGCA SignTool)
- Sinh hình ảnh chữ ký số cá nhân, con dấu cơ quan và mẫu kết hợp chuẩn NĐ 30/2020/NĐ-CP
- Tự động định vị ô ký thông minh (Smart Signature Locator)
- Thử nghiệm ký số thực tế qua Windows CryptoAPI và dọn dẹp sạch sẽ
"""

import os
import unittest
import tempfile
import shutil
from pypdf import PdfWriter, PdfReader
from core.signature_engine import (
    WindowsCertStoreHelper, VGCABridge, VisualStampGenerator,
    CertificateInfo, PdfSignatureLocator, PdfDigitalSigner
)


class TestSignatureEngine(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="af_test_sig_unit_")

    def tearDown(self):
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_cert_store_enumeration(self):
        certs = WindowsCertStoreHelper.list_certificates()
        self.assertIsInstance(certs, list)
        for c in certs:
            self.assertIsInstance(c, CertificateInfo)
            self.assertTrue(c.has_private_key)
            self.assertIsNotNone(c.display_name)

    def test_vgca_bridge_check(self):
        ok, msg, port = VGCABridge.check_service_status()
        self.assertIsInstance(ok, bool)
        self.assertIsInstance(msg, str)
        self.assertIsInstance(port, int)

    def test_visual_stamp_personal(self):
        temp_pdf = os.path.join(self.temp_dir, "temp_personal_stamp.pdf")
        ok = VisualStampGenerator.create_stamp_pdf(
            output_stamp_pdf=temp_pdf,
            sign_type="personal",
            signer_name="Nguyễn Văn A",
            signer_title="Chủ tịch UBND",
            agency_name="Ủy ban nhân dân xã"
        )
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(temp_pdf))
        self.assertGreater(os.path.getsize(temp_pdf), 1000)

    def test_visual_stamp_agency(self):
        temp_pdf = os.path.join(self.temp_dir, "temp_agency_stamp.pdf")
        ok = VisualStampGenerator.create_stamp_pdf(
            output_stamp_pdf=temp_pdf,
            sign_type="agency",
            signer_name="Ủy ban nhân dân xã",
            signer_title="",
            agency_name="ỦY BAN NHÂN DÂN XÃ"
        )
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(temp_pdf))
        self.assertGreater(os.path.getsize(temp_pdf), 1000)

    def test_visual_stamp_combined(self):
        temp_pdf = os.path.join(self.temp_dir, "temp_combined_stamp.pdf")
        ok = VisualStampGenerator.create_stamp_pdf(
            output_stamp_pdf=temp_pdf,
            sign_type="combined",
            signer_name="Nguyễn Văn An",
            signer_title="Chủ tịch",
            agency_name="ỦY BAN NHÂN DÂN XÃ"
        )
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(temp_pdf))
        self.assertGreater(os.path.getsize(temp_pdf), 1000)

    def test_signature_locator_multipage(self):
        test_pdf = os.path.join(self.temp_dir, "sample_multipage.pdf")
        w = PdfWriter()
        w.add_blank_page(width=595.28, height=841.89)
        w.add_blank_page(width=595.28, height=841.89)
        with open(test_pdf, "wb") as f:
            w.write(f)

        anchor = PdfSignatureLocator.find_signature_anchor(test_pdf)
        self.assertIsInstance(anchor, dict)
        self.assertIn("page_index", anchor)
        self.assertIn("norm_x", anchor)
        self.assertIn("norm_y", anchor)
        self.assertEqual(anchor["page_index"], 1)

    def test_cert_filter_excludes_dev_and_guid_certs(self):
        """Kiểm tra bộ lọc chứng thư: loại bỏ 100% chứng thư dev, IIS, ASP.NET, localhost, GUID"""
        self.assertTrue(bool(WindowsCertStoreHelper.GUID_REGEX.match("a7e9ce00-9d76-480e-95d0-f87ae615d9dc")))
        self.assertTrue(bool(WindowsCertStoreHelper.GUID_REGEX.match("trust_a7e9ce00-9d76-480e-95d0-f87ae615d9dc")))
        self.assertFalse(bool(WindowsCertStoreHelper.GUID_REGEX.match("NGUYỄN VĂN AN")))

        certs = WindowsCertStoreHelper.list_certificates()
        # Trên môi trường test không cắm token thật, danh sách certs trả về phải rỗng (không chứa dev GUID cert)
        for c in certs:
            self.assertFalse(bool(WindowsCertStoreHelper.GUID_REGEX.match(c.common_name)))
            self.assertNotIn("localhost", c.common_name.lower())

    def test_signature_anchor_on_real_administrative_doc(self):
        """Kiểm tra định vị khối ký chuẩn xác trên tài liệu hành chính mẫu thật"""
        sample_pdf = "Tai_Lieu_Mau_Test/Van_Ban_Hanh_Chinh_Mau_Loi_ChuanHoa.pdf"
        if not os.path.exists(sample_pdf):
            self.skipTest("Không tìm thấy tệp PDF mẫu kiểm tra.")

        anchor = PdfSignatureLocator.find_signature_anchor(sample_pdf)
        self.assertTrue(anchor["found"])
        self.assertEqual(anchor["authority"], "PHÓ CHỦ TỊCH")
        self.assertEqual(anchor["signer_name"], "Nguyễn Văn An")
        self.assertAlmostEqual(anchor["norm_x"], 0.55, delta=0.05)
        self.assertGreater(anchor["norm_y"], 0.60)
        self.assertLess(anchor["norm_y"], 0.70)

    def test_hybrid_sign_with_available_cert(self):
        certs = WindowsCertStoreHelper.list_certificates()
        if not certs:
            self.skipTest("Không có chứng thư số nào trong Windows Cert Store để thử ký số.")

        test_pdf = os.path.join(self.temp_dir, "doc_to_sign.pdf")
        out_pdf = os.path.join(self.temp_dir, "doc_signed.pdf")

        w = PdfWriter()
        w.add_blank_page(width=595.28, height=841.89)
        with open(test_pdf, "wb") as f:
            w.write(f)

        cert = certs[0]
        ok, msg = PdfDigitalSigner.sign_pdf_hybrid(
            input_pdf_path=test_pdf,
            output_pdf_path=out_pdf,
            cert_info=cert,
            sign_type="combined",
            signer_name="Trần Văn Thử Nghiệm",
            signer_title="Chủ tịch",
            agency_name="ỦY BAN NHÂN DÂN XÃ",
            target_page=0,
            norm_x=0.55,
            norm_y=0.70
        )
        self.assertTrue(ok, f"Ký số thất bại: {msg}")
        self.assertTrue(os.path.exists(out_pdf))
        r = PdfReader(out_pdf)
        self.assertEqual(len(r.pages), 1)
        self.assertGreater(os.path.getsize(out_pdf), 1000)


if __name__ == "__main__":
    unittest.main()
