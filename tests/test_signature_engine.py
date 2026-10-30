"""
tests/test_signature_engine.py: Kiểm tra tính đúng đắn của Bộ máy Ký số Hybrid
- Quét kho chứng thư số Windows
- Cầu nối Ban Cơ yếu Chính phủ (VGCA SignTool)
- Sinh hình ảnh chữ ký số cá nhân và con dấu cơ quan chuẩn NĐ 30/2020/NĐ-CP
"""

import os
import unittest
from core.signature_engine import (
    WindowsCertStoreHelper, VGCABridge, VisualStampGenerator, CertificateInfo
)


class TestSignatureEngine(unittest.TestCase):
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
        temp_pdf = "temp_test_personal_stamp.pdf"
        try:
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
        finally:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)

    def test_visual_stamp_agency(self):
        temp_pdf = "temp_test_agency_stamp.pdf"
        try:
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
        finally:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)


if __name__ == "__main__":
    unittest.main()
