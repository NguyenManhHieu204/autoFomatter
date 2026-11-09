"""
tests/test_updater.py: Kiểm thử đơn vị cho hệ thống kiểm tra và cập nhật tự động (Auto-Updater).
"""

import unittest
from unittest.mock import patch, MagicMock
import json
import io
import sys
from PyQt6.QtWidgets import QApplication

from core.updater import parse_version, is_newer_version, check_for_updates, APP_VERSION
from ui.update_dialog import UpdateDialog


class TestUpdater(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_version_parsing(self):
        self.assertEqual(parse_version("2.0.0"), (2, 0, 0))
        self.assertEqual(parse_version("v2.0.1"), (2, 0, 1))
        self.assertEqual(parse_version("v3.1.2-beta"), (3, 1, 2))
        self.assertEqual(parse_version(""), (0, 0, 0))
        self.assertEqual(parse_version(None), (0, 0, 0))

    def test_version_comparison(self):
        self.assertTrue(is_newer_version("2.0.1", "2.0.0"))
        self.assertTrue(is_newer_version("v2.1.0", "2.0.9"))
        self.assertTrue(is_newer_version("3.0.0", "2.9.9"))
        self.assertFalse(is_newer_version("2.0.0", "2.0.0"))
        self.assertFalse(is_newer_version("v2.0.0", "2.0.0"))
        self.assertFalse(is_newer_version("1.9.9", "2.0.0"))

    @patch("urllib.request.urlopen")
    def test_check_for_updates_has_new_version(self, mock_urlopen):
        mock_response = MagicMock()
        mock_response.status = 200
        fake_data = {
            "tag_name": "v2.1.0",
            "name": "Auto-Formatter v2.1",
            "body": "Cải tiến tính năng ký số và OCR",
            "html_url": "https://github.com/NguyenManhHieu204/autoFomatter/releases/tag/v2.1.0",
            "published_at": "2026-12-01T10:00:00Z",
            "assets": [
                {
                    "name": "AutoFormatter_v2.1.0.zip",
                    "browser_download_url": "https://github.com/NguyenManhHieu204/autoFomatter/releases/download/v2.1.0/AutoFormatter_v2.1.0.zip",
                    "size": 135000000
                }
            ]
        }
        mock_response.read.return_value = json.dumps(fake_data).encode("utf-8")
        mock_urlopen.return_value.__enter__.return_value = mock_response

        res = check_for_updates(current_version="2.0.0")
        self.assertTrue(res["success"])
        self.assertTrue(res["has_update"])
        self.assertEqual(res["latest_version"], "v2.1.0")
        self.assertEqual(res["asset_name"], "AutoFormatter_v2.1.0.zip")
        self.assertIn("AutoFormatter_v2.1.0.zip", res["download_url"])

    @patch("urllib.request.urlopen")
    def test_check_for_updates_already_latest(self, mock_urlopen):
        mock_response = MagicMock()
        mock_response.status = 200
        fake_data = {
            "tag_name": "v2.0.0",
            "name": "Auto-Formatter v2.0",
            "body": "Phiên bản thương mại",
            "html_url": "https://github.com/NguyenManhHieu204/autoFomatter/releases/tag/v2.0.0",
            "assets": []
        }
        mock_response.read.return_value = json.dumps(fake_data).encode("utf-8")
        mock_urlopen.return_value.__enter__.return_value = mock_response

        res = check_for_updates(current_version="2.0.0")
        self.assertTrue(res["success"])
        self.assertFalse(res["has_update"])

    def test_update_dialog_ui(self):
        info = {
            "current_version": "2.0.0",
            "latest_version": "v2.1.0",
            "release_name": "Bản phát hành v2.1.0",
            "release_notes": "- Thêm phông chữ mới\n- Tăng tốc xuất PDF",
            "download_url": "https://example.com/update.zip",
            "asset_name": "AutoFormatter_v2.1.0.zip",
            "asset_size": 120000000
        }
        dlg = UpdateDialog(None, update_info=info)
        self.assertEqual(dlg.txt_notes.toPlainText(), "- Thêm phông chữ mới\n- Tăng tốc xuất PDF")
        self.assertTrue(dlg.btn_update.isEnabled())
        dlg.close()


if __name__ == "__main__":
    unittest.main()
