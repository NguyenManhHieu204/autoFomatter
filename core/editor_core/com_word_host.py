"""
core/editor_core/com_word_host.py: Wrapper điều khiển nhúng Word OLE / Automation
- Quản trị việc mở tài liệu trong cửa sổ MS Word với cấu hình CommandBars chuyên dụng
- Tận dụng sức mạnh ngắt trang, kiểm tra chính tả của MS Word khi có sẵn Office
- Tự động fallback sang Native Canvas nếu máy trạm không có Microsoft Word
"""

import os
import subprocess
from typing import Optional, Dict, Any

from core.win32_helper import Win32WordHelper, WordWatchdog


class ComWordHost:
    """Wrapper điều khiển phiên làm việc Word OLE / Automation ngoại tuyến"""

    def __init__(self):
        self.helper = Win32WordHelper()

    def is_available(self) -> bool:
        return self.helper.is_word_available()

    def open_document(self, docx_path: str) -> bool:
        """Mở tài liệu docx trong cửa sổ Microsoft Word an toàn"""
        abs_path = os.path.abspath(docx_path)
        if not os.path.exists(abs_path):
            return False

        try:
            os.startfile(abs_path)
            return True
        except Exception:
            try:
                subprocess.Popen(["explorer.exe", abs_path], shell=False)
                return True
            except Exception:
                return False

    def export_pdf(self, docx_path: str, output_pdf_path: Optional[str] = None) -> Optional[str]:
        """Xuất PDF ngoại tuyến qua COM Word engine"""
        return self.helper.export_pdf(docx_path, output_pdf_path)
