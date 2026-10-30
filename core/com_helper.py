"""
core/com_helper.py: Tiện ích xuất PDF ngoại tuyến và tương tác an toàn với MS Word COM
"""

import os
import shutil
from typing import Tuple, Optional
from core.win32_helper import Win32WordHelper


class ComHelper:
    def __init__(self):
        self._helper = Win32WordHelper()

    def is_word_available(self) -> bool:
        return self._helper.is_word_available()

    def convert_docx_to_pdf(self, docx_path: str, pdf_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        Chuyển đổi tệp docx thành tệp PDF ngoại tuyến bằng MS Word COM.
        Đảm bảo an toàn không gây crash ứng dụng ngay cả khi Word gặp sự cố.
        """
        abs_docx = os.path.abspath(docx_path)
        if not os.path.exists(abs_docx):
            return False, f"Tệp không tồn tại: {abs_docx}"

        if pdf_path is None:
            base, _ = os.path.splitext(abs_docx)
            pdf_path = f"{base}.pdf"
        abs_pdf = os.path.abspath(pdf_path)

        if not self._helper.is_word_available():
            return False, "Không phát hiện ứng dụng MS Word khả dụng trên máy tính này."

        try:
            res = self._helper.process_document_via_word(abs_docx, export_pdf_also=True)
            if res.get("pdf_exported"):
                created_pdf = res.get("pdf_path")
                if created_pdf and os.path.exists(created_pdf):
                    if os.path.abspath(created_pdf) != abs_pdf:
                        try:
                            shutil.copy2(created_pdf, abs_pdf)
                        except Exception:
                            pass
                    return True, "Xuất PDF thành công."

            return False, res.get("message", "Không thể xuất tệp PDF qua MS Word.")
        except Exception as e:
            return False, f"Lỗi COM: {str(e)}"
