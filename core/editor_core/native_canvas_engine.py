"""
core/editor_core/native_canvas_engine.py: Bộ kết xuất Trang in WYSIWYG trên Qt Canvas (Host Mode 2)
- Dự phòng độc lập 100% khi máy trạm không cài đặt Microsoft Word
- Tính toán phân trang A4 thực tế (210 x 297 mm tại 96 DPI màn hình = 794 x 1123 px)
- Quản trị chuyển đổi hai chiều giữa mô hình OpenXML (python-docx) và QTextDocument
"""

from typing import Dict, Any, List, Optional
import docx
from docx.shared import Mm, Pt


class NativeCanvasEngine:
    """Mô hình dữ liệu và thuật toán kết xuất trang in WYSIWYG A4 độc lập"""

    DPI = 96.0
    MM_TO_PX = DPI / 25.4  # ~ 3.7795 px / mm

    PAGE_WIDTH_MM = 210.0
    PAGE_HEIGHT_MM = 297.0

    MARGIN_TOP_MM = 20.0
    MARGIN_BOTTOM_MM = 20.0
    MARGIN_LEFT_MM = 30.0
    MARGIN_RIGHT_MM = 15.0

    @classmethod
    def get_page_dimensions_px(cls) -> Dict[str, int]:
        return {
            "page_width_px": int(cls.PAGE_WIDTH_MM * cls.MM_TO_PX),
            "page_height_px": int(cls.PAGE_HEIGHT_MM * cls.MM_TO_PX),
            "margin_left_px": int(cls.MARGIN_LEFT_MM * cls.MM_TO_PX),
            "margin_right_px": int(cls.MARGIN_RIGHT_MM * cls.MM_TO_PX),
            "margin_top_px": int(cls.MARGIN_TOP_MM * cls.MM_TO_PX),
            "margin_bottom_px": int(cls.MARGIN_BOTTOM_MM * cls.MM_TO_PX),
            "printable_width_px": int((cls.PAGE_WIDTH_MM - cls.MARGIN_LEFT_MM - cls.MARGIN_RIGHT_MM) * cls.MM_TO_PX),
            "printable_height_px": int((cls.PAGE_HEIGHT_MM - cls.MARGIN_TOP_MM - cls.MARGIN_BOTTOM_MM) * cls.MM_TO_PX)
        }

    @classmethod
    def docx_to_canvas_data(cls, docx_path: str) -> Dict[str, Any]:
        """Trích xuất dữ liệu từ file docx để nạp vào DocumentCanvas"""
        import os
        from core.doc_parser import DocParser

        result = {
            "standard_id": "ND30",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "",
            "body_html": "",
            "body_paragraphs": []
        }

        if not os.path.exists(docx_path):
            return result

        doc = docx.Document(docx_path)
        parser = DocParser(doc)
        structure = parser.parse()

        result["standard_id"] = structure.detected_standard
        if structure.doc_type:
            result["doc_type"] = structure.doc_type
        if structure.doc_subject:
            result["doc_subject"] = structure.doc_subject

        html_paras = []
        body_paras = []
        for p_idx in structure.body_paragraphs_indices:
            p = doc.paragraphs[p_idx]
            t = p.text.strip()
            if not t:
                continue
            body_paras.append(t)
            is_bold = any(r.font.bold for r in p.runs)
            is_italic = any(r.font.italic for r in p.runs)

            style = "text-align: justify; text-indent: 36px;"
            if is_bold:
                html_paras.append(f"<p style='{style} font-weight: bold;'>{t}</p>")
            elif is_italic:
                html_paras.append(f"<p style='{style} font-style: italic;'>{t}</p>")
            else:
                html_paras.append(f"<p style='{style}'>{t}</p>")

        result["body_paragraphs"] = body_paras
        result["body_html"] = "".join(html_paras)
        return result
