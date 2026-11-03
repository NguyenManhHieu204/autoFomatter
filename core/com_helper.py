"""
core/com_helper.py: Tiện ích xuất PDF ngoại tuyến và tương tác an toàn với MS Word COM
Hỗ trợ cơ chế Fallback tự động:
1. Thử xuất PDF chuẩn xác qua MS Word COM
2. Nếu MS Word bị treo, quá thời gian chờ (watchdog) hoặc không được cài đặt:
   Tự động kích hoạt Engine chuyển đổi PDF ngoại tuyến độc lập (Qt PDF Engine)
   để đảm bảo không bao giờ làm gián đoạn công việc của cán bộ.
"""

import os
import shutil
import html
from typing import Tuple, Optional
from core.win32_helper import Win32WordHelper


def convert_docx_to_pdf_offline_qt(docx_path: str, pdf_target: str) -> bool:
    """
    Chuyển đổi tệp DOCX sang PDF ngoại tuyến bằng engine Qt (QPdfWriter + QTextDocument)
    Đảm bảo 100% độc lập, không phụ thuộc vào Microsoft Word, không bao giờ bị treo tiến trình.
    Tuân thủ nghiêm ngặt kích thước trang A4 và lề NĐ 30 (Trái 30mm, Trên 20mm, Phải 15mm, Dưới 20mm).
    """
    try:
        import docx
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtGui import QPdfWriter, QTextDocument, QPageSize, QPageLayout, QFont
        from PyQt6.QtCore import QMarginsF

        app = QApplication.instance()
        if app is None:
            app = QApplication([])

        doc = docx.Document(docx_path)
        writer = QPdfWriter(pdf_target)
        writer.setResolution(300)
        page_size = QPageSize(QPageSize.PageSizeId.A4)
        orientation = QPageLayout.Orientation.Portrait
        # Lề chuẩn Nghị định 30/2020/NĐ-CP
        margins = QMarginsF(30, 20, 15, 20)
        layout = QPageLayout(page_size, orientation, margins, QPageLayout.Unit.Millimeter)
        writer.setPageLayout(layout)

        html_parts = []
        html_parts.append("""
        <html>
        <head>
        <meta charset="utf-8">
        <style>
            body { font-family: 'Times New Roman', serif; color: #000000; font-size: 13pt; line-height: 1.25; }
            table { border-collapse: collapse; width: 100%; margin: 6px 0; }
            p { margin: 3px 0; }
        </style>
        </head>
        <body>
        """)

        def _format_paragraph(p: Paragraph) -> str:
            align_css = "text-align: justify;"
            if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                align_css = "text-align: center;"
            elif p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                align_css = "text-align: right;"
            elif p.alignment == WD_ALIGN_PARAGRAPH.LEFT:
                align_css = "text-align: left;"

            indent_css = ""
            if p.paragraph_format.first_line_indent and p.paragraph_format.first_line_indent.pt > 0:
                indent_css = "text-indent: 36px;"

            runs_html = []
            for r in p.runs:
                txt = html.escape(r.text)
                txt = txt.replace("\n", "<br/>")
                style_bits = []
                if r.font.size:
                    style_bits.append(f"font-size: {r.font.size.pt:.1f}pt;")
                if r.bold:
                    txt = f"<b>{txt}</b>"
                if r.italic:
                    txt = f"<i>{txt}</i>"
                if r.underline:
                    txt = f"<u>{txt}</u>"
                if style_bits:
                    txt = f"<span style='{' '.join(style_bits)}'>{txt}</span>"
                runs_html.append(txt)

            inner = "".join(runs_html).strip()
            if not inner:
                return "<p style='margin: 2px 0;'><br/></p>"
            return f"<p style='{align_css} {indent_css} margin: 3px 0;'>{inner}</p>"

        # Duyệt qua các phần tử cấp cao nhất của tài liệu
        for element in doc.element.body:
            tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
            if tag == "p":
                p = Paragraph(element, doc)
                html_parts.append(_format_paragraph(p))
            elif tag == "tbl":
                tbl = Table(element, doc)
                # Kiểm tra viền bảng: nếu có 'Table Grid' thì hiển thị viền đen mảnh
                style_name = tbl.style.name if tbl.style else ""
                has_border = "grid" in style_name.lower() or "table" in style_name.lower()
                border_attr = "border='1' cellpadding='4' cellspacing='0' style='border: 1px solid #000000;'" if has_border else "border='0' cellpadding='2' cellspacing='0'"
                tbl_html = [f"<table width='100%' {border_attr}>"]
                for row in tbl.rows:
                    tbl_html.append("<tr>")
                    for cell in row.cells:
                        c_paras = [_format_paragraph(cp) for cp in cell.paragraphs]
                        cell_content = "".join(c_paras) or "&nbsp;"
                        c_style = "border: 1px solid #000000; padding: 4px;" if has_border else "border: none; padding: 2px;"
                        tbl_html.append(f"<td valign='top' style='{c_style}'>{cell_content}</td>")
                    tbl_html.append("</tr>")
                tbl_html.append("</table>")
                html_parts.append("".join(tbl_html))

        html_parts.append("</body></html>")
        full_html = "".join(html_parts)

        qt_doc = QTextDocument()
        qt_doc.setDefaultFont(QFont("Times New Roman", 13))
        qt_doc.setHtml(full_html)
        qt_doc.print(writer)
        return os.path.exists(pdf_target) and os.path.getsize(pdf_target) > 0
    except Exception as e:
        print(f"[Offline Qt PDF Error] {e}")
        return False


class ComHelper:
    def __init__(self):
        self._helper = Win32WordHelper()

    def is_word_available(self) -> bool:
        return self._helper.is_word_available()

    def convert_docx_to_pdf(self, docx_path: str, pdf_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        Chuyển đổi tệp docx thành tệp PDF an toàn, 100% ngoại tuyến.
        Cơ chế:
        1. Thử xuất bằng MS Word COM nếu khả dụng.
        2. Nếu Word COM gặp sự cố, bị treo quá thời gian chờ, hoặc không cài Office:
           Tự động chuyển đổi ngay bằng Engine Qt nội bộ (Fail-safe Fallback).
        """
        abs_docx = os.path.abspath(docx_path)
        if not os.path.exists(abs_docx):
            return False, f"Tệp không tồn tại: {abs_docx}"

        if pdf_path is None:
            base, _ = os.path.splitext(abs_docx)
            pdf_path = f"{base}.pdf"
        abs_pdf = os.path.abspath(pdf_path)

        # 1. Thử chuyển đổi qua MS Word COM nếu khả dụng
        if self._helper.is_word_available():
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
                        return True, "Xuất PDF thành công bằng MS Word."
            except Exception:
                pass

        # 2. Cơ chế cứu sinh tự động (Fail-safe Fallback): Chuyển đổi ngoại tuyến bằng Qt PDF Engine
        ok_qt = convert_docx_to_pdf_offline_qt(abs_docx, abs_pdf)
        if ok_qt and os.path.exists(abs_pdf):
            return True, "Xuất PDF thành công (Engine ngoại tuyến độc lập)."

        return False, "Không thể xuất tệp PDF. Vui lòng kiểm tra quyền truy cập tệp."
