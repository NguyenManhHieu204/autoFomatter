"""
ui/editor/document_canvas.py: Khung vẽ Soạn thảo A4 Trực quan (WYSIWYG Document Canvas)
- Kết xuất trang in A4 chuẩn (210 x 297 mm) với đổ bóng (Drop Shadow) chân thực
- Hiển thị ranh giới lề in NĐ 30 (Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm)
- Phân tách 4 vùng rõ rệt:
  + Vùng 1: Header Quốc hiệu - Tên cơ quan (Khóa cấu trúc)
  + Vùng 2: Tiêu đề Loại văn bản & Trích yếu
  + Vùng 3: Thân văn bản (Body - Toàn quyền soạn thảo như Word)
  + Vùng 4: Footer Nơi nhận & Chữ ký (Khóa cấu trúc)
- Tích hợp công cụ chỉnh sửa văn bản phong phú (QTextEdit rich-text engine)
"""

import os
import re
from typing import Dict, Any, Optional, List
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTextEdit, QLabel,
    QFrame, QScrollArea, QSizePolicy, QMessageBox, QFileDialog
)
from PyQt6.QtGui import (
    QPainter, QColor, QFont, QTextCursor, QTextCharFormat,
    QTextBlockFormat, QTextDocument, QPixmap, QTextTableFormat,
    QTextLength
)
from PyQt6.QtCore import Qt, pyqtSignal, QRectF, QSize, QMarginsF


class DocumentCanvas(QWidget):
    """Bề mặt hiển thị trang giấy A4 phân vùng soạn thảo"""
    document_modified = pyqtSignal()
    cursor_position_changed = pyqtSignal(int, int)  # line, col

    # Kích thước A4 tại 96 DPI màn hình tiêu chuẩn (1 inch = 25.4mm = 96px)
    A4_WIDTH_PX = 794
    A4_HEIGHT_PX = 1123

    def __init__(self, parent=None):
        super().__init__(parent)
        self.zoom = 1.0
        self.lock_scaffolding = True
        self.metadata: Dict[str, Any] = {
            "standard_id": "ND30",
            "superior_agency": "ỦY BAN NHÂN DÂN",
            "issuing_agency": "THÀNH PHỐ HÀ NỘI",
            "doc_number": "Số: 125/QĐ-UBND",
            "location": "Hà Nội",
            "date_str": "Hà Nội, ngày ... tháng ... năm 2026",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "Về việc phê duyệt kế hoạch chuyển đổi số năm 2026",
            "recipients": ["- Như Điều 3;", "- Lưu: VT, TH."],
            "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
            "signer_position": "CHỦ TỊCH",
            "signer_name": "Nguyễn Văn A"
        }

        # Bố cục trang giấy A4 trung tâm
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.setContentsMargins(40, 30, 40, 30)

        # Trang giấy A4 (Page Sheet Card)
        self.page_sheet = QWidget()
        self.page_sheet.setFixedSize(self.A4_WIDTH_PX, self.A4_HEIGHT_PX)
        self.page_sheet.setStyleSheet("""
            QWidget#page_sheet {
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 2px;
            }
        """)
        self.page_sheet.setObjectName("page_sheet")

        # Đặt lề trang in chuẩn (Lề trái 30mm = 113px, Phải 15mm = 57px, Trên 20mm = 76px, Dưới 20mm = 76px)
        sheet_layout = QVBoxLayout(self.page_sheet)
        sheet_layout.setContentsMargins(113, 76, 57, 76)
        sheet_layout.setSpacing(6)

        # 1. KHUNG HEADER THỂ THỨC (QUỐC HIỆU & CƠ QUAN)
        self.header_frame = QFrame()
        self.header_frame.setStyleSheet("""
            QFrame {
                background-color: #f8fafc;
                border: 1px dashed #94a3b8;
                border-radius: 4px;
                padding: 4px;
            }
        """)
        header_layout = QHBoxLayout(self.header_frame)
        header_layout.setContentsMargins(6, 6, 6, 6)

        self.lbl_header_left = QLabel()
        self.lbl_header_left.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_header_left.setStyleSheet("color: #0f172a; font-family: 'Times New Roman';")
        header_layout.addWidget(self.lbl_header_left, 1)

        self.lbl_header_right = QLabel()
        self.lbl_header_right.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_header_right.setStyleSheet("color: #0f172a; font-family: 'Times New Roman';")
        header_layout.addWidget(self.lbl_header_right, 1)

        sheet_layout.addWidget(self.header_frame)

        # 2. KHUNG TIÊU ĐỀ LOẠI VĂN BẢN VÀ TRÍCH YẾU
        self.title_frame = QFrame()
        title_layout = QVBoxLayout(self.title_frame)
        title_layout.setContentsMargins(0, 8, 0, 8)
        title_layout.setSpacing(2)

        self.lbl_doc_type = QLabel("QUYẾT ĐỊNH")
        self.lbl_doc_type.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_doc_type.setStyleSheet("font-family: 'Times New Roman'; font-size: 15pt; font-weight: bold; color: #0f172a;")
        title_layout.addWidget(self.lbl_doc_type)

        self.lbl_doc_subject = QLabel("Về việc phê duyệt kế hoạch chuyển đổi số năm 2026")
        self.lbl_doc_subject.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_doc_subject.setStyleSheet("font-family: 'Times New Roman'; font-size: 13.5pt; font-weight: bold; color: #1e293b;")
        self.lbl_doc_subject.setWordWrap(True)
        title_layout.addWidget(self.lbl_doc_subject)

        sheet_layout.addWidget(self.title_frame)

        # 3. KHUNG THÂN VĂN BẢN (BODY EDITOR - VÙNG BIÊN TẬP MỞ 100%)
        self.body_editor = QTextEdit()
        self.body_editor.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: transparent;
                font-family: 'Times New Roman';
                font-size: 13pt;
                color: #0f172a;
                line-height: 1.25;
            }
        """)
        self.body_editor.textChanged.connect(self.document_modified.emit)
        self.body_editor.cursorPositionChanged.connect(self._on_cursor_changed)
        sheet_layout.addWidget(self.body_editor, 1)

        # 4. KHUNG FOOTER THỂ THỨC (NƠI NHẬN & CHỮ KÝ)
        self.footer_frame = QFrame()
        self.footer_frame.setStyleSheet("""
            QFrame {
                background-color: #f8fafc;
                border: 1px dashed #94a3b8;
                border-radius: 4px;
                padding: 4px;
            }
        """)
        footer_layout = QHBoxLayout(self.footer_frame)
        footer_layout.setContentsMargins(6, 6, 6, 6)

        self.lbl_footer_left = QLabel()
        self.lbl_footer_left.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.lbl_footer_left.setStyleSheet("color: #0f172a; font-family: 'Times New Roman';")
        footer_layout.addWidget(self.lbl_footer_left, 1)

        self.lbl_footer_right = QLabel()
        self.lbl_footer_right.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignTop)
        self.lbl_footer_right.setStyleSheet("color: #0f172a; font-family: 'Times New Roman';")
        footer_layout.addWidget(self.lbl_footer_right, 1)

        sheet_layout.addWidget(self.footer_frame)

        main_layout.addWidget(self.page_sheet)

        self._refresh_scaffolding_display()
        self._load_sample_body()

    def set_zoom(self, zoom: float):
        self.zoom = max(0.4, min(2.5, zoom))
        new_w = int(self.A4_WIDTH_PX * self.zoom)
        new_h = int(self.A4_HEIGHT_PX * self.zoom)
        self.page_sheet.setFixedSize(new_w, new_h)

    def set_lock_scaffolding(self, locked: bool):
        self.lock_scaffolding = locked
        border_style = "1px dashed #94a3b8;" if locked else "1px dashed #f59e0b;"
        bg_style = "#f8fafc;" if locked else "#fffbeb;"
        qss = f"QFrame {{ background-color: {bg_style} border: {border_style} border-radius: 4px; padding: 4px; }}"
        self.header_frame.setStyleSheet(qss)
        self.footer_frame.setStyleSheet(qss)

    def set_metadata(self, meta: Dict[str, Any]):
        self.metadata.update(meta)
        self._refresh_scaffolding_display()

    def _refresh_scaffolding_display(self):
        """Cập nhật hiển thị Quốc hiệu, Số hiệu, Nơi nhận và Khối ký theo metadata"""
        std = self.metadata.get("standard_id", "ND30").upper()
        sup = self.metadata.get("superior_agency", "").strip().upper()
        iss = self.metadata.get("issuing_agency", "").strip().upper()
        num = self.metadata.get("doc_number", "Số: .../QĐ-UBND")
        date = self.metadata.get("date_str", "Hà Nội, ngày ... tháng ... năm 2026")

        div_sym = "———*———" if std == "HD36" else "—————"

        # Cột Trái Header
        html_l = f"<div style='text-align: center; font-size: 11pt;'>"
        if sup:
            html_l += f"<div style='font-size: 11pt;'>{sup}</div>"
        html_l += f"<div style='font-weight: bold; font-size: 11.5pt;'>{iss or 'CƠ QUAN BAN HÀNH'}</div>"
        html_l += f"<div style='font-weight: bold; font-size: 8pt; margin: 1px 0;'>{div_sym}</div>"
        html_l += f"<div style='font-size: 11.5pt; margin-top: 2px;'>{num}</div>"
        html_l += "</div>"
        self.lbl_header_left.setText(html_l)

        # Cột Phải Header
        html_r = f"<div style='text-align: center; font-size: 11.5pt;'>"
        if std == "HD36":
            html_r += "<div style='font-weight: bold; font-size: 12pt;'>ĐẢNG CỘNG SẢN VIỆT NAM</div>"
            html_r += "<div style='font-weight: bold; font-size: 9pt; margin: 1px 0;'>———————————————</div>"
        else:
            html_r += "<div style='font-weight: bold; font-size: 11.5pt;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>"
            html_r += "<div style='font-weight: bold; font-size: 12pt;'>Độc lập - Tự do - Hạnh phúc</div>"
            html_r += "<div style='font-weight: bold; font-size: 9pt; margin: 1px 0;'>———————————————</div>"
        html_r += f"<div style='font-style: italic; font-size: 12pt; margin-top: 3px;'>{date}</div>"
        html_r += "</div>"
        self.lbl_header_right.setText(html_r)

        # Tiêu đề & Trích yếu
        self.lbl_doc_type.setText(self.metadata.get("doc_type", "QUYẾT ĐỊNH").upper())
        self.lbl_doc_subject.setText(self.metadata.get("doc_subject", ""))

        # Cột Trái Footer: Nơi nhận
        recs = self.metadata.get("recipients", [])
        html_fl = "<div style='font-size: 10.5pt; line-height: 1.2;'>"
        html_fl += "<div style='font-weight: bold; font-style: italic;'>Nơi nhận:</div>"
        for r in recs:
            html_fl += f"<div>{r}</div>"
        html_fl += "</div>"
        self.lbl_footer_left.setText(html_fl)

        # Cột Phải Footer: Thẩm quyền & Chữ ký
        auth = self.metadata.get("authority_prefix", "TM. ỦY BAN NHÂN DÂN")
        pos = self.metadata.get("signer_position", "CHỦ TỊCH")
        name = self.metadata.get("signer_name", "Nguyễn Văn A")

        html_fr = "<div style='text-align: center; font-size: 12pt; line-height: 1.25;'>"
        html_fr += f"<div style='font-weight: bold;'>{auth}</div>"
        html_fr += f"<div style='font-weight: bold;'>{pos}</div>"
        html_fr += "<div style='height: 48px;'>&nbsp;<br/>&nbsp;</div>"  # Khoảng trống ký
        html_fr += f"<div style='font-weight: bold; font-size: 12.5pt;'>{name}</div>"
        html_fr += "</div>"
        self.lbl_footer_right.setText(html_fr)

    def _load_sample_body(self):
        sample_html = """
        <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;</p>
        <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Theo đề nghị của Chánh Văn phòng,</p>
        <p style='text-align: center; font-weight: bold; margin: 12px 0;'>QUYẾT ĐỊNH:</p>
        <p style='text-align: justify; text-indent: 36px;'><b>Điều 1.</b> Ban hành kèm theo Quyết định này Kế hoạch chuyển đổi số và chuẩn hóa văn bản hành chính năm 2026.</p>
        <p style='text-align: justify; text-indent: 36px;'><b>Điều 2.</b> Quyết định này có hiệu lực thi hành kể từ ngày ký.</p>
        <p style='text-align: justify; text-indent: 36px;'><b>Điều 3.</b> Chánh Văn phòng, Thủ trưởng các cơ quan, đơn vị có liên quan chịu trách nhiệm thi hành Quyết định này./.</p>
        """
        self.body_editor.setHtml(sample_html)

    def _on_cursor_changed(self):
        cursor = self.body_editor.textCursor()
        line = cursor.blockNumber() + 1
        col = cursor.positionInBlock() + 1
        self.cursor_position_changed.emit(line, col)

    # ==================== ĐIỀU KHIỂN ĐỊNH DẠNG TỪ RIBBON ====================
    def set_font_bold(self):
        cursor = self.body_editor.textCursor()
        fmt = cursor.charFormat()
        weight = QFont.Weight.Normal if fmt.fontWeight() == QFont.Weight.Bold else QFont.Weight.Bold
        fmt.setFontWeight(weight)
        cursor.mergeCharFormat(fmt)

    def set_font_italic(self):
        cursor = self.body_editor.textCursor()
        fmt = cursor.charFormat()
        fmt.setFontItalic(not fmt.fontItalic())
        cursor.mergeCharFormat(fmt)

    def set_font_underline(self):
        cursor = self.body_editor.textCursor()
        fmt = cursor.charFormat()
        fmt.setFontUnderline(not fmt.fontUnderline())
        cursor.mergeCharFormat(fmt)

    def set_font_size(self, size_pt: float):
        cursor = self.body_editor.textCursor()
        fmt = cursor.charFormat()
        fmt.setFontPointSize(size_pt)
        cursor.mergeCharFormat(fmt)

    def set_font_color(self, color: QColor):
        cursor = self.body_editor.textCursor()
        fmt = cursor.charFormat()
        fmt.setForeground(color)
        cursor.mergeCharFormat(fmt)

    def set_alignment(self, align: Qt.AlignmentFlag):
        self.body_editor.setAlignment(align)

    def set_line_spacing(self, spacing_val: float):
        cursor = self.body_editor.textCursor()
        block_fmt = cursor.blockFormat()
        block_fmt.setLineHeight(int(spacing_val * 100), 2)  # ProportionalHeight
        cursor.setBlockFormat(block_fmt)

    def set_first_line_indent(self, mm: float = 12.7):
        cursor = self.body_editor.textCursor()
        block_fmt = cursor.blockFormat()
        # 1mm ~ 3.78px tại 96 DPI
        px = int(mm * 3.7795)
        block_fmt.setTextIndent(px)
        cursor.setBlockFormat(block_fmt)

    def insert_symbol(self, sym: str):
        cursor = self.body_editor.textCursor()
        cursor.insertText(sym)

    def insert_table(self, rows: int, cols: int, border_type: str = "grid"):
        cursor = self.body_editor.textCursor()
        tbl_fmt = QTextTableFormat()
        tbl_fmt.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tbl_fmt.setCellPadding(4)
        tbl_fmt.setCellSpacing(0)
        tbl_fmt.setBorder(1 if border_type == "grid" else 0)
        tbl_fmt.setBorderStyle(QTextTableFormat.BorderStyle.BorderStyle_Solid)
        tbl_fmt.setBorderBrush(QColor("#000000"))
        tbl_fmt.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        cursor.insertTable(rows, cols, tbl_fmt)

    def insert_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Chọn hình ảnh chèn vào văn bản", "", "Hình ảnh (*.png *.jpg *.jpeg *.bmp)"
        )
        if file_path:
            cursor = self.body_editor.textCursor()
            img = QPixmap(file_path)
            # Giới hạn chiều rộng ảnh không vượt quá bề rộng trang soạn thảo
            max_w = self.page_sheet.width() - 170
            if img.width() > max_w:
                img = img.scaledToWidth(max_w, Qt.TransformationMode.SmoothTransformation)
            cursor.insertImage(img.toImage())

    def insert_page_break(self):
        cursor = self.body_editor.textCursor()
        cursor.insertHtml("<div style='page-break-before: always; margin-top: 20px; border-top: 1px dashed #94a3b8;'></div><br/>")

    def get_body_paragraphs(self) -> List[str]:
        """Trích xuất danh sách các đoạn văn bản trong Body để chuyển giao sang OpenXmlBuilder"""
        raw_text = self.body_editor.toPlainText()
        return [p.strip() for p in raw_text.splitlines() if p.strip()]
