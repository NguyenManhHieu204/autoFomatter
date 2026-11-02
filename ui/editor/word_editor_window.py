"""
ui/editor/word_editor_window.py: Cửa sổ Soạn thảo Văn bản Tích hợp Toàn diện
(Embedded Administrative Word Processor Window)
- Tích hợp trọn bộ: Ribbon Bar, Thước căn lề Ruler, Khung vẽ A4 Canvas, Thanh trạng thái Status Bar
- Các hành động tác nghiệp văn phòng:
  + Lưu tệp (.docx) tương thích 100% Microsoft Word
  + Xuất PDF chuẩn in ấn qua Word COM Engine
  + Ký số trực tiếp (mở SignatureDialog với định vị khối ký chuẩn NĐ 30)
  + Rà soát tuân thủ thể thức thời gian thực (Legal Compliance Audit)
"""

import os
import sys
from typing import Dict, Any, Optional

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QScrollArea, QSlider, QStatusBar, QMessageBox,
    QFileDialog, QFrame, QApplication
)
from PyQt6.QtGui import QIcon, QFont, QColor
from PyQt6.QtCore import Qt, QSize

from ui.editor.administrative_ribbon import AdministrativeRibbonBar
from ui.editor.document_ruler import HorizontalRuler
from ui.editor.document_canvas import DocumentCanvas
from ui.editor.table_properties_dlg import TablePropertiesDialog
from core.editor_core.openxml_builder import OpenXmlBuilder
from core.legal_constraint_guard import LegalConstraintGuard
from core.win32_helper import Win32WordHelper
from ui.signature_dialog import SignatureDialog
from ui.icon_factory import get_ui_icon


class WordEditorWindow(QMainWindow):
    """Cửa sổ Trình Soạn thảo Hành chính Tích hợp (Embedded Administrative Word Processor)"""

    def __init__(self, parent=None, initial_metadata: Optional[Dict[str, Any]] = None, current_file_path: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Auto-Formatter 1.2 — Trình Soạn thảo Văn bản Hành chính & Đảng Tích hợp")
        self.resize(1180, 860)
        self.current_file_path = current_file_path
        self.win32_helper = Win32WordHelper()

        # Widget trung tâm
        central_widget = QWidget()
        central_widget.setStyleSheet("background-color: #cbd5e1;")
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 1. THANH TÁC VỤ ĐỈNH (TOP TITLE & ACTION BAR)
        self._setup_top_action_bar(main_layout)

        # 2. DẢI BĂNG ĐIỀU KHIỂN (RIBBON BAR)
        self.ribbon = AdministrativeRibbonBar()
        main_layout.addWidget(self.ribbon)

        # 3. THƯỚC CĂN LỀ NGANG (RULER)
        self.ruler_container = QWidget()
        self.ruler_container.setStyleSheet("background-color: #e2e8f0; border-bottom: 1px solid #cbd5e1;")
        ruler_layout = QHBoxLayout(self.ruler_container)
        ruler_layout.setContentsMargins(0, 0, 0, 0)
        self.ruler = HorizontalRuler()
        ruler_layout.addWidget(self.ruler)
        main_layout.addWidget(self.ruler_container)

        # 4. VÙNG KHUNG VẼ SOẠN THẢO (SCROLLABLE CANVAS)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #94a3b8; }")

        self.canvas = DocumentCanvas()
        if initial_metadata:
            self.canvas.set_metadata(initial_metadata)

        self.scroll_area.setWidget(self.canvas)
        main_layout.addWidget(self.scroll_area, 1)

        # 5. THANH TRẠNG THÁI (STATUS BAR)
        self._setup_status_bar()

        # Kết nối các sự kiện tương tác
        self._connect_signals()

    def _setup_top_action_bar(self, parent_layout: QVBoxLayout):
        """Thanh tiêu đề chứa các nút hành động cốt lõi: Lưu DOCX, Xuất PDF, Ký số ngay"""
        bar = QFrame()
        bar.setFixedHeight(48)
        bar.setStyleSheet("background-color: #1e293b; border-bottom: 1px solid #0f172a;")
        bar_layout = QHBoxLayout(bar)
        bar_layout.setContentsMargins(16, 6, 16, 6)
        bar_layout.setSpacing(10)

        lbl_app = QLabel("SOẠN THẢO VĂN BẢN (WYSIWYG)")
        lbl_app.setStyleSheet("color: #38bdf8; font-weight: bold; font-size: 13px; letter-spacing: 0.5px;")
        bar_layout.addWidget(lbl_app)

        bar_layout.addStretch()

        # Nút Lưu file DOCX
        self.btn_save = QPushButton("Lưu tệp (.docx)")
        self.btn_save.setIcon(get_ui_icon("file_word"))
        self.btn_save.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: white; font-weight: bold;
                padding: 6px 14px; border-radius: 4px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #1d4ed8; }
        """)
        self.btn_save.clicked.connect(self._save_document)
        bar_layout.addWidget(self.btn_save)

        # Nút Xuất PDF
        self.btn_export_pdf = QPushButton("Xuất PDF")
        self.btn_export_pdf.setIcon(get_ui_icon("file_pdf"))
        self.btn_export_pdf.setStyleSheet("""
            QPushButton {
                background-color: #dc2626; color: white; font-weight: bold;
                padding: 6px 14px; border-radius: 4px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #b91c1c; }
        """)
        self.btn_export_pdf.clicked.connect(self._export_pdf)
        bar_layout.addWidget(self.btn_export_pdf)

        # Nút Ký số ngay
        self.btn_sign = QPushButton("Ký số ngay (NĐ 30)")
        self.btn_sign.setIcon(get_ui_icon("digital_sign"))
        self.btn_sign.setStyleSheet("""
            QPushButton {
                background-color: #7c3aed; color: white; font-weight: bold;
                padding: 6px 16px; border-radius: 4px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #6d28d9; }
        """)
        self.btn_sign.clicked.connect(self._sign_document_now)
        bar_layout.addWidget(self.btn_sign)

        parent_layout.addWidget(bar)

    def _setup_status_bar(self):
        status = QStatusBar()
        status.setStyleSheet("background-color: #f1f5f9; color: #475569; font-size: 12px; border-top: 1px solid #cbd5e1;")
        self.setStatusBar(status)

        self.lbl_cursor_pos = QLabel("Dòng: 1, Cột: 1")
        status.addWidget(self.lbl_cursor_pos)

        self.lbl_compliance_badge = QLabel("✔ Khung Thể thức: Đã Khóa Bảo vệ (NĐ 30 / HD 36)")
        self.lbl_compliance_badge.setStyleSheet("color: #15803d; font-weight: 600; margin-left: 20px;")
        status.addWidget(self.lbl_compliance_badge)

        # Thanh trượt Zoom
        zoom_container = QWidget()
        z_layout = QHBoxLayout(zoom_container)
        z_layout.setContentsMargins(0, 0, 10, 0)
        z_layout.setSpacing(6)

        self.lbl_zoom = QLabel("100%")
        self.lbl_zoom.setFixedWidth(45)

        self.slider_zoom = QSlider(Qt.Orientation.Horizontal)
        self.slider_zoom.setRange(50, 180)
        self.slider_zoom.setValue(100)
        self.slider_zoom.setFixedWidth(100)
        self.slider_zoom.valueChanged.connect(self._on_zoom_changed)

        z_layout.addWidget(QLabel("Thu phóng:"))
        z_layout.addWidget(self.slider_zoom)
        z_layout.addWidget(self.lbl_zoom)

        status.addPermanentWidget(zoom_container)

    def _connect_signals(self):
        # 1. Định dạng văn bản từ Ribbon sang Canvas
        self.ribbon.format_bold_triggered.connect(self.canvas.set_font_bold)
        self.ribbon.format_italic_triggered.connect(self.canvas.set_font_italic)
        self.ribbon.format_underline_triggered.connect(self.canvas.set_font_underline)
        self.ribbon.font_size_changed.connect(self.canvas.set_font_size)
        self.ribbon.font_color_changed.connect(self.canvas.set_font_color)

        self.ribbon.align_left_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignLeft))
        self.ribbon.align_center_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignCenter))
        self.ribbon.align_right_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignRight))
        self.ribbon.align_justify_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignJustify))

        self.ribbon.line_spacing_changed.connect(self.canvas.set_line_spacing)
        self.ribbon.indent_first_line_triggered.connect(self.canvas.set_first_line_indent)
        self.ribbon.insert_symbol_requested.connect(self.canvas.insert_symbol)

        # 2. Thao tác Chèn (Insert)
        self.ribbon.insert_table_requested.connect(self._show_insert_table_dialog)
        self.ribbon.insert_image_requested.connect(self.canvas.insert_image)
        self.ribbon.page_break_requested.connect(self.canvas.insert_page_break)

        # 3. Bố cục & Thước kẻ
        self.ribbon.toggle_ruler_triggered.connect(self.ruler_container.setVisible)
        self.ribbon.lock_scaffolding_toggled.connect(self.canvas.set_lock_scaffolding)

        # 4. Công cụ thể thức
        self.ribbon.check_compliance_requested.connect(self._audit_document_compliance)
        self.ribbon.switch_doc_type_requested.connect(self._on_switch_doc_type)

        # 5. Cập nhật vị trí con trỏ
        self.canvas.cursor_position_changed.connect(self._on_canvas_cursor_moved)

    def _on_canvas_cursor_moved(self, line: int, col: int):
        self.lbl_cursor_pos.setText(f"Dòng: {line}, Cột: {col}")

    def _on_zoom_changed(self, val: int):
        self.lbl_zoom.setText(f"{val}%")
        factor = val / 100.0
        self.canvas.set_zoom(factor)
        self.ruler.set_zoom(factor)

    def _show_insert_table_dialog(self):
        dlg = TablePropertiesDialog(self)
        if dlg.exec():
            cfg = dlg.get_table_config()
            self.canvas.insert_table(cfg["rows"], cfg["cols"], cfg["border_type"])

    def _on_switch_doc_type(self, new_type: str):
        self.canvas.metadata["doc_type"] = new_type
        self.canvas._refresh_scaffolding_display()
        self.canvas.document_modified.emit()

    def _audit_document_compliance(self):
        """Rà soát tuân thủ thể thức pháp lý thời gian thực"""
        temp_docx = os.path.join(os.environ.get("TEMP", "."), "af_audit_temp.docx")
        try:
            doc = self._build_docx_object()
            doc.save(temp_docx)
            audit_res = LegalConstraintGuard.validate_document(doc, standard_id=self.canvas.metadata.get("standard_id", "ND30"))

            score = audit_res["score"]
            errs = audit_res["error_count"]
            warns = audit_res["warning_count"]

            msg = f"<b>ĐIỂM CHUẨN HÓA THỂ THỨC: {score}/100</b><br/><br/>"
            if audit_res["is_compliant"]:
                msg += "✔ Văn bản tuân thủ 100% quy chuẩn Nghị định 30/2020/NĐ-CP và Hướng dẫn 36-HD/VPTW.<br/>Sẵn sàng xuất bản và ký số điện tử."
            else:
                msg += f"Phát hiện {errs} lỗi nghiêm trọng và {warns} cảnh báo thể thức:<br/>"
                for idx, v in enumerate(audit_res["violations"][:5]):
                    msg += f"<br/><b>{idx + 1}. [{v['severity'].upper()}]</b> {v['message']}<br/><i>👉 Gợi ý: {v['suggestion']}</i>"

            QMessageBox.information(self, "Kết quả Rà soát Thể thức", msg)
        except Exception as e:
            QMessageBox.warning(self, "Lỗi kiểm tra", f"Không thể rà soát: {str(e)}")
        finally:
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass

    def _build_docx_object(self):
        """Đóng gói dữ liệu Canvas hiện tại thành tài liệu docx chuẩn hóa"""
        meta = dict(self.canvas.metadata)
        meta["body_paragraphs"] = self.canvas.get_body_paragraphs()
        return OpenXmlBuilder.create_scaffolded_document(meta, lock_scaffolding=self.canvas.lock_scaffolding)

    def _save_document(self) -> Optional[str]:
        target_path = self.current_file_path
        if not target_path:
            suggested = f"{self.canvas.metadata.get('doc_type', 'VanBan')}_ChuanHoa.docx"
            target_path, _ = QFileDialog.getSaveFileName(
                self, "Lưu tài liệu Word chuẩn hóa", suggested, "Tệp Word (*.docx)"
            )

        if not target_path:
            return None

        try:
            doc = self._build_docx_object()
            doc.save(target_path)
            self.current_file_path = target_path
            QMessageBox.information(
                self, "Lưu tệp thành công",
                f"Đã lưu văn bản chuẩn thể thức tại:\n{target_path}\n\nTương thích 100% Microsoft Word và các hệ thống QLVB&ĐH."
            )
            return target_path
        except Exception as e:
            QMessageBox.critical(self, "Lỗi lưu tệp", f"Không thể lưu tài liệu: {str(e)}")
            return None

    def _export_pdf(self) -> Optional[str]:
        docx_path = self._save_document()
        if not docx_path:
            return None

        base, _ = os.path.splitext(docx_path)
        pdf_target = f"{base}.pdf"

        self.btn_export_pdf.setEnabled(False)
        self.btn_export_pdf.setText("Đang xuất...")
        QApplication.processEvents()

        try:
            res = self.win32_helper.process_document_via_word(docx_path, export_pdf_also=True)
            if res.get("pdf_exported") and os.path.exists(pdf_target):
                QMessageBox.information(
                    self, "Xuất PDF Thành công",
                    f"Đã xuất tệp PDF ngoại tuyến chuẩn in ấn:\n{pdf_target}"
                )
                return pdf_target
            else:
                QMessageBox.warning(self, "Xuất PDF", res.get("message", "Không thể xuất PDF qua MS Word COM."))
                return None
        finally:
            self.btn_export_pdf.setEnabled(True)
            self.btn_export_pdf.setText("Xuất PDF")

    def _sign_document_now(self):
        """Xuất PDF và mở trực tiếp Module Ký số điện tử với vị trí khung ký định vị tự động"""
        pdf_path = self._export_pdf()
        if not pdf_path or not os.path.exists(pdf_path):
            return

        agency_name = self.canvas.metadata.get("issuing_agency", "")
        sign_dlg = SignatureDialog(self, pdf_path=pdf_path, default_agency=agency_name)
        # Tự động điền thông tin người ký từ khung soạn thảo
        sign_dlg.txt_signer_name.setText(self.canvas.metadata.get("signer_name", ""))
        sign_dlg.txt_signer_title.setText(self.canvas.metadata.get("signer_position", ""))
        sign_dlg.exec()
