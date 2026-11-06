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
from PyQt6.QtGui import QIcon, QFont, QColor, QShortcut, QKeySequence
from PyQt6.QtCore import Qt, QSize

from ui.editor.administrative_ribbon import AdministrativeRibbonBar
from ui.editor.document_ruler import HorizontalRuler
from ui.editor.document_canvas import DocumentCanvas
from ui.editor.table_properties_dlg import TablePropertiesDialog
from core.editor_core.openxml_builder import OpenXmlBuilder
from core.legal_constraint_guard import LegalConstraintGuard
from core.win32_helper import Win32WordHelper
from ui.signature_dialog import SignatureDialog


class WordEditorWindow(QMainWindow):
    """Cửa sổ Trình Soạn thảo Hành chính Tích hợp (Embedded Administrative Word Processor)"""

    def __init__(self, parent=None, initial_metadata: Optional[Dict[str, Any]] = None, current_file_path: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Soạn thảo Văn bản — Auto-Formatter")
        self.resize(1180, 860)
        # Luôn mở sẵn toàn màn hình (Maximized) theo trải nghiệm soạn thảo chuyên nghiệp
        self.setWindowState(Qt.WindowState.WindowMaximized)
        self.current_file_path = current_file_path
        self.win32_helper = Win32WordHelper()

        # Widget trung tâm (Bàn làm việc màu xám đậm chuyên nghiệp chuẩn Word)
        central_widget = QWidget()
        central_widget.setObjectName("central_workspace")
        central_widget.setStyleSheet("QWidget#central_workspace { background-color: #718096; }")
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
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: #718096;
            }
        """)
        self.scroll_area.viewport().setStyleSheet("background-color: #718096;")

        self.canvas = DocumentCanvas()
        if initial_metadata:
            self.canvas.set_metadata(initial_metadata)

        self.scroll_area.setWidget(self.canvas)
        main_layout.addWidget(self.scroll_area, 1)

        # 5. THANH TRẠNG THÁI (STATUS BAR)
        self._setup_status_bar()

        # Kết nối các sự kiện tương tác
        self._connect_signals()

        # Nạp tệp hiện có nếu được chỉ định
        if self.current_file_path and os.path.exists(self.current_file_path):
            self._load_document_from_file(self.current_file_path)

    def _setup_top_action_bar(self, parent_layout: QVBoxLayout):
        """Thanh tiêu đề chứa các nút hành động cốt lõi: Lưu DOCX, Xuất PDF, Ký số ngay"""
        bar = QFrame()
        bar.setFixedHeight(50)
        bar.setStyleSheet("background-color: #ffffff; border-bottom: 1px solid #cbd5e1;")
        bar_layout = QHBoxLayout(bar)
        bar_layout.setContentsMargins(16, 6, 16, 6)
        bar_layout.setSpacing(10)

        # Tiêu đề tài liệu
        doc_title = "Trình soạn thảo văn bản hành chính"
        if self.current_file_path:
            doc_title = os.path.basename(self.current_file_path)
        self.lbl_doc_title = QLabel(doc_title)
        self.lbl_doc_title.setStyleSheet("color: #0f172a; font-weight: 700; font-size: 13.5px;")
        bar_layout.addWidget(self.lbl_doc_title)

        bar_layout.addStretch()

        # Nút Lưu văn bản (Hành động chính - Royal Blue)
        self.btn_save = QPushButton("Lưu văn bản")
        self.btn_save.setObjectName("btnSaveDocx")
        self.btn_save.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: #ffffff; font-weight: 600;
                font-size: 12.5px; padding: 7px 18px; border-radius: 6px;
                border: 1px solid #1d4ed8;
            }
            QPushButton:hover { background-color: #1d4ed8; }
            QPushButton:pressed { background-color: #1e40af; }
        """)
        self.btn_save.setToolTip("Lưu tài liệu định dạng Word (.docx) chuẩn thể thức")
        self.btn_save.clicked.connect(self._save_document)
        bar_layout.addWidget(self.btn_save)

        # Nút Xuất PDF (Thứ cấp - Trang nhã, chuyên nghiệp)
        self.btn_export_pdf = QPushButton("Xuất PDF")
        self.btn_export_pdf.setObjectName("btnExportPdf")
        self.btn_export_pdf.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #1e293b; font-weight: 600;
                font-size: 12.5px; padding: 7px 18px; border-radius: 6px;
                border: 1px solid #cbd5e1;
            }
            QPushButton:hover { background-color: #f8fafc; border-color: #dc2626; color: #dc2626; }
            QPushButton:pressed { background-color: #fee2e2; }
            QPushButton:disabled { background-color: #f1f5f9; color: #94a3b8; border-color: #e2e8f0; }
        """)
        self.btn_export_pdf.setToolTip("Xuất văn bản sang tệp PDF chuẩn in ấn và lưu trữ theo NĐ 30")
        self.btn_export_pdf.clicked.connect(self._export_pdf)
        bar_layout.addWidget(self.btn_export_pdf)

        # Nút Ký số (Ký số điện tử - Indigo hài hòa)
        self.btn_sign = QPushButton("Ký số")
        self.btn_sign.setObjectName("btnSignNow")
        self.btn_sign.setStyleSheet("""
            QPushButton {
                background-color: #4f46e5; color: #ffffff; font-weight: 600;
                font-size: 12.5px; padding: 7px 18px; border-radius: 6px;
                border: 1px solid #4338ca;
            }
            QPushButton:hover { background-color: #4338ca; }
            QPushButton:pressed { background-color: #3730a3; }
        """)
        self.btn_sign.setToolTip("Xuất PDF và ký số điện tử với Token Ban Cơ yếu / Chữ ký số")
        self.btn_sign.clicked.connect(self._sign_document_now)
        bar_layout.addWidget(self.btn_sign)

        parent_layout.addWidget(bar)

    def _setup_status_bar(self):
        status = QStatusBar()
        status.setStyleSheet("background-color: #ffffff; color: #475569; font-size: 11.5px; border-top: 1px solid #e2e8f0;")
        self.setStatusBar(status)

        self.lbl_cursor_pos = QLabel("Dòng: 1, Cột: 1")
        status.addWidget(self.lbl_cursor_pos)

        self.lbl_compliance_badge = QLabel("Chuẩn thể thức: Đã khóa bảo vệ (NĐ 30 / HD 36)")
        self.lbl_compliance_badge.setStyleSheet("color: #16a34a; font-weight: 500; margin-left: 20px;")
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
        # 0. Tác vụ Tệp từ Ribbon
        self.ribbon.save_document_requested.connect(self._save_document)
        self.ribbon.export_pdf_requested.connect(self._export_pdf)
        self.ribbon.sign_document_requested.connect(self._sign_document_now)

        # 1. Định dạng văn bản từ Ribbon sang Canvas
        self.ribbon.format_bold_triggered.connect(self.canvas.set_font_bold)
        self.ribbon.format_italic_triggered.connect(self.canvas.set_font_italic)
        self.ribbon.format_underline_triggered.connect(self.canvas.set_font_underline)
        self.ribbon.font_size_changed.connect(self.canvas.set_font_size)
        self.ribbon.font_family_changed.connect(self.canvas.set_font_family)
        self.ribbon.font_color_changed.connect(self.canvas.set_font_color)

        self.ribbon.align_left_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignLeft))
        self.ribbon.align_center_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignCenter))
        self.ribbon.align_right_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignRight))
        self.ribbon.align_justify_triggered.connect(lambda: self.canvas.set_alignment(Qt.AlignmentFlag.AlignJustify))

        self.ribbon.line_spacing_changed.connect(self.canvas.set_line_spacing)
        self.ribbon.indent_first_line_triggered.connect(self.canvas.set_first_line_indent)
        self.ribbon.increase_indent_triggered.connect(self.canvas.increase_indent)
        self.ribbon.decrease_indent_triggered.connect(self.canvas.decrease_indent)
        self.ribbon.remove_space_before_triggered.connect(self.canvas.remove_space_before)
        self.ribbon.remove_space_after_triggered.connect(self.canvas.remove_space_after)
        self.ribbon.remove_all_spacing_triggered.connect(self.canvas.remove_all_spacing)
        self.ribbon.add_space_before_triggered.connect(self.canvas.add_space_before)
        self.ribbon.add_space_after_triggered.connect(self.canvas.add_space_after)
        self.ribbon.apply_nd30_spacing_triggered.connect(self.canvas.apply_nd30_spacing)
        self.ribbon.bullet_list_triggered.connect(self.canvas.insert_bullet_list)
        self.ribbon.numbered_list_triggered.connect(self.canvas.insert_numbered_list)
        self.ribbon.insert_symbol_requested.connect(self.canvas.insert_symbol)

        # 2. Thao tác Chèn & Bảng biểu
        self.ribbon.insert_table_requested.connect(self._show_insert_table_dialog)
        self.ribbon.insert_image_requested.connect(self.canvas.insert_image)
        self.ribbon.page_break_requested.connect(self.canvas.insert_page_break)

        self.ribbon.table_insert_row_above_triggered.connect(self.canvas.insert_table_row_above)
        self.ribbon.table_insert_row_below_triggered.connect(self.canvas.insert_table_row_below)
        self.ribbon.table_delete_row_triggered.connect(self.canvas.delete_table_row)
        self.ribbon.table_insert_col_left_triggered.connect(self.canvas.insert_table_col_left)
        self.ribbon.table_insert_col_right_triggered.connect(self.canvas.insert_table_col_right)
        self.ribbon.table_delete_col_triggered.connect(self.canvas.delete_table_col)
        self.ribbon.table_merge_cells_triggered.connect(self.canvas.merge_table_cells)
        self.ribbon.table_split_cell_triggered.connect(self.canvas.split_table_cell)

        # 3. Bố cục & Thước kẻ
        self.ribbon.toggle_ruler_triggered.connect(self.ruler_container.setVisible)
        self.ribbon.page_orientation_changed.connect(self.canvas.set_page_orientation)
        self.ribbon.margins_preset_triggered.connect(lambda p: self._on_apply_margins_preset())
        self.ribbon.apply_nd30_spacing_triggered.connect(self._on_apply_margins_preset)
        self.ribbon.lock_scaffolding_toggled.connect(self.canvas.set_lock_scaffolding)

        # 4. Công cụ thể thức & Sửa khung
        self.ribbon.edit_metadata_requested.connect(self.canvas.open_metadata_editor)
        self.ribbon.check_compliance_requested.connect(self._audit_document_compliance)
        self.ribbon.switch_doc_type_requested.connect(self._on_switch_doc_type)

        # Phím tắt F2 để mở sửa thể thức
        self.shortcut_f2 = QShortcut(QKeySequence("F2"), self)
        self.shortcut_f2.activated.connect(self.canvas.open_metadata_editor)

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

    def _on_apply_margins_preset(self):
        """Tự động căn lề và định dạng chuẩn NĐ 30, loại bỏ thụt dòng thủ công xô lệch"""
        stats = self.canvas.apply_nd30_standards_and_alignment()
        QMessageBox.information(
            self, "Căn lề & Định dạng chuẩn NĐ 30/2020",
            f"<b>ĐÃ TỰ ĐỘNG CĂN CHỈNH TOÀN BỘ VĂN BẢN VỀ CHUẨN NĐ 30/2020:</b><br/><br/>"
            f"• <b>Kích thước lề in:</b> Trái 30 mm, Phải 15 mm, Trên 20 mm, Dưới 20 mm.<br/>"
            f"• <b>Đoạn văn bản đã chuẩn hóa:</b> {stats['fixed_paragraphs']} đoạn (căn đều hai bên - Justify, giãn dòng 1.25, giãn đoạn 0/6 pt).<br/>"
            f"• <b>Khoảng trắng thụt lề thủ công đã xóa:</b> {stats['stripped_spaces']} vị trí xô lệch dòng.<br/>"
            f"• <b>Thụt đầu dòng:</b> 1.0 - 1.27 cm chuẩn quy định."
        )

    def _on_switch_doc_type(self, new_type: str):
        tpl = self.canvas.get_template_for_type(new_type)
        # Giữ lại thông tin cơ quan, người ký tùy biến nếu có
        old_signer = self.canvas.metadata.get("signer_name", "")
        old_issuing = self.canvas.metadata.get("issuing_agency", "")
        old_superior = self.canvas.metadata.get("superior_agency", "")
        old_loc = self.canvas.metadata.get("location", "")

        self.canvas.metadata.update(tpl)
        if old_signer and old_signer != "Nguyễn Văn A":
            self.canvas.metadata["signer_name"] = old_signer
        if old_issuing and old_issuing not in ["ỦY BAN NHÂN DÂN XÃ", "ĐẢNG BỘ XÃ", "CƠ QUAN BAN HÀNH"]:
            self.canvas.metadata["issuing_agency"] = old_issuing
        if old_superior and old_superior not in ["ỦY BAN NHÂN DÂN HUYỆN", "HUYỆN ỦY"]:
            self.canvas.metadata["superior_agency"] = old_superior
        if old_loc and old_loc not in ["Hà Nội", "Xã"]:
            self.canvas.metadata["location"] = old_loc

        # Nạp cấu trúc nội dung mẫu chuẩn của loại văn bản mới
        if self.canvas.pages and tpl.get("body_html"):
            self.canvas.pages[0].editor.setHtml(tpl["body_html"])

        self.canvas._refresh_scaffolding_display()
        self.canvas.document_modified.emit()

    def _audit_document_compliance(self):
        """Rà soát tuân thủ thể thức pháp lý thời gian thực với bảng phân tích chi tiết và nút sửa nhanh"""
        temp_docx = os.path.join(os.environ.get("TEMP", "."), "af_audit_temp.docx")
        try:
            doc = self._build_docx_object()
            doc.save(temp_docx)
            audit_res = LegalConstraintGuard.validate_document(doc, standard_id=self.canvas.metadata.get("standard_id", "ND30"))

            from ui.editor.compliance_audit_dlg import ComplianceAuditDialog
            dlg = ComplianceAuditDialog(self, audit_result=audit_res, canvas=self.canvas)

            def _refresh_audit():
                try:
                    d = self._build_docx_object()
                    d.save(temp_docx)
                    new_res = LegalConstraintGuard.validate_document(d, standard_id=self.canvas.metadata.get("standard_id", "ND30"))
                    dlg.update_audit_result(new_res)
                except Exception:
                    pass

            dlg.audit_refresh_requested.connect(_refresh_audit)
            dlg.exec()
        except Exception as e:
            QMessageBox.warning(self, "Lỗi kiểm tra", f"Không thể rà soát: {str(e)}")
        finally:
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass

    def _load_document_from_file(self, file_path: str):
        """Nạp nội dung và thể thức từ tệp docx vào trình soạn thảo"""
        try:
            import docx
            import html
            import re
            from core.doc_parser import DocParser

            parser = DocParser(file_path)
            st = parser.parse()
            doc = parser.doc

            # 1. Cập nhật Metadata chuẩn thể thức từ tài liệu thực tế
            meta_update = {}
            if st.detected_standard:
                meta_update["standard_id"] = st.detected_standard
            if st.header.superior_agency:
                meta_update["superior_agency"] = st.header.superior_agency
            elif st.header.has_existing_table:
                meta_update["superior_agency"] = ""

            if st.header.issuing_agency:
                meta_update["issuing_agency"] = st.header.issuing_agency
            if st.header.doc_number:
                meta_update["doc_number"] = st.header.doc_number
            if st.header.location_date:
                meta_update["date_str"] = st.header.location_date
            if st.doc_type:
                meta_update["doc_type"] = st.doc_type
            if st.doc_subject:
                meta_update["doc_subject"] = st.doc_subject
            if st.footer.recipients:
                meta_update["recipients"] = st.footer.recipients
            if st.footer.authority_prefix:
                meta_update["authority_prefix"] = st.footer.authority_prefix
            if st.footer.signer_position:
                meta_update["signer_position"] = st.footer.signer_position
            if st.footer.signer_name:
                meta_update["signer_name"] = st.footer.signer_name

            self.canvas.metadata.update(meta_update)
            self.canvas._refresh_scaffolding_display()

            # Đồng bộ combobox chọn loại văn bản trên thanh Ribbon
            if hasattr(self, 'ribbon') and hasattr(self.ribbon, 'cb_type') and st.doc_type:
                self.ribbon.cb_type.blockSignals(True)
                idx = self.ribbon.cb_type.findText(st.doc_type.upper())
                if idx >= 0:
                    self.ribbon.cb_type.setCurrentIndex(idx)
                else:
                    self.ribbon.cb_type.addItem(st.doc_type.upper())
                    self.ribbon.cb_type.setCurrentText(st.doc_type.upper())
                self.ribbon.cb_type.blockSignals(False)

            # 2. Xác định các đoạn thân bài (Body) cần hiển thị trong editor
            # LOẠI BỎ: Header (Quốc hiệu, cơ quan), Title (Tên loại VB, trích yếu), Footer (Nơi nhận, chữ ký)
            # vì các thành phần này đã được hiển thị chuẩn xác trong các Frame widget chuyên biệt của Canvas!
            body_indices = set(st.body_paragraphs_indices)
            if not body_indices:
                excluded = set(st.title_paragraphs_indices)
                if st.header.raw_paragraphs_count > 0:
                    excluded.update(range(st.header.raw_paragraphs_count))
                if st.footer_para_indices:
                    excluded.update(st.footer_para_indices)
                body_indices = [i for i in range(len(doc.paragraphs)) if i not in excluded]

            # 3. Chuyển đổi các đoạn văn thân bài sang HTML có định dạng (đậm, nghiêng, thụt lề chuẩn NĐ 30)
            html_parts = []
            for i in sorted(body_indices):
                if i < 0 or i >= len(doc.paragraphs):
                    continue
                p = doc.paragraphs[i]
                txt = p.text.strip()
                if not txt:
                    continue

                # Tạo HTML từ các runs để giữ nguyên in đậm / in nghiêng
                run_htmls = []
                for r in p.runs:
                    r_text = r.text
                    if not r_text:
                        continue
                    esc = html.escape(r_text)
                    if r.bold:
                        esc = f"<b>{esc}</b>"
                    if r.italic:
                        esc = f"<i>{esc}</i>"
                    if r.underline:
                        esc = f"<u>{esc}</u>"
                    run_htmls.append(esc)
                inner = "".join(run_htmls) if run_htmls else html.escape(txt)

                # Căn lề chuẩn NĐ 30 cho thân bài
                u_txt = txt.upper()
                if txt.endswith(":") and any(k in u_txt for k in ["QUYẾT ĐỊNH", "QUYẾT NGHỊ", "THÔNG BÁO"]):
                    html_parts.append(f"<p style='text-align: center; font-weight: bold; margin: 10px 0;'>{inner}</p>")
                elif txt.lower().startswith("kính gửi"):
                    html_parts.append(f"<p style='text-align: left; text-indent: 0px; margin: 6px 0;'>{inner}</p>")
                elif re.match(r'^[-–—•*]?\s*(căn\s+cứ|theo\s+đề\s+nghị)\b', txt, re.IGNORECASE):
                    html_parts.append(f"<p style='text-align: justify; text-indent: 36px; font-style: italic; margin: 4px 0;'>{inner}</p>")
                elif re.match(r'^(điều|dieu)\s+\d+', txt, re.IGNORECASE) or re.match(r'^[IVXLCDM]+\.\s+', txt):
                    html_parts.append(f"<p style='text-align: justify; text-indent: 36px; font-weight: bold; margin: 6px 0;'>{inner}</p>")
                else:
                    html_parts.append(f"<p style='text-align: justify; text-indent: 36px; margin: 4px 0;'>{inner}</p>")

            # 4. Kiểm tra các bảng trong thân bài (bỏ qua bảng Header đầu và bảng Footer cuối)
            if doc.tables:
                for idx_tbl, tbl in enumerate(doc.tables):
                    if idx_tbl == 0 and st.header.has_existing_table:
                        continue
                    if idx_tbl == len(doc.tables) - 1 and st.footer.has_existing_table:
                        continue
                    tbl_html = ["<table border='1' cellspacing='0' cellpadding='4' style='width: 100%; border-collapse: collapse; margin: 10px 0;'>"]
                    for row in tbl.rows:
                        tbl_html.append("<tr>")
                        for cell in row.cells:
                            c_text = "<br/>".join(html.escape(cp.text.strip()) for cp in cell.paragraphs if cp.text.strip())
                            tbl_html.append(f"<td style='border: 1px solid #000000; padding: 4px;'>{c_text}</td>")
                        tbl_html.append("</tr>")
                    tbl_html.append("</table>")
                    html_parts.append("".join(tbl_html))

            if html_parts:
                self.canvas.body_editor.setHtml("".join(html_parts))
            else:
                self.canvas.body_editor.setPlainText("")

        except Exception as e:
            try:
                import docx
                doc = docx.Document(file_path)
                html_parts = []
                for p in doc.paragraphs:
                    txt = p.text.strip()
                    if txt:
                        html_parts.append(f"<p style='text-indent: 36px; margin: 4px 0;'>{txt}</p>")
                if html_parts:
                    self.canvas.body_editor.setHtml("".join(html_parts))
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

    def _export_pdf_via_qt(self, pdf_target: str) -> bool:
        """Xuất tệp PDF ngoại tuyến độc lập bằng QPdfWriter chuẩn kích thước A4 và lề NĐ 30, hỗ trợ đa trang và xoay ngang độc lập"""
        try:
            from PyQt6.QtGui import QPdfWriter, QTextDocument, QPageSize, QPageLayout, QFont
            from PyQt6.QtCore import QMarginsF

            writer = QPdfWriter(pdf_target)
            writer.setResolution(300)
            page_size = QPageSize(QPageSize.PageSizeId.A4)
            margins = QMarginsF(30, 20, 15, 20)
            layout_portrait = QPageLayout(page_size, QPageLayout.Orientation.Portrait, margins, QPageLayout.Unit.Millimeter)
            layout_landscape = QPageLayout(page_size, QPageLayout.Orientation.Landscape, margins, QPageLayout.Unit.Millimeter)

            meta = self.canvas.metadata
            std = meta.get("standard_id", "ND30").upper()
            sup = meta.get("superior_agency", "").strip().upper()
            iss = meta.get("issuing_agency", "").strip().upper()
            num = meta.get("doc_number", "Số: .../QĐ-UBND")
            date_str = meta.get("date_str", "..., ngày ... tháng ... năm 2026")
            doc_type = meta.get("doc_type", "QUYẾT ĐỊNH")
            doc_subject = meta.get("doc_subject", "")
            auth_prefix = meta.get("authority_prefix", "")
            signer_pos = meta.get("signer_position", "")
            signer_name = meta.get("signer_name", "")
            recipients = meta.get("recipients", [])

            # Header left
            h_left = "<div style='text-align: center; font-family: \"Times New Roman\";'>"
            if sup:
                h_left += f"<div style='font-size: 11pt;'>{sup}</div>"
            h_left += f"<div style='font-weight: bold; font-size: 11.5pt;'>{iss}</div>"
            div_sym = "———*———" if std == "HD36" else "—————"
            h_left += f"<div style='font-size: 8pt; margin: 2px 0;'>{div_sym}</div>"
            h_left += f"<div style='font-size: 11pt; margin-top: 2px;'>{num}</div>"
            h_left += "</div>"

            # Header right
            h_right = "<div style='text-align: center; font-family: \"Times New Roman\";'>"
            if std == "HD36":
                h_right += "<div style='font-weight: bold; font-size: 12pt;'>ĐẢNG CỘNG SẢN VIỆT NAM</div>"
                h_right += "<div style='font-size: 9pt; margin: 2px 0;'>———————————————</div>"
            else:
                h_right += "<div style='font-weight: bold; font-size: 11.5pt;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>"
                h_right += "<div style='font-weight: bold; font-size: 12pt;'>Độc lập - Tự do - Hạnh phúc</div>"
                h_right += "<div style='font-size: 8pt; margin: 2px 0;'>—————————————</div>"
            h_right += f"<div style='font-style: italic; font-size: 12pt; margin-top: 4px;'>{date_str}</div>"
            h_right += "</div>"

            # Footer left (Nơi nhận)
            f_left = "<div style='font-family: \"Times New Roman\"; font-size: 10.5pt; text-align: left;'>"
            f_left += "<b><i>Nơi nhận:</i></b><br/>"
            for r in recipients:
                f_left += f"<div style='font-size: 10pt; line-height: 1.2;'>{r}</div>"
            f_left += "</div>"

            # Footer right (Chữ ký)
            f_right = "<div style='font-family: \"Times New Roman\"; font-size: 12.5pt; text-align: center;'>"
            if auth_prefix:
                f_right += f"<div style='font-weight: bold;'>{auth_prefix}</div>"
            f_right += f"<div style='font-weight: bold;'>{signer_pos}</div>"
            f_right += "<div style='height: 55px;'></div>"
            f_right += f"<div style='font-weight: bold;'>{signer_name}</div>"
            f_right += "</div>"

            footer_html = f"""
            <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-top: 24px;">
                <tr>
                    <td width="45%" valign="top">{f_left}</td>
                    <td width="55%" valign="top">{f_right}</td>
                </tr>
            </table>
            """

            css = """
            <style>
                body { font-family: 'Times New Roman', serif; color: #000000; font-size: 13pt; line-height: 1.3; }
                table { border-collapse: collapse; }
                p { margin: 4px 0; text-indent: 36px; text-align: justify; }
            </style>
            """

            pages = getattr(self.canvas, "pages", [])
            if not pages:
                # Nếu không có danh sách pages, kết xuất dạng đơn lẻ
                body_html = self.canvas.body_editor.toHtml() if hasattr(self.canvas, "body_editor") else ""
                writer.setPageLayout(layout_portrait)
                full_html = f"<html><head><meta charset='utf-8'>{css}</head><body>"
                full_html += f"<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td width='42%' valign='top'>{h_left}</td><td width='58%' valign='top'>{h_right}</td></tr></table>"
                full_html += f"<div style='text-align: center; margin: 18px 0 14px 0;'><div style='font-size: 15pt; font-weight: bold;'>{doc_type}</div><div style='font-size: 13.5pt; font-weight: bold; margin-top: 4px;'>{doc_subject}</div></div>"
                full_html += f"<div style='margin: 12px 0 24px 0;'>{body_html}</div>"
                full_html += footer_html
                full_html += "</body></html>"
                doc = QTextDocument()
                doc.setDefaultFont(QFont("Times New Roman", 13))
                doc.setHtml(full_html)
                doc.print(writer)
                return True

            total_pages = len(pages)
            for idx, p in enumerate(pages):
                if idx > 0:
                    writer.newPage()

                orient = getattr(p, "orientation", "portrait")
                writer.setPageLayout(layout_landscape if orient == "landscape" else layout_portrait)

                p_html = f"<html><head><meta charset='utf-8'>{css}</head><body>"

                if idx == 0:
                    # Trang 1: Khối Quốc hiệu, Cơ quan & Tiêu đề
                    p_html += f"<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td width='42%' valign='top'>{h_left}</td><td width='58%' valign='top'>{h_right}</td></tr></table>"
                    p_html += f"<div style='text-align: center; margin: 18px 0 14px 0;'><div style='font-size: 15pt; font-weight: bold;'>{doc_type}</div><div style='font-size: 13.5pt; font-weight: bold; margin-top: 4px;'>{doc_subject}</div></div>"
                else:
                    # Trang 2+: Đánh số trang căn giữa lề trên per NĐ 30
                    p_html += f"<div style='text-align: center; font-size: 13pt; margin-bottom: 12px;'>- {idx + 1} -</div>"

                # Thân trang
                raw_body = p.editor.toHtml()
                p_html += f"<div style='margin: 8px 0 16px 0;'>{raw_body}</div>"

                # Nếu là trang cuối cùng: Đính kèm Khối Nơi nhận & Chữ ký
                if idx == total_pages - 1:
                    p_html += footer_html

                p_html += "</body></html>"

                doc = QTextDocument()
                doc.setDefaultFont(QFont("Times New Roman", 13))
                doc.setHtml(p_html)
                doc.print(writer)

            return True
        except Exception as e:
            print(f"[Qt PDF Export Error] {e}")
            return False

    def _export_pdf(self) -> Optional[str]:
        docx_path = self._save_document()
        if not docx_path:
            return None

        base, _ = os.path.splitext(docx_path)
        pdf_target = f"{base}.pdf"

        if self.btn_export_pdf:
            self.btn_export_pdf.setEnabled(False)
            self.btn_export_pdf.setText("Đang xuất...")
        if self.btn_sign:
            self.btn_sign.setEnabled(False)
            self.btn_sign.setText("Đang xử lý...")
        QApplication.processEvents()

        try:
            pdf_ok = False
            # 1. Thử xuất qua Word COM nếu khả dụng
            if self.win32_helper.is_word_available():
                try:
                    res = self.win32_helper.process_document_via_word(docx_path, export_pdf_also=True)
                    if res.get("pdf_exported") and os.path.exists(pdf_target):
                        pdf_ok = True
                except Exception:
                    pdf_ok = False

            # 2. Cơ chế cứu sinh tự động (Fail-safe Fallback): Nếu Word COM không xuất được hoặc bị treo/thiếu Office, dùng ngay Qt PDF Writer
            if not pdf_ok:
                pdf_ok = self._export_pdf_via_qt(pdf_target)

            if pdf_ok and os.path.exists(pdf_target):
                QMessageBox.information(
                    self, "Xuất PDF Thành công",
                    f"Đã xuất tệp PDF chuẩn thể thức NĐ 30 thành công:\n{pdf_target}\n\nSẵn sàng ký số điện tử hoặc in ấn."
                )
                return pdf_target
            else:
                QMessageBox.warning(self, "Lỗi xuất PDF", "Không thể tạo tệp PDF. Vui lòng kiểm tra quyền ghi tệp.")
                return None
        finally:
            if self.btn_export_pdf:
                self.btn_export_pdf.setEnabled(True)
                self.btn_export_pdf.setText("Xuất tệp PDF")
            if self.btn_sign:
                self.btn_sign.setEnabled(True)
                self.btn_sign.setText("Ký số")

    def _sign_document_now(self):
        """Xuất PDF và mở trực tiếp Module Ký số điện tử với vị trí khung ký định vị tự động"""
        pdf_path = self._export_pdf()
        if not pdf_path or not os.path.exists(pdf_path):
            return

        agency_name = self.canvas.metadata.get("issuing_agency", "")
        sign_dlg = SignatureDialog(self, pdf_path=pdf_path, default_agency=agency_name)
        # Tự động điền thông tin người ký từ khung soạn thảo nếu có
        name = self.canvas.metadata.get("signer_name", "").strip()
        pos = self.canvas.metadata.get("signer_position", "").strip()
        if name and name != "Nguyễn Văn A":
            sign_dlg.txt_signer_name.setText(name)
        if pos and pos != "CHỦ TỊCH":
            sign_dlg.txt_signer_title.setText(pos)
        sign_dlg.exec()
