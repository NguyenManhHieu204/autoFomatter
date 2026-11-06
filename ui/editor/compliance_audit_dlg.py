"""
ui/editor/compliance_audit_dlg.py: Hộp thoại Rà soát & Tương tác Thể thức Pháp lý
- Hiển thị điểm số nguyên chuẩn (ví dụ: 100 / 100 điểm)
- Giao diện đồ họa hiện đại, sang trọng, tinh tế theo phong cách phần mềm hành chính công vụ cao cấp
- Bảng checklist quy chuẩn dạng bảng thẻ gọn gàng, nhãn badge [ĐẠT] / [LỖI] thanh lịch
- Nút bấm [Sửa nhanh] (105px) và [Sửa chi tiết] (110px) căn chỉnh chuẩn xác, tuyệt đối không bị cắt chữ
- Trạng thái 100% hoàn thiện được thiết kế thẻ tinh gọn, không bị giãn nở rỗng
- Hoàn toàn không có thanh cuộn ngang, hỗ trợ cuộn dọc siêu mượt (Slim Scrollbar)
"""

import os
from datetime import datetime
from typing import Dict, Any, List, Optional
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QScrollArea, QWidget, QInputDialog, QMessageBox, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont, QColor


class ComplianceAuditDialog(QDialog):
    """Hộp thoại rà soát thể thức với chức năng sửa nhanh trực tiếp"""

    metadata_edit_requested = pyqtSignal()
    audit_refresh_requested = pyqtSignal()

    def __init__(self, parent=None, audit_result: Dict[str, Any] = None, canvas=None):
        super().__init__(parent)
        self.setWindowTitle("Kết quả Rà soát Thể thức Pháp lý")
        self.resize(880, 640)
        self.setMinimumSize(820, 560)
        self.setStyleSheet("""
            QDialog {
                background-color: #f8fafc;
                color: #0f172a;
                font-family: 'Segoe UI', -apple-system, Arial, sans-serif;
            }
            QLabel {
                color: #0f172a;
                font-size: 13px;
                background: transparent;
            }
            QPushButton {
                background-color: #ffffff;
                color: #334155;
                font-weight: 600;
                font-size: 12.5px;
                padding: 6px 16px;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #f1f5f9;
                border-color: #94a3b8;
            }
            QPushButton#btnQuickFix {
                background-color: #ecfdf5;
                color: #065f46;
                border: 1px solid #10b981;
                font-weight: 600;
                font-size: 12px;
                padding: 5px 12px;
                border-radius: 5px;
            }
            QPushButton#btnQuickFix:hover {
                background-color: #d1fae5;
                border-color: #059669;
            }
            QPushButton#btnDetailFix {
                background-color: #eff6ff;
                color: #1e40af;
                border: 1px solid #3b82f6;
                font-weight: 600;
                font-size: 12px;
                padding: 5px 12px;
                border-radius: 5px;
            }
            QPushButton#btnDetailFix:hover {
                background-color: #dbeafe;
                border-color: #2563eb;
            }
            QPushButton#btnPrimary {
                background-color: #2563eb;
                color: #ffffff;
                font-weight: 600;
                border: 1px solid #1d4ed8;
            }
            QPushButton#btnPrimary:hover {
                background-color: #1d4ed8;
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
            QScrollBar:vertical {
                width: 6px;
                background: transparent;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #cbd5e1;
                border-radius: 3px;
                min-height: 24px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94a3b8;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)

        self.audit_result = audit_result or {}
        self.canvas = canvas

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(20, 16, 20, 16)
        self.main_layout.setSpacing(12)

        # 1. Thẻ điểm số tổng quan
        self.banner_frame = QFrame()
        self.banner_frame.setObjectName("bannerFrame")
        self.banner_layout = QVBoxLayout(self.banner_frame)
        self.banner_layout.setContentsMargins(18, 14, 18, 14)
        self.banner_layout.setSpacing(4)
        self.main_layout.addWidget(self.banner_frame)

        # 2. Vùng cuộn chứa danh mục & danh sách lỗi (chỉ cuộn dọc, tắt hẳn cuộn ngang)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        self.scroll_widget = QWidget()
        self.scroll_widget.setStyleSheet("background: transparent;")
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.setContentsMargins(0, 0, 4, 0)
        self.scroll_layout.setSpacing(12)
        self.scroll_area.setWidget(self.scroll_widget)
        self.main_layout.addWidget(self.scroll_area, 1)

        # 3. Thanh nút điều khiển dưới cùng
        line_div = QFrame()
        line_div.setFrameShape(QFrame.Shape.HLine)
        line_div.setFrameShadow(QFrame.Shadow.Sunken)
        line_div.setStyleSheet("color: #e2e8f0;")
        self.main_layout.addWidget(line_div)

        btn_box = QHBoxLayout()
        btn_box.setSpacing(10)

        self.btn_auto_all = QPushButton("Tự động sửa toàn bộ")
        self.btn_auto_all.setObjectName("btnQuickFix")
        self.btn_auto_all.setToolTip("Tự động điền ngày thực tế, chuẩn hóa dấu câu và loại bỏ dữ liệu giữ chỗ mặc định")
        self.btn_auto_all.clicked.connect(self._on_auto_fix_all)
        btn_box.addWidget(self.btn_auto_all)

        btn_box.addStretch()

        self.btn_open_f2 = QPushButton("Sửa thể thức (F2)")
        self.btn_open_f2.setObjectName("btnDetailFix")
        self.btn_open_f2.clicked.connect(self._on_open_metadata_editor)
        btn_box.addWidget(self.btn_open_f2)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.clicked.connect(self.accept)
        btn_box.addWidget(self.btn_close)

        self.main_layout.addLayout(btn_box)

        # Render dữ liệu ban đầu
        self._render_audit_data()

    def update_audit_result(self, new_result: Dict[str, Any]):
        """Cập nhật dữ liệu đánh giá mới khi người dùng sửa lỗi"""
        self.audit_result = new_result
        self._render_audit_data()

    def _render_audit_data(self):
        # Điểm số làm tròn số nguyên (xx / 100 điểm, không để số thập phân xx.x)
        raw_score = self.audit_result.get("score", 0.0)
        int_score = int(round(raw_score))

        errs = self.audit_result.get("error_count", 0)
        warns = self.audit_result.get("warning_count", 0)
        is_compliant = self.audit_result.get("is_compliant", False)

        # Cập nhật banner điểm
        if is_compliant:
            border_color = "#86efac"
            bg_color = "#f0fdf4"
            score_color = "#15803d"
            status_text = "VĂN BẢN ĐẠT CHUẨN THỂ THỨC XUẤT BẢN"
        elif int_score >= 70:
            border_color = "#fde68a"
            bg_color = "#fefce8"
            score_color = "#b45309"
            status_text = "VĂN BẢN CẦN CHỈNH SỬA BỔ SUNG"
        else:
            border_color = "#fca5a5"
            bg_color = "#fef2f2"
            score_color = "#b91c1c"
            status_text = "VĂN BẢN CHƯA ĐẠT CHUẨN - CÓ XUNG ĐỘT"

        self.banner_frame.setStyleSheet(f"""
            QFrame#bannerFrame {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 8px;
            }}
        """)

        # Xóa các widget cũ trong banner_layout
        while self.banner_layout.count():
            item = self.banner_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        lbl_score = QLabel(f"ĐIỂM CHUẨN HÓA THỂ THỨC: {int_score} / 100 ({status_text})")
        lbl_score.setStyleSheet(f"font-size: 15px; font-weight: 800; color: {score_color};")
        self.banner_layout.addWidget(lbl_score)

        lbl_summary = QLabel(f"Phát hiện <b>{errs}</b> lỗi nghiêm trọng và <b>{warns}</b> cảnh báo cần hoàn thiện.")
        lbl_summary.setStyleSheet("font-size: 12.5px; color: #475569;")
        self.banner_layout.addWidget(lbl_summary)

        # Xóa các widget cũ trong scroll_layout
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 2. Danh mục quy chuẩn (Checklist Card)
        card_check = QFrame()
        card_check.setObjectName("cardCheck")
        card_check.setStyleSheet("QFrame#cardCheck { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; }")
        l_chk = QVBoxLayout(card_check)
        l_chk.setContentsMargins(18, 14, 18, 16)
        l_chk.setSpacing(8)

        lbl_chk_title = QLabel("DANH MỤC KIỂM TRA QUY CHUẨN:")
        lbl_chk_title.setStyleSheet("font-size: 12.5px; font-weight: 700; color: #1e293b; margin-bottom: 2px;")
        l_chk.addWidget(lbl_chk_title)

        for item in self.audit_result.get("checklist", []):
            row_chk = QHBoxLayout()
            row_chk.setSpacing(12)
            passed = item["passed"]
            icon_badge = QLabel("[ĐẠT]" if passed else "[LỖI]")
            badge_bg = "#dcfce7" if passed else "#fee2e2"
            badge_fg = "#15803d" if passed else "#b91c1c"
            badge_bd = "#86efac" if passed else "#fca5a5"
            icon_badge.setStyleSheet(f"""
                font-weight: 800;
                font-size: 11px;
                color: {badge_fg};
                background-color: {badge_bg};
                border: 1px solid {badge_bd};
                border-radius: 4px;
                padding: 2px 4px;
            """)
            icon_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
            icon_badge.setFixedWidth(54)
            row_chk.addWidget(icon_badge)

            lbl_name = QLabel(item["name"])
            lbl_name.setStyleSheet("font-weight: 600; font-size: 12.5px; color: #1e293b;")
            lbl_name.setFixedWidth(230)
            row_chk.addWidget(lbl_name)

            lbl_det = QLabel(item['detail'])
            lbl_det.setStyleSheet("font-size: 12px; color: #64748b;")
            lbl_det.setWordWrap(True)
            row_chk.addWidget(lbl_det, 1)

            l_chk.addLayout(row_chk)
        self.scroll_layout.addWidget(card_check)

        # 3. Danh sách lỗi chi tiết có nút bấm sửa nhanh & sửa chi tiết
        violations = self.audit_result.get("violations", [])
        if violations:
            card_viols = QFrame()
            card_viols.setObjectName("cardViols")
            card_viols.setStyleSheet("QFrame#cardViols { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; }")
            l_v = QVBoxLayout(card_viols)
            l_v.setContentsMargins(18, 14, 18, 16)
            l_v.setSpacing(10)

            lbl_v_title = QLabel("CHI TIẾT LỖI & THAO TÁC SỬA:")
            lbl_v_title.setStyleSheet("font-size: 12.5px; font-weight: 700; color: #1e293b; margin-bottom: 2px;")
            l_v.addWidget(lbl_v_title)

            for idx, v in enumerate(violations):
                item_frame = QFrame()
                item_frame.setObjectName("itemViolFrame")
                item_frame.setStyleSheet("QFrame#itemViolFrame { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; }")
                item_layout = QVBoxLayout(item_frame)
                item_layout.setContentsMargins(14, 10, 14, 10)
                item_layout.setSpacing(6)

                row_top = QHBoxLayout()
                row_top.setSpacing(12)

                is_err = (v.get("severity", "error") == "error")
                badge_text = "[LỖI]" if is_err else "[CẢNH BÁO]"
                lbl_badge = QLabel(f"<b>{idx + 1}.</b> <span style='color: {'#dc2626' if is_err else '#d97706'}; font-weight: 800;'>{badge_text}</span>")
                lbl_badge.setFixedWidth(105)
                row_top.addWidget(lbl_badge)

                lbl_msg = QLabel(v.get("message", ""))
                lbl_msg.setWordWrap(True)
                lbl_msg.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
                lbl_msg.setStyleSheet("font-size: 12.5px; font-weight: 600; color: #1e293b;")
                row_top.addWidget(lbl_msg, 1)

                # Nút Sửa nhanh (105px rộng rãi, không bị cắt chữ)
                btn_fix = QPushButton("Sửa nhanh")
                btn_fix.setObjectName("btnQuickFix")
                btn_fix.setFixedWidth(105)
                btn_fix.clicked.connect(lambda checked, viol=v: self._on_quick_fix(viol))
                row_top.addWidget(btn_fix)

                # Nút Sửa chi tiết (110px rộng rãi)
                btn_detail = QPushButton("Sửa chi tiết")
                btn_detail.setObjectName("btnDetailFix")
                btn_detail.setFixedWidth(110)
                btn_detail.clicked.connect(lambda checked, viol=v: self._on_detail_fix(viol))
                row_top.addWidget(btn_detail)

                item_layout.addLayout(row_top)

                lbl_sug = QLabel(f"<span style='color: #2563eb;'><i>→ Gợi ý: {v.get('suggestion', '')}</i></span>")
                lbl_sug.setWordWrap(True)
                lbl_sug.setStyleSheet("font-size: 11.5px; margin-left: 24px;")
                item_layout.addWidget(lbl_sug)

                l_v.addWidget(item_frame)

            self.scroll_layout.addWidget(card_viols)
        else:
            # Thẻ chúc mừng khi đạt 100% chuẩn định - thiết kế tinh gọn, đẹp mắt, không kéo giãn rỗng
            card_empty = QFrame()
            card_empty.setObjectName("cardEmpty")
            card_empty.setStyleSheet("""
                QFrame#cardEmpty {
                    background-color: #f0fdf4;
                    border: 1px solid #86efac;
                    border-radius: 8px;
                }
            """)
            l_e = QVBoxLayout(card_empty)
            l_e.setContentsMargins(18, 16, 18, 16)
            l_e.setSpacing(4)
            lbl_ok_title = QLabel("VĂN BẢN ĐẠT CHUẨN THỂ THỨC 100%")
            lbl_ok_title.setStyleSheet("font-size: 13.5px; font-weight: 800; color: #15803d;")
            lbl_ok_sub = QLabel("Không phát hiện bất kỳ lỗi thể thức nào. Văn bản đã hoàn toàn sẵn sàng để xuất tệp PDF hoặc ký số điện tử.")
            lbl_ok_sub.setStyleSheet("font-size: 12.5px; color: #166534;")
            lbl_ok_sub.setWordWrap(True)
            l_e.addWidget(lbl_ok_title)
            l_e.addWidget(lbl_ok_sub)
            self.scroll_layout.addWidget(card_empty)

        # Thêm khoảng đệm stretch ở đáy để các thẻ co giãn tự nhiên, không bị kéo phình rỗng
        self.scroll_layout.addStretch(1)

        # Ẩn nút tự động sửa toàn bộ nếu không còn lỗi
        self.btn_auto_all.setVisible(len(violations) > 0)

    # ==================== CÁC HÀNH ĐỘNG SỬA NHANH TRỰC TIẾP ====================
    def _on_quick_fix(self, viol: Dict[str, str]):
        rule_id = viol.get("rule_id", "")
        if not self.canvas:
            return

        now = datetime.now()
        loc = self.canvas.metadata.get("location", "Hà Nội") or "Hà Nội"

        if rule_id == "PLACEHOLDER_DATE":
            # Tự động thay ngày tháng thực tế hiện tại
            real_date = f"{loc}, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"
            self.canvas.metadata["date_str"] = real_date
            self.canvas._refresh_scaffolding_display()
            self.canvas.document_modified.emit()
            self.audit_refresh_requested.emit()

        elif rule_id == "PLACEHOLDER_SUBJECT":
            # Mở hộp thoại nhập trích yếu nhanh
            dt = self.canvas.metadata.get("doc_type", "QUYẾT ĐỊNH")
            def_subj = f"Về việc ban hành Kế hoạch phát triển công tác năm {now.year}"
            text, ok = QInputDialog.getText(self, "Sửa trích yếu nhanh", "Nhập trích yếu nội dung văn bản (Về việc...):", text=def_subj)
            if ok and text.strip():
                self.canvas.metadata["doc_subject"] = text.strip()
                self.canvas._refresh_scaffolding_display()
                self.canvas.document_modified.emit()
                self.audit_refresh_requested.emit()

        elif rule_id == "DOC_TYPE_STRUCTURE_MISMATCH":
            cur_type = self.canvas.metadata.get("doc_type", "QUYẾT ĐỊNH").upper()
            tpl = self.canvas.get_template_for_type(cur_type)
            if self.canvas.pages and tpl.get("body_html"):
                self.canvas.pages[0].editor.setHtml(tpl["body_html"])
            self.canvas._refresh_scaffolding_display()
            self.canvas.document_modified.emit()
            self.audit_refresh_requested.emit()

        elif rule_id in ["QD_MISSING_COMMAND", "QD_MISSING_ARTICLES"]:
            tpl = self.canvas.get_template_for_type("QUYẾT ĐỊNH")
            if self.canvas.pages and tpl.get("body_html"):
                self.canvas.pages[0].editor.setHtml(tpl["body_html"])
            self.canvas._refresh_scaffolding_display()
            self.canvas.document_modified.emit()
            self.audit_refresh_requested.emit()

        elif rule_id == "TTR_MISSING_SUBMISSION":
            if self.canvas.pages:
                html = self.canvas.pages[0].editor.toHtml()
                subm = "<p style='text-align: left; text-indent: 0px;'><b>Kính gửi:</b> Ủy ban nhân dân huyện.</p>"
                self.canvas.pages[0].editor.setHtml(subm + html)
            self.canvas.document_modified.emit()
            self.audit_refresh_requested.emit()

        elif rule_id == "PLACEHOLDER_AGENCY":
            agency_text, ok = QInputDialog.getText(
                self, "Sửa tên cơ quan ban hành",
                "Nhập tên cơ quan ban hành cụ thể (ví dụ: ỦY BAN NHÂN DÂN HUYỆN CƯ JÚT):",
                text="ỦY BAN NHÂN DÂN XÃ"
            )
            if ok and agency_text.strip():
                self.canvas.metadata["issuing_agency"] = agency_text.strip().upper()
                self.canvas._refresh_scaffolding_display()
                self.canvas.document_modified.emit()
                self.audit_refresh_requested.emit()

        elif rule_id == "PLACEHOLDER_DOC_NUMBER":
            num_text, ok = QInputDialog.getText(
                self, "Sửa số hiệu văn bản",
                "Nhập số và ký hiệu văn bản:",
                text="Số: 15/QĐ-UBND"
            )
            if ok and num_text.strip():
                self.canvas.metadata["doc_number"] = num_text.strip()
                self.canvas._refresh_scaffolding_display()
                self.canvas.document_modified.emit()
                self.audit_refresh_requested.emit()

        elif rule_id == "SAMPLE_SIGNER_NAME":
            name_text, ok = QInputDialog.getText(
                self, "Sửa họ tên người ký",
                "Nhập họ và tên người ký ban hành văn bản:",
                text="Nguyễn Văn An"
            )
            if ok and name_text.strip():
                self.canvas.metadata["signer_name"] = name_text.strip()
                self.canvas._refresh_scaffolding_display()
                self.canvas.document_modified.emit()
                self.audit_refresh_requested.emit()

        elif rule_id in ["LEGAL_GROUNDS_SYNTAX", "PROPOSAL_SYNTAX"]:
            # Tự động căn chỉnh và chuẩn hóa dấu câu
            self.canvas.apply_nd30_standards_and_alignment()
            self.audit_refresh_requested.emit()

        else:
            self._on_detail_fix(viol)

    def _on_detail_fix(self, viol: Dict[str, str]):
        """Mở hộp thoại chỉnh sửa thể thức chi tiết (F2)"""
        self._on_open_metadata_editor()

    def _on_auto_fix_all(self):
        """Tự động sửa nhanh toàn bộ các lỗi mẫu: ngày tháng, dấu câu, số hiệu, trích yếu rác"""
        if not self.canvas:
            return

        now = datetime.now()
        loc = self.canvas.metadata.get("location", "Hà Nội") or "Hà Nội"

        # 1. Sửa ngày tháng nếu có ...
        cur_date = self.canvas.metadata.get("date_str", "")
        if "..." in cur_date or not cur_date:
            self.canvas.metadata["date_str"] = f"{loc}, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"

        # 2. Sửa trích yếu nếu là 0 hoặc Về việc 0
        cur_subj = self.canvas.metadata.get("doc_subject", "")
        if cur_subj in ["0", "Về việc 0"] or not cur_subj:
            self.canvas.metadata["doc_subject"] = f"Về việc ban hành Kế hoạch phát triển công tác năm {now.year}"

        # 3. Sửa số hiệu nếu là QĐ-UBND không có số hoặc có ...
        cur_num = self.canvas.metadata.get("doc_number", "")
        if "..." in cur_num or cur_num == "QĐ-UBND" or not cur_num:
            std = self.canvas.metadata.get("standard_id", "ND30")
            self.canvas.metadata["doc_number"] = "Số 01-QĐ/ĐU" if std == "HD36" else "Số: 15/QĐ-UBND"

        # 3b. Sửa họ tên người ký nếu còn để tên mẫu mặc định
        cur_name = self.canvas.metadata.get("signer_name", "")
        if cur_name in ["Nguyễn Văn A", "Nguyễn Văn ...", ""]:
            self.canvas.metadata["signer_name"] = "Nguyễn Văn An"

        # 3c. Sửa cơ quan ban hành nếu còn là mặc định CƠ QUAN BAN HÀNH
        cur_iss = self.canvas.metadata.get("issuing_agency", "")
        if "CƠ QUAN BAN HÀNH" in cur_iss.upper() or not cur_iss:
            std = self.canvas.metadata.get("standard_id", "ND30")
            self.canvas.metadata["issuing_agency"] = "ĐẢNG ỦY XÃ" if std == "HD36" else "ỦY BAN NHÂN DÂN XÃ"

        # 4. Sửa xung đột cấu trúc nội dung với loại văn bản hiện tại
        cur_type = self.canvas.metadata.get("doc_type", "QUYẾT ĐỊNH").upper()
        cur_body = self.canvas.pages[0].editor.toPlainText() if self.canvas.pages else ""
        has_qd_cmd = ("QUYẾT ĐỊNH:" in cur_body or "QUYẾT NGHỊ:" in cur_body)

        if cur_type in ["TỜ TRÌNH", "BÁO CÁO", "CÔNG VĂN", "THÔNG BÁO"] and has_qd_cmd:
            tpl = self.canvas.get_template_for_type(cur_type)
            if self.canvas.pages and tpl.get("body_html"):
                self.canvas.pages[0].editor.setHtml(tpl["body_html"])
        elif cur_type == "QUYẾT ĐỊNH" and not has_qd_cmd:
            tpl = self.canvas.get_template_for_type("QUYẾT ĐỊNH")
            if self.canvas.pages and tpl.get("body_html"):
                self.canvas.pages[0].editor.setHtml(tpl["body_html"])

        # 5. Căn chỉnh lề và dấu câu
        self.canvas.apply_nd30_standards_and_alignment()

        self.canvas._refresh_scaffolding_display()
        self.canvas.document_modified.emit()
        self.audit_refresh_requested.emit()

    def _on_open_metadata_editor(self):
        """Mở hộp thoại chỉnh sửa thể thức F2"""
        if self.canvas:
            self.canvas.open_metadata_editor()
            self.audit_refresh_requested.emit()
