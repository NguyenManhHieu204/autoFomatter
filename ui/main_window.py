"""
ui/main_window.py: Giao diện chính của ứng dụng Auto-Formatter (PyQt6)
Cung cấp trải nghiệm kéo-thả trực quan, xử lý bất đồng bộ (QThread),
thống kê báo cáo theo thẻ và bảng cảnh báo chi tiết.
"""

import os
import sys
import json
from typing import List, Dict, Any, Optional

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QRadioButton, QButtonGroup, QComboBox, QCheckBox,
    QProgressBar, QGroupBox, QFrame, QMessageBox, QSplitter,
    QAbstractScrollArea, QSizePolicy, QLineEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMimeData, QSize, QTimer
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QFont, QColor, QPixmap, QAction

from core.formatter_engine import FormatterEngine, FormattingReport
from core.win32_helper import Win32WordHelper
from core.path_utils import get_config_dir, get_app_icon_path
from ui.styles import MAIN_STYLE
from ui.quick_wizard_dialog import QuickWizardDialog
from ui.editor.word_editor_window import WordEditorWindow


class ProcessingWorker(QThread):
    progress_updated = pyqtSignal(int, str)
    file_started = pyqtSignal(int, str)
    file_finished = pyqtSignal(int, FormattingReport, dict)
    all_finished = pyqtSignal(bool)

    def __init__(
        self,
        files: List[str],
        standard_id: str,
        profile_data: Optional[Dict[str, Any]],
        export_pdf: bool,
        config_dir: str,
        auto_ocr_scans: bool = True,
        output_dir: Optional[str] = None
    ):
        super().__init__()
        self.files = files
        self.standard_id = standard_id
        self.profile_data = profile_data
        self.export_pdf = export_pdf
        self.config_dir = config_dir
        self.auto_ocr_scans = auto_ocr_scans
        self.output_dir = output_dir
        self.engine = FormatterEngine(config_dir=config_dir)
        self.win32_helper = Win32WordHelper()
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        total = len(self.files)
        # Tối ưu hóa hiệu năng: Chỉ khởi tạo tiến trình Word nếu người dùng thực sự chọn xuất PDF
        needs_word = self.export_pdf and any(not f.lower().endswith(".pdf") for f in self.files)
        word_session_ctx = self.win32_helper.word_session() if needs_word else None

        try:
            word_app = word_session_ctx.__enter__() if word_session_ctx else None
            for idx, file_path in enumerate(self.files):
                if self._is_cancelled:
                    break

                pct = int((idx / total) * 100)
                base_name = os.path.basename(file_path)
                self.file_started.emit(idx, base_name)

                # Nếu là file PDF, tự động nhận dạng OCR để bóc tách thành văn bản
                if file_path.lower().endswith(".pdf"):
                    self.progress_updated.emit(pct, f"Đang xử lý tệp PDF: {base_name} ({idx + 1}/{total})...")
                    report = self.engine.format_scanned_document_with_ocr(
                        input_file=file_path,
                        standard_id=self.standard_id,
                        profile_override=self.profile_data,
                        progress_callback=lambda p, msg: self.progress_updated.emit(pct + int(p * 0.5 / total), msg),
                        output_dir=self.output_dir
                    )
                else:
                    self.progress_updated.emit(pct, f"Đang chuẩn hóa: {base_name} ({idx + 1}/{total})...")
                    # 1. Chạy engine định dạng chính siêu tốc (OpenXML thuần túy, < 0.05s)
                    report = self.engine.format_document(
                        input_docx_path=file_path,
                        standard_id=self.standard_id,
                        profile_override=self.profile_data,
                        output_dir=self.output_dir
                    )

                    # Tự động chuyển giao sang OCR nếu phát hiện ảnh scan và người dùng bật auto_ocr_scans
                    if report.is_scanned_image and self.auto_ocr_scans and not self._is_cancelled:
                        self.progress_updated.emit(pct, f"Tự động chạy OCR bóc tách chữ: {base_name}...")
                        report = self.engine.format_scanned_document_with_ocr(
                            input_file=file_path,
                            standard_id=self.standard_id,
                            profile_override=self.profile_data,
                            progress_callback=lambda p, msg: self.progress_updated.emit(pct + int(p * 0.5 / total), msg),
                            output_dir=self.output_dir
                        )

                if self._is_cancelled:
                    break

                # 2. Xử lý xuất PDF qua MS Word COM CHỈ KHI người dùng chọn xuất PDF
                com_res = {"success": False, "pdf_exported": False, "pdf_path": ""}
                if report.success and self.export_pdf and not file_path.lower().endswith(".pdf"):
                    com_res = self.win32_helper.process_document_via_word(
                        docx_path=report.output_file,
                        export_pdf_also=True,
                        standard_id=self.standard_id,
                        word_app=word_app
                    )

                self.file_finished.emit(idx, report, com_res)

        finally:
            if word_session_ctx:
                try:
                    word_session_ctx.__exit__(None, None, None)
                except Exception:
                    pass

        if self._is_cancelled:
            self.progress_updated.emit(0, "Đã dừng xử lý theo yêu cầu.")
        else:
            self.progress_updated.emit(100, "Hoàn tất xử lý tất cả các tài liệu!")
        self.all_finished.emit(self._is_cancelled)


class OcrWorker(QThread):
    progress_updated = pyqtSignal(int, str)
    ocr_finished = pyqtSignal(FormattingReport, dict)

    def __init__(
        self,
        file_path: str,
        standard_id: str,
        profile_data: Optional[Dict[str, Any]],
        export_pdf: bool,
        config_dir: str,
        output_dir: Optional[str] = None
    ):
        super().__init__()
        self.file_path = file_path
        self.standard_id = standard_id
        self.profile_data = profile_data
        self.export_pdf = export_pdf
        self.config_dir = config_dir
        self.output_dir = output_dir
        self.engine = FormatterEngine(config_dir=config_dir)
        self.win32_helper = Win32WordHelper()
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        base_name = os.path.basename(self.file_path)
        self.progress_updated.emit(10, f"Đang bóc tách ảnh và chạy OCR ngoại tuyến: {base_name}...")

        report = self.engine.format_scanned_document_with_ocr(
            input_file=self.file_path,
            standard_id=self.standard_id,
            profile_override=self.profile_data,
            progress_callback=lambda pct, msg: self.progress_updated.emit(pct, msg),
            output_dir=self.output_dir
        )

        com_res = {"success": False, "pdf_exported": False, "pdf_path": ""}
        if report.success and not self._is_cancelled:
            com_res = self.win32_helper.process_document_via_word(
                docx_path=report.output_file,
                export_pdf_also=self.export_pdf,
                standard_id=self.standard_id
            )

        self.progress_updated.emit(100, "Hoàn tất nhận dạng OCR và chuẩn hóa!")
        self.ocr_finished.emit(report, com_res)


class DropZoneWidget(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(6)

        self.lbl_badge = QLabel("TIẾP NHẬN DỮ LIỆU ĐẦU VÀO")
        self.lbl_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_badge.setStyleSheet("""
            color: #2563eb; font-size: 11px; font-weight: 700;
            background-color: #eff6ff; border: 1px solid #bfdbfe;
            border-radius: 10px; padding: 3px 12px; margin-bottom: 2px;
        """)

        self.lbl_text = QLabel("KÉO VÀ THẢ TỆP DỰ THẢO HOẶC THƯ MỤC VÀO ĐÂY")
        self.lbl_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_text.setStyleSheet("color: #1e293b; font-weight: 700; font-size: 13px;")

        self.lbl_sub = QLabel("Hỗ trợ tệp văn bản .docx và tệp scan .pdf (Tự động quét các tệp bên trong thư mục)")
        self.lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sub.setStyleSheet("color: #64748b; font-size: 11px;")

        layout.addWidget(self.lbl_badge, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_text)
        layout.addWidget(self.lbl_sub)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        file_paths = []
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isdir(p):
                # Tối ưu: Kéo thả cả thư mục -> tự động quét đệ quy các tệp bên trong
                for root, _, files in os.walk(p):
                    for f in files:
                        if f.lower().endswith((".docx", ".pdf")):
                            file_paths.append(os.path.join(root, f))
            elif p.lower().endswith((".docx", ".pdf")):
                file_paths.append(p)

        if file_paths:
            self.files_dropped.emit(file_paths)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Auto-Formatter — Chuẩn hóa Thể thức Văn bản")
        self.resize(1200, 800)
        self.setMinimumSize(980, 640)
        self.setStyleSheet(MAIN_STYLE)

        icon_path = get_app_icon_path()
        if icon_path and os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.config_dir = get_config_dir()
        self.files_queue: List[str] = []
        self.reports_map: Dict[int, FormattingReport] = {}
        self.com_results_map: Dict[int, dict] = {}
        self.profiles_data: Dict[str, Any] = {}
        self.current_report: Optional[FormattingReport] = None
        self.active_issue_filter: str = "all"
        self.stat_cards: List[QFrame] = []

        self._load_profiles_data()
        self._init_ui()
        self._refresh_profiles_dropdown()
        QTimer.singleShot(3500, self._check_for_updates_silent)

    def _load_profiles_data(self):
        p_path = os.path.join(self.config_dir, "agency_profiles.json")
        if os.path.exists(p_path):
            try:
                with open(p_path, "r", encoding="utf-8") as f:
                    self.profiles_data = json.load(f)
            except Exception:
                self.profiles_data = {"profiles": []}
        else:
            self.profiles_data = {"profiles": []}

    def _get_selected_standard(self) -> str:
        """Lấy chuẩn thể thức đang chọn: 'HD36' (Đảng) hoặc 'ND30' (Nhà nước)"""
        return "HD36" if hasattr(self, "rb_hd36") and self.rb_hd36.isChecked() else "ND30"

    def _setup_menu_bar(self):
        """Khởi tạo thanh Menu Bar tiện ích ở đầu ứng dụng"""
        menubar = self.menuBar()
        menubar.setStyleSheet("""
            QMenuBar {
                background-color: #ffffff;
                color: #0f172a;
                font-size: 12px;
                padding: 2px 8px;
                border-bottom: 1px solid #e2e8f0;
            }
            QMenuBar::item {
                background-color: transparent;
                padding: 4px 10px;
                border-radius: 4px;
            }
            QMenuBar::item:selected {
                background-color: #f1f5f9;
                color: #2563eb;
            }
            QMenu {
                background-color: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 24px 6px 10px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #f1f5f9;
                color: #2563eb;
            }
            QMenu::separator {
                height: 1px;
                background-color: #e2e8f0;
                margin: 4px 6px;
            }
        """)

        # 1. Menu Tệp
        menu_file = menubar.addMenu("Tệp")
        act_new = menu_file.addAction("Soạn thảo văn bản mới...")
        act_new.setShortcut("Ctrl+N")
        act_new.triggered.connect(self._open_quick_wizard)

        act_add_file = menu_file.addAction("Thêm tệp dự thảo...")
        act_add_file.setShortcut("Ctrl+O")
        act_add_file.triggered.connect(self._choose_files_dialog)

        act_add_folder = menu_file.addAction("Thêm thư mục...")
        act_add_folder.triggered.connect(self._choose_folder_dialog)

        menu_file.addSeparator()
        act_exit = menu_file.addAction("Thoát ứng dụng")
        act_exit.setShortcut("Alt+F4")
        act_exit.triggered.connect(self.close)

        # 2. Menu Soạn thảo
        menu_edit = menubar.addMenu("Soạn thảo")
        act_wiz = menu_edit.addAction("Khởi tạo nhanh từ mẫu chuẩn (Template Wizard)...")
        act_wiz.setShortcut("Ctrl+Shift+N")
        act_wiz.triggered.connect(self._open_quick_wizard)

        act_emb = menu_edit.addAction("Trình soạn thảo văn bản tích hợp...")
        act_emb.setShortcut("F4")
        act_emb.triggered.connect(self._open_in_embedded_editor)

        # 3. Menu Công cụ
        menu_tools = menubar.addMenu("Công cụ")
        act_sign = menu_tools.addAction("Ký số...")
        act_sign.setShortcut("Ctrl+K")
        act_sign.triggered.connect(self._open_signature_dialog)

        act_prof = menu_tools.addAction("Quản lý Hồ sơ Cơ quan & Đơn vị...")
        act_prof.triggered.connect(self._open_profile_dialog)

        act_batch = menu_tools.addAction("Xử lý hàng loạt văn bản (Nội dung & Đổi tên)...")
        act_batch.setShortcut("Ctrl+B")
        act_batch.triggered.connect(self._open_batch_dialog)

        act_ocr = menu_tools.addAction("Nhận dạng lại bảng biểu (Table OCR)...")
        act_ocr.triggered.connect(self._run_ocr_on_selected_file)

        # 4. Menu Trợ giúp
        menu_help = menubar.addMenu("Trợ giúp")
        act_about = menu_help.addAction("Thông tin phiên bản Auto-Formatter")
        act_about.triggered.connect(self._show_about_dialog)

        act_update = menu_help.addAction("Kiểm tra bản cập nhật...")
        act_update.triggered.connect(self._check_for_updates_manual)

    def _show_about_dialog(self):
        """Hiển thị hộp thoại giới thiệu phiên bản 2.0"""
        QMessageBox.about(
            self,
            "Về Auto-Formatter v2.0",
            "<h3>Auto-Formatter v2.0 — Soạn Thảo & Chuẩn Hóa Thể Thức Văn Bản</h3>"
            "<p><b>Giải pháp chuẩn hóa thể thức và ký số văn bản điện tử chuyên nghiệp:</b></p>"
            "<ul>"
            "<li><b>Khóa cứng Thể thức (Layout Guard):</b> Tự động bảo vệ Quốc hiệu, Tiêu ngữ, Tên cơ quan, Khối ký tên theo NĐ 30/2020 và HD 36-HD/VPTW.</li>"
            "<li><b>Thước căn lề milimét & Khung vẽ A4 rời rạc:</b> Lề chuẩn Trái 30mm, Phải 15mm, Trên 20mm, Dưới 20mm, Thụt dòng 1.27cm, hỗ trợ độc lập xoay ngang từ trang 2+.</li>"
            "<li><b>Dải băng điều khiển Ribbon Bar:</b> Thiết kế Typography thanh lịch, chuẩn mực công vụ văn phòng.</li>"
            "<li><b>Bảng biểu hành chính:</b> Lặp tiêu đề (tblHeader) và chống gãy trang ngang dòng (cantSplit).</li>"
            "<li><b>Ký số điện tử PAdES CAPI:</b> Tương thích 100% USB Token Ban Cơ yếu Chính phủ và các CA công cộng theo Điều 13 NĐ 30.</li>"
            "<li><b>Rà soát thể thức thời gian thực:</b> Chấm điểm 0 - 100 điểm kèm nút bấm sửa nhanh 1 chạm.</li>"
            "</ul>"
            "<p><i>Bản quyền © 2026 Auto-Formatter Team. Hoạt động ngoại tuyến 100% (Air-gapped) bảo mật dữ liệu tuyệt đối.</i></p>"
        )

    def _check_for_updates_manual(self):
        """Kiểm tra bản cập nhật mới từ GitHub theo yêu cầu người dùng."""
        from core.updater import CheckUpdateThread, APP_VERSION
        from ui.update_dialog import UpdateDialog

        self._update_checker_thread = CheckUpdateThread()

        def on_result(info: dict):
            if not info.get("success"):
                QMessageBox.warning(
                    self,
                    "Kiểm tra cập nhật",
                    info.get("error", "Không thể kết nối đến máy chủ GitHub để kiểm tra cập nhật.")
                )
                return

            if info.get("has_update"):
                dlg = UpdateDialog(self, update_info=info)
                dlg.exec()
            else:
                QMessageBox.information(
                    self,
                    "Kiểm tra cập nhật",
                    f"Bạn đang sử dụng phiên bản mới nhất: v{APP_VERSION}."
                )

        self._update_checker_thread.result_ready.connect(on_result)
        self._update_checker_thread.start()

    def _check_for_updates_silent(self):
        """Kiểm tra cập nhật chạy ngầm khi mở ứng dụng (chỉ hiện thông báo khi có bản mới)."""
        from core.updater import CheckUpdateThread
        from ui.update_dialog import UpdateDialog

        self._silent_update_thread = CheckUpdateThread()

        def on_silent_result(info: dict):
            if info.get("success") and info.get("has_update"):
                dlg = UpdateDialog(self, update_info=info)
                dlg.exec()

        self._silent_update_thread.result_ready.connect(on_silent_result)
        self._silent_update_thread.start()

    def _init_ui(self):
        self._setup_menu_bar()

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(14, 10, 14, 14)
        main_layout.setSpacing(10)

        # Splitter chia đôi màn hình: Trái (Đầu vào & Cấu hình) - Phải (Kết quả & Cảnh báo)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ====================== CỘT TRÁI ======================
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(10)

        # Drop Zone
        self.drop_zone = DropZoneWidget()
        self.drop_zone.files_dropped.connect(self._add_files)
        left_layout.addWidget(self.drop_zone)

        # Thanh công cụ quản lý tệp tinh gọn (1 hàng chuyên nghiệp, không lặp lại)
        file_toolbar = QHBoxLayout()
        file_toolbar.setSpacing(6)

        self.btn_add_files = QPushButton("Thêm tệp...")
        self.btn_add_files.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: #ffffff; font-weight: 600;
                font-size: 12px; padding: 6px 14px; border-radius: 5px; border: 1px solid #1d4ed8;
            }
            QPushButton:hover { background-color: #1d4ed8; }
            QPushButton:pressed { background-color: #1e40af; }
        """)
        self.btn_add_files.setToolTip("Chọn một hoặc nhiều tệp Word (.docx) hoặc PDF để chuẩn hóa")
        self.btn_add_files.clicked.connect(self._choose_files_dialog)

        self.btn_add_folder = QPushButton("Thêm thư mục...")
        self.btn_add_folder.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #334155; font-weight: 500;
                font-size: 12px; padding: 6px 12px; border-radius: 5px; border: 1px solid #cbd5e1;
            }
            QPushButton:hover { background-color: #f8fafc; border-color: #94a3b8; color: #0f172a; }
            QPushButton:pressed { background-color: #f1f5f9; }
        """)
        self.btn_add_folder.setToolTip("Quét và nạp toàn bộ tệp tài liệu trong một thư mục")
        self.btn_add_folder.clicked.connect(self._choose_folder_dialog)

        file_toolbar.addWidget(self.btn_add_files)
        file_toolbar.addWidget(self.btn_add_folder)

        self.btn_batch = QPushButton("Xử lý hàng loạt (Nội dung & Tên tệp)...")
        self.btn_batch.setStyleSheet("""
            QPushButton {
                background-color: #f8fafc; color: #1e3a8a; font-weight: 600;
                font-size: 12px; padding: 6px 14px; border-radius: 5px; border: 1px solid #cbd5e1;
            }
            QPushButton:hover { background-color: #eff6ff; border-color: #93c5fd; color: #1d4ed8; }
            QPushButton:pressed { background-color: #dbeafe; }
        """)
        self.btn_batch.setToolTip("Chỉnh sửa nội dung theo chuẩn NĐ 30 và đổi tên tệp hàng loạt đồng thời")
        self.btn_batch.clicked.connect(self._open_batch_dialog)
        file_toolbar.addWidget(self.btn_batch)

        file_toolbar.addStretch(1)

        self.btn_remove_file = QPushButton("Xóa")
        self.btn_remove_file.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #64748b; font-weight: 500;
                font-size: 12px; padding: 6px 12px; border-radius: 5px; border: 1px solid #e2e8f0;
            }
            QPushButton:hover { background-color: #fef2f2; border-color: #fca5a5; color: #dc2626; }
            QPushButton:pressed { background-color: #fee2e2; }
        """)
        self.btn_remove_file.setToolTip("Xóa tệp đang chọn khỏi danh sách")
        self.btn_remove_file.clicked.connect(self._remove_selected_file)

        self.btn_clear_all = QPushButton("Làm mới")
        self.btn_clear_all.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #64748b; font-weight: 500;
                font-size: 12px; padding: 6px 12px; border-radius: 5px; border: 1px solid #e2e8f0;
            }
            QPushButton:hover { background-color: #f1f5f9; border-color: #cbd5e1; color: #1e293b; }
            QPushButton:pressed { background-color: #e2e8f0; }
        """)
        self.btn_clear_all.setToolTip("Xóa toàn bộ tệp trong danh sách để làm mới")
        self.btn_clear_all.clicked.connect(self._clear_all_files)

        file_toolbar.addWidget(self.btn_remove_file)
        file_toolbar.addWidget(self.btn_clear_all)

        # Giữ tham chiếu btn_quick_create để tránh lỗi nếu có module gọi
        self.btn_quick_create = None

        left_layout.addLayout(file_toolbar)

        # Bảng danh sách file
        self.table_files = QTableWidget(0, 3)
        self.table_files.setHorizontalHeaderLabels(["Tên tệp dự thảo", "Kích thước", "Trạng thái"])
        self.table_files.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.table_files.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_files.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        self.table_files.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table_files.setColumnWidth(1, 85)
        self.table_files.setColumnWidth(2, 115)
        self.table_files.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_files.cellClicked.connect(self._on_file_selected)
        left_layout.addWidget(self.table_files, 1)

        # Khung Cấu hình
        config_group = QGroupBox("Cấu hình Thể thức & Cơ quan Ban hành")
        cfg_layout = QVBoxLayout(config_group)
        cfg_layout.setSpacing(8)

        # Lựa chọn quy chuẩn (Bố cục rộng rãi, không bị cắt chữ)
        rule_box = QVBoxLayout()
        rule_box.setSpacing(4)
        lbl_rule = QLabel("Quy chuẩn thể thức áp dụng:")
        lbl_rule.setStyleSheet("font-weight: 600; color: #1e293b; font-size: 12px;")

        rb_row = QHBoxLayout()
        rb_row.setSpacing(14)
        self.rb_nd30 = QRadioButton("Hành chính Nhà nước (NĐ 30)")
        self.rb_hd36 = QRadioButton("Cơ quan Đảng (HD 36)")
        self.rb_nd30.setChecked(True)
        self.rb_nd30.setStyleSheet("font-size: 12px; font-weight: 500; color: #1e293b;")
        self.rb_hd36.setStyleSheet("font-size: 12px; font-weight: 500; color: #1e293b;")
        self.rb_nd30.setToolTip("Quy chuẩn thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP")
        self.rb_hd36.setToolTip("Quy chuẩn thể thức văn bản Đảng theo Hướng dẫn 36-HD/VPTW")
        self.rb_nd30.toggled.connect(self._on_standard_changed)
        rb_row.addWidget(self.rb_nd30)
        rb_row.addWidget(self.rb_hd36)
        rb_row.addStretch()

        rule_box.addWidget(lbl_rule)
        rule_box.addLayout(rb_row)
        cfg_layout.addLayout(rule_box)

        # Hồ sơ đơn vị (Tiêu đề + nút Quản lý hàng trên, Hộp chọn full hàng dưới)
        prof_box = QVBoxLayout()
        prof_box.setSpacing(4)

        prof_header = QHBoxLayout()
        lbl_prof = QLabel("Hồ sơ cơ quan ban hành:")
        lbl_prof.setStyleSheet("font-weight: 600; color: #1e293b; font-size: 12px;")

        self.btn_manage_prof = QPushButton("Quản lý hồ sơ...")
        self.btn_manage_prof.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_manage_prof.setStyleSheet("""
            QPushButton {
                background-color: #f1f5f9; color: #0284c7; font-weight: 600;
                font-size: 11.5px; padding: 4px 10px; border-radius: 4px; border: 1px solid #cbd5e1;
            }
            QPushButton:hover { background-color: #e2e8f0; color: #0369a1; }
        """)
        self.btn_manage_prof.clicked.connect(self._open_profile_dialog)
        prof_header.addWidget(lbl_prof)
        prof_header.addStretch()
        prof_header.addWidget(self.btn_manage_prof)
        prof_box.addLayout(prof_header)

        self.cb_profiles = QComboBox()
        self.cb_profiles.setMinimumHeight(30)
        self.cb_profiles.currentIndexChanged.connect(self._on_profile_changed)
        prof_box.addWidget(self.cb_profiles)
        cfg_layout.addLayout(prof_box)

        # Tùy chọn nâng cao
        opts_layout = QVBoxLayout()
        opts_row1 = QHBoxLayout()
        self.chk_export_pdf = QCheckBox("Tự động xuất tệp PDF khi xử lý")
        self.chk_export_pdf.setChecked(False)

        self.chk_open_folder = QCheckBox("Mở thư mục khi hoàn tất")
        self.chk_open_folder.setChecked(True)

        opts_row1.addWidget(self.chk_export_pdf)
        opts_row1.addWidget(self.chk_open_folder)
        opts_row1.addStretch()

        opts_row2 = QHBoxLayout()
        self.chk_auto_ocr = QCheckBox("Tự động nhận dạng OCR khi phát hiện tệp ảnh scan")
        self.chk_auto_ocr.setChecked(True)
        self.chk_auto_ocr.setToolTip("Khi phát hiện tệp scan dạng ảnh, tự động bóc tách thành văn bản chữ chuẩn hóa")
        opts_row2.addWidget(self.chk_auto_ocr)
        opts_row2.addStretch()

        opts_layout.addLayout(opts_row1)
        opts_layout.addLayout(opts_row2)

        # Tùy chọn nơi lưu tệp xuất
        opts_out = QVBoxLayout()
        opts_out.setSpacing(4)
        lbl_out_hdr = QLabel("Nơi lưu tệp xuất ra:")
        lbl_out_hdr.setStyleSheet("font-weight: 600; color: #1e293b; font-size: 11.5px;")
        opts_out.addWidget(lbl_out_hdr)

        row_out = QHBoxLayout()
        self.txt_output_dir = QLineEdit()
        self.txt_output_dir.setPlaceholderText("Mặc định: Cùng thư mục với tệp gốc")
        self.txt_output_dir.setReadOnly(True)
        self.txt_output_dir.setStyleSheet("""
            QLineEdit {
                background-color: #f8fafc;
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 11.5px;
                color: #334155;
            }
        """)
        self.btn_browse_output = QPushButton("Chọn...")
        self.btn_browse_output.setStyleSheet("padding: 4px 10px; font-size: 11.5px;")
        self.btn_browse_output.setToolTip("Chọn thư mục lưu tệp Word/PDF xuất ra")
        self.btn_browse_output.clicked.connect(self._choose_output_directory)

        self.btn_clear_output = QPushButton("Mặc định")
        self.btn_clear_output.setStyleSheet("padding: 4px 8px; font-size: 11.5px; color: #64748b;")
        self.btn_clear_output.setToolTip("Xóa thư mục riêng để lưu cùng thư mục với từng tệp gốc")
        self.btn_clear_output.clicked.connect(self._reset_output_directory)

        row_out.addWidget(self.txt_output_dir)
        row_out.addWidget(self.btn_browse_output)
        row_out.addWidget(self.btn_clear_output)
        opts_out.addLayout(row_out)
        opts_layout.addLayout(opts_out)

        cfg_layout.addLayout(opts_layout)

        left_layout.addWidget(config_group)

        # Nút Chạy và Nút Hủy (Hiện đại, tương phản cao, nổi bật, không dùng icon)
        run_btns_layout = QHBoxLayout()
        self.btn_start = QPushButton("BẮT ĐẦU CHUẨN HÓA THỂ THỨC")
        self.btn_start.setObjectName("btnRun")
        self.btn_start.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #10b981, stop:1 #059669);
                color: #ffffff;
                font-weight: bold;
                font-size: 13.5px;
                padding: 10px 18px;
                border-radius: 6px;
                border: 1px solid #047857;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #34d399, stop:1 #059669);
            }
            QPushButton:pressed {
                background-color: #047857;
            }
            QPushButton:disabled {
                background-color: #e2e8f0;
                color: #94a3b8;
                border: 1px solid #cbd5e1;
            }
        """)
        self.btn_start.clicked.connect(self._start_processing)

        self.btn_cancel = QPushButton("Dừng lại")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #fee2e2; color: #dc2626; font-weight: bold;
                font-size: 12.5px; padding: 10px 14px; border: 1px solid #f87171; border-radius: 6px;
            }
            QPushButton:hover { background-color: #fecaca; }
            QPushButton:disabled { background-color: #f8fafc; color: #cbd5e1; border: 1px solid #e2e8f0; }
        """)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_processing)

        run_btns_layout.addWidget(self.btn_start, 3)
        run_btns_layout.addWidget(self.btn_cancel, 1)
        left_layout.addLayout(run_btns_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        left_layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("Sẵn sàng xử lý tài liệu.")
        self.lbl_status.setStyleSheet("color: #64748b; font-size: 12px;")
        left_layout.addWidget(self.lbl_status)

        # ====================== CỘT PHẢI ======================
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(10)

        # 4 Thẻ thống kê với biểu tượng nhận diện chuyên biệt và hỗ trợ lọc danh mục
        stats_layout = QHBoxLayout()

        self.card_paragraphs = self._create_stat_card("Đoạn văn chuẩn hóa", "0", "all", "#2563eb", "#eff6ff")
        self.card_spaces = self._create_stat_card("Khoảng trắng dọn", "0", "whitespace", "#0284c7", "#ecfeff")
        self.card_punct = self._create_stat_card("Chính tả & Dấu câu", "0", "spelling_punct", "#059669", "#f0fdf4")
        self.card_warnings = self._create_stat_card("Cảnh báo viết hoa", "0", "capitalization", "#d97706", "#fffbeb")

        self.stat_cards = [self.card_paragraphs, self.card_spaces, self.card_punct, self.card_warnings]

        stats_layout.addWidget(self.card_paragraphs)
        stats_layout.addWidget(self.card_spaces)
        stats_layout.addWidget(self.card_punct)
        stats_layout.addWidget(self.card_warnings)
        right_layout.addLayout(stats_layout)

        # Bảng cảnh báo & Linter
        self.report_group = QGroupBox("Chi tiết lỗi")
        rep_layout = QVBoxLayout(self.report_group)

        self.table_issues = QTableWidget(0, 4)
        self.table_issues.setHorizontalHeaderLabels(["Loại lỗi", "Nội dung gốc", "Chuẩn hóa / Gợi ý", "Quy định & Giải thích"])
        self.table_issues.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.table_issues.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table_issues.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table_issues.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table_issues.setWordWrap(True)

        header = self.table_issues.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        self.table_issues.setColumnWidth(0, 130)
        self.table_issues.setColumnWidth(1, 140)
        self.table_issues.setColumnWidth(2, 140)
        self.table_issues.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        rep_layout.addWidget(self.table_issues)

        right_layout.addWidget(self.report_group, 1)

        # Thanh nút thao tác trên văn bản (Thiết kế hài hòa, màu sắc văn phòng tinh tế, phân cấp rõ ràng)
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(6)

        # 1. Nút Soạn thảo tích hợp (Hành động chính - Royal Blue)
        self.btn_embedded_edit = QPushButton("Soạn thảo tích hợp")
        self.btn_embedded_edit.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: #ffffff; font-weight: 600;
                font-size: 12px; padding: 7px 14px; border-radius: 5px; border: 1px solid #1d4ed8;
            }
            QPushButton:hover { background-color: #1d4ed8; }
            QPushButton:pressed { background-color: #1e40af; }
            QPushButton:disabled { background-color: #f1f5f9; color: #94a3b8; border-color: #e2e8f0; }
        """)
        self.btn_embedded_edit.setEnabled(True)
        self.btn_embedded_edit.setToolTip("Mở văn bản trong Trình Soạn thảo Hành chính tích hợp (hoặc tạo mới)")
        self.btn_embedded_edit.clicked.connect(self._open_in_embedded_editor)

        # 2. Nút Ký số (Hành động ký số - Elegant Indigo)
        self.btn_sign_pdf = QPushButton("Ký số")
        self.btn_sign_pdf.setStyleSheet("""
            QPushButton {
                background-color: #4f46e5; color: #ffffff; font-weight: 600;
                font-size: 12px; padding: 7px 14px; border-radius: 5px; border: 1px solid #4338ca;
            }
            QPushButton:hover { background-color: #4338ca; }
            QPushButton:pressed { background-color: #3730a3; }
            QPushButton:disabled { background-color: #f1f5f9; color: #94a3b8; border-color: #e2e8f0; }
        """)
        self.btn_sign_pdf.setToolTip("Ký số điện tử bằng USB Token Ban Cơ yếu hoặc VGCA SignTool theo Nghị định 30/2020/NĐ-CP")
        self.btn_sign_pdf.setEnabled(True)
        self.btn_sign_pdf.clicked.connect(self._open_signature_dialog)

        # 3. Nút Mở file Word (Thứ cấp - Nền trắng trang nhã, viền mềm mại)
        self.btn_open_word = QPushButton("Mở file Word")
        self.btn_open_word.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #1e293b; font-weight: 500;
                font-size: 12px; padding: 7px 11px; border-radius: 5px; border: 1px solid #cbd5e1;
            }
            QPushButton:hover { background-color: #eff6ff; border-color: #93c5fd; color: #1d4ed8; }
            QPushButton:pressed { background-color: #dbeafe; }
            QPushButton:disabled { background-color: #f8fafc; color: #cbd5e1; border-color: #e2e8f0; }
        """)
        self.btn_open_word.setEnabled(False)
        self.btn_open_word.clicked.connect(self._open_output_word)

        # 4. Nút Mở file PDF (Thứ cấp - Nền trắng trang nhã)
        self.btn_open_pdf = QPushButton("Mở file PDF")
        self.btn_open_pdf.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #1e293b; font-weight: 500;
                font-size: 12px; padding: 7px 11px; border-radius: 5px; border: 1px solid #cbd5e1;
            }
            QPushButton:hover { background-color: #fef2f2; border-color: #fca5a5; color: #b91c1c; }
            QPushButton:pressed { background-color: #fee2e2; }
            QPushButton:disabled { background-color: #f8fafc; color: #cbd5e1; border-color: #e2e8f0; }
        """)
        self.btn_open_pdf.setEnabled(False)
        self.btn_open_pdf.clicked.connect(self._open_output_pdf)

        # 5. Nút Mở thư mục chứa (Thứ cấp)
        self.btn_open_dir = QPushButton("Mở thư mục")
        self.btn_open_dir.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #475569; font-weight: 500;
                font-size: 12px; padding: 7px 11px; border-radius: 5px; border: 1px solid #cbd5e1;
            }
            QPushButton:hover { background-color: #f8fafc; border-color: #94a3b8; color: #0f172a; }
            QPushButton:pressed { background-color: #f1f5f9; }
            QPushButton:disabled { background-color: #f8fafc; color: #cbd5e1; border-color: #e2e8f0; }
        """)
        self.btn_open_dir.setEnabled(False)
        self.btn_open_dir.clicked.connect(self._open_output_folder)

        # 6. Nút Nhận dạng OCR bổ trợ (Viền hổ phách nhẹ)
        self.btn_run_ocr = QPushButton("Nhận dạng OCR")
        self.btn_run_ocr.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #b45309; font-weight: 500;
                font-size: 12px; padding: 7px 11px; border-radius: 5px; border: 1px solid #fcd34d;
            }
            QPushButton:hover { background-color: #fffbeb; border-color: #f59e0b; color: #92400e; }
            QPushButton:pressed { background-color: #fef3c7; }
            QPushButton:disabled { background-color: #f8fafc; color: #cbd5e1; border-color: #e2e8f0; }
        """)
        self.btn_run_ocr.setEnabled(False)
        self.btn_run_ocr.clicked.connect(self._run_ocr_on_selected_file)

        # Bố trí thứ tự logic: Tác vụ chính -> Xem xuất bản -> Tiện ích bổ trợ
        actions_layout.addWidget(self.btn_embedded_edit)
        actions_layout.addWidget(self.btn_sign_pdf)
        actions_layout.addWidget(self.btn_open_word)
        actions_layout.addWidget(self.btn_open_pdf)
        actions_layout.addWidget(self.btn_open_dir)
        actions_layout.addWidget(self.btn_run_ocr)
        right_layout.addLayout(actions_layout)

        # Thêm 2 cột vào Splitter
        left_widget.setMinimumWidth(320)
        right_widget.setMinimumWidth(340)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setChildrenCollapsible(False)
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 6)

        main_layout.addWidget(splitter)

    def _create_stat_card(self, title: str, initial_value: str, category: str, accent_color: str = "#2563eb", bg_tint: str = "#eff6ff") -> QFrame:
        card = QFrame()
        card.setObjectName("statCard")
        card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        card.setMinimumWidth(80)
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card.setProperty("active", "false")
        card.category = category
        card.accent_color = accent_color

        if category == "all":
            card.setToolTip("Bấm để hiển thị toàn bộ các lỗi & đoạn văn")
        elif category == "whitespace":
            card.setToolTip("Bấm để lọc và chỉ hiển thị lỗi khoảng trắng thừa")
        elif category == "spelling_punct":
            card.setToolTip("Bấm để lọc và chỉ hiển thị lỗi chính tả & dấu câu")
        elif category == "capitalization":
            card.setToolTip("Bấm để lọc và chỉ hiển thị cảnh báo quy tắc viết hoa")

        card.setStyleSheet(f"""
            QFrame#statCard {{
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                border-top: 3.5px solid {accent_color};
                border-radius: 6px;
            }}
            QFrame#statCard:hover {{
                background-color: {bg_tint};
                border-color: {accent_color};
            }}
            QFrame#statCard[active="true"] {{
                background-color: {bg_tint};
                border: 2px solid {accent_color};
                border-top: 4px solid {accent_color};
            }}
        """)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(2)

        lbl_t = QLabel(title.upper())
        lbl_t.setStyleSheet(f"font-size: 10.5px; font-weight: 700; color: {accent_color}; letter-spacing: 0.3px;")
        lbl_v = QLabel(initial_value)
        lbl_v.setStyleSheet("font-size: 21px; font-weight: 800; color: #0f172a;")

        lay.addWidget(lbl_t)
        lay.addWidget(lbl_v)

        card.value_label = lbl_v  # tham chiếu nhanh
        card.mousePressEvent = lambda e, cat=category: self._on_stat_card_clicked(cat)
        return card

    def _on_stat_card_clicked(self, category: str):
        if self.active_issue_filter == category and category != "all":
            self.active_issue_filter = "all"
        else:
            self.active_issue_filter = category

        for c in self.stat_cards:
            is_act = (c.category == self.active_issue_filter and self.active_issue_filter != "all")
            c.setProperty("active", "true" if is_act else "false")
            c.style().unpolish(c)
            c.style().polish(c)

        self._render_current_issues_table()

    def _refresh_profiles_dropdown(self):
        self.cb_profiles.clear()
        self.cb_profiles.addItem("Tự động nhận diện từ tài liệu gốc (Khuyên dùng)", None)
        profiles = self.profiles_data.get("profiles", [])
        current_std = "HD36" if self.rb_hd36.isChecked() else "ND30"

        for p in profiles:
            if p.get("standard_id", "ND30") == current_std:
                self.cb_profiles.addItem(f"Áp dụng mẫu: {p.get('name', p.get('id'))}", p)

        self.cb_profiles.setCurrentIndex(0)

    def _on_standard_changed(self):
        self._refresh_profiles_dropdown()

    def _on_profile_changed(self, idx: int):
        pass

    def _open_profile_dialog(self):
        from ui.profile_dialog import ProfileDialog
        dlg = ProfileDialog(self, config_dir=self.config_dir)
        if dlg.exec():
            self._load_profiles_data()
            self._refresh_profiles_dropdown()

    def _choose_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn các tệp Word hoặc PDF cần chuẩn hóa thể thức",
            "",
            "Văn bản & Tài liệu scan (*.docx *.pdf);;Word Documents (*.docx);;PDF Documents (*.pdf)"
        )
        if files:
            self._add_files(files)

    def _choose_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục chứa các tệp văn bản dự thảo (.docx, .pdf)"
        )
        if folder:
            collected = []
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith((".docx", ".pdf")):
                        collected.append(os.path.join(root, f))
            if collected:
                self._add_files(collected)
            else:
                QMessageBox.information(
                    self,
                    "Không tìm thấy tệp",
                    "Không tìm thấy tệp .docx hoặc .pdf nào trong thư mục đã chọn."
                )

    def _add_files(self, paths: List[str]):
        from core.security_guard import SecurityGuard
        for path in paths:
            ok, msg = SecurityGuard.validate_input_file(path)
            if not ok:
                QMessageBox.warning(self, "Cảnh báo An toàn Tệp", f"Từ chối tệp '{os.path.basename(path)}':\n{msg}")
                continue

            if path not in self.files_queue:
                self.files_queue.append(path)
                row = self.table_files.rowCount()
                self.table_files.insertRow(row)

                file_name = os.path.basename(path)
                size_kb = f"{os.path.getsize(path) / 1024:.1f} KB" if os.path.exists(path) else "-"

                self.table_files.setItem(row, 0, QTableWidgetItem(file_name))
                self.table_files.setItem(row, 1, QTableWidgetItem(size_kb))

                status_item = QTableWidgetItem("Chờ xử lý")
                status_item.setForeground(QColor("#64748b"))
                self.table_files.setItem(row, 2, status_item)

        self.lbl_status.setText(f"Đã nạp {len(self.files_queue)} tài liệu trong danh sách.")

    def _remove_selected_file(self):
        row = self.table_files.currentRow()
        if 0 <= row < len(self.files_queue):
            self.files_queue.pop(row)
            self.table_files.removeRow(row)
            self.reports_map.pop(row, None)
            self.com_results_map.pop(row, None)

    def _clear_all_files(self):
        self.files_queue.clear()
        self.table_files.setRowCount(0)
        self.reports_map.clear()
        self.com_results_map.clear()
        self._clear_reports_view()

    def _clear_reports_view(self):
        self.current_report = None
        self.active_issue_filter = "all"
        for c in getattr(self, "stat_cards", []):
            c.setProperty("active", "false")
            c.style().unpolish(c)
            c.style().polish(c)
        self.card_paragraphs.value_label.setText("0")
        self.card_spaces.value_label.setText("0")
        self.card_punct.value_label.setText("0")
        self.card_warnings.value_label.setText("0")
        self.table_issues.setRowCount(0)
        if hasattr(self, "report_group"):
            self.report_group.setTitle("Chi tiết lỗi")
        self.btn_open_word.setEnabled(False)
        self.btn_embedded_edit.setEnabled(True)
        self.btn_open_pdf.setEnabled(False)
        self.btn_sign_pdf.setEnabled(True)
        self.btn_open_dir.setEnabled(False)
        self.btn_run_ocr.setEnabled(False)

    def _cancel_processing(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.cancel()
            self.lbl_status.setText("Đang dừng tiến trình...")
            self.btn_cancel.setEnabled(False)
        if hasattr(self, 'ocr_worker') and self.ocr_worker.isRunning():
            self.ocr_worker.cancel()
            self.lbl_status.setText("Đang dừng nhận dạng OCR...")
            self.btn_cancel.setEnabled(False)

    def _start_processing(self):
        if not self.files_queue:
            QMessageBox.warning(self, "Chưa chọn tệp", "Vui lòng kéo-thả hoặc nhấn 'Thêm tệp...' để chọn tài liệu cần chuẩn hóa!")
            return

        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_run_ocr.setEnabled(False)
        self.progress_bar.setValue(0)
        self._clear_reports_view()

        standard_id = "HD36" if self.rb_hd36.isChecked() else "ND30"
        profile_data = self.cb_profiles.currentData()
        export_pdf = self.chk_export_pdf.isChecked()
        auto_ocr_scans = self.chk_auto_ocr.isChecked()
        custom_out_dir = self.txt_output_dir.text().strip() or None

        # Tạo worker chạy ngầm
        self.worker = ProcessingWorker(
            files=self.files_queue,
            standard_id=standard_id,
            profile_data=profile_data,
            export_pdf=export_pdf,
            config_dir=self.config_dir,
            auto_ocr_scans=auto_ocr_scans,
            output_dir=custom_out_dir
        )
        self.worker.progress_updated.connect(self._on_worker_progress)
        self.worker.file_started.connect(self._on_file_started)
        self.worker.file_finished.connect(self._on_file_finished)
        self.worker.all_finished.connect(self._on_all_finished)
        self.worker.start()

    def _on_file_started(self, row: int, base_name: str):
        if 0 <= row < self.table_files.rowCount():
            status_item = self.table_files.item(row, 2)
            if status_item:
                status_item.setText("Đang xử lý...")
                status_item.setForeground(QColor("#2563eb"))
            self.table_files.selectRow(row)

    def _on_worker_progress(self, pct: int, msg: str):
        self.progress_bar.setValue(pct)
        self.lbl_status.setText(msg)

    def _on_file_finished(self, row: int, report: FormattingReport, com_res: dict):
        self.reports_map[row] = report
        self.com_results_map[row] = com_res

        status_item = self.table_files.item(row, 2)
        if report.success:
            if report.is_scanned_image and not any(i.issue_type == "ocr_processed" for i in report.all_issues):
                status_item.setText("Tệp ảnh scan (Bảo vệ)")
                status_item.setForeground(QColor("#d97706"))
            elif any(i.issue_type == "ocr_processed" for i in report.all_issues):
                status_item.setText("Đã OCR & Chuẩn hóa")
                status_item.setForeground(QColor("#16a34a"))
            else:
                status_item.setText("Đã chuẩn hóa")
                status_item.setForeground(QColor("#16a34a"))
        else:
            status_item.setText("Lỗi")
            status_item.setForeground(QColor("#dc2626"))

        # Nếu là dòng được chọn hiện tại hoặc dòng đầu tiên, hiển thị báo cáo
        if self.table_files.currentRow() == row or self.table_files.currentRow() == -1:
            self._display_report(report, com_res)

    def _on_all_finished(self, was_cancelled: bool = False):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_run_ocr.setEnabled(True)

        if was_cancelled:
            QMessageBox.information(
                self,
                "Đã Dừng",
                "Đã dừng tiến trình chuẩn hóa theo yêu cầu của bạn."
            )
            return

        QMessageBox.information(
            self,
            "Hoàn tất Xử lý",
            f"Đã xử lý xong {len(self.files_queue)} tài liệu!\n"
            "Các tệp văn bản đã được căn chỉnh thể thức chuẩn xác theo quy định."
        )

        if self.chk_open_folder.isChecked():
            out_custom = self.txt_output_dir.text().strip()
            if out_custom and os.path.exists(out_custom):
                try:
                    os.startfile(out_custom)
                except Exception:
                    pass
            elif self.files_queue:
                first_rep = self.reports_map.get(0)
                if first_rep and os.path.exists(first_rep.output_file):
                    folder = os.path.dirname(first_rep.output_file)
                    try:
                        os.startfile(folder)
                    except Exception:
                        pass

    def _on_file_selected(self, row: int, col: int):
        report = self.reports_map.get(row)
        com_res = self.com_results_map.get(row, {})
        if report:
            self._display_report(report, com_res)
        else:
            self._clear_reports_view()

    def _display_report(self, report: FormattingReport, com_res: dict):
        self.current_report = report
        # Cập nhật số liệu trên thẻ
        self.card_paragraphs.value_label.setText(str(report.paragraphs_formatted))
        self.card_spaces.value_label.setText(str(report.spaces_fixed))
        total_spelling_punct = report.punctuation_fixed + report.spelling_fixed + report.dates_fixed + report.quotes_fixed
        self.card_punct.value_label.setText(str(total_spelling_punct))
        self.card_warnings.value_label.setText(str(report.cap_warnings_count))

        # Hiển thị bảng chi tiết lỗi theo bộ lọc đang chọn
        self._render_current_issues_table()

        # Kích hoạt các nút mở file
        self.btn_open_word.setEnabled(os.path.exists(report.output_file))
        self.btn_embedded_edit.setEnabled(True)
        # Nút PDF luôn kích hoạt nếu đã có tệp Word kết quả (để tạo PDF tức thì theo yêu cầu)
        self.btn_open_pdf.setEnabled(os.path.exists(report.output_file))
        self.btn_sign_pdf.setEnabled(True)
        self.btn_open_dir.setEnabled(os.path.exists(report.output_file))

        # Bật nút chạy OCR nếu là tệp ảnh scan hoặc khuyến nghị OCR
        self.btn_run_ocr.setEnabled(True)
        if report.is_scanned_image or report.ocr_recommended:
            self.btn_run_ocr.setText("Nhận dạng OCR & Chuẩn hóa")
            self.btn_run_ocr.setToolTip("Khuyên dùng: Bóc tách hình ảnh scan thành văn bản có thể chỉnh sửa")
        else:
            self.btn_run_ocr.setText("Nhận dạng lại bằng OCR")
            self.btn_run_ocr.setToolTip("Bóc tách lại nội dung tài liệu bằng công nghệ OCR ngoại tuyến")

    def _render_current_issues_table(self):
        if not self.current_report:
            self.table_issues.setRowCount(0)
            self.report_group.setTitle("Chi tiết lỗi")
            return

        all_issues = self.current_report.all_issues
        f = self.active_issue_filter

        if f == "whitespace":
            filtered = [i for i in all_issues if i.issue_type == "whitespace"]
            self.report_group.setTitle(f"Chi tiết lỗi — Khoảng trắng thừa ({len(filtered)})")
        elif f == "spelling_punct":
            filtered = [i for i in all_issues if i.issue_type in ["spelling", "punctuation", "quote", "date"]]
            self.report_group.setTitle(f"Chi tiết lỗi — Chính tả & Dấu câu ({len(filtered)})")
        elif f == "capitalization":
            filtered = [i for i in all_issues if i.issue_type == "capitalization"]
            self.report_group.setTitle(f"Chi tiết lỗi — Cảnh báo viết hoa ({len(filtered)})")
        else:
            filtered = all_issues
            self.report_group.setTitle(f"Chi tiết lỗi ({len(filtered)})")

        self.table_issues.setRowCount(0)
        for issue in filtered:
            r = self.table_issues.rowCount()
            self.table_issues.insertRow(r)

            type_name = {
                "spelling": "Chính tả & Dấu chữ",
                "punctuation": "Dấu câu hành chính",
                "whitespace": "Khoảng trắng thừa",
                "quote": "Ngoặc kép chuẩn",
                "date": "Thể thức ngày tháng",
                "capitalization": "Quy tắc Viết hoa",
                "scanned_image": "Ảnh scan",
                "ocr_processed": "Xử lý OCR"
            }.get(issue.issue_type, issue.issue_type)

            orig_text = issue.original.strip()
            repl_text = issue.replacement.strip()
            orig_disp = orig_text if len(orig_text) <= 50 else orig_text[:47] + "..."
            repl_disp = repl_text if len(repl_text) <= 50 else repl_text[:47] + "..."

            it_type = QTableWidgetItem(type_name)
            it_orig = QTableWidgetItem(orig_disp)
            it_orig.setToolTip(orig_text)
            it_repl = QTableWidgetItem(repl_disp)
            it_repl.setToolTip(repl_text)
            it_msg = QTableWidgetItem(issue.message)
            it_msg.setToolTip(issue.message)

            if issue.severity == "warning":
                it_type.setForeground(QColor("#d97706"))
                it_msg.setForeground(QColor("#b45309"))
            else:
                it_type.setForeground(QColor("#16a34a"))
                it_repl.setForeground(QColor("#16a34a"))

            self.table_issues.setItem(r, 0, it_type)
            self.table_issues.setItem(r, 1, it_orig)
            self.table_issues.setItem(r, 2, it_repl)
            self.table_issues.setItem(r, 3, it_msg)

    def _run_ocr_on_selected_file(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0

        if 0 <= row < len(self.files_queue):
            file_path = self.files_queue[row]
            standard_id = "HD36" if self.rb_hd36.isChecked() else "ND30"
            profile_data = self.cb_profiles.currentData()
            export_pdf = self.chk_export_pdf.isChecked()

            self.btn_run_ocr.setEnabled(False)
            self.btn_start.setEnabled(False)
            self.btn_cancel.setEnabled(True)
            self.progress_bar.setValue(5)
            self.lbl_status.setText(f"Đang bóc tách ảnh scan và chạy OCR: {os.path.basename(file_path)}...")

            custom_out_dir = self.txt_output_dir.text().strip() or None
            self.ocr_worker = OcrWorker(
                file_path=file_path,
                standard_id=standard_id,
                profile_data=profile_data,
                export_pdf=export_pdf,
                config_dir=self.config_dir,
                output_dir=custom_out_dir
            )
            self.ocr_worker.progress_updated.connect(self._on_worker_progress)
            self.ocr_worker.ocr_finished.connect(lambda rep, com: self._on_single_ocr_finished(row, rep, com))
            self.ocr_worker.start()

    def _on_single_ocr_finished(self, row: int, report: FormattingReport, com_res: dict):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_run_ocr.setEnabled(True)
        self.reports_map[row] = report
        self.com_results_map[row] = com_res

        status_item = self.table_files.item(row, 2)
        if report.success:
            status_item.setText("Đã OCR & Chuẩn hóa")
            status_item.setForeground(QColor("#16a34a"))
            self._display_report(report, com_res)
            QMessageBox.information(
                self,
                "Nhận dạng OCR Thành công",
                f"Đã nhận dạng thành công các dòng chữ từ ảnh scan và tự động chuẩn hóa thể thức văn bản!\n"
                f"Tệp kết quả: {os.path.basename(report.output_file)}"
            )
        else:
            status_item.setText("Lỗi OCR")
            status_item.setForeground(QColor("#dc2626"))
            QMessageBox.critical(self, "Lỗi OCR", f"Không thể nhận dạng OCR: {report.error_message}")

    def _open_quick_wizard(self):
        """Mở Wizard tạo nhanh văn bản chuẩn thể thức NĐ 30 / HD 36"""
        try:
            wiz = QuickWizardDialog(self)
            if wiz.exec() and getattr(wiz, 'result_metadata', None):
                self._active_editor = WordEditorWindow(None, initial_metadata=wiz.result_metadata)
                self._active_editor.showMaximized()
                self._active_editor.raise_()
                self._active_editor.activateWindow()
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở trình soạn thảo: {str(e)}")

    def _open_in_embedded_editor(self):
        """Mở văn bản hiện tại trong Trình Soạn thảo Hành chính WYSIWYG tích hợp (Mặc định toàn màn hình)"""
        try:
            row = self.table_files.currentRow()
            if row < 0 and self.files_queue:
                row = 0
            report = self.reports_map.get(row)
            target = report.output_file if (report and os.path.exists(report.output_file)) else (self.files_queue[row] if (row >= 0 and row < len(self.files_queue)) else "")
            if target and os.path.exists(target):
                std = self._get_selected_standard()
                self._active_editor = WordEditorWindow(None, initial_metadata={"standard_id": std}, current_file_path=target)
                self._active_editor.showMaximized()
                self._active_editor.raise_()
                self._active_editor.activateWindow()
            else:
                self._open_quick_wizard()
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở trình soạn thảo: {str(e)}")

    def _open_output_word(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0
        report = self.reports_map.get(row)
        if report and os.path.exists(report.output_file):
            target = os.path.abspath(report.output_file)
            try:
                os.startfile(target)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(["explorer.exe", target], shell=False)
                except Exception as e:
                    QMessageBox.warning(self, "Mở Word", f"Không thể mở tệp Word: {str(e)}")

    def _set_override_wait_cursor(self, busy: bool):
        try:
            if busy:
                QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            else:
                QApplication.restoreOverrideCursor()
        except Exception:
            pass

    def _open_output_pdf(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0
        report = self.reports_map.get(row)
        if not report or not os.path.exists(report.output_file):
            return

        base, _ = os.path.splitext(report.output_file)
        pdf_path = os.path.abspath(f"{base}.pdf")

        # Nếu file PDF chưa tồn tại -> Tự động chuyển đổi từ file Word kết quả ngay lập tức
        if not os.path.exists(pdf_path):
            self.lbl_status.setText("Đang tự động chuyển đổi sang tệp PDF bằng MS Word...")
            self._set_override_wait_cursor(True)
            try:
                from core.com_helper import ComHelper
                com = ComHelper()
                ok, msg = com.convert_docx_to_pdf(report.output_file, pdf_path)
                if not ok or not os.path.exists(pdf_path):
                    self._set_override_wait_cursor(False)
                    QMessageBox.warning(self, "Xuất PDF", f"Không thể xuất file PDF (yêu cầu máy có cài MS Word): {msg}")
                    return
                self.com_results_map.setdefault(row, {})["pdf_path"] = pdf_path
            except Exception as e:
                self._set_override_wait_cursor(False)
                QMessageBox.warning(self, "Lỗi xuất PDF", f"Lỗi khi chuyển đổi PDF: {str(e)}")
                return
            finally:
                self._set_override_wait_cursor(False)
                self.lbl_status.setText("Đã xuất tệp PDF thành công.")

        if os.path.exists(pdf_path):
            try:
                os.startfile(pdf_path)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(["explorer.exe", pdf_path], shell=False)
                except Exception:
                    QMessageBox.information(
                        self, "Mở PDF",
                        f"Tệp PDF đã sẵn sàng tại:\n{pdf_path}"
                    )

    def _open_output_folder(self):
        out_custom = self.txt_output_dir.text().strip()
        folder = ""
        if out_custom and os.path.exists(out_custom):
            folder = os.path.abspath(out_custom)
        else:
            row = self.table_files.currentRow()
            if row < 0 and self.files_queue:
                row = 0
            report = self.reports_map.get(row)
            if report and os.path.exists(report.output_file):
                folder = os.path.abspath(os.path.dirname(report.output_file))
            elif self.files_queue and 0 <= row < len(self.files_queue):
                folder = os.path.abspath(os.path.dirname(self.files_queue[row]))

        if folder and os.path.exists(folder):
            try:
                os.startfile(folder)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(["explorer.exe", folder], shell=False)
                except Exception:
                    pass

    def _open_signature_dialog(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0

        pdf_path = ""
        report = self.reports_map.get(row)

        if report and report.output_file and os.path.exists(report.output_file):
            base, ext = os.path.splitext(report.output_file)
            if ext.lower() == ".pdf":
                pdf_path = report.output_file
            else:
                pdf_path = f"{base}.pdf"
                # Nếu tệp PDF chưa tồn tại -> Tự động chuyển đổi từ file Word kết quả ngay
                if not os.path.exists(pdf_path):
                    self.lbl_status.setText("Đang tự động xuất tệp PDF để chuẩn bị ký số...")
                    self._set_override_wait_cursor(True)
                    try:
                        from core.com_helper import ComHelper
                        com = ComHelper()
                        ok, msg = com.convert_docx_to_pdf(report.output_file, pdf_path)
                        if not ok or not os.path.exists(pdf_path):
                            self._set_override_wait_cursor(False)
                            QMessageBox.warning(self, "Xuất PDF", f"Không thể xuất file PDF để ký số: {msg}")
                            return
                    except Exception as e:
                        self._set_override_wait_cursor(False)
                        QMessageBox.warning(self, "Lỗi xuất PDF", f"Lỗi chuyển đổi PDF: {str(e)}")
                        return
                    finally:
                        self._set_override_wait_cursor(False)
        elif 0 <= row < len(self.files_queue):
            orig = self.files_queue[row]
            if orig.lower().endswith(".pdf") and os.path.exists(orig):
                pdf_path = orig

        if not pdf_path or not os.path.exists(pdf_path):
            chosen, _ = QFileDialog.getOpenFileName(
                self, "Chọn tệp PDF cần ký số điện tử (NĐ 30/2020)",
                "", "Tài liệu PDF (*.pdf);;Tất cả tệp (*.*)"
            )
            if chosen and os.path.exists(chosen):
                pdf_path = chosen
            else:
                return

        agency_name = ""
        prof = self.cb_profiles.currentData()
        if prof and isinstance(prof, dict):
            agency_name = prof.get("name", prof.get("id", ""))

        from ui.signature_dialog import SignatureDialog
        dlg = SignatureDialog(self, pdf_path=pdf_path, default_agency=agency_name)
        if dlg.exec() and dlg.signed_pdf_path and os.path.exists(dlg.signed_pdf_path):
            self.lbl_status.setText(f"Đã hoàn tất ký số văn bản: {os.path.basename(dlg.signed_pdf_path)}")
            try:
                os.startfile(dlg.signed_pdf_path)
            except Exception:
                pass

    def _choose_output_directory(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu tệp xuất")
        if folder:
            self.txt_output_dir.setText(folder)

    def _reset_output_directory(self):
        self.txt_output_dir.clear()

    def _open_batch_dialog(self, initial_tab: int = 0):
        from ui.batch_editor_dialog import BatchEditorDialog
        selected_rows = sorted(set(index.row() for index in self.table_files.selectedIndexes()))
        if selected_rows:
            target_files = [self.files_queue[r] for r in selected_rows if 0 <= r < len(self.files_queue)]
        else:
            target_files = list(self.files_queue)

        dlg = BatchEditorDialog(self, initial_files=target_files, initial_tab=initial_tab)
        dlg.exec()

    def _open_batch_editor_dialog(self):
        self._open_batch_dialog(initial_tab=0)

    def _open_batch_rename_dialog(self):
        self._open_batch_dialog(initial_tab=1)
