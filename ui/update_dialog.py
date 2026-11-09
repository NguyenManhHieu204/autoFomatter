"""
ui/update_dialog.py: Giao diện thông báo và thực hiện cập nhật phiên bản mới cho Auto-Formatter.
Thiết kế theo chuẩn Typography thanh lịch, không dùng icon/emoji lòe loẹt.
"""

import sys
import webbrowser
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QTextBrowser, QMessageBox, QApplication
)
from PyQt6.QtCore import Qt
from core.updater import DownloadUpdateThread, apply_update_and_restart


class UpdateDialog(QDialog):
    """Hộp thoại thông báo có bản cập nhật mới và hỗ trợ tải về / tự động nâng cấp."""

    def __init__(self, parent=None, update_info: dict = None):
        super().__init__(parent)
        self.update_info = update_info or {}
        self.download_thread = None

        self.setWindowTitle("Cập nhật phần mềm — Auto-Formatter")
        self.setMinimumSize(540, 420)
        self.resize(560, 460)

        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        # Tiêu đề phiên bản
        lbl_title = QLabel("ĐÃ CÓ PHIÊN BẢN MỚI")
        lbl_title.setStyleSheet("font-size: 15pt; font-weight: bold; color: #0F172A;")
        layout.addWidget(lbl_title)

        cur_ver = self.update_info.get("current_version", "2.0.0")
        new_ver = self.update_info.get("latest_version", "")
        rel_name = self.update_info.get("release_name", "")

        sub_text = f"Phiên bản hiện tại: <b>v{cur_ver}</b>  →  Phiên bản phát hành: <b>{new_ver}</b>"
        if rel_name and rel_name != new_ver:
            sub_text += f"<br><i>{rel_name}</i>"

        lbl_sub = QLabel(sub_text)
        lbl_sub.setStyleSheet("font-size: 10pt; color: #334155;")
        layout.addWidget(lbl_sub)

        # Ghi chú phát hành (Release Notes)
        lbl_notes = QLabel("Nội dung thay đổi & Cải tiến:")
        lbl_notes.setStyleSheet("font-weight: 600; font-size: 9.5pt; color: #1E293B;")
        layout.addWidget(lbl_notes)

        self.txt_notes = QTextBrowser()
        notes_content = self.update_info.get("release_notes", "") or "Không có ghi chú chi tiết cho bản phát hành này."
        self.txt_notes.setPlainText(notes_content)
        self.txt_notes.setStyleSheet("""
            QTextBrowser {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 10px;
                font-size: 9.5pt;
                line-height: 1.4;
                color: #0F172A;
            }
        """)
        layout.addWidget(self.txt_notes, stretch=1)

        # Khu vực tiến trình tải về
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("font-size: 9pt; color: #475569;")
        self.lbl_status.setVisible(False)
        layout.addWidget(self.lbl_status)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CBD5E1;
                border-radius: 3px;
                text-align: center;
                height: 18px;
                font-size: 8.5pt;
                color: #0F172A;
            }
            QProgressBar::chunk {
                background-color: #0284C7;
            }
        """)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Hàng nút bấm tác vụ
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.btn_github = QPushButton("Xem trên GitHub")
        self.btn_github.setStyleSheet("""
            QPushButton {
                padding: 7px 14px;
                border: 1px solid #94A3B8;
                border-radius: 4px;
                background-color: #FFFFFF;
                color: #334155;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #64748B;
            }
        """)
        self.btn_github.clicked.connect(self._open_github)
        btn_layout.addWidget(self.btn_github)

        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Để sau")
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                padding: 7px 16px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                color: #475569;
            }
            QPushButton:hover {
                background-color: #F8FAFC;
            }
        """)
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_update = QPushButton("Tải & Cập nhật tự động")
        self.btn_update.setStyleSheet("""
            QPushButton {
                padding: 7px 20px;
                border: none;
                border-radius: 4px;
                background-color: #0284C7;
                color: #FFFFFF;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #0369A1;
            }
            QPushButton:disabled {
                background-color: #94A3B8;
            }
        """)
        self.btn_update.clicked.connect(self._start_download_and_update)
        btn_layout.addWidget(self.btn_update)

        layout.addLayout(btn_layout)

    def _open_github(self):
        url = self.update_info.get("html_url") or "https://github.com/NguyenManhHieu204/autoFomatter/releases"
        webbrowser.open(url)

    def _start_download_and_update(self):
        download_url = self.update_info.get("download_url")
        if not download_url:
            QMessageBox.warning(
                self,
                "Không có gói cập nhật tự động",
                "Phiên bản này chưa đính kèm tệp nén cập nhật trực tiếp (.zip).\n\n"
                "Bạn sẽ được chuyển hướng đến trang GitHub để tải thủ công."
            )
            self._open_github()
            return

        self.btn_update.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self.lbl_status.setVisible(True)
        self.progress_bar.setVisible(True)
        self.lbl_status.setText("Đang kết nối máy chủ GitHub để tải gói cập nhật...")

        self.download_thread = DownloadUpdateThread(download_url)
        self.download_thread.progress.connect(self._on_download_progress)
        self.download_thread.finished.connect(self._on_download_finished)
        self.download_thread.error.connect(self._on_download_error)
        self.download_thread.start()

    def _on_download_progress(self, downloaded: int, total: int):
        if total > 0:
            pct = int((downloaded / total) * 100)
            self.progress_bar.setValue(pct)
            mb_down = downloaded / (1024 * 1024)
            mb_total = total / (1024 * 1024)
            self.lbl_status.setText(f"Đang tải bản cập nhật: {mb_down:.1f} MB / {mb_total:.1f} MB ({pct}%)")
        else:
            mb_down = downloaded / (1024 * 1024)
            self.lbl_status.setText(f"Đang tải bản cập nhật: {mb_down:.1f} MB...")

    def _on_download_finished(self, zip_path: str):
        self.progress_bar.setValue(100)
        self.lbl_status.setText("Tải hoàn tất! Đang chuẩn bị cài đặt...")

        is_frozen = getattr(sys, "frozen", False)
        if not is_frozen:
            QMessageBox.information(
                self,
                "Chế độ phát triển (Source Code)",
                f"Đã tải gói cập nhật về:\n{zip_path}\n\n"
                "Do ứng dụng đang chạy từ mã nguồn Python trực tiếp, "
                "bạn vui lòng dùng lệnh 'git pull' để đồng bộ mã nguồn mới nhất."
            )
            self.accept()
            return

        reply = QMessageBox.question(
            self,
            "Cài đặt bản cập nhật",
            "Đã tải thành công bản cập nhật!\n\n"
            "Ứng dụng sẽ tự động đóng và khởi động lại phiên bản mới trong vài giây.\n"
            "Bạn có muốn khởi động lại ngay không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )

        if reply == QMessageBox.StandardButton.Yes:
            success, msg = apply_update_and_restart(zip_path)
            if success:
                QApplication.quit()
            else:
                QMessageBox.critical(self, "Lỗi cài đặt", msg)
                self.btn_update.setEnabled(True)
                self.btn_cancel.setEnabled(True)
        else:
            self.accept()

    def _on_download_error(self, err_msg: str):
        self.lbl_status.setText("Tải thất bại.")
        QMessageBox.critical(
            self,
            "Lỗi tải cập nhật",
            f"Không thể tải gói cập nhật tự động:\n{err_msg}\n\n"
            "Vui lòng thử lại hoặc tải thủ công qua trang GitHub."
        )
        self.btn_update.setEnabled(True)
        self.btn_cancel.setEnabled(True)
