"""
core/updater.py: Module kiểm tra và tự động cập nhật phiên bản mới từ GitHub Releases.
Hỗ trợ:
- So sánh semantic version (ví dụ: 2.0.0 < 2.0.1)
- Truy vấn GitHub API ngoại tuyến an toàn (không gây nghẽn, timeout ngắn)
- Tải gói cập nhật (.zip) với thanh tiến trình
- Kích hoạt batch updater tự động ghi đè và khởi động lại ứng dụng
"""

import os
import sys
import re
import json
import shutil
import tempfile
import subprocess
import urllib.request
import urllib.error
from typing import Dict, Any, Optional, Tuple
from PyQt6.QtCore import QThread, pyqtSignal

APP_VERSION = "2.0.0"
GITHUB_REPO = "NguyenManhHieu204/autoFomatter"
USER_AGENT = "AutoFormatter-Desktop-App/2.0"


def parse_version(v_str: str) -> Tuple[int, ...]:
    """Phân tách chuỗi phiên bản dạng 'v2.0.1' hoặc '2.1' thành tuple số (2, 0, 1)."""
    if not v_str:
        return (0, 0, 0)
    nums = re.findall(r"\d+", str(v_str))
    return tuple(map(int, nums)) if nums else (0, 0, 0)


def is_newer_version(latest_str: str, current_str: str = APP_VERSION) -> bool:
    """Trả về True nếu latest_str mới hơn current_str."""
    return parse_version(latest_str) > parse_version(current_str)


def check_for_updates(repo: str = GITHUB_REPO, current_version: str = APP_VERSION, timeout: int = 6) -> Dict[str, Any]:
    """Truy vấn GitHub Releases API để kiểm tra phiên bản mới nhất.
    
    Returns:
        dict chứa:
            - success: bool
            - has_update: bool
            - current_version: str
            - latest_version: str
            - release_name: str
            - release_notes: str
            - html_url: str
            - download_url: str (URL tải file .zip)
            - asset_name: str
            - asset_size: int (bytes)
            - error: str (nếu có lỗi)
    """
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "application/vnd.github.v3+json"
        }
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            if response.status == 200:
                data = json.loads(response.read().decode("utf-8"))
                tag_name = data.get("tag_name", "").strip()
                latest_ver_clean = tag_name.lstrip("v")
                has_update = is_newer_version(latest_ver_clean, current_version)

                assets = data.get("assets", [])
                download_url = None
                asset_name = ""
                asset_size = 0

                for asset in assets:
                    name = asset.get("name", "")
                    if name.lower().endswith(".zip") or name.lower().endswith(".exe"):
                        download_url = asset.get("browser_download_url")
                        asset_name = name
                        asset_size = asset.get("size", 0)
                        if "setup" in name.lower() or name.lower().endswith(".exe"):
                            break

                return {
                    "success": True,
                    "has_update": has_update,
                    "current_version": current_version,
                    "latest_version": tag_name or latest_ver_clean,
                    "release_name": data.get("name") or tag_name,
                    "release_notes": data.get("body", ""),
                    "html_url": data.get("html_url") or f"https://github.com/{repo}/releases",
                    "download_url": download_url,
                    "asset_name": asset_name,
                    "asset_size": asset_size,
                    "published_at": data.get("published_at", "")
                }
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return {
                "success": True,
                "has_update": False,
                "current_version": current_version,
                "latest_version": current_version,
                "message": "Chưa có bản phát hành nào trên GitHub."
            }
        return {
            "success": False,
            "has_update": False,
            "error": f"Lỗi phản hồi từ GitHub: HTTP {e.code}"
        }
    except Exception as e:
        return {
            "success": False,
            "has_update": False,
            "error": f"Không thể kết nối máy chủ GitHub ({type(e).__name__}): {str(e)}"
        }


class CheckUpdateThread(QThread):
    """Luồng chạy ngầm kiểm tra cập nhật để không làm treo giao diện chính."""
    result_ready = pyqtSignal(dict)

    def __init__(self, repo: str = GITHUB_REPO, current_version: str = APP_VERSION, timeout: int = 6):
        super().__init__()
        self.repo = repo
        self.current_version = current_version
        self.timeout = timeout

    def run(self):
        info = check_for_updates(self.repo, self.current_version, self.timeout)
        self.result_ready.emit(info)


class DownloadUpdateThread(QThread):
    """Luồng tải file cập nhật (.zip) kèm báo tiến trình theo phần trăm."""
    progress = pyqtSignal(int, int)  # (bytes_downloaded, total_bytes)
    finished = pyqtSignal(str)       # local_file_path
    error = pyqtSignal(str)          # error_message

    def __init__(self, download_url: str, dest_dir: Optional[str] = None):
        super().__init__()
        self.download_url = download_url
        self.dest_dir = dest_dir or tempfile.gettempdir()
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        try:
            req = urllib.request.Request(self.download_url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=30) as resp:
                total_size = int(resp.headers.get("content-length", 0))
                filename = os.path.basename(self.download_url.split("?")[0]) or "update.zip"
                target_path = os.path.join(self.dest_dir, filename)

                downloaded = 0
                chunk_size = 64 * 1024  # 64 KB

                with open(target_path, "wb") as f:
                    while True:
                        if self._is_cancelled:
                            return
                        chunk = resp.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        self.progress.emit(downloaded, total_size)

                self.finished.emit(target_path)
        except Exception as e:
            self.error.emit(f"Lỗi khi tải bản cập nhật: {str(e)}")


def apply_update_and_restart(zip_path: str) -> Tuple[bool, str]:
    """Giải nén bản cập nhật và khởi chạy batch script tự động cập nhật.
    
    Quy trình:
    1. Kiểm tra ứng dụng có đang chạy dưới dạng PyInstaller exe (frozen) hay không.
    2. Giải nén zip vào thư mục tạm `AutoFormatter_Staging`.
    3. Tạo script `run_auto_update.bat`.
    4. Khởi chạy script ngầm và thoát ứng dụng hiện tại để tránh xung đột khoá tệp Windows.
    """
    is_frozen = getattr(sys, "frozen", False)
    if not is_frozen:
        return False, "Ứng dụng đang chạy ở chế độ mã nguồn phát triển (Python Script). Vui lòng dùng git pull để cập nhật."

    import zipfile
    app_dir = os.path.dirname(sys.executable)
    staging_dir = os.path.join(tempfile.gettempdir(), "AutoFormatter_Update_Staging")

    try:
        if os.path.exists(staging_dir):
            shutil.rmtree(staging_dir, ignore_errors=True)
        os.makedirs(staging_dir, exist_ok=True)

        # Giải nén
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(staging_dir)

        # Nếu tệp zip chứa thư mục con 'AutoFormatter'
        extracted_root = staging_dir
        subdirs = [os.path.join(staging_dir, d) for d in os.listdir(staging_dir) if os.path.isdir(os.path.join(staging_dir, d))]
        for d in subdirs:
            if os.path.exists(os.path.join(d, "AutoFormatter.exe")):
                extracted_root = d
                break

        pid = os.getpid()
        exe_path = sys.executable
        bat_path = os.path.join(tempfile.gettempdir(), "apply_autoformatter_update.bat")

        # Viết kịch bản batch chờ tiến trình hiện tại tắt, copy đè và bật lại
        bat_content = f"""@echo off
chcp 65001 > nul
echo Dang cap nhat Auto-Formatter len phien ban moi...
echo Vui long cho trong giay lat...

:: Cho tien trinh ung dung cu ket thuc
:wait_process
timeout /t 1 /nobreak > nul
tasklist /fi "pid eq {pid}" | find "{pid}" > nul
if not errorlevel 1 goto wait_process

:: Cho them 1s de giai phong hoan toan khoa tap tin tren Windows
timeout /t 1 /nobreak > nul

:: Ghi de cac tep tin moi vao thu muc cai dat ung dung
xcopy /s /e /y /q "{extracted_root}\\*" "{app_dir}\\"

:: Khoi dong lai ung dung moi
start "" "{exe_path}"

:: Don dep thu muc giai nen va tu xoa ban than
rmdir /s /q "{staging_dir}"
del "{zip_path}" > nul 2>&1
(goto) 2>nul & del "%~f0"
"""

        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(bat_content)

        # Chạy batch script độc lập ở chế độ detached
        CREATE_NO_WINDOW = 0x08000000
        DETACHED_PROCESS = 0x00000008
        subprocess.Popen(
            ["cmd.exe", "/c", bat_path],
            creationflags=CREATE_NO_WINDOW | DETACHED_PROCESS,
            close_fds=True
        )

        return True, "Kịch bản cập nhật đã được khởi chạy thành công."
    except Exception as e:
        return False, f"Không thể áp dụng bản cập nhật: {str(e)}"
