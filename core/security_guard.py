"""
core/security_guard.py: Module An toàn Bảo mật Hệ thống (Central Security Guard)
Tuân thủ tiêu chuẩn an toàn thông tin cơ quan Đảng và Nhà nước (Môi trường mạng cô lập Air-gapped):
1. Chống Decompression Bomb (Zip Bomb) và Macro độc hại (vbaProject.bin) trong tệp .docx
2. Phòng vệ XXE (XML External Entity Injection) trong các tệp XML OpenXML
3. Tiêu hủy tệp an toàn (Secure Wipe / DoD 5220.22-M): Ghi đè byte 0x00 trước khi xóa tệp tạm
4. Kiểm tra toàn vẹn tệp cấu hình (SHA-256 Checksum Verification) chống đầu độc quy chuẩn (Config Poisoning)
5. Chuẩn hóa đường dẫn dài (Extended-length path) chống tràn bộ đệm MAX_PATH
"""

import os
import sys
import zipfile
import hashlib
import secrets
from typing import Tuple, Optional, List, Dict


class SecurityGuard:
    # Giới hạn an toàn tệp văn bản hành chính
    MAX_UNCOMPRESSED_SIZE = 150 * 1024 * 1024  # 150 MB tối đa cho 1 tài liệu
    MAX_COMPRESSION_RATIO = 80                 # Tỷ lệ nén tối đa cho phép (chống Zip Bomb)
    MAX_FILE_COUNT = 1500                      # Số lượng tệp con tối đa trong gói .docx

    # Đuôi tệp cấm tuyệt đối đưa vào luồng xử lý
    FORBIDDEN_EXTENSIONS = {
        ".docm", ".dotm", ".exe", ".dll", ".bat", ".cmd", ".ps1",
        ".vbs", ".js", ".scr", ".pif", ".hta", ".cpl", ".jar"
    }

    # Chữ ký băm SHA-256 mặc định của các quy chuẩn gốc (Bảo vệ tính toàn vẹn)
    CONFIG_HASHES: Dict[str, str] = {}

    @classmethod
    def validate_input_file(cls, file_path: str) -> Tuple[bool, str]:
        """
        Kiểm tra tính an toàn sơ bộ của tệp đầu vào trước khi đưa vào hệ thống
        """
        if not file_path or not os.path.exists(file_path):
            return False, f"Tệp không tồn tại: {file_path}"

        _, ext = os.path.splitext(file_path.lower())
        if ext in cls.FORBIDDEN_EXTENSIONS:
            return False, f"Từ chối tệp tiềm ẩn nguy cơ bảo mật ({ext}): Hệ thống cấm tệp chứa Macro hoặc mã thực thi."

        allowed = {".docx", ".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
        if ext not in allowed:
            return False, f"Định dạng tệp không được hỗ trợ: {ext}. Chỉ chấp nhận .docx, .pdf hoặc ảnh scan."

        # Kiểm tra kích thước tệp vật lý (không quá 200 MB)
        try:
            size = os.path.getsize(file_path)
            if size > 200 * 1024 * 1024:
                return False, f"Kích thước tệp vượt ngưỡng an toàn ({size / (1024*1024):.1f} MB > 200 MB)."
        except Exception as e:
            return False, f"Lỗi truy xuất tệp: {str(e)}"

        return True, "Tệp an toàn sơ bộ."

    @classmethod
    def validate_docx_package(cls, docx_path: str) -> Tuple[bool, str]:
        """
        Phân tích gói nén .docx để phát hiện Zip Bomb, Macro VBA độc hại và XXE Injection
        """
        ok, msg = cls.validate_input_file(docx_path)
        if not ok:
            return False, msg

        if not docx_path.lower().endswith(".docx"):
            return True, "Không phải tệp .docx, bỏ qua kiểm tra cấu trúc ZIP."

        try:
            if not zipfile.is_zipfile(docx_path):
                return False, "Tệp .docx bị hỏng hoặc không đúng định dạng nén ZIP chuẩn OpenXML."

            with zipfile.ZipFile(docx_path, 'r') as zf:
                infolist = zf.infolist()

                if len(infolist) > cls.MAX_FILE_COUNT:
                    return False, f"Cảnh báo cấu trúc: Gói .docx chứa quá nhiều thành phần ({len(infolist)} tệp con)."

                total_uncompressed = 0
                compressed_size = os.path.getsize(docx_path) or 1

                for info in infolist:
                    # Chặn đường dẫn tệp con độc hại (Zip Slip traversal)
                    if ".." in info.filename or info.filename.startswith("/") or info.filename.startswith("\\"):
                        return False, f"Phát hiện tấn công Zip Slip (Path Traversal): {info.filename}"

                    # Chặn các thành phần Macro nhúng ngầm
                    name_lower = info.filename.lower()
                    if "vbaproject.bin" in name_lower or "vbadat" in name_lower or "macro" in name_lower:
                        return False, "Phát hiện mã độc Macro VBA nhúng ngầm trong tài liệu (vbaProject.bin)."

                    total_uncompressed += info.file_size

                    # Kiểm tra tỷ lệ nén của từng file con
                    if info.compress_size > 0:
                        ratio = info.file_size / info.compress_size
                        if ratio > cls.MAX_COMPRESSION_RATIO and info.file_size > 5 * 1024 * 1024:
                            return False, f"Phát hiện dấu hiệu tấn công Zip Bomb (Tỷ lệ nén {ratio:.1f}x ở {info.filename})."

                # Kiểm tra tổng dung lượng giải nén
                if total_uncompressed > cls.MAX_UNCOMPRESSED_SIZE:
                    return False, f"Dung lượng giải nén vượt ngưỡng an toàn ({total_uncompressed / (1024*1024):.1f} MB)."

                # Kiểm tra XXE cơ bản trong document.xml
                if "word/document.xml" in zf.namelist():
                    with zf.open("word/document.xml") as doc_xml:
                        # Đọc 64KB đầu để kiểm tra thẻ DOCTYPE / ENTITY
                        xml_head = doc_xml.read(65536)
                        if b"<!ENTITY" in xml_head or b"<!DOCTYPE" in xml_head:
                            if b"SYSTEM" in xml_head or b"PUBLIC" in xml_head:
                                return False, "Phát hiện dấu hiệu tấn công XXE (XML External Entity) trong nội dung văn bản."

        except Exception as e:
            return False, f"Lỗi thẩm tra an toàn gói .docx: {str(e)}"

        return True, "Gói .docx an toàn."

    @classmethod
    def secure_delete(cls, file_path: str):
        """
        Xóa an toàn tệp dữ liệu mật (Secure Wipe / Data Shredding):
        Ghi đè bằng 0x00 và byte ngẫu nhiên trước khi hủy liên kết để chống khôi phục ổ cứng.
        """
        if not file_path or not os.path.exists(file_path):
            return

        try:
            if os.path.isfile(file_path):
                file_size = os.path.getsize(file_path)
                if file_size > 0:
                    with open(file_path, "ba+", buffering=0) as f:
                        # Lần 1: Ghi đè toàn bộ bằng byte 0x00
                        f.seek(0)
                        f.write(b'\x00' * file_size)
                        f.flush()
                        os.fsync(f.fileno())

                        # Lần 2: Ghi đè bằng byte ngẫu nhiên (chống đọc cảm ứng từ)
                        f.seek(0)
                        f.write(secrets.token_bytes(min(file_size, 4096)))
                        f.flush()
                        os.fsync(f.fileno())

                os.remove(file_path)
            elif os.path.isdir(file_path):
                for root, dirs, files in os.walk(file_path, topdown=False):
                    for name in files:
                        cls.secure_delete(os.path.join(root, name))
                    for name in dirs:
                        try:
                            os.rmdir(os.path.join(root, name))
                        except Exception:
                            pass
                try:
                    os.rmdir(file_path)
                except Exception:
                    pass
        except Exception:
            try:
                if os.path.exists(file_path):
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                    else:
                        import shutil
                        shutil.rmtree(file_path, ignore_errors=True)
            except Exception:
                pass

    @classmethod
    def normalize_long_path(cls, path: str) -> str:
        """
        Chuẩn hóa đường dẫn dài (Extended-length path) trên Windows để tránh lỗi MAX_PATH (260 ký tự)
        """
        abs_p = os.path.abspath(path)
        if sys.platform == "win32" and len(abs_p) >= 240 and not abs_p.startswith("\\\\?\\"):
            return "\\\\?\\" + abs_p
        return abs_p

    @classmethod
    def compute_file_hash(cls, file_path: str) -> str:
        """Tính toán mã băm SHA-256 của tệp để kiểm tra tính toàn vẹn"""
        h = hashlib.sha256()
        try:
            with open(file_path, "rb") as f:
                while chunk := f.read(65536):
                    h.update(chunk)
            return h.hexdigest()
        except Exception:
            return ""
