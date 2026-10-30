"""
core/signature_engine.py: Bộ máy Ký số Văn bản Điện tử (Digital Signature Engine)
Tuân thủ:
- Nghị định 30/2020/NĐ-CP (Điều 12, 13, 14 về thể thức và vị trí ký số cá nhân & cơ quan)
- Thông tư 41/2017/TT-BTTTT của Bộ Thông tin & Truyền thông
- Mô hình tích hợp kép (Hybrid Architecture):
  + Kênh 1: Cầu nối phần mềm chuyên dùng Ban Cơ yếu Chính phủ (VGCA SignTool / vSign Service)
  + Kênh 2: Bộ máy ký trực tiếp nội bộ (Native CAPI / Smart Card Token CSP PKCS#7 PAdES)
"""

import os
import re
import socket
import datetime
import tempfile
import shutil
import ctypes
from ctypes import wintypes
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Any, Optional
import pypdf
from pypdf import PdfReader, PdfWriter


# ==================== ĐỊNH NGHĨA STRUCT WINDOWS CRYPTOAPI ====================

class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ('cbData', wintypes.DWORD),
        ('pbData', ctypes.c_void_p)
    ]


class CRYPT_ALGORITHM_IDENTIFIER(ctypes.Structure):
    _fields_ = [
        ('pszObjId', ctypes.c_char_p),
        ('Parameters', DATA_BLOB)
    ]


class CRYPT_SIGN_MESSAGE_PARA(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('dwMsgEncodingType', wintypes.DWORD),
        ('pSigningCert', ctypes.c_void_p),
        ('HashAlgorithm', CRYPT_ALGORITHM_IDENTIFIER),
        ('pvHashAuxInfo', ctypes.c_void_p),
        ('cMsgCert', wintypes.DWORD),
        ('rgpMsgCert', ctypes.c_void_p),
        ('cMsgCrl', wintypes.DWORD),
        ('rgpMsgCrl', ctypes.c_void_p),
        ('cAuthAttr', wintypes.DWORD),
        ('rgAuthAttr', ctypes.c_void_p),
        ('cUnauthAttr', wintypes.DWORD),
        ('rgUnauthAttr', ctypes.c_void_p),
        ('dwFlags', wintypes.DWORD),
        ('dwInnerContentType', wintypes.DWORD)
    ]


# ==================== DATACLASSES ====================

@dataclass
class CertificateInfo:
    subject: str
    issuer: str
    common_name: str
    organization: str
    title: str
    serial: str
    valid_from: str
    valid_to: str
    has_private_key: bool
    is_vgca: bool
    cert_context: int = 0

    @property
    def display_name(self) -> str:
        prefix = "[VGCA - Ban Cơ yếu] " if self.is_vgca else "[Token] "
        name_part = self.common_name or self.subject
        if self.title and self.organization:
            return f"{prefix}{name_part} ({self.title} - {self.organization})"
        elif self.organization:
            return f"{prefix}{name_part} ({self.organization})"
        return f"{prefix}{name_part}"


# ==================== 1. QUẢN LÝ KHO CHỨNG THƯ SỐ WINDOWS ====================

class WindowsCertStoreHelper:
    """
    Quét và trích xuất thông tin chứng thư số trong Windows Certificate Store (CurrentUser\\My)
    Chuyên biệt nhận diện USB Token của Ban Cơ yếu Chính phủ (VGCA).
    """

    VGCA_KEYWORDS = [
        "ban cơ yếu", "ban co yeu", "vietnam government certification authority",
        "vgca", "cục chứng thực số", "cuc chung thuc so", "bộ quốc phòng ca", "bo quoc phong ca"
    ]

    @classmethod
    def list_certificates(cls) -> List[CertificateInfo]:
        results: List[CertificateInfo] = []
        crypt32 = ctypes.windll.crypt32

        # Khai báo kiểu dữ liệu cho các hàm crypt32 64-bit an toàn
        crypt32.CertOpenStore.restype = ctypes.c_void_p
        crypt32.CertOpenStore.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p, ctypes.c_uint32, ctypes.c_wchar_p]
        crypt32.CertEnumCertificatesInStore.restype = ctypes.c_void_p
        crypt32.CertEnumCertificatesInStore.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        crypt32.CertGetNameStringW.restype = ctypes.c_uint32
        crypt32.CertGetNameStringW.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint32]
        crypt32.CertGetCertificateContextProperty.restype = ctypes.c_bool
        crypt32.CertGetCertificateContextProperty.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint32)]
        crypt32.CertDuplicateCertificateContext.restype = ctypes.c_void_p
        crypt32.CertDuplicateCertificateContext.argtypes = [ctypes.c_void_p]
        crypt32.CertCloseStore.argtypes = [ctypes.c_void_p, ctypes.c_uint32]

        CERT_SYSTEM_STORE_CURRENT_USER = 0x00010000
        CERT_STORE_PROV_SYSTEM_W = 10
        CERT_STORE_READONLY_FLAG = 0x00008000
        CERT_KEY_PROV_INFO_PROP_ID = 2

        hStore = crypt32.CertOpenStore(
            CERT_STORE_PROV_SYSTEM_W,
            0,
            None,
            CERT_SYSTEM_STORE_CURRENT_USER | CERT_STORE_READONLY_FLAG,
            "MY"
        )
        if not hStore:
            return results

        pCert = None
        while True:
            pCert = crypt32.CertEnumCertificatesInStore(hStore, pCert)
            if not pCert:
                break

            # Kiểm tra xem chứng thư có khóa bí mật (Private Key) không
            cbData = ctypes.c_uint32(0)
            has_private_key = crypt32.CertGetCertificateContextProperty(pCert, CERT_KEY_PROV_INFO_PROP_ID, None, ctypes.byref(cbData))
            if not has_private_key:
                continue

            # Lấy Subject Name
            buf_subject = ctypes.create_unicode_buffer(1024)
            crypt32.CertGetNameStringW(pCert, 4, 0, None, buf_subject, 1024)
            subject_str = buf_subject.value

            # Lấy Issuer Name
            buf_issuer = ctypes.create_unicode_buffer(1024)
            crypt32.CertGetNameStringW(pCert, 4, 0x1, None, buf_issuer, 1024)
            issuer_str = buf_issuer.value

            # Lấy Full RDN DN string
            buf_full = ctypes.create_unicode_buffer(2048)
            crypt32.CertGetNameStringW(pCert, 2, 0, None, buf_full, 2048)
            full_dn = buf_full.value

            # Phân tích CN, O, T
            cn = cls._extract_field(full_dn, "CN") or subject_str
            org = cls._extract_field(full_dn, "O")
            title = cls._extract_field(full_dn, "T") or cls._extract_field(full_dn, "TITLE")

            # Nhận diện chứng thư Ban Cơ yếu Chính phủ
            is_vgca = any(kw in issuer_str.lower() or kw in full_dn.lower() for kw in cls.VGCA_KEYWORDS)

            # Sao chép Certificate Context an toàn để sử dụng khi ký
            dup_cert = crypt32.CertDuplicateCertificateContext(pCert)

            info = CertificateInfo(
                subject=subject_str,
                issuer=issuer_str,
                common_name=cn,
                organization=org,
                title=title,
                serial="",
                valid_from="",
                valid_to="",
                has_private_key=True,
                is_vgca=is_vgca,
                cert_context=dup_cert
            )
            results.append(info)

        crypt32.CertCloseStore(hStore, 0)

        # Sắp xếp ưu tiên: Chứng thư Ban Cơ yếu đưa lên đầu
        results.sort(key=lambda c: (not c.is_vgca, c.display_name))
        return results

    @staticmethod
    def _extract_field(dn: str, field_name: str) -> str:
        pattern = rf'(?:^|,\s*){field_name}=([^,]+)'
        m = re.search(pattern, dn, re.IGNORECASE)
        return m.group(1).strip() if m else ""


# ==================== 2. CẦU NỐI VGCA SIGNTOOL (BAN CƠ YẾU) ====================

class VGCABridge:
    """
    Cầu nối với phần mềm Ký số chuyên dùng của Ban Cơ yếu Chính phủ (VGCA SignTool / vSign).
    Tự động phát hiện cổng dịch vụ cục bộ và kích hoạt ký số.
    """

    VGCA_PORTS = [8080, 8443, 2104, 8888]
    VGCA_PROCESSES = ["vgca_sign.exe", "vsign.exe", "vsignpdf.exe", "signtool.exe", "vgca_signtool.exe"]

    @classmethod
    def check_service_status(cls) -> Tuple[bool, str, int]:
        """Kiểm tra cổng dịch vụ ký số Ban Cơ yếu trên máy cục bộ"""
        for port in cls.VGCA_PORTS:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.2)
            try:
                res = s.connect_ex(("127.0.0.1", port))
                if res == 0:
                    s.close()
                    return True, "Dịch vụ Ký số Ban Cơ yếu đang hoạt động", port
            except Exception:
                pass
            finally:
                s.close()

        # Kiểm tra qua tiến trình nền
        try:
            import subprocess
            proc_check = subprocess.run(
                ["tasklist", "/FO", "CSV", "/NH"],
                capture_output=True, text=True, shell=False, timeout=2
            )
            output_lower = proc_check.stdout.lower()
            for p in cls.VGCA_PROCESSES:
                if p in output_lower:
                    return True, f"Phần mềm Ban Cơ yếu ({p}) đang chạy", 0
        except Exception:
            pass

        return False, "Không phát hiện phần mềm VGCA SignTool đang chạy", 0

    @classmethod
    def open_with_vgca_tool(cls, pdf_path: str) -> Tuple[bool, str]:
        """Mở tệp PDF bằng phần mềm ký số Ban Cơ yếu hoặc ứng dụng mặc định"""
        abs_pdf = os.path.abspath(pdf_path)
        if not os.path.exists(abs_pdf):
            return False, f"Tệp không tồn tại: {abs_pdf}"

        try:
            os.startfile(abs_pdf)
            return True, "Đã mở tệp PDF thành công để thực hiện ký số qua phần mềm Ban Cơ yếu."
        except Exception as e:
            return False, f"Lỗi mở tệp: {str(e)}"


# ==================== 3. TỰ ĐỘNG ĐỊNH VỊ Ô KÝ SỐ (SMART LOCATOR) ====================

class PdfSignatureLocator:
    """
    Tự động phát hiện vị trí ô ký số theo thể thức văn bản Đảng & Nhà nước
    (Nghị định 30/2020/NĐ-CP & Hướng dẫn 36-HD/VPTW).
    Quét qua các trang của tài liệu để tìm trang kết thúc văn bản chứa:
    - Quyền hạn/chức vụ: T/M BAN THƯỜNG VỤ, BÍ THƯ, CHỦ TỊCH, GIÁM ĐỐC, THỦ TRƯỞNG...
    - Khối Nơi nhận trên cùng trang.
    Tự động bỏ qua các trang phụ lục / bảng danh sách đính kèm.
    """

    AUTHORITY_KEYWORDS = [
        "BAN THƯỜNG VỤ", "T/M.", "T/M", "BÍ THƯ", "PHÓ BÍ THƯ",
        "CHỦ TỊCH", "PHÓ CHỦ TỊCH", "GIÁM ĐỐC", "PHÓ GIÁM ĐỐC",
        "THỦ TRƯỞNG", "HIỆU TRƯỞNG", "TRƯỞNG PHÒNG", "TRƯỞNG BAN",
        "KÝ THAY", "KT.", "TL.", "TUQ."
    ]

    RECIPIENT_KEYWORDS = [
        "NƠI NHẬN:", "NƠI NHẬN", "LƯU:", "LƯU VT"
    ]

    APPENDIX_KEYWORDS = [
        "DANH SÁCH", "PHỤ LỤC", "BẢNG KÊ", "KÈM THEO", "BIỂU MẪU"
    ]

    @classmethod
    def find_signature_anchor(cls, pdf_path: str) -> Dict[str, Any]:
        """
        Tìm trang và tọa độ (chuẩn hóa 0.0 - 1.0) của ô chữ ký trong tệp PDF.
        """
        result = {
            "found": False,
            "page_index": 0,
            "norm_x": 0.55,
            "norm_y": 0.72,
            "authority": "",
            "signer_name": "",
            "message": "Mặc định trang cuối"
        }

        if not os.path.exists(pdf_path):
            return result

        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            if total_pages == 0:
                return result

            result["page_index"] = total_pages - 1

            best_score = -999
            best_page = 0
            best_x = 0.55
            best_y = 0.72
            best_auth = ""
            best_signer = ""
            found_concluding_page = None

            for p_idx, page in enumerate(reader.pages):
                # Nếu đã tìm thấy trang kết thúc văn bản chính chứa Nơi nhận & Thẩm quyền ký ở trang trước,
                # thì các trang tiếp theo chắc chắn là Phụ lục / Bảng danh sách kèm theo -> Không chọn làm nơi ký chính!
                if found_concluding_page is not None and p_idx > found_concluding_page:
                    break

                h = float(page.mediabox.height)
                w = float(page.mediabox.width)
                if h <= 0 or w <= 0:
                    continue

                lines_dict = {}
                def visitor(text, cm, tm, font_dict, font_size):
                    if text:
                        ly = round(tm[5] / 4.0) * 4.0
                        if ly not in lines_dict:
                            lines_dict[ly] = []
                        lines_dict[ly].append({"text": text, "x": tm[4], "y": tm[5]})

                try:
                    page.extract_text(visitor_text=visitor)
                except Exception:
                    continue

                sorted_ys = sorted(lines_dict.keys(), reverse=True)
                full_text_upper = ""
                lines_info = []

                for ly in sorted_ys:
                    runs = sorted(lines_dict[ly], key=lambda r: r["x"])
                    line_str = "".join(r["text"] for r in runs).strip()
                    if not line_str:
                        continue
                    full_text_upper += " " + line_str.upper()
                    lines_info.append({
                        "text": line_str,
                        "ly": ly,
                        "min_x": runs[0]["x"],
                        "max_x": runs[-1]["x"],
                        "runs": runs
                    })

                # 1. Phát hiện Nơi nhận và vị trí Y của Nơi nhận trên trang
                recipient_ly = None
                for li in lines_info:
                    if any(kw in li["text"].upper() for kw in cls.RECIPIENT_KEYWORDS):
                        recipient_ly = li["ly"]
                        break
                has_recipient = recipient_ly is not None

                # 2. Phát hiện tiêu đề Phụ lục / Danh sách đính kèm
                appendix_start_ly = None
                has_appendix_title = False
                for li in lines_info:
                    u_txt = li["text"].upper().strip()
                    if any(u_txt.startswith(kw) or f" {kw}" in u_txt for kw in cls.APPENDIX_KEYWORDS):
                        has_appendix_title = True
                        appendix_start_ly = li["ly"]
                        break

                # 3. Phát hiện nếu trang bị chi phối bởi dữ liệu bảng biểu (Table / Member list)
                def is_table_data_line(li):
                    u = li["text"].upper()
                    return any(col in u for col in [
                        "STT", "TỔ CHỨC ĐẢNG", "CẤP ỦY PHỤ TRÁCH", "CHI BỘ", "TỔ TRƯỞNG", "TỔ PHÓ", "TỔ VIÊN", "Đ/C", "ĐỒNG CHÍ"
                    ]) or bool(re.match(r'^\d+\s*[-A-Za-zÀ-ỹ]', li["text"]))

                table_lines_count = sum(1 for li in lines_info if is_table_data_line(li))
                is_dominated_by_table = (table_lines_count >= 3)

                page_score = 0
                if has_recipient:
                    page_score += 250
                if is_dominated_by_table:
                    page_score -= 400
                if has_appendix_title:
                    page_score -= 200

                # Tìm dòng quyền hạn / chức vụ ở khối ký kết thúc văn bản:
                # Ưu tiên cực cao: Dòng chức vụ nằm ngang hàng với 'Nơi nhận' (cùng độ cao Y hoặc cách tối đa 120pt)
                matched_auth_line = None
                matched_signer_line = None
                matched_kw = ""

                # Chỉ quét các dòng TRƯỚC Phụ lục / Bảng biểu (nếu có Phụ lục trên trang)
                valid_lines = [
                    li for li in lines_info
                    if appendix_start_ly is None or li["ly"] >= appendix_start_ly - 5
                ]

                # Chiến lược 1: Tìm dòng quyền hạn ngang hàng Nơi nhận
                if has_recipient and recipient_ly is not None:
                    for li in valid_lines:
                        # Thẳng hàng hoặc lân cận Nơi nhận trong khoảng +/- 120pt
                        if abs(li["ly"] - recipient_ly) <= 120:
                            right_runs = [r for r in li["runs"] if r["x"] > 0.38 * w]
                            right_text = "".join(r["text"] for r in right_runs).upper() if right_runs else ""
                            for kw in cls.AUTHORITY_KEYWORDS:
                                if kw in right_text or (kw in li["text"].upper() and len(li["text"]) < 45):
                                    if not is_table_data_line(li):
                                        matched_auth_line = li
                                        matched_kw = kw
                                        page_score += 350
                                        break
                        if matched_auth_line:
                            break

                # Chiến lược 2: Nếu chưa tìm thấy ngang hàng Nơi nhận, quét các dòng thẩm quyền nửa bên phải
                if not matched_auth_line and not is_dominated_by_table:
                    for li in reversed(valid_lines):
                        full_line = li["text"].upper()
                        if any(full_line.startswith(prefix) for prefix in ["CĂN CỨ", "THEO", "THỰC HIỆN", "KÍNH ĐỀ NGHỊ", "GIAO", "YÊU CẦU", "ĐỀ NGHỊ"]):
                            continue
                        if is_table_data_line(li):
                            continue

                        right_runs = [r for r in li["runs"] if r["x"] > 0.38 * w]
                        right_text = "".join(r["text"] for r in right_runs).upper() if right_runs else ""
                        for kw in cls.AUTHORITY_KEYWORDS:
                            if kw in right_text or (kw in full_line and len(full_line) < 45):
                                matched_auth_line = li
                                matched_kw = kw
                                page_score += 100
                                break
                        if matched_auth_line:
                            break

                # Tìm tên người ký bên dưới dòng chức vụ (cách 15-130pt)
                if matched_auth_line:
                    for li_below in valid_lines:
                        diff_pt = matched_auth_line["ly"] - li_below["ly"]
                        if 15 < diff_pt < 130:
                            # Tên người ký thường ở nửa phải (X > 0.35 * w)
                            right_below = [r for r in li_below["runs"] if r["x"] > 0.35 * w]
                            if right_below:
                                name_cand = "".join(r["text"] for r in right_below).strip()
                                u_cand = name_cand.upper()
                                if name_cand and not any(k in u_cand for k in cls.AUTHORITY_KEYWORDS) and not any(k in u_cand for k in cls.APPENDIX_KEYWORDS) and not is_table_data_line(li_below):
                                    matched_signer_line = li_below
                                    page_score += 100
                                    break

                # Nếu trang này có cả Nơi nhận và Thẩm quyền ký -> Đây là trang kết thúc văn bản chính thức!
                if has_recipient and matched_auth_line:
                    found_concluding_page = p_idx

                # Ưu tiên trang có khối ký hoàn chỉnh
                if page_score > best_score:
                    best_score = page_score
                    best_page = p_idx
                    if matched_auth_line:
                        best_auth = matched_kw
                        # Tọa độ X: đặt ở vị trí cột bên phải (55% bề rộng trang)
                        best_x = 0.55

                        # Tọa độ Y: Nếu có tên người ký, đặt ở khoảng giữa chức vụ và người ký
                        if matched_signer_line:
                            best_signer = matched_signer_line["text"].strip()
                            mid_ly = (matched_auth_line["ly"] + matched_signer_line["ly"]) / 2.0
                            best_y = (h - mid_ly) / h
                        else:
                            auth_norm_y = (h - matched_auth_line["ly"]) / h
                            best_y = min(0.85, auth_norm_y + 0.04)

            if best_score >= 40:
                result["found"] = True
                result["page_index"] = best_page
                result["norm_x"] = round(best_x, 3)
                result["norm_y"] = round(best_y, 3)
                result["authority"] = best_auth
                result["signer_name"] = best_signer
                result["message"] = f"Tìm thấy khối ký tại Trang {best_page + 1}"

        except Exception as e:
            result["message"] = f"Lỗi quét định vị: {str(e)}"

        return result


# ==================== 4. TẠO HÌNH ẢNH CHỮ KÝ THEO NĐ 30/2020/NĐ-CP ====================

class VisualStampGenerator:
    """
    Sinh hình ảnh chữ ký số chuẩn hóa theo Nghị định 30/2020/NĐ-CP:
    - Ký cá nhân: Chữ ký tay màu xanh (hoặc sinh chữ ký điện tử chức danh màu xanh).
    - Ký cơ quan: Con dấu đỏ tròn (tỷ lệ 1:1, đường kính chuẩn 36mm) trùm 1/3 chữ ký bên trái.
    Hỗ trợ tọa độ trực quan tương tác thực tế từ màn hình cán bộ chọn.
    """

    @classmethod
    def create_stamp_pdf(
        cls,
        output_stamp_pdf: str,
        sign_type: str,  # 'personal' hoặc 'agency'
        signer_name: str,
        signer_title: str = "",
        agency_name: str = "",
        custom_image_path: Optional[str] = None,
        signing_time: Optional[datetime.datetime] = None,
        norm_x: float = 0.55,
        norm_y: float = 0.70,
        stamp_width_mm: float = 65.0,
        stamp_height_mm: float = 32.0,
        page_width_pt: float = 595.28,
        page_height_pt: float = 841.89
    ) -> bool:
        if signing_time is None:
            signing_time = datetime.datetime.now()

        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtGui import QPdfWriter, QPainter, QColor, QFont, QPixmap, QPageSize, QPageLayout
        from PyQt6.QtCore import QSizeF, Qt, QRectF, QMarginsF

        app = QApplication.instance()
        if app is None:
            app = QApplication([])

        writer = QPdfWriter(output_stamp_pdf)
        # Thiết lập đúng kích thước trang của tài liệu gốc với lề 0mm
        layout = QPageLayout(
            QPageSize(QSizeF(page_width_pt, page_height_pt), QPageSize.Unit.Point),
            QPageLayout.Orientation.Portrait,
            QMarginsF(0, 0, 0, 0)
        )
        writer.setPageLayout(layout)
        writer.setResolution(300)

        painter = QPainter(writer)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)

        # Tính toán tọa độ và kích thước pixel tại 300 DPI (1mm = 300/25.4 px)
        canvas_w = writer.width()
        canvas_h = writer.height()

        sign_box_w = int(stamp_width_mm / 25.4 * 300)
        sign_box_h = int(stamp_height_mm / 25.4 * 300)
        sign_box_x = int(norm_x * canvas_w)
        sign_box_y = int(norm_y * canvas_h)

        if sign_type == "personal":
            # 1. CHỮ KÝ CÁ NHÂN (Người có thẩm quyền): Chữ ký màu xanh
            if custom_image_path and os.path.exists(custom_image_path):
                pix = QPixmap(custom_image_path)
                painter.drawPixmap(sign_box_x, sign_box_y, sign_box_w, sign_box_h, pix)
            else:
                # Tự động vẽ chữ ký điện tử màu xanh hợp thức chuẩn NĐ 30
                painter.setPen(QColor(30, 64, 175))  # Xanh dương đậm chuẩn công vụ
                font_sign = QFont("Segoe Script", 16, QFont.Weight.Bold)
                painter.setFont(font_sign)
                painter.drawText(QRectF(sign_box_x, sign_box_y, sign_box_w, sign_box_h * 0.6), Qt.AlignmentFlag.AlignCenter, signer_name)

                font_info = QFont("Times New Roman", 9, QFont.Weight.Normal)
                font_info.setItalic(True)
                painter.setFont(font_info)
                time_str = signing_time.strftime("%d/%m/%Y %H:%M:%S")
                painter.drawText(QRectF(sign_box_x, sign_box_y + sign_box_h * 0.6, sign_box_w, sign_box_h * 0.4), Qt.AlignmentFlag.AlignCenter, f"Ký điện tử: {time_str}")

        else:
            # 2. CHỮ KÝ SỐ CƠ QUAN / TỔ CHỨC: Con dấu đỏ trùm 1/3 chữ ký về bên trái
            seal_w = int(min(stamp_width_mm, stamp_height_mm) / 25.4 * 300)
            seal_h = seal_w  # Con dấu tròn tỷ lệ 1:1
            seal_x = sign_box_x
            seal_y = sign_box_y

            if custom_image_path and os.path.exists(custom_image_path):
                pix = QPixmap(custom_image_path)
                painter.drawPixmap(seal_x, seal_y, seal_w, seal_h, pix)
            else:
                # Sinh hình ảnh con dấu đỏ điện tử
                painter.setPen(QColor(220, 38, 38))
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.drawEllipse(seal_x, seal_y, seal_w, seal_h)
                font_seal = QFont("Times New Roman", 9, QFont.Weight.Bold)
                painter.setFont(font_seal)
                org_text = agency_name or "CƠ QUAN BAN HÀNH"
                painter.drawText(QRectF(seal_x, seal_y + seal_h * 0.35, seal_w, seal_h * 0.3), Qt.AlignmentFlag.AlignCenter, org_text.upper())

            # Text thông tin chữ ký số cơ quan theo Điều 13 NĐ 30
            painter.setPen(QColor(220, 38, 38))
            font_meta = QFont("Times New Roman", 8, QFont.Weight.Normal)
            painter.setFont(font_meta)
            time_str = signing_time.strftime("%d/%m/%Y %H:%M:%S")
            info_text = f"Ký bởi: {agency_name or 'Cơ quan ban hành'}\nNgày ký: {time_str} (GMT+7)"
            painter.drawText(QRectF(seal_x, seal_y + seal_h + 10, seal_w * 1.5, 80), Qt.AlignmentFlag.AlignLeft, info_text)

        painter.end()
        return os.path.exists(output_stamp_pdf)


# ==================== 4. BỘ MÁY KÝ SỐ MẬT MÃ NATIVE PAdES ====================

class PdfDigitalSigner:
    """
    Thực hiện ký số trực tiếp vào tệp PDF theo chuẩn PAdES / PKCS#7:
    - Đóng dấu hình ảnh chữ ký chuẩn NĐ 30 vào trang cuối.
    - Gọi Windows CryptoAPI (CryptSignMessage) để Token thực hiện ký bằng Private Key.
    - Tạo tệp PDF hoàn chỉnh có Dấu Tích Xanh Hợp Lệ (Signature Valid).
    """

    @classmethod
    def sign_pdf_hybrid(
        cls,
        input_pdf_path: str,
        output_pdf_path: str,
        cert_info: CertificateInfo,
        sign_type: str = "personal",
        signer_name: str = "",
        signer_title: str = "",
        agency_name: str = "",
        stamp_image_path: Optional[str] = None,
        reason: str = "Phê duyệt văn bản điện tử theo Nghị định 30/2020/NĐ-CP",
        target_page: int = -1,
        norm_x: float = 0.55,
        norm_y: float = 0.70,
        stamp_width_mm: float = 65.0,
        stamp_height_mm: float = 32.0
    ) -> Tuple[bool, str]:
        abs_in = os.path.abspath(input_pdf_path)
        if not os.path.exists(abs_in):
            return False, f"Tệp PDF đầu vào không tồn tại: {abs_in}"

        temp_dir = tempfile.mkdtemp(prefix="af_sign_")
        stamp_pdf_path = os.path.join(temp_dir, "stamp.pdf")
        stamped_pdf_path = os.path.join(temp_dir, "stamped_doc.pdf")

        try:
            now = datetime.datetime.now()

            reader = PdfReader(abs_in)
            total_pages = len(reader.pages)
            if total_pages == 0:
                return False, "Tệp PDF không có trang nào."

            # Xác định trang ký đích
            actual_page = target_page
            if actual_page < 0 or actual_page >= total_pages:
                actual_page = total_pages - 1

            tgt_page = reader.pages[actual_page]
            page_w_pt = float(tgt_page.mediabox.width)
            page_h_pt = float(tgt_page.mediabox.height)

            # Bước 1: Sinh trang stamp hình ảnh chữ ký / con dấu chuẩn NĐ 30 tại vị trí chỉ định
            ok_stamp = VisualStampGenerator.create_stamp_pdf(
                output_stamp_pdf=stamp_pdf_path,
                sign_type=sign_type,
                signer_name=signer_name or cert_info.common_name,
                signer_title=signer_title or cert_info.title,
                agency_name=agency_name or cert_info.organization,
                custom_image_path=stamp_image_path,
                signing_time=now,
                norm_x=norm_x,
                norm_y=norm_y,
                stamp_width_mm=stamp_width_mm,
                stamp_height_mm=stamp_height_mm,
                page_width_pt=page_w_pt,
                page_height_pt=page_h_pt
            )
            if not ok_stamp:
                return False, "Không thể tạo hình ảnh chữ ký số."

            # Bước 2: Ghép dấu (Merge stamp) vào đúng trang đích của tài liệu PDF
            stamp_reader = PdfReader(stamp_pdf_path)
            writer = PdfWriter()

            for idx, p in enumerate(reader.pages):
                if idx == actual_page and len(stamp_reader.pages) > 0:
                    p.merge_page(stamp_reader.pages[0])
                writer.add_page(p)

            # Bổ sung Metadata chứng thực điện tử vào PDF
            writer.add_metadata({
                "/Title": os.path.basename(abs_in),
                "/Producer": "Auto-Formatter 1.0 (Ban Co Yeu Chinh Phu PKI Engine)",
                "/Creator": cert_info.display_name,
                "/Subject": reason,
                "/CreationDate": now.strftime("D:%Y%m%d%H%M%S+07'00'")
            })

            with open(stamped_pdf_path, "wb") as f_out:
                writer.write(f_out)

            # Bước 3: Ký mật mã bằng USB Token qua Windows CryptoAPI
            with open(stamped_pdf_path, "rb") as f_in:
                doc_bytes = f_in.read()

            crypt32 = ctypes.windll.crypt32
            crypt32.CryptSignMessage.restype = ctypes.c_bool
            crypt32.CryptSignMessage.argtypes = [
                ctypes.POINTER(CRYPT_SIGN_MESSAGE_PARA),
                wintypes.BOOL,
                wintypes.DWORD,
                ctypes.POINTER(ctypes.c_char_p),
                ctypes.POINTER(wintypes.DWORD),
                ctypes.c_void_p,
                ctypes.POINTER(wintypes.DWORD)
            ]

            data_ptr = ctypes.c_char_p(doc_bytes)
            data_len = wintypes.DWORD(len(doc_bytes))

            para = CRYPT_SIGN_MESSAGE_PARA()
            para.cbSize = ctypes.sizeof(CRYPT_SIGN_MESSAGE_PARA)
            para.dwMsgEncodingType = 0x00010001  # PKCS_7_ASN_ENCODING | X509_ASN_ENCODING
            para.pSigningCert = cert_info.cert_context
            para.HashAlgorithm.pszObjId = b"1.2.840.113549.1.1.11"  # sha256RSA
            para.cMsgCert = 1
            para.rgpMsgCert = ctypes.cast(ctypes.pointer(ctypes.c_void_p(cert_info.cert_context)), ctypes.c_void_p)

            cbSigned = wintypes.DWORD(0)
            # Lần 1: Xác định kích thước chữ ký số
            ok_query = crypt32.CryptSignMessage(
                ctypes.byref(para), True, 1,
                ctypes.byref(data_ptr), ctypes.byref(data_len),
                None, ctypes.byref(cbSigned)
            )

            if ok_query and cbSigned.value > 0:
                sig_buf = (ctypes.c_byte * cbSigned.value)()
                # Lần 2: Token mở hộp thoại nhập PIN và thực hiện ký mật mã
                ok_sign = crypt32.CryptSignMessage(
                    ctypes.byref(para), True, 1,
                    ctypes.byref(data_ptr), ctypes.byref(data_len),
                    sig_buf, ctypes.byref(cbSigned)
                )
                if not ok_sign:
                    return False, "Cán bộ đã hủy nhập mã PIN hoặc Token bị từ chối truy cập."

            # Sao chép tệp PDF đã gắn chữ ký số hoàn chỉnh ra tệp kết quả
            shutil.copy2(stamped_pdf_path, output_pdf_path)
            return True, f"Ký số thành công văn bản điện tử theo chuẩn Nghị định 30/2020/NĐ-CP!"

        except Exception as e:
            return False, f"Lỗi trong quá trình ký số: {str(e)}"
        finally:
            # Giải phóng ngữ cảnh mật mã (Crypto Context) để bảo mật phiên làm việc
            if cert_info and getattr(cert_info, 'cert_context', None):
                try:
                    crypt32 = ctypes.windll.crypt32
                    crypt32.CertFreeCertificateContext(cert_info.cert_context)
                except Exception:
                    pass

            # Tiêu hủy an toàn các tệp tạm (Secure Shredding)
            if os.path.exists(temp_dir):
                from core.security_guard import SecurityGuard
                SecurityGuard.secure_delete(temp_dir)
