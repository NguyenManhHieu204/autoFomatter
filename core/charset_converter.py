"""
core/charset_converter.py: Bộ Chuyển đổi Bảng mã Tiếng Việt & Chuẩn hóa Unicode
Chức năng:
1. Tự động nhận diện và chuyển đổi bảng mã cũ TCVN3 (ABC / .VnTime) sang Unicode dựng sẵn (NFC)
2. Tự động nhận diện và chuyển đổi bảng mã cũ VNI-Windows (VNI-Times) sang Unicode dựng sẵn (NFC)
3. Chuẩn hóa triệt để Unicode tổ hợp (NFD) sang Unicode dựng sẵn (NFC) tuân thủ TCVN 6909:2001
"""

import unicodedata
import re
from typing import Tuple


class CharsetConverter:
    # Bảng ánh xạ ký tự TCVN3 (ABC) sang Unicode dựng sẵn
    TCVN3_TO_UNICODE = {
        0xB5: 'à', 0xB6: 'ả', 0xB7: 'ã', 0xB8: 'á', 0xB9: 'ạ',
        0xBA: 'ă', 0xBB: 'ằ', 0xBC: 'ẳ', 0xBD: 'ẵ', 0xBE: 'ắ', 0xC6: 'ặ',
        0xC7: 'â', 0xC8: 'ầ', 0xC9: 'ẩ', 0xCA: 'ẫ', 0xCB: 'ấ', 0xCC: 'ậ',
        0xD0: 'đ',
        0xCE: 'è', 0xCF: 'ẻ', 0xD1: 'ẽ', 0xD2: 'é', 0xD3: 'ẹ',
        0xD4: 'ê', 0xD5: 'ề', 0xD6: 'ể', 0xD7: 'ễ', 0xD8: 'ế', 0xDC: 'ệ',
        0xDD: 'ì', 0xDE: 'ỉ', 0xDF: 'ĩ', 0xE1: 'í', 0xE2: 'ị',
        0xE3: 'ò', 0xE4: 'ỏ', 0xE5: 'õ', 0xE6: 'ó', 0xE7: 'ọ',
        0xE8: 'ô', 0xE9: 'ồ', 0xEA: 'ổ', 0xEB: 'ỗ', 0xEC: 'ố', 0xED: 'ộ',
        0xEE: 'ơ', 0xEF: 'ờ', 0xF1: 'ở', 0xF2: 'ỡ', 0xF3: 'ớ', 0xF4: 'ợ',
        0xF5: 'ù', 0xF6: 'ủ', 0xF7: 'ũ', 0xF8: 'ú', 0xF9: 'ụ',
        0xFA: 'ư', 0xFB: 'ừ', 0xFC: 'ử', 0xFD: 'ữ', 0xFE: 'ứ', 0xF0: 'ự',
        0xA1: 'ỳ', 0xA2: 'ỷ', 0xA3: 'ỹ', 0xA4: 'ý', 0xA5: 'ỵ',
        # Chữ hoa TCVN3
        0x80: 'À', 0x81: 'Ả', 0x82: 'Ã', 0x83: 'Á', 0x84: 'Ạ',
        0x85: 'Ă', 0x86: 'Ằ', 0x87: 'Ẳ', 0x88: 'Ẵ', 0x89: 'Ắ', 0x8A: 'Ặ',
        0x8B: 'Â', 0x8C: 'Ầ', 0x8D: 'Ẩ', 0x8E: 'Ẫ', 0x8F: 'Ấ', 0x90: 'Ậ',
        0x91: 'Đ',
        0x92: 'È', 0x93: 'Ẻ', 0x94: 'Ẽ', 0x95: 'É', 0x96: 'Ẹ',
        0x97: 'Ê', 0x98: 'Ề', 0x99: 'Ể', 0x9A: 'Ễ', 0x9B: 'Ế', 0x9C: 'Ệ',
        0x9D: 'Ì', 0x9E: 'Ỉ', 0x9F: 'Ĩ', 0xAA: 'Í', 0xAB: 'Ị',
        0xAC: 'Ò', 0xAD: 'Ỏ', 0xAE: 'Õ', 0xAF: 'Ó', 0xB0: 'Ọ',
        0xB1: 'Ô', 0xB2: 'Ồ', 0xB3: 'Ổ', 0xB4: 'Ỗ', 0x80: 'Ố', 0x81: 'Ộ',
    }

    # Bảng ánh xạ ký tự 2-byte của VNI-Windows sang Unicode
    VNI_PATTERNS = [
        ('aù', 'á'), ('aà', 'à'), ('aả', 'ả'), ('aã', 'ã'), ('aạ', 'ạ'),
        ('aê', 'ă'), ('aé', 'ắ'), ('aè', 'ằ'), ('aẳ', 'ẳ'), ('aẵ', 'ẵ'), ('aặ', 'ặ'),
        ('aâ', 'â'), ('aá', 'ấ'), ('aà', 'ầ'), ('aẩ', 'ẩ'), ('aẫ', 'ẫ'), ('aậ', 'ậ'),
        ('eù', 'é'), ('eà', 'è'), ('eả', 'ẻ'), ('eã', 'ẽ'), ('eạ', 'ẹ'),
        ('eâ', 'ê'), ('eá', 'ế'), ('eà', 'ề'), ('eẩ', 'ể'), ('eẫ', 'ễ'), ('eậ', 'ệ'),
        ('où', 'ó'), ('oà', 'ò'), ('oả', 'ỏ'), ('oã', 'õ'), ('oạ', 'ọ'),
        ('oâ', 'ô'), ('oá', 'ố'), ('oà', 'ồ'), ('oẩ', 'ổ'), ('oẫ', 'ỗ'), ('oậ', 'ộ'),
        ('ôù', 'ớ'), ('ôà', 'ờ'), ('ôả', 'ở'), ('ôã', 'ỡ'), ('ôạ', 'ợ'),
        ('uù', 'ú'), ('uà', 'ù'), ('uả', 'ủ'), ('uã', 'ũ'), ('uạ', 'ụ'),
        ('öù', 'ứ'), ('öà', 'ừ'), ('öả', 'ử'), ('öã', 'ữ'), ('öạ', 'ự'), ('ö', 'ư'),
        ('yù', 'ý'), ('yà', 'ỳ'), ('yả', 'ỷ'), ('yã', 'ỹ'), ('yạ', 'ỵ'),
        ('ñ', 'đ'), ('Ñ', 'Đ'),
        ('AÙ', 'Á'), ('AÀ', 'À'), ('AẢ', 'Ả'), ('AÃ', 'Ã'), ('AẠ', 'Ạ'),
        ('AÊ', 'Ă'), ('AÂ', 'Â'), ('EÙ', 'É'), ('EÀ', 'È'), ('EÂ', 'Ê'),
        ('OÙ', 'Ó'), ('OÀ', 'Ò'), ('OÂ', 'Ô'), ('ÔÙ', 'Ớ'), ('ÔÀ', 'Ờ'),
        ('UÙ', 'Ú'), ('UÀ', 'Ù'), ('Ö', 'Ư'), ('ÖÙ', 'Ứ'), ('YÙ', 'Ý')
    ]

    @classmethod
    def normalize_to_nfc(cls, text: str) -> str:
        """
        Chuẩn hóa văn bản về Unicode dựng sẵn (NFC)
        Khắc phục triệt để lỗi gõ bàn phím macOS hoặc copy web bị dính Unicode tổ hợp (NFD)
        """
        if not text:
            return ""
        # 1. Ép chuyển về chuẩn NFC
        normalized = unicodedata.normalize('NFC', text)
        return normalized

    @classmethod
    def convert_tcvn3_to_unicode(cls, text: str) -> str:
        """Chuyển đổi chuỗi ký tự TCVN3 (.VnTime) sang Unicode"""
        if not text:
            return ""
        chars = []
        for char in text:
            code = ord(char)
            if code in cls.TCVN3_TO_UNICODE:
                chars.append(cls.TCVN3_TO_UNICODE[code])
            else:
                chars.append(char)
        return cls.normalize_to_nfc("".join(chars))

    @classmethod
    def convert_vni_to_unicode(cls, text: str) -> str:
        """Chuyển đổi chuỗi ký tự VNI-Windows (VNI-Times) sang Unicode"""
        if not text:
            return ""
        result = text
        for vni, uni in cls.VNI_PATTERNS:
            result = result.replace(vni, uni)
        return cls.normalize_to_nfc(result)

    @classmethod
    def auto_convert_to_unicode(cls, text: str, font_name: str = "") -> Tuple[str, bool]:
        """
        Tự động phát hiện font hoặc bảng mã và chuyển đổi về Unicode dựng sẵn chuẩn
        Trả về: (chuỗi_đã_chuyển, cờ_có_chuyển_đổi_không)
        """
        if not text:
            return "", False

        f_lower = font_name.lower()
        # Nhận diện font TCVN3
        if ".vn" in f_lower or "vn" in f_lower and any(fn in f_lower for fn in ["vntime", "vnarial", "vntimeh", "vnavant"]):
            return cls.convert_tcvn3_to_unicode(text), True

        # Nhận diện font VNI
        if "vni-" in f_lower or "vni " in f_lower:
            return cls.convert_vni_to_unicode(text), True

        # Kiểm tra theo tần suất ký tự TCVN3 trong chuỗi
        tcvn_count = sum(1 for c in text if ord(c) in cls.TCVN3_TO_UNICODE)
        if tcvn_count > 3:
            return cls.convert_tcvn3_to_unicode(text), True

        # Mặc định chuẩn hóa Unicode NFC
        nfc_text = cls.normalize_to_nfc(text)
        was_changed = (nfc_text != text)
        return nfc_text, was_changed
