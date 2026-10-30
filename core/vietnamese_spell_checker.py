"""
core/vietnamese_spell_checker.py: Bộ máy Rà soát & Phục hồi Chính tả Hành chính Tiếng Việt Toàn diện
Chức năng:
1. Tự động phục hồi dấu tiếng Việt bị thiếu / gõ thiếu dấu trong các cụm từ hành chính
2. Sửa lỗi chính tả thường gặp trong văn bản cơ quan Đảng & Nhà nước (sáp nhập, xử lý, xuất sắc...)
3. Khắc phục triệt để lỗi gõ Telex / Unikey và nhiễu ký tự OCR (s6 -> số, cuia -> của, trurc -> trực...)
4. Khử nhiễu ký tự rác vô hình (Zero-width spaces, BOM, ký tự điều khiển)
5. Giữ nguyên định dạng chữ hoa/chữ thường (Case preservation)
"""

import re
import unicodedata
from typing import List, Tuple, Dict, Any, Optional
from core.charset_converter import CharsetConverter


class VietnameseSpellChecker:
    # 1. Danh mục ký tự rác vô hình cần thanh lọc
    INVISIBLE_CHARS_PATTERN = re.compile(r'[\u200b\u200c\u200d\ufeff\u200e\u200f\x00-\x08\x0b\x0c\x0e-\x1f]')

    # 2. Cặp từ nhầm lẫn chính tả hành chính kinh điển (Dễ sai -> Đúng)
    ADMIN_TYPOS_MAP = {
        # Từ chuẩn mực công vụ
        "sát nhập": "sáp nhập",
        "sử lý": "xử lý",
        "suất sắc": "xuất sắc",
        "sơ xuất": "sơ suất",
        "chuẩn đoán": "chẩn đoán",
        "trao chuốt": "trau chuốt",
        "chín chu": "chỉn chu",
        "thăm quan": "tham quan",
        "xúc tích": "súc tích",
        "chủ chương": "chủ trương",
        "câu truyện": "câu chuyện",
        "rút song nghiệm": "rút kinh nghiệm",
        "nghiêm tút": "nghiêm túc",
        "tập chung": "tập trung",
        "bổ xung": "bổ sung",
        "sắp sếp": "sắp xếp",
        "sắp sấp": "sắp xếp",
        "dành độc lập": "giành độc lập",
        "kỷ thuật": "kỹ thuật",
        "liên luỵ": "liên lụy",
        "qui định": "quy định",
        "qui chế": "quy chế"
    }

    # 3. Danh mục từ/cụm từ hành chính thiếu dấu phổ biến (Không dấu -> Có dấu chuẩn)
    MISSING_DIACRITICS_MAP = {
        # Chức danh, thẩm quyền
        "thu truong": "thủ trưởng",
        "hieu truong": "hiệu trưởng",
        "chu tich": "chủ tịch",
        "pho chu tich": "phó chủ tịch",
        "giam doc": "giám đốc",
        "pho giam doc": "phó giám đốc",
        "bi thu": "bí thư",
        "pho bi thu": "phó bí thư",
        "truong phong": "trưởng phòng",
        "pho truong phong": "phó trưởng phòng",
        "truong ban": "trưởng ban",
        "pho truong ban": "phó trưởng ban",
        "can bo": "cán bộ",
        "cong chuc": "công chức",
        "vien chuc": "viên chức",
        "dong chi": "đồng chí",
        "dang vien": "đảng viên",
        "doan vien": "đoàn viên",
        "hoi vien": "hội viên",
        "chuyen vien": "chuyên viên",
        "thu ky": "thư ký",
        "chanh van phong": "chánh văn phòng",
        "pho chanh van phong": "phó chánh văn phòng",

        # Cơ quan, tổ chức
        "uy ban nhan dan": "ủy ban nhân dân",
        "hoi dong nhan dan": "hội đồng nhân dân",
        "chinh phu": "chính phủ",
        "thu tuong": "thủ tướng",
        "quoc hoi": "quốc hội",
        "bo chinh tri": "bộ chính trị",
        "ban bi thu": "ban bí thư",
        "ban chap hanh": "ban chấp hành",
        "trung uong": "trung ương",
        "tinh uy": "tỉnh ủy",
        "thanh uy": "thành ủy",
        "huyen uy": "huyện ủy",
        "dang bo": "đảng bộ",
        "dang uy": "đảng ủy",
        "chi bo": "chi bộ",
        "thuong truc": "thường trực",
        "thuong vu": "thường vụ",
        "ban thuong vu": "ban thường vụ",
        "to dan pho": "tổ dân phố",
        "thon truong": "thôn trưởng",
        "to truong": "tổ trưởng",

        # Văn bản & Thể thức
        "nghi quyet": "nghị quyết",
        "quyet dinh": "quyết định",
        "chi thi": "chỉ thị",
        "quy che": "quy chế",
        "quy dinh": "quy định",
        "thong bao": "thông báo",
        "thong tri": "thông tri",
        "huong dan": "hướng dẫn",
        "cong van": "công văn",
        "to trinh": "tờ trình",
        "ke hoach": "kế hoạch",
        "bao cao": "báo cáo",
        "bien ban": "biên bản",
        "giay moi": "giấy mời",
        "chuong trinh": "chương trình",
        "de an": "đề án",
        "phuong an": "phương án",
        "trich yeu": "trích yếu",
        "noi nhan": "nơi nhận",
        "luu: vt": "lưu: vt",

        # Hoạt động công vụ & Nghiệp vụ
        "hoan thanh": "hoàn thành",
        "thuc hien": "thực hiện",
        "trien khai": "triển khai",
        "tong ket": "tổng kết",
        "so ket": "sơ kết",
        "danh gia": "đánh giá",
        "phan loai": "phân loại",
        "phe duyet": "phê duyệt",
        "ban hanh": "ban hành",
        "chap thuan": "chấp thuận",
        "thong qua": "thông qua",
        "tham muu": "tham mưu",
        "giup viec": "giúp việc",
        "phoi hop": "phối hợp",
        "chi dao": "chỉ đạo",
        "quan tam": "quan tâm",
        "tao dieu kien": "tạo điều kiện",
        "dieu kien": "điều kiện",
        "trach nhiem": "trách nhiệm",
        "nghia vu": "nghĩa vụ",
        "quyen han": "quyền hạn",
        "nhiem vu": "nhiệm vụ",
        "chuc nang": "chức năng",
        "bo may": "bộ máy",
        "to chuc": "tổ chức",
        "co quan": "cơ quan",
        "don vi": "đơn vị",
        "phong ban": "phòng ban",
        "dia ban": "địa bàn",
        "co so": "cơ sở",
        "cap tren": "cấp trên",
        "truc tiep": "trực tiếp",
        "truc thuoc": "trực thuộc",
        "thanh lap": "thành lập",
        "bo nhiem": "bổ nhiệm",
        "mien nhiem": "miễn nhiệm",
        "dieu dong": "điều động",
        "luan chuyen": "luân chuyển",
        "khen thuong": "khen thưởng",
        "ky luat": "kỷ luật",
        "thanh tra": "thanh tra",
        "kiem tra": "kiểm tra",
        "giam sat": "giám sát",
        "phat trien": "phát triển",
        "kinh te": "kinh tế",
        "xa hoi": "xã hội",
        "quoc phong": "quốc phòng",
        "an ninh": "an ninh",
        "tai chinh": "tài chính",
        "ngan sach": "ngân sách",
        "dau tu": "đầu tư",
        "cai cach": "cải cách",
        "hanh chinh": "hành chính",
        "cai cach hanh chinh": "cải cách hành chính",
        "cong viec": "công việc",
        "chu y": "chú ý",
        "sat nhap": "sáp nhập",
        "su ly": "xử lý",
        "xuat sac": "xuất sắc",
        "so xuat": "sơ suất",
        "chuan doan": "chẩn đoán",
        "trao chuot": "trau chuốt",
        "tham quan": "tham quan",
        "xuc tich": "súc tích",
        "chu chuong": "chủ trương",
        "chuyen doi so": "chuyển đổi số",
        "cong nghe thong tin": "công nghệ thông tin",
        "ung dung": "ứng dụng",
        "hieu qua": "hiệu quả",
        "tien do": "tiến độ",
        "muc tieu": "mục tiêu",
        "chi tieu": "chỉ tiêu",
        "quy trinh": "quy trình",
        "tieu chuan": "tiêu chuẩn",
        "thuc te": "thực tế",
        "ton tai": "tồn tại",
        "han che": "hạn chế",
        "nguyen nhan": "nguyên nhân",
        "giai phap": "giải pháp",
        "phuong huong": "phương hướng",
        "kien nghi": "kiến nghị",
        "de xuat": "đề xuất",
        "thoi gian qua": "thời gian qua",
        "trong thoi gian toi": "trong thời gian tới",
        "tren dia ban tinh": "trên địa bàn tỉnh",
        "co hieu luc ke tu ngay ky": "có hiệu lực kể từ ngày ký",
        "can cu": "căn cứ",
        "theo de nghi": "theo đề nghị",
        "lam viec": "làm việc",
        "quy che lam viec": "quy chế làm việc",
        "co trach nhiem": "có trách nhiệm",
        "bao cao ve": "báo cáo về",
        "ve tinh hinh": "về tình hình",
        "tren cac": "trên các",
        "tai dia ban": "tại địa bàn",
        "nhiemky": "nhiệm kỳ"
    }

    # 4. Lỗi OCR & Gõ dính phím số / ký tự lạ
    OCR_NOISE_PATTERNS = [
        # Số hiệu và từ dính
        (re.compile(r'\bs[6ÓÕ]\b', re.IGNORECASE), 'số'),
        (re.compile(r'\bc[6ÓÕ]\b', re.IGNORECASE), 'có'),
        (re.compile(r'\bcuia\b|\bcưa\b', re.IGNORECASE), 'của'),
        (re.compile(r'\btrurc\b|\btryc\b|\btrục\s+tiếp\b', re.IGNORECASE), 'trực tiếp'),
        (re.compile(r'\btrục\s+thuộc\b|\btryc\s+thuoc\b|\btrực\s+thuọc\b', re.IGNORECASE), 'trực thuộc'),
        (re.compile(r'\bchurc\b', re.IGNORECASE), 'chức'),
        (re.compile(r'\btir\b', re.IGNORECASE), 'từ'),
        (re.compile(r'\bLuru\b|\bLuu:\b', re.IGNORECASE), 'Lưu:'),
        (re.compile(r'\bđươc\b|\bđuợc\b|\bđuoc\b', re.IGNORECASE), 'được'),
        (re.compile(r'\bđổng\s+chí\b|\bđông\s+chí\b', re.IGNORECASE), 'đồng chí'),
        (re.compile(r'\bphê\s+duyêt\b', re.IGNORECASE), 'phê duyệt'),
        (re.compile(r'\bkêt\s+luân\b', re.IGNORECASE), 'kết luận'),
        (re.compile(r'\bvãn\s+bản\b', re.IGNORECASE), 'văn bản'),
        (re.compile(r'\bthòi\s+gian\b', re.IGNORECASE), 'thời gian'),
        (re.compile(r'\bquyêt\s+đinh\b|\bquyết\s+đinh\b', re.IGNORECASE), 'quyết định'),
        (re.compile(r'\bhoan\s+thanh\b', re.IGNORECASE), 'hoàn thành'),
        (re.compile(r'\bkế\s+hoach\b|\bke\s+hoạch\b', re.IGNORECASE), 'kế hoạch'),
        (re.compile(r'\bnhiêm\s+vụ\b|\bnhiệm\s+vu\b', re.IGNORECASE), 'nhiệm vụ'),
        (re.compile(r'\bhuong\s+dẫn\b|\bhướng\s+dân\b', re.IGNORECASE), 'hướng dẫn'),
        (re.compile(r'\bnghi\s+quyết\b|\bnghị\s+quyet\b', re.IGNORECASE), 'nghị quyết'),

        # Cụm từ thiếu dấu ngữ cảnh hành chính
        (re.compile(r'\bco\s+trách\s+nhiệm\b', re.IGNORECASE), 'có trách nhiệm'),
        (re.compile(r'\bbáo\s+cáo\s+ve\b', re.IGNORECASE), 'báo cáo về'),
        (re.compile(r'\bve\s+tình\s+hình\b', re.IGNORECASE), 'về tình hình'),
        (re.compile(r'\b(tren\s+cac|trên\s+cac|tren\s+các)\b', re.IGNORECASE), 'trên các'),
        (re.compile(r'\btai\s+địa\s+bàn\b', re.IGNORECASE), 'tại địa bàn'),
        (re.compile(r'\blam\s+viec\b', re.IGNORECASE), 'làm việc'),
        (re.compile(r'\bkhoa\s+([IVXLCDM\d]+)\b', re.IGNORECASE), r'khóa \1'),
        (re.compile(r'\bnhiemky\s*(\d{4})', re.IGNORECASE), r'nhiệm kỳ \1'),
        (re.compile(r'\bnhiemky\b', re.IGNORECASE), 'nhiệm kỳ'),
        (re.compile(r'\bPho\s+BÍ\s+THƯ\b|\bPho\s+Bí\s+thư\b', re.IGNORECASE), 'Phó Bí thư'),
        (re.compile(r'\bBÍ\s+THƯ\s+Đảng\s+ủy\b', re.IGNORECASE), 'Bí thư Đảng ủy'),
        (re.compile(r'\bCHỦ\s+TỊCH\s+(HĐND|UBND)\b', re.IGNORECASE), r'Chủ tịch \1'),
        (re.compile(r'\bphụ\s+trách\s+các\s+Chi\s+trực\s+thuộc\b', re.IGNORECASE), 'phụ trách các chi bộ trực thuộc'),
        (re.compile(r'(\b[Cc]ăn\s+cứ\s+)QUYẾT\s+ĐỊNH(\s+số\b)'), r'\1Quyết định\2'),
        (re.compile(r'\b(các|những|mỗi|từng)\s+Chi\s+bộ\b'), r'\1 chi bộ'),
        (re.compile(r'\b(các|những|mỗi|từng)\s+Cơ\s+quan\b'), r'\1 cơ quan'),
        (re.compile(r'\b(các|những|mỗi|từng)\s+Đơn\s+vị\b'), r'\1 đơn vị'),
        (re.compile(r'(\d+)\s*-\s*QC/ĐU', re.IGNORECASE), r'\1-QC/ĐU'),
        (re.compile(r'(\d+)\s*-\s*QĐ/', re.IGNORECASE), r'\1-QĐ/'),
        (re.compile(r'([A-Za-zÀ-ỹ]+)\?\s+', re.IGNORECASE), r'\1 '),
        (re.compile(r'(Tổ\s*trưởng|Tổ\s*phó|Tổ\s*viên)(Đ/c|đ/c)', re.IGNORECASE), r'\1, \2'),
        (re.compile(r'\b(Tổ\s*viên)\s*10\b', re.IGNORECASE), r'\1'),
        (re.compile(r'^-([A-Za-zÀ-ỹ])', re.MULTILINE), r'- \1'),

        # Sửa dính chữ và số: số45 -> số 45, ngày05 -> ngày 05, Điều1 -> Điều 1
        (re.compile(r'\b(số|ngày|tháng|năm|Điều|Khoản|Điểm)(\d+)\b', re.IGNORECASE), r'\1 \2'),
    ]

    @classmethod
    def clean_invisible_noise(cls, text: str) -> str:
        """Xóa sạch các ký tự điều khiển và ký tự vô hình gây nhiễu"""
        if not text:
            return ""
        return cls.INVISIBLE_CHARS_PATTERN.sub('', text)

    @classmethod
    def apply_case_pattern(cls, original: str, replacement: str) -> str:
        """
        Giữ nguyên phong cách viết hoa của từ gốc:
        - Nếu gốc là IN HOA TOÀN BỘ -> Thay thế cũng IN HOA TOÀN BỘ
        - Nếu gốc Viết Hoa Toàn Bộ Chữ Đầu (Title Case) -> Thay thế cũng Title Case
        - Nếu gốc chỉ Viết hoa chữ cái đầu câu (Sentence Case) -> Chỉ viết hoa chữ cái đầu tiên
        - Nếu gốc viết thường -> Thay thế viết thường
        """
        if original.isupper():
            return replacement.upper()

        orig_words = original.split()
        repl_words = replacement.split()

        # Nếu replacement đã được định dạng hoa/thường chuẩn (chứa từ viết hoa bên trong như 'Bí thư Đảng ủy')
        if len(repl_words) > 1 and any(w and w[0].isupper() for w in repl_words[1:]):
            if original and original[0].islower():
                return replacement[0].lower() + replacement[1:]
            return replacement

        # Kiểm tra nếu tất cả các từ trong cụm gốc đều viết hoa chữ đầu (ví dụ: Ban Thường Vụ)
        if len(orig_words) > 1 and all(w and w[0].isupper() for w in orig_words):
            return " ".join(w.capitalize() for w in repl_words)

        # Nếu chỉ có từ đầu tiên viết hoa (ví dụ: Cán bộ)
        if original and original[0].isupper():
            if len(repl_words) == 1:
                return repl_words[0].capitalize()
            return repl_words[0].capitalize() + " " + " ".join(w.lower() for w in repl_words[1:])

        return replacement.lower()

    @classmethod
    def correct_text(cls, text: str) -> Tuple[str, List[Dict[str, Any]]]:
        """
        Rà soát và sửa lỗi toàn diện cho đoạn văn bản:
        1. Chuẩn hóa Unicode NFC
        2. Dọn ký tự rác
        3. Khắc phục lỗi OCR & gõ Telex
        4. Sửa lỗi chính tả hành chính
        5. Phục hồi dấu tiếng Việt bị thiếu
        """
        if not text:
            return "", []

        # Bước 1: Chuẩn hóa NFC và khử ký tự vô hình
        current_text = CharsetConverter.normalize_to_nfc(text)
        current_text = cls.clean_invisible_noise(current_text)

        issues: List[Dict[str, Any]] = []

        # Bước 2: Sửa lỗi OCR và gõ dính phím
        for pattern, repl in cls.OCR_NOISE_PATTERNS:
            def _sub_ocr(m):
                orig = m.group(0)
                if "\\" in repl:
                    new_val = m.expand(repl)
                elif m.lastindex and m.lastindex >= 2:
                    new_val = f"{m.group(1)} {m.group(2)}"
                else:
                    new_val = cls.apply_case_pattern(orig, repl)

                orig_nfc = unicodedata.normalize('NFC', orig.strip())
                new_val_nfc = unicodedata.normalize('NFC', new_val.strip())
                if orig_nfc != new_val_nfc:
                    issues.append({
                        "issue_type": "spelling",
                        "original": orig_nfc,
                        "replacement": new_val_nfc,
                        "message": f"Sửa lỗi ký tự/nhiễu: '{orig_nfc}' -> '{new_val_nfc}'",
                        "severity": "auto_fixed"
                    })
                return new_val

            current_text = pattern.sub(_sub_ocr, current_text)

        # Bước 3: Sửa cặp từ nhầm lẫn chính tả hành chính
        for typo, correct in cls.ADMIN_TYPOS_MAP.items():
            pattern = re.compile(rf'\b{re.escape(typo)}\b', re.IGNORECASE)
            matches = list(pattern.finditer(current_text))
            if matches:
                def _sub_typo(m):
                    orig = m.group(0)
                    new_val = cls.apply_case_pattern(orig, correct)
                    orig_nfc = unicodedata.normalize('NFC', orig.strip())
                    new_val_nfc = unicodedata.normalize('NFC', new_val.strip())
                    if orig_nfc != new_val_nfc:
                        issues.append({
                            "issue_type": "spelling",
                            "original": orig_nfc,
                            "replacement": new_val_nfc,
                            "message": f"Chuẩn hóa chính tả hành chính: '{orig_nfc}' -> '{new_val_nfc}'",
                            "severity": "auto_fixed"
                        })
                    return new_val

                current_text = pattern.sub(_sub_typo, current_text)

        # Bước 4: Phục hồi dấu tiếng Việt bị thiếu cho các cụm từ hành chính
        # Sắp xếp các cụm từ theo chiều dài giảm dần để ưu tiên cụm dài trước (Greedy Match)
        sorted_missing = sorted(cls.MISSING_DIACRITICS_MAP.items(), key=lambda x: len(x[0]), reverse=True)
        for no_diacritic, with_diacritic in sorted_missing:
            pattern = re.compile(rf'\b{re.escape(no_diacritic)}\b', re.IGNORECASE)
            matches = list(pattern.finditer(current_text))
            if matches:
                def _sub_diacritic(m):
                    orig = m.group(0)
                    new_val = cls.apply_case_pattern(orig, with_diacritic)
                    orig_nfc = unicodedata.normalize('NFC', orig.strip())
                    new_val_nfc = unicodedata.normalize('NFC', new_val.strip())
                    if orig_nfc != new_val_nfc:
                        issues.append({
                            "issue_type": "spelling",
                            "original": orig_nfc,
                            "replacement": new_val_nfc,
                            "message": f"Bổ sung dấu tiếng Việt chuẩn: '{orig_nfc}' -> '{new_val_nfc}'",
                            "severity": "auto_fixed"
                        })
                    return new_val

                current_text = pattern.sub(_sub_diacritic, current_text)

        return current_text, issues
