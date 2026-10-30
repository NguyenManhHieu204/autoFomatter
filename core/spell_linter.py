"""
Module 4: Bộ máy Rà soát Lỗi Kỹ thuật & Chính tả Hành chính (Linter & Sanitizer)
Quy chuẩn áp dụng:
- Hướng dẫn 36-HD/VPTW của Ban Chấp hành Trung ương Đảng
- Nghị định 30/2020/NĐ-CP của Chính phủ (đặc biệt Phụ lục II về quy tắc viết hoa)
"""

import json
import os
import re
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Any, Optional

from core.path_utils import get_config_dir


@dataclass
class LintIssue:
    issue_type: str        # 'whitespace', 'punctuation', 'quote', 'date', 'capitalization'
    original: str
    replacement: str
    message: str
    context: str = ""
    severity: str = "warning"  # 'auto_fixed', 'warning', 'info'


@dataclass
class LintResult:
    cleaned_text: str
    issues: List[LintIssue] = field(default_factory=list)
    spaces_fixed: int = 0
    punctuation_fixed: int = 0
    quotes_fixed: int = 0
    dates_fixed: int = 0
    capitalization_warnings: int = 0
    spelling_fixed: int = 0


# Pre-compiled static regex patterns for maximum performance
RE_MULTI_SPACES = re.compile(r'[^\S\r\n]{2,}')
RE_PRE_PUNCT_SPACES = re.compile(r'\s+([,.:;!?\)\]\}])')
RE_POST_BRACKET_SPACES = re.compile(r'([\(\[\{“«])\s+')
RE_MISSING_SPACE_PUNCT = re.compile(r'([,;:!?])([A-Za-zÀ-ỹ0-9])')
RE_MISSING_SPACE_PERIOD = re.compile(r'([A-Za-zÀ-ỹ]{3,})\.([A-ZÀ-Ỹ])')
RE_DATE_PATTERN = re.compile(r'\bngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})\b', re.IGNORECASE)
CONTEXT_TRIGGERS = r'(?:các|những|mỗi|toàn thể|giao|yêu cầu|đề nghị|của|tại|trong|với|tới|do|về|được)\s+'


class SpellLinter:
    def __init__(self, config_dir: Optional[str] = None):
        if config_dir is None:
            config_dir = get_config_dir()
        self.config_dir = config_dir
        self.cap_rules = self._load_capitalization_rules()
        self._compile_capitalization_patterns()

    def _load_capitalization_rules(self) -> Dict[str, Any]:
        rules_path = os.path.join(self.config_dir, "capitalization_rules.json")
        if os.path.exists(rules_path):
            try:
                with open(rules_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "proper_agencies": [
                "Đảng Cộng sản Việt Nam", "Ban Chấp hành Trung ương", "Bộ Chính trị",
                "Ban Bí thư", "Quốc hội", "Ủy ban Thường vụ Quốc hội", "Chủ tịch nước",
                "Chính phủ", "Thủ tướng Chính phủ", "Hội đồng nhân dân", "Ủy ban nhân dân"
            ],
            "common_nouns_improperly_capitalized": [
                "Đồng chí", "Cơ quan", "Đơn vị", "Chi bộ", "Đảng viên", "Cán bộ",
                "Công chức", "Viên chức", "Phòng", "Ban", "Kế hoạch", "Báo cáo"
            ],
            "special_titles_capitalized": [
                "Tổng Bí thư", "Chủ tịch nước", "Thủ tướng Chính phủ", "Chủ tịch Quốc hội"
            ]
        }

    def _compile_capitalization_patterns(self):
        """Tối ưu hóa: biên dịch trước các biểu thức chính quy để tăng tốc độ quét hàng chục lần"""
        improper_nouns = self.cap_rules.get("common_nouns_improperly_capitalized", [])
        if improper_nouns:
            # Sắp xếp từ dài nhất đến ngắn nhất để regex ưu tiên khớp đúng từ ghép
            sorted_nouns = sorted(improper_nouns, key=len, reverse=True)
            escaped_nouns = "|".join(re.escape(n) for n in sorted_nouns)
            self._re_improper_nouns = re.compile(rf'({CONTEXT_TRIGGERS})({escaped_nouns})\b', re.UNICODE)
        else:
            self._re_improper_nouns = None

        proper_agencies = self.cap_rules.get("proper_agencies", [])
        if proper_agencies:
            sorted_agencies = sorted(proper_agencies, key=len, reverse=True)
            self._agency_map = {a.lower(): a for a in sorted_agencies}
            escaped_agencies = "|".join(re.escape(a.lower()) for a in sorted_agencies)
            self._re_proper_agencies = re.compile(
                rf'(?<![A-Za-zÀ-ỹ])({escaped_agencies})(?![A-Za-zÀ-ỹ])',
                re.IGNORECASE
            )
        else:
            self._re_proper_agencies = None
            self._agency_map = {}

    def clean_whitespace(self, text: str) -> Tuple[str, int]:
        """
        Làm sạch khoảng trắng thừa:
        - 2 dấu cách liên tiếp trở lên -> 1 dấu cách
        - Chuẩn hóa non-breaking spaces và zero-width spaces
        """
        if not text:
            return text, 0

        # Chuẩn hóa non-breaking spaces và zero-width spaces
        if '\u00a0' in text or '\u200b' in text:
            text = text.replace('\u00a0', ' ').replace('\u200b', '')

        # Xóa nhiều khoảng trắng liên tiếp trong dòng (giữ nguyên \n nếu có)
        text, count = RE_MULTI_SPACES.subn(' ', text)
        return text, count

    def clean_punctuation_spacing(self, text: str) -> Tuple[str, int, List[LintIssue]]:
        """
        Làm sạch dấu câu theo quy chuẩn soạn thảo văn bản:
        - Xóa khoảng trắng đứng trước: , . ; : ! ? ) ] } ” »
        - Xóa khoảng trắng đứng sau: ( [ { “ «
        - Bổ sung khoảng trắng nếu thiếu sau dấu câu (tránh can thiệp số thập phân, giờ, ký hiệu tắt)
        """
        if not text:
            return text, 0, []

        issues: List[LintIssue] = []
        fixes_count = 0

        # 1. Xóa khoảng trắng đứng trước dấu câu đóng: , . ; : ! ? ) ] }
        def remove_pre_punct_space(match):
            nonlocal fixes_count
            fixes_count += 1
            issues.append(LintIssue(
                issue_type='punctuation',
                original=match.group(0),
                replacement=match.group(1),
                message=f"Xóa khoảng trắng thừa trước dấu '{match.group(1)}'",
                severity='auto_fixed'
            ))
            return match.group(1)

        text = RE_PRE_PUNCT_SPACES.sub(remove_pre_punct_space, text)

        # 2. Xóa khoảng trắng sau dấu mở ngoặc: ( [ { “ «
        def remove_post_bracket_space(match):
            nonlocal fixes_count
            fixes_count += 1
            issues.append(LintIssue(
                issue_type='punctuation',
                original=match.group(0),
                replacement=match.group(1),
                message=f"Xóa khoảng trắng thừa sau dấu '{match.group(1)}'",
                severity='auto_fixed'
            ))
            return match.group(1)

        text = RE_POST_BRACKET_SPACES.sub(remove_post_bracket_space, text)

        # 3. Thêm khoảng trắng sau dấu câu nếu thiếu
        def fix_missing_space_after_comma(match):
            nonlocal fixes_count
            punct = match.group(1)
            char_after = match.group(2)
            # Nếu là số, dấu phẩy giữa 2 số như 3,5 hoặc 1,000 thì không tách
            if punct == ',' and char_after.isdigit():
                return match.group(0)
            fixes_count += 1
            issues.append(LintIssue(
                issue_type='punctuation',
                original=match.group(0),
                replacement=f"{punct} {char_after}",
                message=f"Bổ sung khoảng trắng sau dấu '{punct}'",
                severity='auto_fixed'
            ))
            return f"{punct} {char_after}"

        text = RE_MISSING_SPACE_PUNCT.sub(fix_missing_space_after_comma, text)

        # 4. Dấu chấm: chỉ thêm cách nếu sau dấu chấm là chữ cái viết hoa và trước dấu chấm là từ bình thường
        def fix_missing_space_after_period(match):
            nonlocal fixes_count
            word_before = match.group(1)
            char_after = match.group(2)
            # Tránh chữ viết tắt 1-2 ký tự như TP.HCM, TT., NĐ.
            if len(word_before) <= 2 and word_before.isupper():
                return match.group(0)
            fixes_count += 1
            issues.append(LintIssue(
                issue_type='punctuation',
                original=match.group(0),
                replacement=f"{word_before}. {char_after}",
                message="Bổ sung khoảng trắng sau dấu chấm",
                severity='auto_fixed'
            ))
            return f"{word_before}. {char_after}"

        text = RE_MISSING_SPACE_PERIOD.sub(fix_missing_space_after_period, text)

        return text, fixes_count, issues

    def standardize_quotes(self, text: str) -> Tuple[str, int, List[LintIssue]]:
        """
        Chuẩn hóa ngoặc kép thẳng "" thành ngoặc kép cong chuẩn in ấn “ ”
        """
        if not text or '"' not in text:
            return text, 0, []

        issues: List[LintIssue] = []
        fixes_count = 0
        result_chars = []
        open_quote = True

        for i, char in enumerate(text):
            if char == '"':
                fixes_count += 1
                prev_char = text[i - 1] if i > 0 else ' '
                next_char = text[i + 1] if i < len(text) - 1 else ' '

                if prev_char in ' \t\n([{':
                    target_quote = '“'
                    open_quote = False
                elif next_char in ' \t\n.,;:!?)]}':
                    target_quote = '”'
                    open_quote = True
                else:
                    target_quote = '“' if open_quote else '”'
                    open_quote = not open_quote

                result_chars.append(target_quote)
            else:
                result_chars.append(char)

        if fixes_count > 0:
            issues.append(LintIssue(
                issue_type='quote',
                original='"..."',
                replacement='“...”',
                message=f"Đã chuẩn hóa {fixes_count} dấu ngoặc kép thẳng thành ngoặc kép cong “ ”",
                severity='auto_fixed'
            ))

        return "".join(result_chars), fixes_count, issues

    def sanitize_dates(self, text: str, standard_id: str = "ND30") -> Tuple[str, int, List[LintIssue]]:
        """
        Sửa lỗi định dạng ngày tháng theo Nghị định 30/2020/NĐ-CP hoặc Hướng dẫn 36-HD/VPTW
        """
        if not text or 'ngày' not in text.lower():
            return text, 0, []

        issues: List[LintIssue] = []
        fixes_count = 0

        def replace_date(match):
            nonlocal fixes_count
            raw_day = match.group(1)
            raw_month = match.group(2)
            raw_year = match.group(3)

            day_num = int(raw_day)
            month_num = int(raw_month)

            corrected_day = raw_day
            corrected_month = raw_month

            # Chuẩn hóa ngày (1-9 thêm số 0)
            if 1 <= day_num <= 9 and len(raw_day) == 1:
                corrected_day = f"0{day_num}"

            # Chuẩn hóa tháng
            if month_num in (1, 2) and len(raw_month) == 1:
                corrected_month = f"0{month_num}"
            elif 3 <= month_num <= 12 and raw_month.startswith('0') and len(raw_month) == 2:
                corrected_month = str(month_num)

            original_str = match.group(0)
            new_str = f"ngày {corrected_day} tháng {corrected_month} năm {raw_year}"

            if original_str != new_str:
                fixes_count += 1
                issues.append(LintIssue(
                    issue_type='date',
                    original=original_str,
                    replacement=new_str,
                    message=f"Chuẩn hóa thể thức ngày tháng: '{original_str}' -> '{new_str}' theo {standard_id}",
                    severity='auto_fixed'
                ))
                return new_str
            return original_str

        text = RE_DATE_PATTERN.sub(replace_date, text)
        return text, fixes_count, issues

    def check_capitalization_warnings(self, text: str) -> List[LintIssue]:
        """
        Rà soát lỗi quy tắc viết hoa (quét nhanh 1 lượt bằng regex gộp):
        1. Cảnh báo viết hoa tùy tiện danh từ chung khi không đứng đầu câu
        2. Cảnh báo quên viết hoa tên cơ quan, tổ chức chính thức
        """
        if not text:
            return []

        warnings: List[LintIssue] = []

        # 1. Kiểm tra danh từ chung bị viết hoa tùy tiện ở giữa câu (Single Pass)
        if self._re_improper_nouns:
            for m in self._re_improper_nouns.finditer(text):
                prefix = m.group(1)
                detected = m.group(2)

                # Ngoại lệ: Nếu danh từ chung đi kèm tên cơ quan, tổ chức cụ thể
                following = text[m.end():m.end() + 25].strip()
                if detected in ("Ban", "ban"):
                    if any(following.startswith(k) for k in [
                        "Chấp hành", "Bí thư", "Thường vụ", "Xây dựng", "Tuyên giáo",
                        "Kiểm tra", "Tổ chức", "Quản lý", "Chỉ đạo", "Cán sự", "Liên lạc"
                    ]):
                        continue
                elif detected in ("Chi bộ", "chi bộ"):
                    if any(following.startswith(k) for k in [
                        "Cơ quan", "Công an", "Quân sự", "thôn", "Thôn", "bản", "Bản", "trường", "Trường",
                        "Y tế", "y tế", "Giáo dục", "giáo dục", "Tổ", "tổ", "Khu", "khu", "Ấp", "ấp", "Buôn", "buôn"
                    ]):
                        continue

                lower_suggestion = detected.lower()
                warnings.append(LintIssue(
                    issue_type='capitalization',
                    original=f"{prefix}{detected}",
                    replacement=f"{prefix}{lower_suggestion}",
                    message=f"Danh từ chung '{detected}' viết hoa tùy tiện khi không đi kèm tên riêng cụ thể (Phụ lục II NĐ 30/2020/NĐ-CP)",
                    context=m.group(0),
                    severity='auto_fixed'
                ))

        # 2. Kiểm tra quên viết hoa tên cơ quan, tổ chức chính thức (Single Pass)
        if self._re_proper_agencies:
            for m in self._re_proper_agencies.finditer(text):
                matched_text = m.group(1)
                lower_key = matched_text.lower()
                standard_name = self._agency_map.get(lower_key, matched_text)

                if matched_text != standard_name:
                    # Bỏ qua nếu toàn bộ từ/cụm từ đang in hoa hết (tiêu đề in hoa)
                    if matched_text.isupper():
                        continue
                    # Ngoại lệ: 'bộ ...' trong 'Chi bộ ...' (Chi bộ Y tế, Chi bộ Công an, Chi bộ Quân sự...)
                    pre_text = text[:m.start()].rstrip().lower()
                    if (
                        pre_text.endswith("chi") or
                        pre_text.endswith("chi bộ") or
                        pre_text.endswith("- chi") or
                        pre_text.endswith("- chi bộ") or
                        pre_text.endswith("tổ") or
                        pre_text.endswith("ban") or
                        pre_text.endswith("khối")
                    ):
                        continue

                    warnings.append(LintIssue(
                        issue_type='capitalization',
                        original=matched_text,
                        replacement=standard_name,
                        message=f"Quên viết hoa đúng chuẩn tên cơ quan/tổ chức '{standard_name}'",
                        context=m.group(0),
                        severity='warning'
                    ))

        return warnings

    def lint_text(self, text: str, standard_id: str = "ND30") -> LintResult:
        """
        Thực hiện toàn bộ quy trình rà soát và làm sạch cho một đoạn văn bản:
        0. Khôi phục dấu tiếng Việt, sửa lỗi chính tả hành chính & khử nhiễu ký tự
        1. Làm sạch khoảng trắng thừa
        2. Chuẩn hóa dấu ngoặc kép
        3. Làm sạch dấu câu
        4. Chuẩn hóa ngày tháng
        5. Rà soát viết hoa
        """
        all_issues: List[LintIssue] = []

        # 0. Sửa lỗi chính tả, phục hồi dấu tiếng Việt & khử nhiễu ký tự (Bộ máy nâng cao)
        from core.vietnamese_spell_checker import VietnameseSpellChecker
        text, spell_raw_issues = VietnameseSpellChecker.correct_text(text)
        spelling_fixed = len(spell_raw_issues)
        for s_iss in spell_raw_issues:
            all_issues.append(LintIssue(
                issue_type='spelling',
                original=s_iss['original'],
                replacement=s_iss['replacement'],
                message=s_iss['message'],
                severity=s_iss.get('severity', 'auto_fixed')
            ))

        # 1. Làm sạch khoảng trắng thừa
        text, spaces_fixed = self.clean_whitespace(text)
        if spaces_fixed > 0:
            all_issues.append(LintIssue(
                issue_type='whitespace',
                original='  ',
                replacement=' ',
                message=f"Đã dọn {spaces_fixed} vị trí chứa khoảng trắng thừa liên tiếp",
                severity='auto_fixed'
            ))

        # 2. Chuẩn hóa dấu ngoặc kép
        text, quotes_fixed, quote_issues = self.standardize_quotes(text)
        all_issues.extend(quote_issues)

        # 3. Làm sạch dấu câu
        text, punct_fixed, punct_issues = self.clean_punctuation_spacing(text)
        all_issues.extend(punct_issues)

        # 4. Chuẩn hóa ngày tháng
        text, dates_fixed, date_issues = self.sanitize_dates(text, standard_id)
        all_issues.extend(date_issues)

        # 5. Rà soát viết hoa
        cap_warnings = self.check_capitalization_warnings(text)
        all_issues.extend(cap_warnings)
        for w in cap_warnings:
            if w.severity == 'auto_fixed' and w.original in text:
                text = text.replace(w.original, w.replacement)

        return LintResult(
            cleaned_text=text,
            issues=all_issues,
            spaces_fixed=spaces_fixed,
            punctuation_fixed=punct_fixed,
            quotes_fixed=quotes_fixed,
            dates_fixed=dates_fixed,
            capitalization_warnings=len(cap_warnings),
            spelling_fixed=spelling_fixed
        )
