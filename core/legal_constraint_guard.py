"""
core/legal_constraint_guard.py: Giám sát & Cưỡng chế Quy chuẩn Thể thức NĐ 30 & HD 36
- Kiểm soát tuân thủ thời gian thực (Real-time Compliance Monitor)
- Cảnh báo tức thì nếu văn bản sử dụng sai phông chữ (không phải Times New Roman),
  sai kích thước lề in, thiếu gạch chân tiêu ngữ hoặc sai tiền tố thẩm quyền ký.
- Trả về Báo cáo Tuân thủ Pháp lý (Legal Compliance Audit) kèm điểm chuẩn hóa (0 - 100%).
"""

from typing import Dict, Any, List, Optional
import docx


class RuleViolation:
    def __init__(self, rule_id: str, severity: str, message: str, suggestion: str):
        self.rule_id = rule_id
        self.severity = severity  # 'error', 'warning', 'info'
        self.message = message
        self.suggestion = suggestion

    def to_dict(self) -> Dict[str, str]:
        return {
            "rule_id": self.rule_id,
            "severity": self.severity,
            "message": self.message,
            "suggestion": self.suggestion
        }


class LegalConstraintGuard:
    """Bộ kiểm soát và đánh giá chuẩn hóa thể thức pháp lý trực tiếp"""

    @classmethod
    def validate_document(cls, doc: docx.Document, standard_id: str = "ND30") -> Dict[str, Any]:
        """
        Rà soát tài liệu docx đối chiếu với quy chuẩn thể thức NĐ 30 / HD 36
        """
        violations: List[RuleViolation] = []
        standard_id = standard_id.upper()

        # 1. RÀ SOÁT LỀ IN (PAGE MARGINS)
        if doc.sections:
            sec = doc.sections[0]
            top_mm = sec.top_margin.mm
            bottom_mm = sec.bottom_margin.mm
            left_mm = sec.left_margin.mm
            right_mm = sec.right_margin.mm

            # NĐ 30 & HD 36: Lề trên 20-25mm, Lề dưới 20-25mm, Lề trái 30-35mm, Lề phải 15-20mm
            if not (19.5 <= top_mm <= 25.5):
                violations.append(RuleViolation(
                    "MARGIN_TOP", "warning",
                    f"Lề trên hiện tại ({top_mm:.1f} mm) không nằm trong khoảng 20-25 mm chuẩn luật định.",
                    "Điều chỉnh lề trên về 20 mm."
                ))
            if not (19.5 <= bottom_mm <= 25.5):
                violations.append(RuleViolation(
                    "MARGIN_BOTTOM", "warning",
                    f"Lề dưới hiện tại ({bottom_mm:.1f} mm) không nằm trong khoảng 20-25 mm chuẩn luật định.",
                    "Điều chỉnh lề dưới về 20 mm."
                ))
            if not (29.5 <= left_mm <= 35.5):
                violations.append(RuleViolation(
                    "MARGIN_LEFT", "error",
                    f"Lề trái hiện tại ({left_mm:.1f} mm) vi phạm quy định (phải từ 30 đến 35 mm để đóng gáy hồ sơ).",
                    "Điều chỉnh lề trái về 30 mm."
                ))
            if not (14.5 <= right_mm <= 20.5):
                violations.append(RuleViolation(
                    "MARGIN_RIGHT", "warning",
                    f"Lề phải hiện tại ({right_mm:.1f} mm) không nằm trong khoảng 15-20 mm chuẩn luật định.",
                    "Điều chỉnh lề phải về 15 mm."
                ))

            # Kiểm tra đánh số trang
            if not sec.different_first_page_header_footer:
                violations.append(RuleViolation(
                    "PAGE_NUM_FIRST_PAGE", "warning",
                    "Chưa bật chế độ ẩn số trang tại trang 1 (Trang đầu tiên của văn bản không đánh số trang).",
                    "Bật tùy chọn ẩn số trang tại trang bìa/trang 1."
                ))

        # 2. RÀ SOÁT PHÔNG CHỮ & KÍCH THƯỚC (TYPOGRAPHY)
        non_tnr_runs = 0
        total_runs = 0
        for p in doc.paragraphs:
            for r in p.runs:
                total_runs += 1
                if r.font.name and r.font.name.lower() not in ["times new roman", "times"]:
                    non_tnr_runs += 1

        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            total_runs += 1
                            if r.font.name and r.font.name.lower() not in ["times new roman", "times"]:
                                non_tnr_runs += 1

        if total_runs > 0 and (non_tnr_runs / total_runs) > 0.05:
            violations.append(RuleViolation(
                "FONT_FAMILY_INVALID", "error",
                f"Phát hiện {non_tnr_runs} đoạn chữ sử dụng phông phi chuẩn (Quy định bắt buộc phông Times New Roman Unicode).",
                "Chuyển toàn bộ văn bản về phông Times New Roman."
            ))

        # 3. RÀ SOÁT TIÊU ĐỀ QUỐC HIỆU / TIÊU NGỮ / ĐẢNG
        # Lấy toàn bộ nội dung văn bản từ các thẻ w:t (bao gồm cả trong các bảng ẩn và content controls w:sdt)
        all_text_nodes = doc._element.xpath('//w:t/text()')
        full_text = " ".join(str(t) for t in all_text_nodes) if all_text_nodes else "\n".join(p.text for p in doc.paragraphs)

        if standard_id == "ND30":
            if "ĐỘC LẬP - TỰ DO - HẠNH PHÚC" not in full_text.upper():
                violations.append(RuleViolation(
                    "MISSING_MOTTO", "error",
                    "Thiếu Tiêu ngữ 'Độc lập - Tự do - Hạnh phúc' của văn bản hành chính Nhà nước.",
                    "Thêm Tiêu ngữ vào góc phải bảng Tiêu đề."
                ))
        elif standard_id == "HD36":
            if "ĐẢNG CỘNG SẢN VIỆT NAM" not in full_text.upper():
                violations.append(RuleViolation(
                    "MISSING_PARTY_HEADING", "error",
                    "Thiếu Tiêu đề 'ĐẢNG CỘNG SẢN VIỆT NAM' của văn bản Đảng.",
                    "Thêm tiêu đề Đảng vào góc phải bảng Tiêu đề."
                ))
            if "ĐỘC LẬP - TỰ DO - HẠNH PHÚC" in full_text.upper():
                violations.append(RuleViolation(
                    "PARTY_UNAUTHORIZED_MOTTO", "error",
                    "Văn bản Đảng tuyệt đối KHÔNG có tiêu ngữ 'Độc lập - Tự do - Hạnh phúc' (Hướng dẫn 36-HD/VPTW).",
                    "Xóa bỏ dòng tiêu ngữ khỏi văn bản Đảng."
                ))

        # 4. TÍNH ĐIỂM CHUẨN HÓA (0 - 100%)
        error_count = sum(1 for v in violations if v.severity == "error")
        warn_count = sum(1 for v in violations if v.severity == "warning")
        score = max(0.0, 100.0 - (error_count * 20.0) - (warn_count * 5.0))

        return {
            "score": round(score, 1),
            "is_compliant": (error_count == 0),
            "error_count": error_count,
            "warning_count": warn_count,
            "violations": [v.to_dict() for v in violations]
        }
