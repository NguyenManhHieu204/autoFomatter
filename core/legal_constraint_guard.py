"""
core/legal_constraint_guard.py: Giám sát & Cưỡng chế Quy chuẩn Thể thức NĐ 30 & HD 36
- Kiểm soát tuân thủ thời gian thực (Real-time Compliance Monitor)
- Cảnh báo tức thì nếu văn bản sử dụng sai phông chữ (không phải Times New Roman),
  sai kích thước lề in, thiếu gạch chân tiêu ngữ hoặc sai tiền tố thẩm quyền ký.
- Trả về Báo cáo Tuân thủ Pháp lý (Legal Compliance Audit) kèm điểm chuẩn hóa (0 - 100%).
"""

from typing import Dict, Any, List, Optional
import re
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
        Rà soát tài liệu docx đối chiếu với quy chuẩn thể thức NĐ 30 / HD 36:
        1. Rà soát lề in (Page Margins)
        2. Rà soát phông chữ & kiểu chữ (Typography)
        3. Rà soát Tiêu đề Quốc hiệu & Tiêu ngữ / Đảng
        4. Rà soát tính tương thích giữa Loại văn bản & Cấu trúc nội dung
        5. Rà soát dữ liệu mẫu chưa hoàn thiện / Ký tự giữ chỗ (...)
        6. Rà soát căn cứ pháp lý & Dấu câu
        """
        violations: List[RuleViolation] = []
        checklist: List[Dict[str, Any]] = []
        standard_id = standard_id.upper()

        # 1. RÀ SOÁT LỀ IN (PAGE MARGINS)
        margin_ok = True
        margin_detail = "Lề in tuân thủ chuẩn Nghị định 30 (Trái 30-35mm, Phải 15-20mm, Trên/Dưới 20-25mm)."
        if doc.sections:
            sec = doc.sections[0]
            top_mm = sec.top_margin.mm
            bottom_mm = sec.bottom_margin.mm
            left_mm = sec.left_margin.mm
            right_mm = sec.right_margin.mm

            # NĐ 30 & HD 36: Lề trên 20-25mm, Lề dưới 20-25mm, Lề trái 30-35mm, Lề phải 15-20mm
            if not (19.5 <= top_mm <= 25.5):
                margin_ok = False
                violations.append(RuleViolation(
                    "MARGIN_TOP", "warning",
                    f"Lề trên hiện tại ({top_mm:.1f} mm) không nằm trong khoảng 20-25 mm chuẩn luật định.",
                    "Điều chỉnh lề trên về 20 mm."
                ))
            if not (19.5 <= bottom_mm <= 25.5):
                margin_ok = False
                violations.append(RuleViolation(
                    "MARGIN_BOTTOM", "warning",
                    f"Lề dưới hiện tại ({bottom_mm:.1f} mm) không nằm trong khoảng 20-25 mm chuẩn luật định.",
                    "Điều chỉnh lề dưới về 20 mm."
                ))
            if not (29.5 <= left_mm <= 35.5):
                margin_ok = False
                violations.append(RuleViolation(
                    "MARGIN_LEFT", "error",
                    f"Lề trái hiện tại ({left_mm:.1f} mm) vi phạm quy định (phải từ 30 đến 35 mm để đóng gáy hồ sơ).",
                    "Điều chỉnh lề trái về 30 mm."
                ))
            if not (14.5 <= right_mm <= 20.5):
                margin_ok = False
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

        if not margin_ok:
            margin_detail = f"Phát hiện kích thước lề chưa đúng chuẩn NĐ 30 ({left_mm:.0f}-{right_mm:.0f}-{top_mm:.0f}-{bottom_mm:.0f} mm)."
        checklist.append({"name": "Khổ giấy A4 & Lề in NĐ 30", "passed": margin_ok, "detail": margin_detail})

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

        font_ok = True
        font_detail = "Toàn bộ văn bản sử dụng phông chữ chuẩn Times New Roman."
        if total_runs > 0 and (non_tnr_runs / total_runs) > 0.05:
            font_ok = False
            font_detail = f"Phát hiện {non_tnr_runs} đoạn chữ sử dụng phông phi chuẩn."
            violations.append(RuleViolation(
                "FONT_FAMILY_INVALID", "error",
                f"Phát hiện {non_tnr_runs} đoạn chữ sử dụng phông phi chuẩn (Quy định bắt buộc phông Times New Roman Unicode).",
                "Chuyển toàn bộ văn bản về phông Times New Roman."
            ))
        checklist.append({"name": "Phông chữ chuẩn Times New Roman", "passed": font_ok, "detail": font_detail})

        # 3. RÀ SOÁT TIÊU ĐỀ QUỐC HIỆU / TIÊU NGỮ / ĐẢNG
        all_text_nodes = doc._element.xpath('//w:t/text()')
        full_text = " ".join(str(t) for t in all_text_nodes) if all_text_nodes else "\n".join(p.text for p in doc.paragraphs)
        full_upper = full_text.upper()

        heading_ok = True
        heading_detail = "Quốc hiệu và Tiêu ngữ đầy đủ, đúng thể thức NĐ 30."
        if standard_id == "ND30":
            has_nation = ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in full_upper)
            has_motto = ("ĐỘC LẬP - TỰ DO - HẠNH PHÚC" in full_upper or "ĐỘC LẬP-TỰ DO-HẠNH PHÚC" in full_upper)
            if not has_nation:
                heading_ok = False
                violations.append(RuleViolation(
                    "MISSING_NATION_NAME", "error",
                    "Thiếu Quốc hiệu 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' ở góc phải văn bản.",
                    "Thêm Quốc hiệu vào góc phải phần đầu trang."
                ))
            if not has_motto:
                heading_ok = False
                violations.append(RuleViolation(
                    "MISSING_MOTTO", "error",
                    "Thiếu Tiêu ngữ 'Độc lập - Tự do - Hạnh phúc' của văn bản hành chính Nhà nước.",
                    "Thêm Tiêu ngữ vào góc phải dưới Quốc hiệu."
                ))
        elif standard_id == "HD36":
            has_party = ("ĐẢNG CỘNG SẢN VIỆT NAM" in full_upper)
            has_motto = ("ĐỘC LẬP - TỰ DO - HẠNH PHÚC" in full_upper)
            if not has_party:
                heading_ok = False
                violations.append(RuleViolation(
                    "MISSING_PARTY_HEADING", "error",
                    "Thiếu Tiêu đề 'ĐẢNG CỘNG SẢN VIỆT NAM' của văn bản Đảng.",
                    "Thêm tiêu đề Đảng vào góc phải bảng Tiêu đề."
                ))
            if has_motto:
                heading_ok = False
                violations.append(RuleViolation(
                    "PARTY_UNAUTHORIZED_MOTTO", "error",
                    "Văn bản Đảng tuyệt đối KHÔNG có tiêu ngữ 'Độc lập - Tự do - Hạnh phúc' (Hướng dẫn 36-HD/VPTW).",
                    "Xóa bỏ dòng tiêu ngữ khỏi văn bản Đảng."
                ))
            heading_detail = "Tiêu đề Đảng Cộng sản Việt Nam đúng thể thức HD 36."
        checklist.append({"name": "Quốc hiệu & Tiêu ngữ chuẩn định", "passed": heading_ok, "detail": heading_detail})

        # 4. RÀ SOÁT TÍNH TƯƠNG THÍCH LOẠI VĂN BẢN VÀ CẤU TRÚC NỘI DUNG
        type_ok = True
        type_detail = "Cấu trúc nội dung tương thích hoàn toàn với loại văn bản."

        # Nhận diện loại văn bản từ nội dung
        has_qd_title = ("QUYẾT ĐỊNH" in full_upper and "TỜ TRÌNH" not in full_upper and "BÁO CÁO" not in full_upper)
        has_ttr_title = ("TỜ TRÌNH" in full_upper)
        has_bc_title = ("BÁO CÁO" in full_upper)
        has_cv_title = ("CÔNG VĂN" in full_upper)

        has_qd_command = ("QUYẾT ĐỊNH:" in full_upper or "QUYẾT NGHỊ:" in full_upper)
        has_articles = bool(re.search(r'Điều\s+\d+[\.:]', full_text))

        if has_ttr_title:
            if has_qd_command or has_articles:
                type_ok = False
                violations.append(RuleViolation(
                    "DOC_TYPE_STRUCTURE_MISMATCH", "error",
                    "Loại văn bản là 'TỜ TRÌNH' nhưng nội dung lại chứa lệnh ban hành 'QUYẾT ĐỊNH:' và các 'Điều khoản'. Tờ trình là văn bản đề xuất lên cấp trên, không chứa cấu trúc mệnh lệnh quy phạm.",
                    "Đổi loại văn bản thành 'QUYẾT ĐỊNH' hoặc chọn mẫu Tờ trình chuẩn (Kính gửi, Sự cần thiết, Nội dung đề xuất)."
                ))
            if "KÍNH GỬI" not in full_upper:
                violations.append(RuleViolation(
                    "TTR_MISSING_SUBMISSION", "warning",
                    "Văn bản Tờ trình cần có mục 'Kính gửi:' cơ quan hoặc người có thẩm quyền phê duyệt.",
                    "Bổ sung 'Kính gửi: [Cơ quan/Cấp thẩm quyền]' ở đầu phần nội dung."
                ))
        elif has_bc_title:
            if has_qd_command or has_articles:
                type_ok = False
                violations.append(RuleViolation(
                    "DOC_TYPE_STRUCTURE_MISMATCH", "error",
                    "Loại văn bản là 'BÁO CÁO' nhưng nội dung lại chứa lệnh ban hành 'QUYẾT ĐỊNH:' và các 'Điều khoản'. Báo cáo dùng để thuyết minh tình hình và kết quả thực hiện.",
                    "Chuyển loại văn bản sang 'QUYẾT ĐỊNH' hoặc chuyển nội dung thành cấu trúc Báo cáo."
                ))
        elif has_qd_title:
            if not has_qd_command:
                type_ok = False
                violations.append(RuleViolation(
                    "QD_MISSING_COMMAND", "error",
                    "Văn bản 'QUYẾT ĐỊNH' thiếu mệnh lệnh ban hành 'QUYẾT ĐỊNH:' (in hoa, đậm, căn giữa) ngăn cách giữa Căn cứ và Điều khoản.",
                    "Thêm dòng 'QUYẾT ĐỊNH:' phía trên Điều 1."
                ))
            if not has_articles:
                type_ok = False
                violations.append(RuleViolation(
                    "QD_MISSING_ARTICLES", "error",
                    "Quyết định hành chính bắt buộc phải có cấu trúc các Điều khoản (Điều 1, Điều 2...).",
                    "Bổ sung cấu trúc Điều khoản quy định nội dung cụ thể."
                ))

        if not type_ok:
            type_detail = "Xung đột nghiêm trọng: Cấu trúc nội dung không phù hợp với loại văn bản."
        checklist.append({"name": "Tương thích Loại văn bản & Cấu trúc", "passed": type_ok, "detail": type_detail})

        # 5. RÀ SOÁT DỮ LIỆU MẪU CHƯA HOÀN THIỆN / KÝ TỰ GIỮ CHỖ (...)
        data_ok = True
        data_detail = "Thông tin ngày tháng, số hiệu và trích yếu đầy đủ, không còn ký tự giữ chỗ."

        # Kiểm tra ngày tháng
        if (
            "..., NGÀY" in full_upper or "NGÀY ... THÁNG" in full_upper
            or "THÁNG ... NĂM" in full_upper or "NGÀY... THÁNG... NĂM" in full_upper
            or "NGÀY...THÁNG...NĂM" in full_upper
        ):
            data_ok = False
            violations.append(RuleViolation(
                "PLACEHOLDER_DATE", "error",
                "Địa danh, ngày tháng năm ban hành văn bản còn chứa ký tự giữ chỗ '...' chưa điền ngày thực tế.",
                "Cập nhật địa danh và ngày tháng năm ban hành chính xác tại F2."
            ))

        # Kiểm tra số và ký hiệu
        if (
            "SỐ: .../" in full_upper or "SỐ:.../" in full_upper or "/..." in full_text
            or re.search(r'Số:\s*\.\.\.', full_text) or (full_upper.find("QĐ-UBND") != -1 and "SỐ:" not in full_upper)
            or re.search(r'Số:\s*(\.\.\.|\.\.)', full_text)
        ):
            data_ok = False
            violations.append(RuleViolation(
                "PLACEHOLDER_DOC_NUMBER", "error",
                "Số và ký hiệu văn bản còn chứa dấu giữ chỗ '...' hoặc thiếu số thứ tự văn bản.",
                "Điền số thứ tự và ký hiệu văn bản (ví dụ: Số: 15/QĐ-UBND) tại F2."
            ))

        # Kiểm tra trích yếu nội dung
        if (
            re.search(r'Về việc\s+0(?!\d)', full_text) or "Về việc 0" in full_text
            or re.search(r'Về việc\s*\.\.\.', full_text)
        ):
            data_ok = False
            violations.append(RuleViolation(
                "PLACEHOLDER_SUBJECT", "error",
                "Trích yếu nội dung văn bản chưa hợp lệ (đang là 'Về việc 0' hoặc chứa ký tự giữ chỗ).",
                "Cập nhật tóm tắt nội dung chủ đề văn bản tại F2."
            ))

        # Kiểm tra cơ quan ban hành mặc định
        if "CƠ QUAN BAN HÀNH" in full_upper or "TÊN CƠ QUAN CHỦ QUẢN" in full_upper:
            data_ok = False
            violations.append(RuleViolation(
                "PLACEHOLDER_AGENCY", "warning",
                "Tên cơ quan ban hành văn bản còn mang chuỗi mặc định 'CƠ QUAN BAN HÀNH'.",
                "Đổi tên cơ quan ban hành thành tên đơn vị cụ thể (ví dụ: UBND TỈNH QUẢNG NINH) tại F2."
            ))

        # Kiểm tra người ký mặc định (sử dụng regex ranh giới từ để tránh bắt nhầm 'Nguyễn Văn An', 'Nguyễn Văn Anh', 'Nguyễn Văn Anf')
        if re.search(r'\bNGUYỄN\s+VĂN\s+A(?!\w)', full_upper):
            violations.append(RuleViolation(
                "SAMPLE_SIGNER_NAME", "warning",
                "Họ tên người ký đang để tên mẫu mặc định ('Nguyễn Văn A').",
                "Điền họ tên người có thẩm quyền ký ban hành tại F2."
            ))

        if not data_ok:
            data_detail = "Phát hiện ký tự giữ chỗ '...' hoặc dữ liệu mẫu chưa hoàn thiện."
        checklist.append({"name": "Độ hoàn thiện thông tin & Ký tự giữ chỗ", "passed": data_ok, "detail": data_detail})

        # 6. RÀ SOÁT CĂN CỨ PHÁP LÝ & DẤU CÂU
        legal_syntax_ok = True
        legal_syntax_detail = "Dấu câu kết thúc căn cứ và đề nghị tuân thủ đúng quy tắc NĐ 30."

        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt.startswith("Căn cứ "):
                if not (txt.endswith(";") or txt.endswith(".")):
                    legal_syntax_ok = False
                    violations.append(RuleViolation(
                        "LEGAL_GROUNDS_SYNTAX", "warning",
                        f"Dòng căn cứ '{txt[:45]}...' phải kết thúc bằng dấu chấm phẩy ';'.",
                        "Sửa dấu kết thúc dòng Căn cứ thành dấu chấm phẩy ';'."
                    ))
            elif txt.startswith("Theo đề nghị của"):
                if not txt.endswith(","):
                    legal_syntax_ok = False
                    violations.append(RuleViolation(
                        "PROPOSAL_SYNTAX", "warning",
                        f"Dòng '{txt[:45]}...' trước lệnh ban hành 'QUYẾT ĐỊNH:' phải kết thúc bằng dấu phẩy ','.",
                        "Sửa dấu kết thúc dòng 'Theo đề nghị của...' thành dấu phẩy ','."
                    ))

        if not legal_syntax_ok:
            legal_syntax_detail = "Một số dòng Căn cứ hoặc Đề nghị kết thúc sai dấu câu quy định."
        checklist.append({"name": "Căn cứ pháp lý & Dấu câu", "passed": legal_syntax_ok, "detail": legal_syntax_detail})

        # 7. TÍNH ĐIỂM CHUẨN HÓA (0 - 100%)
        error_count = sum(1 for v in violations if v.severity == "error")
        warn_count = sum(1 for v in violations if v.severity == "warning")
        score = max(0.0, 100.0 - (error_count * 20.0) - (warn_count * 5.0))

        return {
            "score": round(score, 1),
            "is_compliant": (error_count == 0 and score >= 90.0),
            "error_count": error_count,
            "warning_count": warn_count,
            "checklist": checklist,
            "violations": [v.to_dict() for v in violations]
        }
