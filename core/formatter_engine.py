"""
Module 1: Bộ máy Định dạng Trang & Đoạn văn (Page & Paragraph Formatter)
Phối hợp cùng Module 2, 3, 4 để hoàn tất quy trình chuẩn hóa:
- Áp đặt thông số trang in (Khổ A4, lề chuẩn)
- Đánh số trang tự động ở đầu trang (ẩn ở trang 1)
- Ép kiểu phông chữ Times New Roman, cỡ chữ, giãn dòng, lùi đầu dòng cho toàn bộ văn bản
- Xử lý các tiêu đề phân mục (Điều, Khoản, Mục, Chương)
- Định dạng bảng biểu nội dung
- Tích hợp SpellLinter để tự động dọn rác ký tự và cảnh báo chính tả hành chính
"""

import os
import re
import shutil
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple
import docx
from docx.shared import Pt, Mm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from core.spell_linter import SpellLinter, LintResult, LintIssue
from core.doc_parser import (
    DocParser, ParsedDocStructure,
    paragraph_has_drawings_or_images, run_has_drawings_or_images
)
from core.header_footer_builder import HeaderFooterBuilder, format_run
from core.administrative_sanitizer import AdministrativeSanitizer


@dataclass
class FormattingReport:
    input_file: str
    output_file: str
    backup_file: str
    standard_id: str
    standard_name: str
    paragraphs_formatted: int = 0
    tables_formatted: int = 0
    spaces_fixed: int = 0
    punctuation_fixed: int = 0
    quotes_fixed: int = 0
    dates_fixed: int = 0
    cap_warnings_count: int = 0
    spelling_fixed: int = 0
    all_issues: List[LintIssue] = field(default_factory=list)
    success: bool = True
    error_message: str = ""
    is_scanned_image: bool = False       # Phát hiện tài liệu dạng ảnh scan
    image_count: int = 0                 # Số lượng hình ảnh scan trong tài liệu
    safe_mode_activated: bool = False    # Chế độ bảo vệ hình ảnh đã được kích hoạt
    ocr_recommended: bool = False


def resolve_clean_output_path(input_path: str, target_ext: str = ".docx") -> str:
    """
    Tạo đường dẫn xuất tệp chuẩn hóa sạch sẽ:
    - Loại bỏ triệt để các hậu tố lặp thừa thãi như _ChuanHoa_ChuanHoa, _KySo...
    - Nếu tệp đầu vào vốn đã là tệp chuẩn hóa (vd: doc_ChuanHoa.docx), giữ nguyên tên
      để cập nhật đè trực tiếp (in-place), không sinh thêm tệp rác mới.
    - Tự động xóa các tệp rác bị lặp đuôi từ các lần chạy trước trong cùng thư mục.
    """
    folder = os.path.dirname(os.path.abspath(input_path))
    base_name = os.path.basename(input_path)
    stem, _ = os.path.splitext(base_name)

    # Loại bỏ sạch sẽ chuỗi lặp _ChuanHoa, _KySo ở đuôi
    clean_stem = re.sub(r'(_ChuanHoa|_KySo|_signed)+$', '', stem, flags=re.IGNORECASE).strip()

    output_filename = f"{clean_stem}_ChuanHoa{target_ext}"
    output_path = os.path.join(folder, output_filename)

    # Quét dọn các tệp rác chuỗi lặp thừa thãi trong thư mục (vd: doc_ChuanHoa_ChuanHoa*.docx/pdf)
    try:
        stale_pattern = re.compile(
            rf'^{re.escape(clean_stem)}_ChuanHoa(_ChuanHoa)+.*\.(docx|pdf)$',
            re.IGNORECASE
        )
        for f in os.listdir(folder):
            if stale_pattern.match(f):
                stale_full = os.path.join(folder, f)
                try:
                    os.remove(stale_full)
                except Exception:
                    pass
    except Exception:
        pass

    return output_path
    cap_warnings_count: int = 0
    spelling_fixed: int = 0
    all_issues: List[LintIssue] = field(default_factory=list)
    success: bool = True
    error_message: str = ""
    is_scanned_image: bool = False       # Phát hiện tài liệu dạng ảnh scan
    image_count: int = 0                 # Số lượng hình ảnh scan trong tài liệu
    safe_mode_activated: bool = False    # Chế độ bảo vệ hình ảnh đã được kích hoạt
    ocr_recommended: bool = False        # Khuyến nghị thực hiện OCR để lấy chữ


HEADING_PATTERNS = [
    re.compile(r'^(CHƯƠNG|CHUONG)\s+[IVXLCDM\d]+[\.:]?', re.IGNORECASE),
    re.compile(r'^(MỤC|MUC|PHẦN|PHAN)\s+\d+[\.:]?', re.IGNORECASE),
    re.compile(r'^(Điều|Dieu)\s+\d+[\.:]?', re.IGNORECASE),
    re.compile(r'^([IVXLCDM]+\.|\d+\.)\s+[A-Za-zÀ-ỹ]', re.UNICODE)
]


from core.path_utils import get_config_dir


class FormatterEngine:
    def __init__(self, config_dir: Optional[str] = None):
        if config_dir is None:
            config_dir = get_config_dir()
        self.config_dir = config_dir
        self.linter = SpellLinter(config_dir=config_dir)

        import json
        with open(os.path.join(config_dir, "hd36_rules.json"), "r", encoding="utf-8") as f:
            self.rules_hd36 = json.load(f)
        with open(os.path.join(config_dir, "nd30_rules.json"), "r", encoding="utf-8") as f:
            self.rules_nd30 = json.load(f)

        self.hf_builder = HeaderFooterBuilder(self.rules_hd36, self.rules_nd30)

    def get_rules(self, standard_id: str) -> Dict[str, Any]:
        return self.rules_hd36 if standard_id.upper() == "HD36" else self.rules_nd30

    def format_document(
        self,
        input_docx_path: str,
        output_docx_path: Optional[str] = None,
        standard_id: str = "ND30",
        profile_override: Optional[Dict[str, Any]] = None
    ) -> FormattingReport:
        """
        Xử lý toàn diện văn bản theo 4 bước được đặc tả.
        """
        if not output_docx_path:
            output_docx_path = resolve_clean_output_path(input_docx_path, target_ext=os.path.splitext(input_docx_path)[1])

        # Bảo vệ toàn vẹn dữ liệu: Cập nhật an toàn qua tệp tạm thời nếu ghi đè tệp gốc
        is_same_file = (os.path.abspath(output_docx_path) == os.path.abspath(input_docx_path))
        temp_target_path = f"{output_docx_path}.tmp" if is_same_file else output_docx_path

        rules = self.get_rules(standard_id)
        report = FormattingReport(
            input_file=input_docx_path,
            output_file=output_docx_path,
            backup_file="",
            standard_id=standard_id,
            standard_name=rules["name"]
        )

        from core.security_guard import SecurityGuard
        ok_sec, sec_msg = SecurityGuard.validate_docx_package(input_docx_path)
        if not ok_sec:
            report.success = False
            report.error_message = f"Từ chối tệp vì lý do an toàn bảo mật: {sec_msg}"
            return report

        try:
            doc = docx.Document(input_docx_path)

            # 0. Chuẩn hóa toàn bộ văn bản sang Unicode NFC chuẩn quốc gia
            import unicodedata
            for p in doc.paragraphs:
                if p.text:
                    p.text = unicodedata.normalize('NFC', p.text)
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if p.text:
                                p.text = unicodedata.normalize('NFC', p.text)

            # 1. Hợp nhất các đoạn văn bị ngắt dòng vụn bởi OCR để phục hồi câu hoàn chỉnh
            AdministrativeSanitizer.merge_fragmented_paragraphs(doc)
            # 2. Nhận diện và chuyển đổi danh sách đính kèm dạng bảng thành Table Word
            AdministrativeSanitizer.convert_ocr_list_to_table(doc)

            parser = DocParser(doc)
            parser.doc_path = input_docx_path
            structure = parser.parse()

            # Tự động đồng bộ chuẩn Đảng (HD36) nếu tài liệu có nội dung Đảng rõ rệt
            if standard_id == "ND30" and structure.detected_standard == "HD36":
                standard_id = "HD36"
                rules = self.get_rules("HD36")
                report.standard_id = "HD36"
                report.standard_name = rules["name"]

            # Kiểm tra nếu là tài liệu dạng ảnh scan (Scanned / Image Document)
            if structure.is_scanned_image:
                report.is_scanned_image = True
                report.image_count = structure.image_count
                report.safe_mode_activated = True
                report.ocr_recommended = True

                # Kích hoạt CHẾ ĐỘ BẢO VỆ ẢNH (Image Safe Mode):
                # Không xóa đoạn văn, không chèn đè bảng Header/Footer làm lệch bố cục ảnh scan
                self._apply_page_setup(doc, rules)

                report.all_issues.append(LintIssue(
                    issue_type="scanned_image",
                    original="[Tài liệu dạng ảnh scan]",
                    replacement="[Khuyến nghị chạy OCR để chuyển thành văn bản chỉnh sửa được]",
                    message=(
                        f"Phát hiện tài liệu dạng ảnh scan (chứa {structure.image_count} hình ảnh chụp trang, "
                        f"chỉ có {structure.text_word_count} từ văn bản có thể chỉnh sửa). Đã tự động kích hoạt "
                        "Chế độ Bảo vệ Hình ảnh (Safe Mode) để bảo toàn 100% hình ảnh không bị xóa hoặc làm lệch bố cục. "
                        "Khuyến nghị sử dụng chức năng 'Nhận dạng OCR & Chuẩn hóa' để bóc tách thành văn bản chuẩn."
                    ),
                    severity="warning"
                ))
                doc.save(temp_target_path)
                if is_same_file and os.path.exists(temp_target_path):
                    os.replace(temp_target_path, output_docx_path)
                report.success = True
                return report

            # 2. Định dạng thông số trang in & đánh số trang
            self._apply_page_setup(doc, rules)

            # 3. Dựng / Cập nhật Khung Đầu Văn bản (Module 2)
            self._apply_header_table(doc, structure, rules, standard_id, profile_override)

            # 4. Định dạng các đoạn văn (Body & Headings) + Chạy Linter (Module 1 & 4)
            self._apply_paragraphs_formatting(doc, structure, rules, standard_id, report)

            # 5. Định dạng các bảng biểu nội dung
            self._apply_tables_formatting(doc, structure, rules, standard_id, report)

            # 6. Dựng / Cập nhật Khung Cuối Văn bản (Nơi nhận & Chữ ký - Module 3)
            self._apply_footer_table(doc, structure, rules, standard_id, profile_override)

            # 7. Lưu file kết quả an toàn
            doc.save(temp_target_path)
            if is_same_file and os.path.exists(temp_target_path):
                os.replace(temp_target_path, output_docx_path)
            report.success = True

        except Exception as e:
            report.success = False
            report.error_message = str(e)

        return report

    def format_scanned_document_with_ocr(
        self,
        input_file: str,
        output_docx_path: Optional[str] = None,
        standard_id: str = "ND30",
        profile_override: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[callable] = None
    ) -> FormattingReport:
        """
        Quy trình xử lý hoàn chỉnh cho tệp PDF scan hoặc DOCX scan:
        1. Sử dụng RapidOCR (chạy ngoại tuyến 100%) để nhận dạng toàn bộ chữ từ các bức ảnh scan
        2. Tự động chuyển giao sang quy trình chuẩn hóa thể thức toàn diện (NĐ 30 / HD 36)
        """
        if not output_docx_path:
            output_docx_path = resolve_clean_output_path(input_file, target_ext=".docx")

        import tempfile
        sec_temp_dir = tempfile.mkdtemp(prefix="af_sec_")
        temp_ocr_docx = os.path.join(sec_temp_dir, "ocr_temp.docx")

        try:
            from core.ocr_engine import OcrEngine
            ocr_engine = OcrEngine()

            if progress_callback:
                progress_callback(10, "Đang trích xuất hình ảnh scan và nhận dạng ký tự quang học (OCR)...")

            ok, msg, line_count = ocr_engine.convert_scanned_file_to_editable_docx(
                input_path=input_file,
                output_docx_path=temp_ocr_docx,
                progress_callback=progress_callback
            )

            if not ok:
                rules = self.get_rules(standard_id)
                return FormattingReport(
                    input_file=input_file,
                    output_file=output_docx_path,
                    backup_file="",
                    standard_id=standard_id,
                    standard_name=rules["name"],
                    success=False,
                    error_message=f"Nhận dạng OCR không thành công: {msg}"
                )

            if progress_callback:
                progress_callback(60, f"Đã nhận dạng {line_count} dòng chữ. Đang chuẩn hóa thể thức theo {standard_id}...")

            # Chạy chuẩn hóa trên file đã nhận dạng chữ
            report = self.format_document(
                input_docx_path=temp_ocr_docx,
                output_docx_path=output_docx_path,
                standard_id=standard_id,
                profile_override=profile_override
            )
            report.input_file = input_file

            # Thêm thông báo vào báo cáo
            report.all_issues.insert(0, LintIssue(
                issue_type="ocr_processed",
                original="[Ảnh scan gốc]",
                replacement=f"[Đã bóc tách thành {line_count} dòng chữ chỉnh sửa được]",
                message=f"Đã thực hiện nhận dạng chữ OCR ngoại tuyến thành công ({line_count} dòng) và chuẩn hóa toàn diện thể thức văn bản.",
                severity="auto_fixed"
            ))

            return report

        finally:
            # Bảo mật dữ liệu: Tiêu hủy hoàn toàn thư mục và tệp tạm bằng ghi đè (Secure Shredding)
            if os.path.exists(sec_temp_dir):
                from core.security_guard import SecurityGuard
                SecurityGuard.secure_delete(sec_temp_dir)

    def _apply_page_setup(self, doc: docx.Document, rules: Dict[str, Any]):
        """
        Thiết lập khổ giấy A4, định lượng lề và đánh số trang tự động
        """
        page_cfg = rules["page"]
        pn_cfg = rules["page_number"]

        for section in doc.sections:
            section.page_width = Mm(page_cfg["width_mm"])
            section.page_height = Mm(page_cfg["height_mm"])
            section.top_margin = Mm(page_cfg["margin_top_mm"])
            section.bottom_margin = Mm(page_cfg["margin_bottom_mm"])
            section.left_margin = Mm(page_cfg["margin_left_mm"])
            section.right_margin = Mm(page_cfg["margin_right_mm"])

            # Thiết lập đánh số trang
            section.different_first_page_header_footer = pn_cfg.get("hide_first_page", True)

            # Header cho các trang sau trang 1
            header = section.header
            p_hdr = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            p_hdr.text = ""
            p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_hdr.paragraph_format.space_before = Pt(0)
            p_hdr.paragraph_format.space_after = Pt(0)

            run_pn = p_hdr.add_run()
            format_run(
                run_pn,
                size_pt=pn_cfg["font_size_pt"],
                bold=pn_cfg.get("bold", False),
                italic=pn_cfg.get("italic", False)
            )

            # Thêm trường trường số trang XML {PAGE}
            fld_xml = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
            run_pn._r.append(fld_xml)

            # Trang đầu tiên để trống header số trang
            first_header = section.first_page_header
            if first_header.paragraphs:
                first_header.paragraphs[0].text = ""

    def _apply_header_table(
        self,
        doc: docx.Document,
        structure: ParsedDocStructure,
        rules: Dict[str, Any],
        standard_id: str,
        profile: Optional[Dict[str, Any]]
    ):
        """
        Dựng hoặc tái cấu trúc bảng Tiêu đề & Quốc hiệu ở đầu văn bản
        """
        # Xác định thông tin cơ quan
        superior = profile.get("superior_agency", "") if profile else structure.header.superior_agency
        issuing = profile.get("issuing_agency", "") if profile else structure.header.issuing_agency
        doc_num = profile.get("doc_number_symbol", "") if profile else structure.header.doc_number
        loc = profile.get("location", "") if profile else (structure.header.location_date.split(",")[0].strip() if "," in structure.header.location_date else "")

        # TỰ ĐỘNG LÀM SẠCH VÀ BẢO ĐẢM TÍNH ĐỒNG BỘ:
        # Nếu là chuẩn Đảng (HD36), tuyệt đối không để sót tên UBND / Nhà nước từ bảng cũ
        if standard_id.upper() == "HD36":
            if any(k in superior.upper() for k in ["ỦY BAN NHÂN DÂN", "UBND", "HỘI ĐỒNG NHÂN DÂN", "THÀNH PHỐ HÀ NỘI"]):
                superior = ""
            if any(k in issuing.upper() for k in ["ỦY BAN NHÂN DÂN", "UBND", "THÀNH PHỐ HÀ NỘI", "CƠ QUAN BAN HÀNH"]):
                issuing = ""
            if "UBND" in doc_num or "Số: ..../...." in doc_num:
                doc_num = ""

            meta = AdministrativeSanitizer.extract_party_doc_metadata(doc, structure.doc_path or "")
            if not superior:
                superior = meta["superior_agency"]
            if not issuing:
                issuing = meta["issuing_agency"]
            if not loc or loc == "Hà Nội":
                loc = meta["location"]
            if not doc_num:
                doc_num = meta["doc_number"]

        date_str = structure.header.location_date
        if not date_str and profile:
            date_str = "ngày ... tháng ... năm 2026"

        if structure.header.has_existing_table:
            # Cập nhật bảng hiện có
            table = doc.tables[0]
            self.hf_builder.build_header_table(
                table, standard_id, superior, issuing, doc_num, loc, date_str
            )
        elif profile or structure.header.raw_paragraphs_count > 0:
            # Chỉ tạo bảng khi có profile do người dùng chọn hoặc trích xuất được từ các đoạn văn gốc
            if structure.header.raw_paragraphs_count > 0:
                for idx in reversed(range(structure.header.raw_paragraphs_count)):
                    p = doc.paragraphs[idx]
                    if paragraph_has_drawings_or_images(p):
                        continue
                    p._element.getparent().remove(p._element)

            # Thêm bảng 2 cột lên đầu
            table = doc.add_table(rows=1, cols=2)
            doc.element.body.insert(0, table._tbl)
            self.hf_builder.build_header_table(
                table, standard_id, superior, issuing, doc_num, loc, date_str
            )

    def _apply_paragraphs_formatting(
        self,
        doc: docx.Document,
        structure: ParsedDocStructure,
        rules: Dict[str, Any],
        standard_id: str,
        report: FormattingReport
    ):
        """
        Định dạng phông chữ, cỡ chữ, căn lề, khoảng cách đoạn và lùi đầu dòng cho từng paragraph
        Đồng thời chạy Linter kiểm tra lỗi và cập nhật thống kê
        """
        typo = rules["typography"]
        font_name = typo["font_name"]
        body_font_size = typo["body_font_size_pt"]
        line_spacing = typo["body_line_spacing"]
        space_before = Pt(typo["body_space_before_pt"])
        space_after = Pt(typo["body_space_after_pt"])
        first_line_indent = Mm(typo["body_first_line_indent_mm"])

        for p in doc.paragraphs:
            raw_text = p.text.strip()
            has_drawing = paragraph_has_drawings_or_images(p)

            # Nếu đoạn văn chỉ có ảnh mà không có chữ, giữ nguyên tuyệt đối
            if not raw_text:
                continue

            # Phục hồi lỗi OCR và chuẩn hóa văn phong hành chính
            healed_text, heal_count, heal_issues = AdministrativeSanitizer.heal_ocr_text_with_issues(p.text)
            if heal_count > 0:
                if not has_drawing:
                    p.text = healed_text
                report.spelling_fixed += heal_count
                report.all_issues.extend(heal_issues)

            # Chạy Linter dọn dẹp nội dung văn bản
            lint_res = self.linter.lint_text(p.text, standard_id=standard_id)
            if lint_res.issues:
                report.all_issues.extend(lint_res.issues)
                report.spaces_fixed += lint_res.spaces_fixed
                report.punctuation_fixed += lint_res.punctuation_fixed
                report.quotes_fixed += lint_res.quotes_fixed
                report.dates_fixed += lint_res.dates_fixed
                report.cap_warnings_count += lint_res.capitalization_warnings
                report.spelling_fixed += lint_res.spelling_fixed

            if lint_res.cleaned_text != p.text:
                if not has_drawing:
                    p.text = lint_res.cleaned_text
                else:
                    # Nếu đoạn văn chứa ảnh, CHỈ cập nhật các run thuần text, không gán p.text
                    for r in p.runs:
                        if not run_has_drawings_or_images(r) and r.text:
                            r_clean, _ = self.linter.clean_whitespace(r.text)
                            r.text = r_clean

            text = p.text.strip()

            # Kiểm tra và tự động sửa dấu câu thể thức hành chính (Administrative Punctuation Rules)
            if not has_drawing and text:
                # A. Căn cứ pháp lý: bắt buộc kết thúc bằng dấu chấm phẩy (;)
                if re.match(r'^[-–—•*]?\s*(căn\s+cứ|can\s+cur|can\s+cu)\b', text, re.IGNORECASE):
                    t_curr = p.text.rstrip()
                    if t_curr and not t_curr.endswith(';'):
                        orig_char = t_curr[-1] if t_curr[-1] in ('.', ',') else "[thiếu dấu ;]"
                        if t_curr[-1] in ('.', ','):
                            p.text = t_curr[:-1] + ';'
                        else:
                            p.text = t_curr + ';'
                        report.punctuation_fixed += 1
                        report.all_issues.append(LintIssue(
                            issue_type="punctuation",
                            original=orig_char,
                            replacement=";",
                            message="Quy định thể thức: Đoạn căn cứ pháp lý phải kết thúc bằng dấu chấm phẩy (;)",
                            severity="auto_fixed"
                        ))

                # B. Theo đề nghị: bắt buộc kết thúc bằng dấu phẩy (,)
                elif re.match(r'^(theo\s+đề\s+nghị|theo\s+de\s+nghi)\b', text, re.IGNORECASE):
                    t_curr = p.text.rstrip()
                    if t_curr and not t_curr.endswith(','):
                        orig_char = t_curr[-1] if t_curr[-1] in ('.', ';', ':') else "[thiếu dấu ,]"
                        if t_curr[-1] in ('.', ';', ':'):
                            p.text = t_curr[:-1] + ','
                        else:
                            p.text = t_curr + ','
                        report.punctuation_fixed += 1
                        report.all_issues.append(LintIssue(
                            issue_type="punctuation",
                            original=orig_char,
                            replacement=",",
                            message="Quy định thể thức: Dòng 'Theo đề nghị của...' phải kết thúc bằng dấu phẩy (,)",
                            severity="auto_fixed"
                        ))

                # C. Tiêu đề Điều khoản (Điều 1., Điều 2....): chữ cái đầu tiên sau Điều khoản phải viết hoa, kết thúc bằng dấu chấm (.)
                elif re.match(r'^(điều|dieu)\s+\d+', text, re.IGNORECASE):
                    m_dieu = re.match(r'^(điều|dieu)\s+(\d+)[\.:]?\s*(.*)$', p.text, re.IGNORECASE | re.DOTALL)
                    if m_dieu:
                        d_num = m_dieu.group(2)
                        rest = m_dieu.group(3)
                        if rest and rest[0].islower():
                            capitalized_rest = rest[0].upper() + rest[1:]
                            p.text = f"Điều {d_num}. {capitalized_rest}"
                            report.cap_warnings_count += 1
                            report.all_issues.append(LintIssue(
                                issue_type="capitalization",
                                original=f"Điều {d_num}. {rest[:15]}...",
                                replacement=f"Điều {d_num}. {capitalized_rest[:15]}...",
                                message="Viết hoa chữ cái đầu tiên sau tiêu đề Điều khoản",
                                severity="auto_fixed"
                            ))
                    t_curr = p.text.rstrip()
                    if t_curr and not t_curr.endswith('.'):
                        orig_char = t_curr[-1] if t_curr[-1] in (';', ',') else "[thiếu dấu .]"
                        if t_curr[-1] in (';', ','):
                            p.text = t_curr[:-1] + '.'
                        else:
                            p.text = t_curr + '.'
                        report.punctuation_fixed += 1
                        report.all_issues.append(LintIssue(
                            issue_type="punctuation",
                            original=orig_char,
                            replacement=".",
                            message="Quy định thể thức: Đoạn kết thúc Điều khoản phải có dấu chấm (.)",
                            severity="auto_fixed"
                        ))

                # D. Quyết định có hiệu lực: kết thúc bằng dấu chấm (.)
                elif re.search(r'có hiệu lực kể từ ngày ký', text, re.IGNORECASE):
                    t_curr = p.text.rstrip()
                    if t_curr and not t_curr.endswith('.'):
                        orig_char = t_curr[-1] if t_curr[-1] in (';', ',') else "[thiếu dấu .]"
                        if t_curr[-1] in (';', ','):
                            p.text = t_curr[:-1] + '.'
                        else:
                            p.text = t_curr + '.'
                        report.punctuation_fixed += 1
                        report.all_issues.append(LintIssue(
                            issue_type="punctuation",
                            original=orig_char,
                            replacement=".",
                            message="Quy định thể thức: Câu hiệu lực văn bản phải kết thúc bằng dấu chấm (.)",
                            severity="auto_fixed"
                        ))

            text = p.text.strip()

            # 1. Kiểm tra Tiêu đề văn bản (QUYẾT ĐỊNH, NGHỊ QUYẾT, BAN THƯỜNG VỤ... QUYẾT ĐỊNH)
            is_title = (
                any(t in text.upper() for t in [
                    "QUYẾT ĐỊNH", "NGHỊ QUYẾT", "QUYẾT NGHỊ", "CHỈ THỊ", "KẾT LUẬN",
                    "KẾ HOẠCH", "BÁO CÁO", "TỜ TRÌNH", "THÔNG BÁO", "CÔNG VĂN"
                ]) and (len(text.split()) <= 15 or text.upper().endswith("QUYẾT ĐỊNH") or text.upper().endswith("QUYẾT NGHỊ"))
            )
            if is_title and not text.startswith("Căn cứ") and not text.startswith("Theo đề nghị"):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Mm(0)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.15
                for run in p.runs:
                    if not run_has_drawings_or_images(run):
                        format_run(run, font_name=font_name, size_pt=rules["subject"]["doc_type_font_size_pt"], bold=True)
                report.paragraphs_formatted += 1
                continue

            # 2. Kiểm tra DANH SÁCH / PHỤ LỤC / BẢNG BIỂU (Chỉ khớp tiêu đề độc lập, không khớp câu trong bài)
            is_danh_sach = (
                re.match(r'^(DANH\s*SÁCH|DANH\s*SACH|PHỤ\s*LỤC|PHU\s*LUC|BẢNG|BANG)(\s+[IVXLCDM\d]+)?\s*$', text, re.IGNORECASE) or
                (re.match(r'^(DANH\s*SÁCH|DANH\s*SACH|PHỤ\s*LỤC|PHU\s*LUC|BẢNG|BANG)\b', text, re.IGNORECASE) and len(text.split()) <= 4 and "kèm theo" not in text.lower() and "kem theo" not in text.lower())
            )
            if is_danh_sach:
                p.paragraph_format.page_break_before = True
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Mm(0)
                p.paragraph_format.space_before = Pt(14)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.15
                for run in p.runs:
                    if not run_has_drawings_or_images(run):
                        format_run(run, font_name=font_name, size_pt=14, bold=True)
                report.paragraphs_formatted += 1
                continue

            # 3. Kiểm tra Trích yếu nội dung (Về việc...) hoặc phụ đề danh sách
            if text.lower().startswith("về việc") or text.lower().startswith("v/v") or (text.lower().startswith("phân công") and ("nhiệm kỳ" in text.lower() or "nhiem ky" in text.lower() or "chi bộ" in text.lower())):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Mm(0)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(12)
                p.paragraph_format.line_spacing = 1.15
                for run in p.runs:
                    if not run_has_drawings_or_images(run):
                        format_run(run, font_name=font_name, size_pt=rules["subject"]["summary_font_size_pt"], bold=True)
                report.paragraphs_formatted += 1
                continue

            # 3b. Kiểm tra Dòng ghi chú dưới bảng (Lưu ý: ...)
            if re.match(r'^(lưu\s*ý|luu\s*y)\s*:', text, re.IGNORECASE):
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Mm(0)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.15
                for run in p.runs:
                    if not run_has_drawings_or_images(run):
                        format_run(run, font_name=font_name, size_pt=12, italic=True)
                report.paragraphs_formatted += 1
                continue

            # 4. Kiểm tra Căn cứ pháp lý ("Căn cứ Luật...", "Căn cứ Nghị định...", "Theo đề nghị...")
            if re.match(r'^[-–—•*]?\s*(căn\s+cứ|theo\s+đề\s+nghị)\b', text, re.IGNORECASE):
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = first_line_indent
                p.paragraph_format.space_before = space_before
                p.paragraph_format.space_after = space_after
                p.paragraph_format.line_spacing = line_spacing
                for run in p.runs:
                    if not run_has_drawings_or_images(run):
                        format_run(run, font_name=font_name, size_pt=body_font_size, italic=True)
                report.paragraphs_formatted += 1
                continue

            # 5. Kiểm tra Tiêu đề mục (Điều 1, Chương I, 1. ...)
            is_heading = any(pattern.match(text) for pattern in HEADING_PATTERNS)
            if is_heading:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = first_line_indent
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = line_spacing
                for run in p.runs:
                    if not run_has_drawings_or_images(run):
                        format_run(run, font_name=font_name, size_pt=body_font_size, bold=True)
                report.paragraphs_formatted += 1
                continue

            # 6. Các đoạn văn thân bài thông thường (Body)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = first_line_indent
            p.paragraph_format.space_before = space_before
            p.paragraph_format.space_after = space_after
            p.paragraph_format.line_spacing = line_spacing

            for run in p.runs:
                if run_has_drawings_or_images(run):
                    continue
                # Giữ nguyên in đậm/in nghiêng gốc của văn bản nếu có, chỉ chuẩn hóa font & size
                is_bold = run.font.bold if run.font.bold is not None else False
                is_italic = run.font.italic if run.font.italic is not None else False
                format_run(run, font_name=font_name, size_pt=body_font_size, bold=is_bold, italic=is_italic)

            report.paragraphs_formatted += 1

    def _apply_tables_formatting(
        self,
        doc: docx.Document,
        structure: ParsedDocStructure,
        rules: Dict[str, Any],
        standard_id: str,
        report: FormattingReport
    ):
        """
        Định dạng bảng biểu nội dung (bỏ qua Header Table và Footer Table)
        """
        font_name = rules["typography"]["font_name"]

        for idx, table in enumerate(doc.tables):
            # Bỏ qua bảng đầu nếu là Header Table
            if idx == 0 and structure.header.has_existing_table:
                continue
            # Bỏ qua bảng cuối nếu là Footer Table
            if idx == len(doc.tables) - 1 and structure.footer.has_existing_table:
                continue

            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if not p.text.strip():
                            continue

                        # Phục hồi lỗi OCR và chính tả trong từng ô bảng
                        healed_text, heal_count, heal_issues = AdministrativeSanitizer.heal_ocr_text_with_issues(p.text)
                        if heal_count > 0:
                            p.text = healed_text
                            report.spelling_fixed += heal_count
                            report.all_issues.extend(heal_issues)

                        lint_res = self.linter.lint_text(p.text, standard_id=standard_id)
                        if lint_res.issues:
                            report.all_issues.extend(lint_res.issues)
                            report.spaces_fixed += lint_res.spaces_fixed
                            report.punctuation_fixed += lint_res.punctuation_fixed
                            report.quotes_fixed += lint_res.quotes_fixed
                            report.dates_fixed += lint_res.dates_fixed
                            report.cap_warnings_count += lint_res.capitalization_warnings
                            report.spelling_fixed += lint_res.spelling_fixed

                        if lint_res.cleaned_text != p.text:
                            p.text = lint_res.cleaned_text

                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(1)
                        p.paragraph_format.line_spacing = 1.1

                        for run in p.runs:
                            is_b = run.font.bold if run.font.bold is not None else False
                            is_i = run.font.italic if run.font.italic is not None else False
                            format_run(run, font_name=font_name, size_pt=12.5, bold=is_b, italic=is_i)

            report.tables_formatted += 1

    def _apply_footer_table(
        self,
        doc: docx.Document,
        structure: ParsedDocStructure,
        rules: Dict[str, Any],
        standard_id: str,
        profile: Optional[Dict[str, Any]]
    ):
        """
        Dựng hoặc tái cấu trúc khung Nơi nhận & Chữ ký ở cuối văn bản
        """
        recipients = profile.get("recipients", []) if profile else structure.footer.recipients
        if not recipients and structure.footer.recipients:
            recipients = structure.footer.recipients

        auth = profile.get("authority_prefix", "") if profile else structure.footer.authority_prefix
        pos = profile.get("signer_position", "") if profile else structure.footer.signer_position
        name = profile.get("signer_name", "") if profile else structure.footer.signer_name

        # 1. Tìm trực tiếp vị trí paragraph bắt đầu phần Nơi nhận trong tài liệu
        footer_start_p = None
        footer_paras = []
        for p in list(doc.paragraphs):
            t_low = p.text.strip().lower()
            if footer_start_p is None:
                if t_low.startswith("nơi nhận") or t_low.startswith("noi nhan"):
                    footer_start_p = p
                    footer_paras.append(p)
            else:
                if re.match(r'^(danh\s*sách|danh\s*sach|phụ\s*lục|phu\s*luc|bảng|bang|điều\s+\d+)\b', t_low, re.IGNORECASE) and "kèm theo" not in t_low and "kem theo" not in t_low:
                    break
                footer_paras.append(p)

        if footer_start_p is not None:
            # Phân tích nội dung nơi nhận và chữ ký từ các đoạn văn tìm thấy
            lines_text = [p.text for p in footer_paras]
            parsed_footer = AdministrativeSanitizer.parse_in_text_footer(lines_text)
            if parsed_footer:
                recipients = parsed_footer["recipients"] or recipients
                auth = parsed_footer["authority_prefix"] or auth
                pos = parsed_footer["signer_position"] or pos
                name = parsed_footer["signer_name"] or name

            # Chèn Bảng 2 cột ẩn viền ngay trước đoạn văn Nơi nhận
            table = doc.add_table(rows=1, cols=2)
            footer_start_p._p.addprevious(table._tbl)

            # Xóa triệt để các đoạn văn footer cũ bị dính chữ
            for p in footer_paras:
                if not paragraph_has_drawings_or_images(p):
                    try:
                        p._element.getparent().remove(p._element)
                    except Exception:
                        pass

            # Xóa các bảng footer rác hoặc bảng cũ từ lần chuẩn hóa trước
            header_tbl_elem = doc.tables[0]._tbl if doc.tables else None
            for tbl in list(doc.tables):
                if tbl._tbl is not table._tbl and tbl._tbl is not header_tbl_elem:
                    t_str = " ".join(c.text for r in tbl.rows for c in r.cells)
                    if any(k in t_str for k in ["Lê Văn Cường", "Le Van Cuong", "Nguyễn Văn A"]) or (standard_id == "HD36" and "ỦY BAN NHÂN DÂN" in t_str):
                        try:
                            tbl._tbl.getparent().remove(tbl._tbl)
                        except Exception:
                            pass

            self.hf_builder.build_footer_table(
                table, standard_id, recipients, auth, pos, name
            )
            return

        if structure.footer.has_existing_table:
            # Cập nhật bảng cuối hiện có
            table = doc.tables[-1]
            self.hf_builder.build_footer_table(
                table, standard_id, recipients, auth, pos, name
            )
            return
        elif profile:
            # Thêm bảng 2 cột vào cuối tài liệu nếu người dùng chủ động chọn áp dụng profile
            table = doc.add_table(rows=1, cols=2)
            self.hf_builder.build_footer_table(
                table, standard_id, recipients, auth, pos, name
            )
