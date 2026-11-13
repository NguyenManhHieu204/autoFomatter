"""
core/batch_editor.py: Bộ xử lý chỉnh sửa hàng loạt văn bản Word (Batch Document Editor)
- Hỗ trợ đổi địa danh hàng loạt (Thuận Hạnh, Đức An, Nam Bình...)
- Hỗ trợ xóa ngày/đổi ngày hàng loạt (chuyển sang dự thảo ngày    tháng    năm 2026, ngày ..., hoặc ngày cố định)
- Hỗ trợ chỉnh sửa số hiệu văn bản hàng loạt:
  Chuyển dấu chấm thành khoảng trống dự thảo (Số: ............../abc -> Số:               /abc)
  Đổi ký hiệu viết tắt, gán số hiệu cụ thể
- Tự động chuẩn hóa ngắt dòng cơ quan ban hành theo chuẩn NĐ 30:
  Dòng 1: ỦY BAN NHÂN DÂN
  Dòng 2: XÃ THUẬN HẠNH (hoặc XÃ ĐỨC AN...)
- Hỗ trợ Tìm kiếm & Thay thế văn bản hàng loạt trong toàn bộ nội dung và bảng biểu
- Hỗ trợ xuất sang thư mục riêng hoặc cập nhật đè an toàn
"""

import os
import re
import datetime
import unicodedata
import uuid
from dataclasses import dataclass, field
from typing import List, Optional, Callable, Dict, Any, Tuple
import docx
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from core.administrative_sanitizer import AdministrativeSanitizer
from core.header_footer_builder import format_run


@dataclass
class BatchRenameOptions:
    """Tùy chọn đổi tên tệp hàng loạt (Batch Rename Options)"""
    remove_chuanhoa: bool = True          # Xóa '_ChuanHoa', '_chuanhoa', '_chuan hoa', chuẩn hóa
    remove_custom_text: str = ""          # Xóa cụm từ tùy chỉnh (vd: - copy, _fixed)
    enable_find_replace: bool = False     # Bật Tìm kiếm & Thay thế trong tên tệp
    find_text: str = ""                   # Từ khóa cần tìm
    replace_text: str = ""                # Thay thế bằng
    match_case: bool = False              # Phân biệt chữ hoa / thường
    prefix: str = ""                      # Tiền tố (thêm vào đầu tên)
    suffix: str = ""                      # Hậu tố (thêm vào cuối tên trước phần mở rộng)
    case_transform: str = "keep"          # "keep", "lower", "upper", "title"
    enable_numbering: bool = False        # Đánh số thứ tự tự động
    numbering_start: int = 1              # Số bắt đầu
    numbering_digits: int = 2             # Số chữ số (2 -> 01, 02; 3 -> 001, 002)
    numbering_pos: str = "prefix"         # "prefix" (01_Doc.docx) hoặc "suffix" (Doc_01.docx)
    collision_policy: str = "auto_index"  # "auto_index" (thêm (1), (2)), "skip", "overwrite"


@dataclass
class BatchRenameItemResult:
    old_path: str
    new_path: str
    old_name: str
    new_name: str
    success: bool
    status: str       # "renamed", "unchanged", "collision_resolved", "skipped", "failed"
    message: str


@dataclass
class BatchEditOptions:
    # 1. Địa danh
    enable_location: bool = False
    new_location: str = ""
    sync_agency_location: bool = True  # Đồng bộ tên đơn vị cơ quan (ví dụ: XÃ ĐỨC AN) theo địa danh mới

    # 2. Số hiệu văn bản
    enable_doc_number: bool = False
    doc_number_mode: str = "dots_to_spaces"  # "dots_to_spaces", "spaces_to_dots", "custom"
    new_symbol: str = ""                     # Thay đổi ký hiệu sau dấu / (vd: TB-ĐKĐĐ, QĐ-UBND)
    custom_doc_number: str = ""              # Nhập nguyên chuỗi số hiệu tùy chỉnh

    # 3. Ngày tháng
    enable_date: bool = False
    date_mode: str = "draft_spaces"  # "draft_spaces", "draft_dots", "clear", "today", "custom"
    custom_date_text: str = ""

    # 4. Cơ quan ban hành
    enable_agency: bool = False
    new_superior_agency: str = ""
    new_issuing_agency: str = ""
    auto_fix_agency_split: bool = True  # Chuẩn hóa ngắt dòng đúng: Dòng 1: ỦY BAN NHÂN DÂN, Dòng 2: XÃ [TÊN]

    # 5. Tìm & Thay thế văn bản
    enable_find_replace: bool = False
    find_text: str = ""
    replace_text: str = ""
    match_case: bool = False

    # 6. Nơi lưu tệp
    output_dir: Optional[str] = None  # None = Ghi đè tệp gốc
    auto_remove_chuanhoa_in_output: bool = False  # Tự động xóa hậu tố _ChuanHoa ở tên tệp khi xuất sang thư mục riêng


@dataclass
class BatchEditItemResult:
    file_path: str
    output_path: str
    success: bool
    changes_count: int = 0
    message: str = ""
    details: List[str] = field(default_factory=list)


class BatchEditor:
    """Bộ máy chỉnh sửa hàng loạt tệp DOCX hiệu năng cao"""

    @staticmethod
    def format_doc_number(cur_num: str, mode: str, new_symbol: str = "", custom_text: str = "") -> str:
        """
        Chuẩn hóa số hiệu văn bản hàng loạt:
        - "dots_to_spaces": Chuyển dấu chấm thành khoảng trống dự thảo (Số: ............../abc -> Số:               /abc)
        - "spaces_to_dots": Chuyển khoảng trống thành dấu chấm (Số: ............../abc)
        - "custom": Gán chuỗi số hiệu cụ thể
        """
        if mode == "custom" and custom_text.strip():
            res = custom_text.strip()
            if not re.search(r'^(số|so)\b', res, re.IGNORECASE):
                res = f"Số: {res}"
            return res

        raw = cur_num.strip() if cur_num else "Số: /TB-ĐKĐĐ"
        m_prefix = re.match(r'^(số\s*[:\.]?\s*|so\s*[:\.]?\s*)', raw, re.IGNORECASE)
        prefix = "Số: "
        if m_prefix:
            prefix = "Số: "
            raw_rest = raw[m_prefix.end():].strip()
        else:
            raw_rest = raw

        # Tách phần số và phần ký hiệu (sau dấu / hoặc -)
        separator = "/"
        if "/" in raw_rest:
            parts = raw_rest.split("/", 1)
            num_part = parts[0].strip()
            sym_part = parts[1].strip()
            separator = "/"
        elif "-" in raw_rest and any(k in raw_rest.upper() for k in ["QĐ", "NQ", "ĐU", "UBND", "HDND"]):
            parts = raw_rest.split("-", 1)
            num_part = parts[0].strip()
            sym_part = parts[1].strip()
            separator = "-"
        else:
            num_part = raw_rest.strip()
            sym_part = ""

        # Nếu người dùng muốn đổi ký hiệu mới
        if new_symbol.strip():
            clean_sym = new_symbol.strip().lstrip("/").lstrip("-").strip()
            sym_part = clean_sym
            separator = "/"

        # Tạo chuỗi số hiệu: 15 khoảng trắng giữa dấu : và / chuẩn theo Nghị định 30
        blank_spaces = "              "  # 14 khoảng trắng (kết hợp với 'Số: ' thành đúng 15 khoảng trắng)
        dots = ".............."

        if mode == "dots_to_spaces":
            if sym_part:
                return f"{prefix}{blank_spaces}{separator}{sym_part}"
            else:
                return f"{prefix}{blank_spaces}"
        elif mode == "spaces_to_dots":
            if sym_part:
                return f"{prefix}{dots}{separator}{sym_part}"
            else:
                return f"{prefix}{dots}"
        else:
            if sym_part:
                return f"{prefix}{num_part}{separator}{sym_part}"
            else:
                return f"{prefix}{num_part}"

    @classmethod
    def process_file(cls, file_path: str, options: BatchEditOptions, target_output_override: Optional[str] = None) -> BatchEditItemResult:
        """Thực hiện cập nhật một tệp Word duy nhất theo BatchEditOptions"""
        if not os.path.exists(file_path):
            return BatchEditItemResult(
                file_path=file_path,
                output_path="",
                success=False,
                message=f"Tệp không tồn tại: {file_path}"
            )

        if not file_path.lower().endswith(".docx"):
            return BatchEditItemResult(
                file_path=file_path,
                output_path="",
                success=False,
                message="Chỉ hỗ trợ tệp định dạng .docx"
            )

        try:
            doc = docx.Document(file_path)
        except Exception as e:
            return BatchEditItemResult(
                file_path=file_path,
                output_path="",
                success=False,
                message=f"Lỗi mở tệp DOCX: {e}"
            )

        details = []
        changes = 0

        # Xác định thư mục lưu
        if target_output_override:
            target_output = target_output_override
        elif options.output_dir and os.path.exists(options.output_dir):
            out_name = os.path.basename(file_path)
            if getattr(options, 'auto_remove_chuanhoa_in_output', False):
                out_name = cls.compute_new_filename(out_name, BatchRenameOptions(remove_chuanhoa=True))
            target_output = os.path.join(options.output_dir, out_name)
        else:
            target_output = file_path

        # 1. XỬ LÝ KHUNG TIÊU ĐỀ (HEADER TABLE: CƠ QUAN, SỐ HIỆU, ĐỊA DANH, NGÀY THÁNG)
        header_table = None
        for tbl in doc.tables[:2]:
            if len(tbl.columns) >= 2 and len(tbl.rows) >= 1:
                t_check = " ".join(c.text for r in tbl.rows for c in r.cells).upper()
                if any(k in t_check for k in ["CỘNG HÒA", "CONG HOA", "ĐỘC LẬP", "DOC LAP", "ĐẢNG CỘNG SẢN", "ỦY BAN", "UBND"]):
                    header_table = tbl
                    break

        if header_table:
            cell_left = header_table.cell(0, 0)
            cell_right = header_table.cell(0, 1)

            # --- A. Cột Phải: Địa danh & Ngày tháng ---
            need_date_update = options.enable_date or options.enable_location
            if need_date_update:
                date_p = None
                date_p_has_motto = False
                for p in cell_right.paragraphs:
                    text = p.text.strip()
                    if re.search(r'\b(?:ngày|ngay)\b', text, re.IGNORECASE) or any(k in text for k in ["năm 20", "nam 20", "năm..."]):
                        date_p = p
                        if any(k in text.upper() for k in ["CỘNG HÒA", "CONG HOA", "ĐỘC LẬP", "DOC LAP"]):
                            date_p_has_motto = True
                        break

                current_text = date_p.text.strip() if date_p else ""
                cur_loc, cur_date = AdministrativeSanitizer.extract_location_and_date(current_text)

                # Xác định địa danh mới
                if options.enable_location and options.new_location.strip():
                    target_loc = AdministrativeSanitizer.clean_location_name(options.new_location.strip())
                else:
                    target_loc = cur_loc or "Thuận Hạnh"

                # Xác định ngày tháng mới
                if options.enable_date:
                    # Trích xuất năm cũ nếu có
                    m_year = re.search(r'\b(20\d{2})\b', current_text)
                    target_year = m_year.group(1) if m_year else str(datetime.datetime.now().year)

                    if options.date_mode == "draft_spaces":
                        new_date_str = f"{target_loc}, ngày    tháng    năm {target_year}"
                    elif options.date_mode == "draft_dots":
                        new_date_str = f"{target_loc}, ngày ... tháng ... năm {target_year}"
                    elif options.date_mode == "clear":
                        new_date_str = f"{target_loc}, ngày ... tháng ... năm ..."
                    elif options.date_mode == "today":
                        now = datetime.datetime.now()
                        new_date_str = f"{target_loc}, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"
                    elif options.date_mode == "custom" and options.custom_date_text.strip():
                        c_text = options.custom_date_text.strip()
                        if not re.search(r'\b(?:ngày|ngay)\b', c_text, re.IGNORECASE):
                            c_text = f"ngày {c_text}"
                        new_date_str = f"{target_loc}, {c_text}"
                    else:
                        new_date_str = f"{target_loc}, ngày    tháng    năm {target_year}"
                else:
                    # Chỉ đổi địa danh, bảo lưu ngày tháng gốc
                    new_date_str = f"{target_loc}, {cur_date}" if cur_date else f"{target_loc}, ngày ... tháng ... năm 2026"

                if date_p_has_motto:
                    # Trường hợp toàn bộ quốc hiệu và ngày dính trong 1 paragraph
                    for p in cell_right.paragraphs[1:]:
                        p._element.getparent().remove(p._element)
                    p_m0 = cell_right.paragraphs[0]
                    p_m0.text = ""
                    p_m0.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r1 = p_m0.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
                    format_run(r1, font_name="Times New Roman", size_pt=12.0, bold=True)
                    r2 = p_m0.add_run("Độc lập - Tự do - Hạnh phúc")
                    format_run(r2, font_name="Times New Roman", size_pt=13.0, bold=True)

                    p_div = cell_right.add_paragraph()
                    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_div.paragraph_format.space_before = Pt(1)
                    p_div.paragraph_format.space_after = Pt(4)
                    r_dv = p_div.add_run("—————————")
                    format_run(r_dv, font_name="Times New Roman", size_pt=9.0, bold=True)

                    date_p = cell_right.add_paragraph()

                if date_p is None:
                    date_p = cell_right.add_paragraph()

                date_p.text = ""
                date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                date_p.paragraph_format.space_before = Pt(4)
                date_p.paragraph_format.space_after = Pt(0)
                r_dt = date_p.add_run(new_date_str)
                format_run(r_dt, font_name="Times New Roman", size_pt=13.5, italic=True)
                changes += 1
                details.append(f"Cập nhật ngày/địa danh: {new_date_str}")

            # --- B. Cột Trái: Chuẩn hóa ngắt dòng, Tên Cơ quan & Số hiệu ---
            need_agency_update = options.enable_agency or options.auto_fix_agency_split or (options.enable_location and options.sync_agency_location)
            need_num_update = options.enable_doc_number

            # Trích xuất thông tin cơ quan và số hiệu hiện tại (tách theo từng dòng kể cả khi có \n)
            p_lines = []
            for p in cell_left.paragraphs:
                for subline in p.text.split("\n"):
                    sub = subline.strip()
                    if sub and not sub.startswith("—") and not sub.startswith("-") and not sub.startswith("_"):
                        p_lines.append(sub)

            cur_sup = ""
            cur_iss = ""
            cur_num = ""

            for line in p_lines:
                l_u = unicodedata.normalize('NFC', line).upper()
                if (l_u.startswith("SỐ") or "Số:" in line or "Số " in line or 
                    l_u.startswith("SO ") or l_u.startswith("S6 ") or 
                    "SỐ:" in l_u or "SO:" in l_u or "SỐ." in l_u or "SO." in l_u or
                    ("/" in line and any(k in l_u for k in ["TB", "QĐ", "QD", "UBND", "BC", "KH", "TT", "NQ", "CV"]))):
                    cur_num = line
                elif not cur_sup:
                    cur_sup = line
                elif not cur_iss:
                    cur_iss = line

            if need_agency_update or need_num_update:
                # 1. Xác định tên cơ quan mới
                if options.enable_agency:
                    target_sup = options.new_superior_agency.strip().upper()
                    target_iss = options.new_issuing_agency.strip().upper()
                elif options.auto_fix_agency_split:
                    # Tự động chuẩn hóa ngắt dòng theo chuẩn NĐ 30
                    target_sup, target_iss = AdministrativeSanitizer.split_and_normalize_agency_header(cur_sup, cur_iss)
                else:
                    target_sup, target_iss = cur_sup, cur_iss

                # Đồng bộ tên đơn vị cơ quan theo địa danh mới nếu được bật
                if options.enable_location and options.sync_agency_location and options.new_location.strip():
                    loc_u = options.new_location.strip().upper()
                    if re.match(r'^(XÃ|XA)\b', target_iss, re.IGNORECASE) or not target_iss:
                        target_iss = f"XÃ {loc_u}"
                    elif re.match(r'^(PHƯỜNG|PHUONG)\b', target_iss, re.IGNORECASE):
                        target_iss = f"PHƯỜNG {loc_u}"
                    elif re.match(r'^(THỊ TRẤN|THI TRAN)\b', target_iss, re.IGNORECASE):
                        target_iss = f"THỊ TRẤN {loc_u}"
                    elif re.match(r'^(HUYỆN|HUYEN)\b', target_iss, re.IGNORECASE):
                        target_iss = f"HUYỆN {loc_u}"
                    elif target_iss == "ỦY BAN NHÂN DÂN" or not target_iss:
                        target_sup = "ỦY BAN NHÂN DÂN"
                        target_iss = f"XÃ {loc_u}"

                # 2. Xác định số hiệu mới
                if need_num_update:
                    final_num = cls.format_doc_number(
                        cur_num,
                        mode=options.doc_number_mode,
                        new_symbol=options.new_symbol,
                        custom_text=options.custom_doc_number
                    )
                else:
                    final_num = cur_num if cur_num else "Số:      /TB-ĐKĐĐ"

                # Tái cấu trúc ô cell_left sạch sẽ
                for p in cell_left.paragraphs[1:]:
                    p._element.getparent().remove(p._element)

                p_l0 = cell_left.paragraphs[0]
                p_l0.text = ""
                p_l0.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l0.paragraph_format.space_before = Pt(0)
                p_l0.paragraph_format.space_after = Pt(2)
                p_l0.paragraph_format.line_spacing = 1.1

                if target_sup:
                    r_s = p_l0.add_run(target_sup + "\n")
                    format_run(r_s, font_name="Times New Roman", size_pt=12.5, bold=False)

                r_i = p_l0.add_run(target_iss if target_iss else "ỦY BAN NHÂN DÂN")
                format_run(r_i, font_name="Times New Roman", size_pt=12.5, bold=True)

                p_l_div = cell_left.add_paragraph()
                p_l_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l_div.paragraph_format.space_before = Pt(1)
                p_l_div.paragraph_format.space_after = Pt(4)
                r_d = p_l_div.add_run("—————")
                format_run(r_d, font_name="Times New Roman", size_pt=9.0, bold=True)

                p_l_num = cell_left.add_paragraph()
                p_l_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l_num.paragraph_format.space_before = Pt(2)
                p_l_num.paragraph_format.space_after = Pt(0)
                r_n = p_l_num.add_run(final_num)
                format_run(r_n, font_name="Times New Roman", size_pt=12.5, bold=False)

                changes += 1
                if need_agency_update:
                    details.append(f"Chuẩn hóa cơ quan: {target_sup} | {target_iss}")
                if need_num_update:
                    details.append(f"Chuẩn hóa số hiệu: {final_num}")

        # 2. TÌM & THAY THẾ VĂN BẢN TOÀN TÀI LIỆU
        if options.enable_find_replace and options.find_text:
            f_txt = options.find_text
            r_txt = options.replace_text
            pattern = re.compile(re.escape(f_txt), 0 if options.match_case else re.IGNORECASE)

            # Thay thế trong paragraphs
            for p in doc.paragraphs:
                if pattern.search(p.text):
                    p.text = pattern.sub(r_txt, p.text)
                    changes += 1

            # Thay thế trong tất cả tables
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if pattern.search(p.text):
                                p.text = pattern.sub(r_txt, p.text)
                                changes += 1
            details.append(f"Tìm & thay thế: '{f_txt}' -> '{r_txt}'")

        # Lưu tệp
        try:
            # Ghi an toàn nguyên tử (atomic write)
            is_same = (os.path.normcase(os.path.abspath(target_output)) == os.path.normcase(os.path.abspath(file_path)))
            temp_path = f"{file_path}.tmp_{uuid.uuid4().hex[:6]}.docx"
            doc.save(temp_path)
            if is_same:
                os.replace(temp_path, target_output)
            else:
                os.replace(temp_path, target_output)
                if target_output_override and (os.path.dirname(os.path.abspath(target_output)) == os.path.dirname(os.path.abspath(file_path))):
                    # Khi lưu đổi tên trong cùng thư mục: xóa file cũ để tránh nhân bản tệp rác
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                        except Exception:
                            pass
        except Exception as e:
            return BatchEditItemResult(
                file_path=file_path,
                output_path=target_output,
                success=False,
                message=f"Lỗi lưu tệp DOCX: {e}"
            )

        return BatchEditItemResult(
            file_path=file_path,
            output_path=target_output,
            success=True,
            changes_count=changes,
            message=f"Đã cập nhật thành công ({changes} mục)",
            details=details
        )

    @classmethod
    def process_batch(
        cls,
        file_paths: List[str],
        options: BatchEditOptions,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[BatchEditItemResult]:
        """Thực hiện xử lý hàng loạt trên danh sách tệp"""
        results = []
        total = len(file_paths)
        for idx, fp in enumerate(file_paths):
            fname = os.path.basename(fp)
            if progress_callback:
                progress_callback(idx, total, f"Đang xử lý ({idx + 1}/{total}): {fname}")

            res = cls.process_file(fp, options)
            results.append(res)

        if progress_callback:
            progress_callback(total, total, f"Hoàn tất xử lý {total} tệp.")

        return results

    @classmethod
    def process_combined_batch(
        cls,
        file_paths: List[str],
        edit_options: Optional[BatchEditOptions],
        rename_options: Optional[BatchRenameOptions],
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[BatchEditItemResult]:
        """
        Thực hiện xử lý hàng loạt KẾT HỢP:
        Vừa chỉnh sửa nội dung văn bản NĐ 30, vừa đổi tên tệp cùng lúc trong 1 lần chạy duy nhất.
        Tiết kiệm gấp đôi thao tác và thời gian cho người dùng.
        """
        results = []
        total = len(file_paths)

        # Lấy trước ánh xạ tên tệp mới nếu có rename_options
        previews = cls.preview_rename(file_paths, rename_options) if rename_options else None
        target_map = {p["old_path"]: p["target_path"] for p in previews} if previews else {}

        for idx, fp in enumerate(file_paths):
            fname = os.path.basename(fp)
            if progress_callback:
                progress_callback(idx, total, f"Đang xử lý ({idx + 1}/{total}): {fname}")

            target_out = target_map.get(fp, None)

            # Trường hợp 1: Có chỉnh sửa nội dung và là tệp DOCX
            if edit_options and fp.lower().endswith(".docx"):
                res = cls.process_file(fp, edit_options, target_output_override=target_out)
                results.append(res)
            # Trường hợp 2: Chỉ đổi tên tệp (hoặc tệp không phải docx)
            elif target_out and os.path.abspath(target_out) != os.path.abspath(fp):
                dir_name = os.path.dirname(fp)
                orig_norm = os.path.normcase(os.path.abspath(fp))
                target_norm = os.path.normcase(os.path.abspath(target_out))
                try:
                    if orig_norm == target_norm and fname != os.path.basename(target_out):
                        temp_name = f"__tmp_ren_{uuid.uuid4().hex[:8]}_{os.path.basename(target_out)}"
                        temp_path = os.path.join(dir_name, temp_name)
                        os.replace(fp, temp_path)
                        os.replace(temp_path, target_out)
                    else:
                        os.replace(fp, target_out)
                    results.append(BatchEditItemResult(
                        file_path=fp,
                        output_path=target_out,
                        success=True,
                        changes_count=1,
                        message=f"Đã đổi tên tệp: {os.path.basename(target_out)}",
                        details=[f"Đổi tên: {fname} -> {os.path.basename(target_out)}"]
                    ))
                except Exception as e:
                    results.append(BatchEditItemResult(
                        file_path=fp,
                        output_path=fp,
                        success=False,
                        changes_count=0,
                        message=f"Lỗi đổi tên: {e}"
                    ))
            else:
                results.append(BatchEditItemResult(
                    file_path=fp,
                    output_path=fp,
                    success=True,
                    changes_count=0,
                    message="Giữ nguyên (không có thay đổi)",
                    details=[]
                ))

        if progress_callback:
            progress_callback(total, total, f"Hoàn tất xử lý kết hợp {total} tệp.")

        return results

    @classmethod
    def compute_new_filename(cls, filename: str, options: BatchRenameOptions, index: int = 0) -> str:
        """
        Tính toán tên tệp mới dựa trên các tùy chọn đổi tên:
        - Xóa _ChuanHoa, _chuanhoa, chuẩn hóa
        - Xóa chuỗi tùy chọn
        - Tìm kiếm & Thay thế
        - Đổi chữ hoa / chữ thường
        - Thêm tiền tố / hậu tố
        - Đánh số thứ tự
        """
        stem, ext = os.path.splitext(filename)

        # 1. Xóa hậu tố _ChuanHoa, _chuanhoa, _chuẩn hóa
        if options.remove_chuanhoa:
            stem = re.sub(r'[\s_-]*(chu\u1ea9n[\s_-]*h\u00f3a|chuan[\s_-]*hoa)', '', stem, flags=re.IGNORECASE)
            stem = re.sub(r'_+', '_', stem)
            stem = re.sub(r'\s+', ' ', stem)
            stem = re.sub(r'^[\s_-]+|[\s_-]+$', '', stem).strip()

        # 2. Xóa cụm từ tùy chọn khác (vd: - copy, _fixed)
        if options.remove_custom_text.strip():
            pat = re.compile(re.escape(options.remove_custom_text.strip()), 0 if options.match_case else re.IGNORECASE)
            stem = pat.sub('', stem)
            stem = re.sub(r'_+', '_', stem)
            stem = re.sub(r'\s+', ' ', stem)
            stem = re.sub(r'^[\s_-]+|[\s_-]+$', '', stem).strip()

        # 3. Tìm & Thay thế trong tên tệp
        if options.enable_find_replace and options.find_text:
            pat = re.compile(re.escape(options.find_text), 0 if options.match_case else re.IGNORECASE)
            stem = pat.sub(options.replace_text, stem)

        # 4. Định dạng chữ hoa / chữ thường
        if options.case_transform == "lower":
            stem = stem.lower()
        elif options.case_transform == "upper":
            stem = stem.upper()
        elif options.case_transform == "title":
            stem = stem.title()

        # 5. Tiền tố & Hậu tố
        if options.prefix:
            stem = f"{options.prefix}{stem}"
        if options.suffix:
            stem = f"{stem}{options.suffix}"

        # 6. Đánh số thứ tự
        if options.enable_numbering:
            num_str = str(options.numbering_start + index).zfill(options.numbering_digits)
            if options.numbering_pos == "prefix":
                stem = f"{num_str}_{stem}"
            else:
                stem = f"{stem}_{num_str}"

        # 7. Fallback an toàn nếu tên bị rỗng
        if not stem.strip():
            orig = os.path.splitext(filename)[0]
            stem = orig if orig.strip() else "Van_Ban"

        return f"{stem}{ext}"

    @classmethod
    def preview_rename(cls, file_paths: List[str], options: BatchRenameOptions) -> List[Dict[str, Any]]:
        """
        Mô phỏng đổi tên tệp hàng loạt và phát hiện trùng lặp trước khi thực hiện.
        Trả về danh sách dict:
        - old_path, old_name, new_name, target_path, changed, status, status_type
        """
        results = []
        claimed_paths = set()

        for idx, fp in enumerate(file_paths):
            old_name = os.path.basename(fp)
            dir_name = os.path.dirname(fp)

            if not os.path.exists(fp):
                results.append({
                    "old_path": fp,
                    "old_name": old_name,
                    "new_name": old_name,
                    "target_path": fp,
                    "changed": False,
                    "status": "Tệp không tồn tại",
                    "status_type": "error"
                })
                continue

            base_new_name = cls.compute_new_filename(old_name, options, idx)
            stem, ext = os.path.splitext(base_new_name)

            # Kiểm tra nếu tên không thay đổi
            if base_new_name == old_name:
                results.append({
                    "old_path": fp,
                    "old_name": old_name,
                    "new_name": base_new_name,
                    "target_path": fp,
                    "changed": False,
                    "status": "Giữ nguyên",
                    "status_type": "unchanged"
                })
                claimed_paths.add(os.path.normcase(os.path.abspath(fp)))
                continue

            # Kiểm tra va chạm tên tệp
            tentative_path = os.path.join(dir_name, base_new_name)
            tentative_norm = os.path.normcase(os.path.abspath(tentative_path))
            orig_norm = os.path.normcase(os.path.abspath(fp))

            is_case_only = (tentative_norm == orig_norm)

            if is_case_only:
                # Đổi chữ hoa/thường trên cùng tệp
                results.append({
                    "old_path": fp,
                    "old_name": old_name,
                    "new_name": base_new_name,
                    "target_path": tentative_path,
                    "changed": True,
                    "status": "Đổi kiểu chữ",
                    "status_type": "rename"
                })
                claimed_paths.add(tentative_norm)
                continue

            # Kiểm tra nếu file đích đã tồn tại hoặc đã được gán bởi tệp khác trong danh sách
            collision = (tentative_norm in claimed_paths or os.path.exists(tentative_path))

            if collision:
                if options.collision_policy == "auto_index":
                    cnt = 1
                    while True:
                        cand_name = f"{stem} ({cnt}){ext}"
                        cand_path = os.path.join(dir_name, cand_name)
                        cand_norm = os.path.normcase(os.path.abspath(cand_path))
                        if cand_norm not in claimed_paths and not os.path.exists(cand_path):
                            claimed_paths.add(cand_norm)
                            results.append({
                                "old_path": fp,
                                "old_name": old_name,
                                "new_name": cand_name,
                                "target_path": cand_path,
                                "changed": True,
                                "status": f"Trùng tên -> thêm ({cnt})",
                                "status_type": "collision"
                            })
                            break
                        cnt += 1
                elif options.collision_policy == "skip":
                    results.append({
                        "old_path": fp,
                        "old_name": old_name,
                        "new_name": base_new_name,
                        "target_path": tentative_path,
                        "changed": False,
                        "status": "Bỏ qua (Trùng tên)",
                        "status_type": "skipped"
                    })
                elif options.collision_policy == "overwrite":
                    claimed_paths.add(tentative_norm)
                    results.append({
                        "old_path": fp,
                        "old_name": old_name,
                        "new_name": base_new_name,
                        "target_path": tentative_path,
                        "changed": True,
                        "status": "Ghi đè tệp cũ",
                        "status_type": "rename"
                    })
            else:
                claimed_paths.add(tentative_norm)
                results.append({
                    "old_path": fp,
                    "old_name": old_name,
                    "new_name": base_new_name,
                    "target_path": tentative_path,
                    "changed": True,
                    "status": "Sẽ đổi tên",
                    "status_type": "rename"
                })

        return results

    @classmethod
    def execute_rename(
        cls,
        file_paths: List[str],
        options: BatchRenameOptions,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[BatchRenameItemResult]:
        """
        Thực hiện đổi tên tệp hàng loạt an toàn trên hệ thống tệp.
        Hỗ trợ giải quyết trùng lặp, an toàn với Windows NTFS case-insensitivity,
        bắt lỗi khi tệp đang bị khóa bởi Word.
        """
        previews = cls.preview_rename(file_paths, options)
        results = []
        total = len(previews)

        for idx, item in enumerate(previews):
            old_path = item["old_path"]
            old_name = item["old_name"]
            new_name = item["new_name"]
            target_path = item["target_path"]
            status_type = item["status_type"]

            if progress_callback:
                progress_callback(idx, total, f"Đang đổi tên ({idx + 1}/{total}): {old_name} -> {new_name}")

            if status_type == "error":
                results.append(BatchRenameItemResult(
                    old_path=old_path,
                    new_path=old_path,
                    old_name=old_name,
                    new_name=new_name,
                    success=False,
                    status="failed",
                    message="Tệp không tồn tại"
                ))
                continue

            if not item["changed"] or status_type == "unchanged":
                results.append(BatchRenameItemResult(
                    old_path=old_path,
                    new_path=old_path,
                    old_name=old_name,
                    new_name=new_name,
                    success=True,
                    status="unchanged",
                    message="Tên tệp không thay đổi"
                ))
                continue

            if status_type == "skipped":
                results.append(BatchRenameItemResult(
                    old_path=old_path,
                    new_path=old_path,
                    old_name=old_name,
                    new_name=new_name,
                    success=False,
                    status="skipped",
                    message="Bỏ qua do trùng lặp tên"
                ))
                continue

            # Thực hiện đổi tên
            dir_name = os.path.dirname(old_path)
            orig_norm = os.path.normcase(os.path.abspath(old_path))
            target_norm = os.path.normcase(os.path.abspath(target_path))

            try:
                if orig_norm == target_norm and old_name != new_name:
                    # Đổi chữ hoa / chữ thường trên Windows (cần bước trung gian)
                    temp_name = f"__tmp_ren_{uuid.uuid4().hex[:8]}_{new_name}"
                    temp_path = os.path.join(dir_name, temp_name)
                    os.replace(old_path, temp_path)
                    os.replace(temp_path, target_path)
                else:
                    os.replace(old_path, target_path)

                results.append(BatchRenameItemResult(
                    old_path=old_path,
                    new_path=target_path,
                    old_name=old_name,
                    new_name=new_name,
                    success=True,
                    status="renamed" if status_type == "rename" else "collision_resolved",
                    message=f"Đã đổi tên thành công: {new_name}"
                ))
            except PermissionError:
                results.append(BatchRenameItemResult(
                    old_path=old_path,
                    new_path=old_path,
                    old_name=old_name,
                    new_name=new_name,
                    success=False,
                    status="failed",
                    message="Tệp đang được mở bởi ứng dụng khác (Word) - Vui lòng đóng tệp trước"
                ))
            except Exception as e:
                results.append(BatchRenameItemResult(
                    old_path=old_path,
                    new_path=old_path,
                    old_name=old_name,
                    new_name=new_name,
                    success=False,
                    status="failed",
                    message=f"Lỗi: {e}"
                ))

        if progress_callback:
            progress_callback(total, total, f"Hoàn tất đổi tên {total} tệp.")

        return results
