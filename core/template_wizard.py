"""
core/template_wizard.py: Bộ điều phối Mẫu văn bản & Khởi tạo Nhanh (Template Wizard Controller)
- Nạp thông tin định danh cơ quan từ config/agency_profiles.json
- Tự động điền khung mẫu chuẩn theo từng loại văn bản (Quyết định, Công văn, Tờ trình, Thông báo, Nghị quyết)
- Sinh tệp DOCX sẵn sàng đưa vào môi trường soạn thảo tích hợp (Embedded Word Processor)
"""

import os
import json
import datetime
from typing import Dict, Any, List, Optional
import docx

from core.path_utils import get_config_dir
from core.editor_core.openxml_builder import OpenXmlBuilder


class TemplateWizardController:
    """Bộ điều phối nạp mẫu và khởi tạo văn bản nhanh"""

    DOCUMENT_TYPES_ND30 = [
        "Quyết định (Cá biệt)",
        "Quyết định (Quy phạm pháp luật)",
        "Công văn hành chính",
        "Tờ trình",
        "Thông báo",
        "Kế hoạch",
        "Báo cáo",
        "Biên bản"
    ]

    DOCUMENT_TYPES_HD36 = [
        "Quyết định Đảng ủy",
        "Nghị quyết Đảng ủy",
        "Chỉ thị Đảng ủy",
        "Công văn Đảng ủy",
        "Kế hoạch Đảng ủy",
        "Báo cáo Đảng ủy",
        "Thông báo Đảng ủy"
    ]

    def __init__(self, config_dir: Optional[str] = None):
        self.config_dir = config_dir or get_config_dir()
        self.profiles_cache: List[Dict[str, Any]] = []
        self._load_profiles()

    def _load_profiles(self):
        prof_file = os.path.join(self.config_dir, "agency_profiles.json")
        if os.path.exists(prof_file):
            try:
                with open(prof_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.profiles_cache = data.get("profiles", [])
            except Exception:
                self.profiles_cache = []

    def get_profiles(self) -> List[Dict[str, Any]]:
        return self.profiles_cache

    def get_profile_by_id(self, profile_id: str) -> Optional[Dict[str, Any]]:
        for p in self.profiles_cache:
            if p.get("id") == profile_id:
                return p
        return None

    def get_available_document_types(self, standard_id: str = "ND30") -> List[str]:
        if standard_id.upper() == "HD36":
            return self.DOCUMENT_TYPES_HD36
        return self.DOCUMENT_TYPES_ND30

    def generate_document(
        self,
        profile_data: Dict[str, Any],
        doc_type_display: str,
        doc_subject: str,
        output_docx_path: str,
        custom_doc_number: str = "",
        custom_location: str = "",
        lock_scaffolding: bool = True
    ) -> docx.Document:
        """
        Khởi tạo tệp .docx hoàn chỉnh từ lựa chọn của người dùng trong Wizard
        """
        standard_id = profile_data.get("standard_id", "ND30").upper()
        # Chuẩn hóa tên loại văn bản (bỏ phần trong ngoặc nếu có)
        clean_doc_type = doc_type_display.split("(")[0].strip().upper()
        if clean_doc_type.endswith("CỦA ĐẢNG"):
            clean_doc_type = clean_doc_type.replace("CỦA ĐẢNG", "").strip()

        from core.administrative_sanitizer import AdministrativeSanitizer
        loc = custom_location or profile_data.get("location", "").strip()
        if not loc or loc in ["Xã", "Huyện", "Tỉnh", "Thành phố", "...", "…"]:
            loc = AdministrativeSanitizer.infer_location_from_agencies(
                profile_data.get("issuing_agency", ""),
                profile_data.get("superior_agency", "")
            )
        now = datetime.datetime.now()
        if loc and loc not in ["...", "…"]:
            date_str = f"{loc}, ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"
        else:
            date_str = f"..., ngày {now.day:02d} tháng {now.month:02d} năm {now.year}"

        # Xác định số hiệu văn bản
        doc_num = custom_doc_number
        if not doc_num:
            symbol = profile_data.get("doc_number_symbol", "UBND")
            if standard_id == "HD36":
                doc_num = f"...-{clean_doc_type[:2]}/{symbol}"
            else:
                doc_num = f".../{clean_doc_type[:2]}-{symbol}"

        meta = {
            "standard_id": standard_id,
            "superior_agency": profile_data.get("superior_agency", ""),
            "issuing_agency": profile_data.get("issuing_agency", ""),
            "doc_number": doc_num,
            "location": loc,
            "date_str": date_str,
            "doc_type": clean_doc_type,
            "doc_subject": doc_subject.strip(),
            "recipients": profile_data.get("recipients", []),
            "authority_prefix": profile_data.get("authority_prefix", ""),
            "signer_position": profile_data.get("signer_position", ""),
            "signer_name": profile_data.get("signer_name", "")
        }

        doc = OpenXmlBuilder.create_scaffolded_document(meta, lock_scaffolding=lock_scaffolding)
        
        # Đảm bảo thư mục lưu tồn tại
        out_dir = os.path.dirname(os.path.abspath(output_docx_path))
        os.makedirs(out_dir, exist_ok=True)
        doc.save(output_docx_path)
        return doc
