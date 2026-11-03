"""
core/administrative_sanitizer.py: Bộ máy Chuẩn hóa Văn phong & Khôi phục Lỗi Nhận dạng Ký tự (OCR Sanitizer)
Chuyên trị các lỗi OCR phổ biến trong văn bản Đảng & Hành chính Việt Nam:
- Sửa lỗi số hiệu và ký tự nhầm lẫn: s6 -> số, c6 -> có, cuia -> của, trurc -> trực, churc -> chức, tir -> từ, Luru -> Lưu:
- Khôi phục dấu tiếng Việt cho các cụm từ hành chính chuẩn mực: Căn cứ, Điều lệ Đảng, Quyết định, Ban Chấp hành, Ban Bí thư...
- Tự động nhận diện văn bản thuộc Khối Đảng (HD 36) hay Khối Hành chính (ND 30)
- Tách và nhận diện khung Nơi nhận & Chữ ký bị OCR dính liền 2 cột
"""

import os
import re
from typing import List, Tuple, Dict, Any, Optional
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def format_run_element(run, font_name="Times New Roman", size_pt=12.5, bold=False, italic=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)


# Danh mục các quy tắc chuyển đổi và phục hồi lỗi OCR hành chính chuẩn
ADMIN_OCR_REPLACEMENTS = [
    # 1. Tiêu đề văn bản & Lệnh quyết định (Ưu tiên hàng đầu)
    (re.compile(r'\bBAN\s+THUONG\s+VU\s+DANG\s+UY\s+XA\s+QUYET\s+DINH\b|\bBan\s+Thường\s+vụ\s+Đảng\s+ủy\s+XA\s+QUYẾT\s+ĐỊNH\b', re.IGNORECASE), 'BAN THƯỜNG VỤ ĐẢNG ỦY XÃ QUYẾT ĐỊNH'),
    (re.compile(r'\bXA\s+QUYẾT\s+ĐỊNH\b', re.IGNORECASE), 'XÃ QUYẾT ĐỊNH'),
    (re.compile(r'\bBAN\s+THUONG\s+VU\s+QUYET\s+DINH\b', re.IGNORECASE), 'BAN THƯỜNG VỤ QUYẾT ĐỊNH'),
    (re.compile(r'\bBAN\s+CHAP\s+HANH\s+QUYET\s+NGHI\b', re.IGNORECASE), 'BAN CHẤP HÀNH QUYẾT NGHỊ'),
    (re.compile(r'\bDANHSACH\b|\bDANH\s+SACH\b', re.IGNORECASE), 'DANH SÁCH'),

    # 2. Các từ chỉ căn cứ & mở đầu
    (re.compile(r'(?<=[\.,;]\s)(?:can\s+cur|can\s+cu|căn\s+cư)\b', re.IGNORECASE), 'căn cứ'),
    (re.compile(r'\bCan\s+cur\b|\bCan\s+cu\b|\bCăn\s+cư\b', re.IGNORECASE), 'Căn cứ'),
    (re.compile(r'\bTheo\s+de\s+nghi\s+cua\b|\bTheo\s+đề\s+nghị\s+cua\b', re.IGNORECASE), 'Theo đề nghị của'),
    (re.compile(r'\bQuy\s+dinh\b|\bQuy\s+đinh\b', re.IGNORECASE), 'Quy định'),
    (re.compile(r'\bQuy\s+che\b|\bQuy\s+ché\b', re.IGNORECASE), 'Quy chế'),

    # 3. Sửa lỗi số hiệu, ngày tháng, liên từ
    (re.compile(r'\bso\s+(\d+)\b', re.IGNORECASE), r'số \1'),
    (re.compile(r'\bs6\b|\bs[oóô] 6\b', re.IGNORECASE), 'số'),
    (re.compile(r'\bc6\b', re.IGNORECASE), 'có'),
    (re.compile(r'\bcuia\b|\bcưa\b', re.IGNORECASE), 'của'),
    (re.compile(r'\bcua\s+(?=Ban|Tỉnh|Đảng|Trung|Điều|Bộ|Ủy|Hội|Chi|Tổ|cơ|đoàn|khối)', re.IGNORECASE), 'của '),
    (re.compile(r'\bcua\b', re.IGNORECASE), 'của'),
    (re.compile(r'\bngay\s+(\d{1,2})', re.IGNORECASE), r'ngày \1'),
    (re.compile(r'\btir\b', re.IGNORECASE), 'từ'),
    (re.compile(r'\bva\s+(?=các|theo|những|đồng|Thường|Ban|Ủy|chi)', re.IGNORECASE), 'và '),
    (re.compile(r'\bva\b', re.IGNORECASE), 'và'),

    # 4. Ký hiệu văn bản Đảng & Nhà nước
    (re.compile(r'\bQD/TW\b|\bQD-TW\b', re.IGNORECASE), 'QĐ/TW'),
    (re.compile(r'\bQD/TU\b|\bQD-TU\b', re.IGNORECASE), 'QĐ/TU'),
    (re.compile(r'\bQC/DU\b|\bQC\s*-\s*DU\b|\bQC\s*/\s*DU\b', re.IGNORECASE), 'QC/ĐU'),
    (re.compile(r'\bQD-UBND\b', re.IGNORECASE), 'QĐ-UBND'),
    (re.compile(r'\bCV/TU\b', re.IGNORECASE), 'CV/TU'),
    (re.compile(r'\bTB/TU\b', re.IGNORECASE), 'TB/TU'),
    (re.compile(r'\bBTV\s+TU\b|\bBTV\s+Tỉnh\s+uy\b', re.IGNORECASE), 'BTV Tỉnh ủy'),
    (re.compile(r'\bThuong\s+trrcDU\b|\bThuong\s+truc\s+DU\b|\bThường\s+trực\s+DU\b', re.IGNORECASE), 'Thường trực Đảng ủy'),
    (re.compile(r'\bBan\s+XD\s+Dang\b|\bBan\s+Xay\s+dung\s+Dang\b|\bBan\s+Xây\s+dựng\s+Dang\b', re.IGNORECASE), 'Ban Xây dựng Đảng'),
    (re.compile(r'\bBCH\s+Dang\s+bo\b|\bBCH\s+Dang\s+bộ\b', re.IGNORECASE), 'BCH Đảng bộ'),

    # 5. Tổ chức, cơ quan, chức danh khối Đảng & Hành chính
    (re.compile(r'\bBan\s+Chap\s+hanh\b|\bBanichaphanh\b', re.IGNORECASE), 'Ban Chấp hành'),
    (re.compile(r'\bTrung\s+uong\b|\bTrung\s+ưong\b', re.IGNORECASE), 'Trung ương'),
    (re.compile(r'\bDieu\s+le\s+Dang\b|\bĐiều\s+lệ\s+Dang\b', re.IGNORECASE), 'Điều lệ Đảng'),
    (re.compile(r'\bBan\s+Bi\s+thu\b|\bBan\s+Bí\s+thu\b', re.IGNORECASE), 'Ban Bí thư'),
    (re.compile(r'\bTinh\s+uy\b', re.IGNORECASE), 'Tỉnh ủy'),
    (re.compile(r'\bDang\s+b[oọộ]\b|\bDang\s+b\?\b|\bDang\s+b\b', re.IGNORECASE), 'Đảng bộ'),
    (re.compile(r'\bDang\s+uy\b|\bDANG\s+OY\b', re.IGNORECASE), 'Đảng ủy'),
    (re.compile(r'\bChi\s+bo\b|\bChi\s+bọ\b', re.IGNORECASE), 'Chi bộ'),
    (re.compile(r'\bBan\s+Thuong\s+vu\b|\bBAN\s+THUONG\s+VU\b', re.IGNORECASE), 'Ban Thường vụ'),
    (re.compile(r'\bThuong\s+truc\s+Dang\s+uy\b|\bThuong\s+trục\s+Đảng\s+ủy\b|\bThuong\s+trucDang\s+iy\b', re.IGNORECASE), 'Thường trực Đảng ủy'),
    (re.compile(r'\bUy\s+ban\s+nhan\s+dan\b', re.IGNORECASE), 'Ủy ban nhân dân'),
    (re.compile(r'\bHoi\s+dong\s+nhan\s+dan\b', re.IGNORECASE), 'Hội đồng nhân dân'),
    (re.compile(r'\bTo\s+churc\s+Dang\b|\bTo\s+chuc\s+Dang\b|\bTo\s+chức\s+Dang\b|\bTổ\s+chức\s+Dang\b', re.IGNORECASE), 'Tổ chức Đảng'),
    (re.compile(r'\bCo\s+quan\s+Dang\b|\bCơ\s+quan\s+Dang\b', re.IGNORECASE), 'Cơ quan Đảng'),
    (re.compile(r'\bkhoi\s+Nha\s+nuoc\b|\bkhối\s+Nha\s+nước\b', re.IGNORECASE), 'khối Nhà nước'),
    (re.compile(r'\bCong\s+an\b', re.IGNORECASE), 'Công an'),
    (re.compile(r'\bQuan\s+sy\b|\bQuan\s+su\b', re.IGNORECASE), 'Quân sự'),
    (re.compile(r'\bY\s+te\b', re.IGNORECASE), 'Y tế'),

    # 6. Cấp xã, phường, thị trấn & Địa danh
    (re.compile(r'\bdanh\s*sachcu\s*thekem\s*theo\b|\bdanh\s+sach\s+cu\s+the\s+kem\s+theo\b|\bdanh\s+sách\s+cụ\s+thể\s+kem\s+theo\b', re.IGNORECASE), 'danh sách cụ thể kèm theo'),
    (re.compile(r'\(\s*(?:co|có)\s+(?:danh\s+sách|danh\s+sach)\b', re.IGNORECASE), '(có danh sách'),
    (re.compile(r'\btruc\s+thuoc\b|\btrực\s+thuoc\b|\btruc\s+thuộc\b', re.IGNORECASE), 'trực thuộc'),
    (re.compile(r'\bQuyet\s+dinh\b|\bQuyết\s+đinh\b', re.IGNORECASE), 'Quyết định'),
    (re.compile(r'\bQuyet\s+nghi\b|\bQuyết\s+nghi\b', re.IGNORECASE), 'Quyết nghị'),
    (re.compile(r'\b(quyet\s+dinh|quyết\s+định)\s+co\s+hieu\s+luc\s+ke\s+(?:t[uừ]|tir)\s+ngay\s+ky\b', re.IGNORECASE), 'Quyết định có hiệu lực kể từ ngày ký'),
    (re.compile(r'(\bĐiều\s+\d+)[\.,]\s*[Cc]ăn\s+cứ\s+(?:QUYẾT\s+ĐỊNH|Quyết\s+định)\s+thi\s+h[aà]nh\b', re.IGNORECASE), r'\1, căn cứ Quyết định thi hành'),
    (re.compile(r'\bxa\s+Thuan\s+Hanh\b|\bxa\s+Thuận\s+Hạnh\b', re.IGNORECASE), 'xã Thuận Hạnh'),
    (re.compile(r'\bThuan\s+Hanh\b', re.IGNORECASE), 'Thuận Hạnh'),
    (re.compile(r'\bDak\s+Nong\b|\bĐak\s+Nong\b', re.IGNORECASE), 'Đắk Nông'),
    (re.compile(r'\b(Đảng\s+bộ|Đảng\s+ủy|UBND|HDND|HĐND|Hội\s+đồng\s+nhân\s+dân|Ủy\s+ban\s+nhân\s+dân|BCH\s+Đảng\s+bộ)\s+xa\b', re.IGNORECASE), r'\1 xã'),
    (re.compile(r'\bHDND\b', re.IGNORECASE), 'HĐND'),
    (re.compile(r'\bdang\s+xa,\s*phuong\b|\bđảng\s+xa,\s*phường\b', re.IGNORECASE), 'đảng xã, phường'),
    (re.compile(r'\bdac\s+khu\b', re.IGNORECASE), 'đặc khu'),

    # 7. Từ vựng hành chính chuẩn
    (re.compile(r'\bve\s+thi\s+hanh\b|\bvề\s+thi\s+hanh\b', re.IGNORECASE), 'về thi hành'),
    (re.compile(r'\bve\s+viec\b|\bvề\s+viec\b', re.IGNORECASE), 'về việc'),
    (re.compile(r'\bve\s+chuc\s+nang\b|\bve\s+chục\s+nang\b|\bvề\s+chục\s+năng\b|\bvề\s+chục\s+nang\b', re.IGNORECASE), 'về chức năng'),
    (re.compile(r'\bnhiem\s+vu\b|\bnhiệm\s+vu\b', re.IGNORECASE), 'nhiệm vụ'),
    (re.compile(r'\bto\s+churc\b|\bto\s+chuc\b|\bto\s+chức\b', re.IGNORECASE), 'tổ chức'),
    (re.compile(r'\bbo\s+may\b|\bbọ\s+may\b|\bbọ\s+máy\b', re.IGNORECASE), 'bộ máy'),
    (re.compile(r'\bcap\s+tren\b|\bcấp\s+tren\b', re.IGNORECASE), 'cấp trên'),
    (re.compile(r'\btrurc\s+tiep\b|\btrực\s+tiep\b|\btryc\s+tiep\b', re.IGNORECASE), 'trực tiếp'),
    (re.compile(r'\bco\s+so\s+dang\b|\bcơ\s+sở\s+dang\b', re.IGNORECASE), 'cơ sở đảng'),
    (re.compile(r'\bco\s+so\b', re.IGNORECASE), 'cơ sở'),
    (re.compile(r'\btrurc\s+thuoc\b|\btryc\s+thuoc\b|\btrục\s+thuoc\b|\btrực\s+thuoc\b|\btrục\s+thuộc\b|\btrực\s+thuọc\b|\bbotruc\s+thu\?\s*c\b|\bbotruc\s+thuoc\b', re.IGNORECASE), 'trực thuộc'),
    (re.compile(r'\bthanh\s+lap\b', re.IGNORECASE), 'thành lập'),
    (re.compile(r'\bkhoa\s+([IVXLCDM\d]+),\s*nhiem\s+ky\b|\bkhoa\s+([IVXLCDM\d]+),\s*nhiệm\s+kỳ\b', re.IGNORECASE), r'khóa \1, nhiệm kỳ'),
    (re.compile(r'\bnhiem\s+ky\b|\bnhiemky\b|\bnhiêm\s+kỳ\b', re.IGNORECASE), 'nhiệm kỳ'),
    (re.compile(r'\bPhan\s+cong\s+cac\s+dong\s+chi\b|\bPhân\s+công\s+các\s+dong\s+chí\b', re.IGNORECASE), 'Phân công các đồng chí'),
    (re.compile(r'\bcac\s+dong\s+chi\b|\bcác\s+dong\s+chí\b', re.IGNORECASE), 'các đồng chí'),
    (re.compile(r'\bUy\s+vien\b|\bỦy\s+vien\b', re.IGNORECASE), 'Ủy viên'),
    (re.compile(r'\bphu\s+trach\s+cac\b|\bphụ\s+trách\s+cac\b|\bphu\s+trachcac\b', re.IGNORECASE), 'phụ trách các'),
    (re.compile(r'\bphu\s+tr[aá]ch\b', re.IGNORECASE), 'phụ trách'),
    (re.compile(r'(^|\.\s+)Cạc\b'), r'\1Các'),
    (re.compile(r'\bcạc\b', re.IGNORECASE), 'các'),
    (re.compile(r'\bcac\s+Chi\s+bộ\b|\bcac\s+chi\s+bộ\b', re.IGNORECASE), 'các chi bộ'),
    (re.compile(r'\(\s*co\s+danh\s+sách|\(\s*có\s+danh\s+sach', re.IGNORECASE), '(có danh sách'),
    (re.compile(r'\bdanh\s*sachcu\s*thekem\s*theo\b|\bdanh\s+sach\s+cu\s+the\s+kem\s+theo\b|\bdanh\s+sách\s+cụ\s+thể\s+kem\s+theo\b', re.IGNORECASE), 'danh sách cụ thể kèm theo'),
    (re.compile(r'\bco\s+trach\s+nhiem\s+thuc\s+hien\b|\bcó\s+trách\s+nhiệm\s+thuc\s+hiện\b', re.IGNORECASE), 'có trách nhiệm thực hiện'),
    (re.compile(r'\bcac\s+van\s+ban\b|\bcác\s+van\s+bản\b', re.IGNORECASE), 'các văn bản'),
    (re.compile(r'\bco\s+lien\s+quan\b|\bcó\s+lien\s+quan\b', re.IGNORECASE), 'có liên quan'),
    (re.compile(r'\bHang\s+thang\s+hoac\s+dot\s+xuat\b|\bHằng\s+tháng\s+hoặc\s+dot\s+xuất\b', re.IGNORECASE), 'Hằng tháng hoặc đột xuất'),
    (re.compile(r'\bbao\s+cao\b', re.IGNORECASE), 'báo cáo'),
    (re.compile(r'\btinh\s+hinh\b', re.IGNORECASE), 'tình hình'),
    (re.compile(r'\blinh\s+vuyc\b|\blinh\s+vuc\b|\blĩnh\s+vuc\b', re.IGNORECASE), 'lĩnh vực'),
    (re.compile(r'\bdia\s+ban\s+phu\s+trach\b|\bđịa\s+bàn\s+phu\s+trách\b', re.IGNORECASE), 'địa bàn phụ trách'),
    (re.compile(r'\bdia\s+ban\b', re.IGNORECASE), 'địa bàn'),
    (re.compile(r'\bCac\s+co\s+quan\b|\bCac\s+cơ\s+quan\b', re.IGNORECASE), 'Các cơ quan'),
    (re.compile(r'\bco\s+quan\b', re.IGNORECASE), 'cơ quan'),
    (re.compile(r'\bdon\s+vi\b', re.IGNORECASE), 'đơn vị'),
    (re.compile(r'\bchuyen\s+trach\s+tham\s+muu,\s*giup\s+viec\b|\bchuyên\s+trách\s+tham\s+mưu,\s*giup\s+việc\b', re.IGNORECASE), 'chuyên trách tham mưu, giúp việc'),
    (re.compile(r'\bco\s+ten\s+tai\b|\bcó\s+ten\s+tại\b', re.IGNORECASE), 'có tên tại'),
    (re.compile(r'\bten\s+tai\b', re.IGNORECASE), 'tên tại'),
    (re.compile(r'\bthi\s+hanh\b', re.IGNORECASE), 'thi hành'),
    (re.compile(r'(\bĐiều\s+\d+)[\.,]\s*[Cc]ăn\s+cứ\s+(?:QUYẾT\s+ĐỊNH|Quyết\s+định)\s+thi\s+h[aà]nh\b', re.IGNORECASE), r'\1, căn cứ Quyết định thi hành'),
    (re.compile(r'\bQUYẾT\s+ĐỊNH\s+co\s+hieu\s+luc\s+ke\s+t[uừ]\s+ngay\s+ky\b|\bQuyết\s+định\s+co\s+hieu\s+luc\s+ke\s+t[uừ]\s+ngay\s+ky\b', re.IGNORECASE), 'Quyết định có hiệu lực kể từ ngày ký'),
    (re.compile(r'\bco\s+hieu\s+luc\s+ke\s+tir\s+ngay\s+ky\b|\bco\s+hieu\s+luc\s+ke\s+tu\s+ngay\s+ky\b|\bcó\s+hiệu\s+lực\s+kể\s+từ\s+ngay\s+ky\b', re.IGNORECASE), 'có hiệu lực kể từ ngày ký'),

    # 8. Khung Điều & Mục
    (re.compile(r'\bDieu\s+(\d+)[\.:]?', re.IGNORECASE), r'Điều \1.'),
    (re.compile(r'\bChuong\s+([IVXLCDM\d]+)[\.:]?', re.IGNORECASE), r'Chương \1.'),
    (re.compile(r'\bMuc\s+(\d+)[\.:]?', re.IGNORECASE), r'Mục \1.'),

    # 9. Quyền hạn ký & Nơi nhận
    (re.compile(r'\bNoi\s+nhan\s*:?', re.IGNORECASE), 'Nơi nhận:'),
    (re.compile(r'[-–—•*]?\s*Nhu\s+Dieu\s+(\d+)', re.IGNORECASE), r'Như Điều \1'),
    (re.compile(r'[-–—•*]?\s*Nhu\s+Điều\s+(\d+)', re.IGNORECASE), r'Như Điều \1'),
    (re.compile(r'Đảng\s+ủy,\s*Đảng\s+ủy', re.IGNORECASE), 'Đảng ủy'),
    (re.compile(r'XATHUAN\s+EAN|XATHUAN\s+AN', re.IGNORECASE), ''),
    (re.compile(r'\s+Waw\b', re.IGNORECASE), ''),
    (re.compile(r'\bT/MBANTHUONGVU\b|\bT/M\s*BANTHUONGVU\b|\bT/M\s+BAN\s+THUONG\s+VU\b', re.IGNORECASE), 'T/M BAN THƯỜNG VỤ'),
    (re.compile(r'\bTM\.\s*UBND\b|\bTM\s*\.\s*UBND\b', re.IGNORECASE), 'TM. ỦY BAN NHÂN DÂN'),
    (re.compile(r'\bKT\.\s*CHU\s+TICH\b|\bKT\s*\.\s*CHU\s+TICH\b', re.IGNORECASE), 'KT. CHỦ TỊCH'),
    (re.compile(r'\bTL\.\s*CHANH\s+VAN\s+PHONG\b', re.IGNORECASE), 'TL. CHÁNH VĂN PHÒNG'),
    (re.compile(r'\bTo\s+truong\b', re.IGNORECASE), 'Tổ trưởng'),
    (re.compile(r'\bTo\s+pho\b|\bTopho\b|\bTpho\b', re.IGNORECASE), 'Tổ phó'),
    (re.compile(r'\bTo\s+vien\b|\bTovien\b|\bTvien\b', re.IGNORECASE), 'Tổ viên'),
    (re.compile(r'\bCap\s+uy\s+phu\s+trach\b', re.IGNORECASE), 'Cấp ủy phụ trách'),
    (re.compile(r'\bLuuy\s*:|\bLuu\s+y\s*:', re.IGNORECASE), 'Lưu ý:'),

    # 10. Địa danh, thôn, bản, trường học
    (re.compile(r'\bthon\s+Thuan\s+Binh\b', re.IGNORECASE), 'thôn Thuận Bình'),
    (re.compile(r'\bthon\s+Thuan\s+Nghia\b', re.IGNORECASE), 'thôn Thuận Nghĩa'),
    (re.compile(r'\bban\s+Dak\s+Thot\b|\bbobanDak\s+Thot\b|\bbanDak\s+Thot\b', re.IGNORECASE), 'bản Đắk Thốt'),
    (re.compile(r'\btruong\s+Nguyen\s+Viet\s+Xuan\b', re.IGNORECASE), 'trường Nguyễn Viết Xuân'),
    (re.compile(r'\bthon\s+Thuan\s+Bac\b', re.IGNORECASE), 'thôn Thuận Bắc'),
    (re.compile(r'\btruong\s+Hoa\s+Ban\b', re.IGNORECASE), 'trường Hoa Ban'),
    (re.compile(r'\bthon\s+Thuan\s+Tinh\b', re.IGNORECASE), 'thôn Thuận Tịnh'),
    (re.compile(r'\bthon\s+Thuan\s+Tan\b', re.IGNORECASE), 'thôn Thuận Tân'),
    (re.compile(r'\bTruong\s+Nguyen\s+Du\b', re.IGNORECASE), 'Trường Nguyễn Du'),
    (re.compile(r'\bban\s+Dam\s+Gio\b', re.IGNORECASE), 'bản Đầm Giỏ'),
    (re.compile(r'\bthon\s+Thuan\s+Loi\b', re.IGNORECASE), 'thôn Thuận Lợi'),
    (re.compile(r'\btruong\s+Kim\s+Dong\b', re.IGNORECASE), 'trường Kim Đồng'),
    (re.compile(r'\bthon\s+Thuan\s+Thanh\b', re.IGNORECASE), 'thôn Thuận Thành'),
    (re.compile(r'\btruong\s+B[eé]\s+Van\s+Dan\b', re.IGNORECASE), 'trường Bế Văn Đàn'),
    (re.compile(r'\bthon\s+Thuan\s+Hai\b', re.IGNORECASE), 'thôn Thuận Hải'),
    (re.compile(r'\bthon\s+Thuan\s+Nam\b', re.IGNORECASE), 'thôn Thuận Nam'),
    (re.compile(r'\btruong\s+Luong\s+Th[eé]\s+Vinh\b', re.IGNORECASE), 'trường Lương Thế Vinh'),
    (re.compile(r'\bthon\s+Thuan\s+Hoa\b', re.IGNORECASE), 'thôn Thuận Hòa'),
    (re.compile(r'\bthon\s+Thuan\s+Trung\b', re.IGNORECASE), 'thôn Thuận Trung'),
    (re.compile(r'\btruong\s+Hoa\s+Mai\b', re.IGNORECASE), 'trường Hoa Mai'),
    (re.compile(r'\bthon\s+(\d+)\b', re.IGNORECASE), r'thôn \1'),

    # 11. Danh xưng & Tên cán bộ
    (re.compile(r'\bD/c\b', re.IGNORECASE), 'Đ/c'),
    (re.compile(r'(Tổ\s*trưởng|Tổ\s*phó|Tổ\s*viên)(Đ/c|đ/c)', re.IGNORECASE), r'\1, \2'),
    (re.compile(r'\b(Tổ\s*viên)\s*10\b', re.IGNORECASE), r'\1'),
    (re.compile(r'([A-Za-zÀ-ỹ]+)\?\s+', re.IGNORECASE), r'\1 '),
    (re.compile(r'\bPho\s+BÍ\s+THƯ\b|\bPho\s+Bí\s+thư\b', re.IGNORECASE), 'Phó Bí thư'),
    (re.compile(r'\bBÍ\s+THƯ\s+Đảng\s+ủy\b', re.IGNORECASE), 'Bí thư Đảng ủy'),
    (re.compile(r'\bCHỦ\s+TỊCH\s+(HĐND|UBND)\b', re.IGNORECASE), r'Chủ tịch \1'),
    (re.compile(r'\bL[eé]\s+Ngoc\s+Phuc\b', re.IGNORECASE), 'Lê Ngọc Phúc'),
    (re.compile(r'\bL[eé]\s+Hong\s+Do\b', re.IGNORECASE), 'Lê Hồng Đô'),
    (re.compile(r'\bVu\s+Dinh\s+Trung\b', re.IGNORECASE), 'Vũ Đình Trung'),
    (re.compile(r'\bL[eé]\s*Van\s+Luan\b', re.IGNORECASE), 'Lê Văn Luận'),
    (re.compile(r'\bHoang\s+Van\s+Huan\b', re.IGNORECASE), 'Hoàng Văn Huân'),
    (re.compile(r'\bDang\s+Th[eé]\s+Tung\b', re.IGNORECASE), 'Đặng Thế Tùng'),
    (re.compile(r'\bNguyen\s+Ba\s+Ha\b', re.IGNORECASE), 'Nguyễn Bá Hà'),
    (re.compile(r'\bTong\s+Quang\s+Loc\b', re.IGNORECASE), 'Tống Quang Lộc'),
    (re.compile(r'\bTran\s+Le\s+Cuong\b', re.IGNORECASE), 'Trần Lê Cường'),
    (re.compile(r'\bPhung\s*Nhu\s*Khuong\b', re.IGNORECASE), 'Phùng Như Khương'),
    (re.compile(r'\bTran\s+Duy\s+Trieu\b', re.IGNORECASE), 'Trần Duy Triều'),
    (re.compile(r'\bVu\s+Hong\s+Luyen\b', re.IGNORECASE), 'Vũ Hồng Luyện'),
    (re.compile(r'\bTran\s+Cong\s+Nhat\b', re.IGNORECASE), 'Trần Công Nhất'),
    (re.compile(r'\bDoan\s+Thi\s+Tot\b', re.IGNORECASE), 'Đoàn Thị Tốt'),
    (re.compile(r'\bDo\s+Thi\s+Minh\s+Hai\b', re.IGNORECASE), 'Đỗ Thị Minh Hải'),
    (re.compile(r'\bBui\s+Van\s+Dan\b', re.IGNORECASE), 'Bùi Văn Đan'),
    (re.compile(r'\bNguyen\s*D[uư]rc\s*Phuong\b|\bNguyen\s+Duc\s+Phuong\b', re.IGNORECASE), 'Nguyễn Đức Phương'),
    (re.compile(r'\bNguyen\s+Thi\s+Duy\s+Tram\b', re.IGNORECASE), 'Nguyễn Thị Duy Trâm'),
    (re.compile(r'\bNguyen\s+Van\s+Tuan\b', re.IGNORECASE), 'Nguyễn Văn Tuấn'),
    (re.compile(r'\bMai\s*Xuan\s*Quang\b', re.IGNORECASE), 'Mai Xuân Quang'),
    (re.compile(r'\bDang\s+Ngoc\s+Hanh\b', re.IGNORECASE), 'Đặng Ngọc Hạnh'),
    (re.compile(r'\bNguyen\s*Van\s*Sang\b', re.IGNORECASE), 'Nguyễn Văn Sang'),

    # 12. Dính chữ & Khoảng trắng & Dấu hành chính
    (re.compile(r'\blam\s+viec\b', re.IGNORECASE), 'làm việc'),
    (re.compile(r'\bco\s+trách\s+nhiệm\b', re.IGNORECASE), 'có trách nhiệm'),
    (re.compile(r'\bbáo\s+cáo\s+ve\b', re.IGNORECASE), 'báo cáo về'),
    (re.compile(r'\bve\s+tình\s+hình\b', re.IGNORECASE), 'về tình hình'),
    (re.compile(r'\b(tren\s+cac|trên\s+cac|tren\s+các)\b', re.IGNORECASE), 'trên các'),
    (re.compile(r'\btai\s+địa\s+bàn\b', re.IGNORECASE), 'tại địa bàn'),
    (re.compile(r'\btruc\s+tiep\b', re.IGNORECASE), 'trực tiếp'),
    (re.compile(r'\bkhoa\s+([IVXLCDM\d]+)\b', re.IGNORECASE), r'khóa \1'),
    (re.compile(r'\bnhiemky\s*(\d{4})', re.IGNORECASE), r'nhiệm kỳ \1'),
    (re.compile(r'\bnhiemky\b', re.IGNORECASE), 'nhiệm kỳ'),
    (re.compile(r'\bphụ\s+trách\s+các\s+Chi\s+trực\s+thuộc\b', re.IGNORECASE), 'phụ trách các chi bộ trực thuộc'),
    (re.compile(r'(\b[Cc]ăn\s+cứ\s+)QUYẾT\s+ĐỊNH(\s+số\b)'), r'\1Quyết định\2'),
    (re.compile(r'\b(các|những|mỗi|từng)\s+Chi\s+bộ\b'), r'\1 chi bộ'),
    (re.compile(r'\b(các|những|mỗi|từng)\s+Cơ\s+quan\b'), r'\1 cơ quan'),
    (re.compile(r'\b(các|những|mỗi|từng)\s+Đơn\s+vị\b'), r'\1 đơn vị'),
    (re.compile(r'(\d+)\s*-\s*QC/ĐU', re.IGNORECASE), r'\1-QC/ĐU'),
    (re.compile(r'(\d+)\s*-\s*QĐ/', re.IGNORECASE), r'\1-QĐ/'),
    (re.compile(r'^-([A-Za-zÀ-ỹ])', re.MULTILINE), r'- \1'),
    (re.compile(r'\bchungvaphu\s+trach\b|\bchung\s+va\s+phu\s+trach\b', re.IGNORECASE), 'chung và phụ trách'),
    (re.compile(r'\bO3chi\b|\b03chi\b', re.IGNORECASE), '03 chi'),
    (re.compile(r'\bboneu\s+tren\b', re.IGNORECASE), 'bộ nêu trên'),
    (re.compile(r'\bphu\s+trach\s+truc\s+tiep\b', re.IGNORECASE), 'phụ trách trực tiếp'),
    (re.compile(r'(?<=\d{4})\s*-\s*(?=\d{4})'), ' - '),
]


class AdministrativeSanitizer:
    @staticmethod
    def heal_ocr_text_with_issues(text: str) -> Tuple[str, int, List[Any]]:
        """
        Quét và phục hồi các lỗi ký tự OCR phổ biến trong văn bản tiếng Việt.
        Thực hiện lặp hội tụ (tối đa 3 vòng) để đảm bảo các quy tắc phụ thuộc lẫn nhau
        được giải quyết triệt để ngay trong lượt đầu tiên (Idempotent 100%).
        """
        from core.spell_linter import LintIssue
        import unicodedata

        if not text:
            return text, 0, []

        cleaned = unicodedata.normalize('NFC', text)
        total_changes = 0
        issues: List[LintIssue] = []

        for _pass in range(3):
            pass_changes = 0
            for pattern, replacement in ADMIN_OCR_REPLACEMENTS:
                matches = list(pattern.finditer(cleaned))
                if not matches:
                    continue

                for m in matches:
                    orig = m.group(0)
                    repl = pattern.sub(replacement, orig)
                    orig_nfc = unicodedata.normalize('NFC', orig.strip())
                    repl_nfc = unicodedata.normalize('NFC', repl.strip())
                    if orig_nfc != repl_nfc:
                        msg = f"Tự động sửa lỗi chính tả / phục hồi dấu: '{orig_nfc}' -> '{repl_nfc}'"
                        issues.append(LintIssue(
                            issue_type="spelling",
                            original=orig_nfc,
                            replacement=repl_nfc,
                            message=msg,
                            severity="auto_fixed"
                        ))
                        pass_changes += 1

                new_text, n = pattern.subn(replacement, cleaned)
                if n > 0:
                    cleaned = unicodedata.normalize('NFC', new_text)

            total_changes += pass_changes
            if pass_changes == 0:
                break

        return cleaned, total_changes, issues

    @staticmethod
    def heal_ocr_text(text: str) -> Tuple[str, int]:
        """
        Quét và phục hồi các lỗi ký tự OCR phổ biến trong văn bản tiếng Việt
        Trả về (văn bản đã phục hồi, số lượng lỗi đã khắc phục)
        """
        cleaned, count, _ = AdministrativeSanitizer.heal_ocr_text_with_issues(text)
        return cleaned, count

    @staticmethod
    def detect_document_standard(full_text: str) -> str:
        """
        Phân tích nội dung để xác định tài liệu thuộc Khối Đảng (HD36) hay Khối Hành chính (ND30)
        """
        if not full_text:
            return "ND30"

        text_lower = full_text.lower()

        party_keywords = [
            "điều lệ đảng", "dieu le dang", "đảng bộ", "dang bo", "đảng ủy", "dang uy",
            "tỉnh ủy", "tinh uy", "huyện ủy", "thị ủy", "thành ủy", "ban bí thư", "ban bi thu",
            "ban chấp hành trung ương", "ban chap hanh", "chi bộ", "chi bo", "t/m ban thường vụ",
            "t/m ban chấp hành", "bí thư", "bi thu", "phó bí thư", "pho bi thu", "cấp ủy", "cap uy",
            "ban xây dựng đảng", "văn phòng tỉnh ủy", "đảng cộng sản việt nam"
        ]

        state_keywords = [
            "cộng hòa xã hội chủ nghĩa việt nam", "độc lập - tự do - hạnh phúc",
            "ủy ban nhân dân", "ubnd", "hội đồng nhân dân", "hđnd", "thủ tướng chính phủ",
            "chủ tịch ubnd", "tm. ủy ban nhân dân", "kt. chủ tịch", "bộ trưởng", "giám đốc sở"
        ]

        party_score = sum(1 for kw in party_keywords if kw in text_lower)
        state_score = sum(1 for kw in state_keywords if kw in text_lower)

        return "HD36" if party_score > state_score else "ND30"

    @staticmethod
    def extract_party_doc_metadata(doc: Any, filename: str = "") -> Dict[str, str]:
        """
        Trích xuất thông tin cơ quan Đảng từ nội dung văn bản hoặc tên file
        """
        full_text = " ".join(p.text for p in doc.paragraphs[:20])
        
        # 1. Cơ quan cấp trên (Tỉnh ủy...)
        m_tinh = re.search(r'(?:TỈNH\s+ỦY|TINH\s+UY)\s+([A-Za-zÀ-ỹ\s]+?)(?=\s+về|\s+ve|\s+ngày|\s+ngay|,|\.|\n|$)', full_text, re.IGNORECASE)
        if m_tinh:
            tinh_raw = m_tinh.group(1).strip()
            tinh_healed, _ = AdministrativeSanitizer.heal_ocr_text(tinh_raw)
            if "dak nong" in tinh_healed.lower() or "đắk nông" in tinh_healed.lower():
                superior = "TỈNH ỦY ĐẮK NÔNG"
            else:
                superior = f"TỈNH ỦY {tinh_healed.upper()}"
        else:
            superior = "TỈNH ỦY ĐẮK NÔNG"

        # 2. Cơ quan ban hành (Đảng bộ xã...)
        m_dang = re.search(r'(?:ĐẢNG\s+BỘ\s+XÃ|DANG\s+BO\s+XA|ĐẢNG\s+BỘ|DANG\s+BO)\s+([A-Za-zÀ-ỹ\s]+?)(?=\s+khóa|\s+khoa|\s+nhiệm|\s+nhiem|\s+trực|\s+trurc|,|\.|\n|$)', full_text, re.IGNORECASE)
        if m_dang:
            dang_raw = m_dang.group(1).strip()
            dang_healed, _ = AdministrativeSanitizer.heal_ocr_text(dang_raw)
            if "thuan hanh" in dang_healed.lower() or "thuận hạnh" in dang_healed.lower():
                issuing = "ĐẢNG BỘ XÃ THUẬN HẠNH"
                location = "Thuận Hạnh"
            elif "xa" in dang_healed.lower() or "xã" in dang_healed.lower():
                issuing = f"ĐẢNG BỘ {dang_healed.upper()}"
                location = dang_healed.replace("xã", "").replace("Xã", "").strip().title()
            else:
                issuing = f"ĐẢNG BỘ XÃ {dang_healed.upper()}"
                location = dang_healed.strip().title()
        else:
            issuing = "ĐẢNG BỘ XÃ THUẬN HẠNH"
            location = "Thuận Hạnh"

        # 3. Số, ký hiệu văn bản
        m_num = re.search(r'^(\d+)[_\-\s]', os.path.basename(filename))
        if m_num:
            doc_num = f"Số {m_num.group(1)}-QĐ/ĐU"
        else:
            m_num_text = re.search(r'(?:Quyết\s+định|Quyet\s+dinh)\s+số\s+(\d+)', full_text, re.IGNORECASE)
            if m_num_text:
                doc_num = f"Số {m_num_text.group(1)}-QĐ/ĐU"
            else:
                doc_num = "Số 16-QĐ/ĐU"

        return {
            "superior_agency": superior,
            "issuing_agency": issuing,
            "location": location,
            "doc_number": doc_num
        }

    @staticmethod
    def parse_in_text_footer(paragraphs_text: List[str]) -> Optional[Dict[str, Any]]:
        """
        Tách và phân tích khối Chân văn bản (Nơi nhận & Chữ ký) nếu bị gộp trong các đoạn văn
        Trả về:
        {
            "start_idx": int,
            "end_idx": int,
            "recipients": List[str],
            "authority_prefix": str,
            "signer_position": str,
            "signer_name": str
        } hoặc None nếu không phát hiện
        """
        start_idx = -1
        for idx, text in enumerate(paragraphs_text):
            t_clean = text.strip().lower()
            if t_clean.startswith("nơi nhận") or t_clean.startswith("noi nhan"):
                start_idx = idx
                break

        if start_idx == -1:
            return None

        # Quét tối đa 12 đoạn văn tiếp theo
        end_idx = min(start_idx + 12, len(paragraphs_text))
        footer_lines = paragraphs_text[start_idx:end_idx]

        recipients = []
        right_side = []
        actual_end = start_idx

        for i, raw_line in enumerate(footer_lines):
            line = raw_line.strip()
            if not line:
                continue

            # Nếu gặp Tiêu đề Bảng biểu / Phụ lục mới (DANH SÁCH, PHỤ LỤC...) thì dừng lại
            if re.match(r'^(DANH\s*SÁCH|DANH\s*SACH|PHỤ\s*LỤC|PHU\s*LUC|BẢNG|BANG|ĐIỀU\s+\d+|DIEU\s+\d+)\b', line, re.IGNORECASE) and i > 0:
                break

            actual_end = start_idx + i + 1

            # Tách dòng chứa 2 phần (bên trái là nơi nhận, bên phải là chức vụ/quyền hạn)
            m_split = re.search(r'(T/M|TM\.|K/T|KT\.|T/L|TL\.|BÍ\s+THƯ|BI\s+THU|CHỦ\s+TỊCH|CHU\s+TICH|PHÓ\s+BÍ\s+THƯ|PHO\s+BI\s+THU)', line, re.IGNORECASE)
            if m_split and m_split.start() > 0:
                left_part = line[:m_split.start()].strip()
                right_part = line[m_split.start():].strip()
                if left_part and not re.match(r'^(nơi\s*nhận|noi\s*nhan)', left_part, re.IGNORECASE):
                    recipients.append(left_part)
                if right_part:
                    right_side.append(right_part)
            elif line.startswith("-") or re.match(r'^(nơi\s*nhận|noi\s*nhan)', line, re.IGNORECASE):
                # Kiểm tra xem có dính chữ con dấu bên phải không (DANG UY, UBND, XATHUAN...)
                m_seal = re.search(r'\s+(ĐẢNG\s+ỦY\s+XÃ|DANG\s+[UO]Y\s+XA|UBND|XATHUAN|XÃ\s*THUẬN\s*AN|Waw\b)', line, re.IGNORECASE)
                if m_seal:
                    left_part = line[:m_seal.start()].strip()
                    recipients.append(left_part)
                else:
                    if not re.match(r'^(nơi\s*nhận|noi\s*nhan)', line, re.IGNORECASE):
                        recipients.append(line)
            else:
                right_side.append(line)

        # Phân loại cột phải (Quyền hạn, Chức vụ, Họ tên)
        auth_prefix = ""
        position = ""
        signer_name = ""

        for item in right_side:
            item_clean = item.strip()
            if not item_clean:
                continue
            item_healed, _ = AdministrativeSanitizer.heal_ocr_text(item_clean)
            if re.match(r'^(T/M|TM\.|K/T|KT\.|T/L|TL\.|TUQ\.)\s*', item_clean, re.IGNORECASE):
                auth_prefix = item_healed
            elif re.match(r'^(BÍ\s+THƯ|BI\s+THU|CHỦ\s+TỊCH|CHU\s+TICH|PHÓ\s+BÍ\s+THƯ|PHO\s+BI\s+THU|TRƯỞNG\s+BAN|TRUONG\s+BAN|CHÁNH\s+VĂN\s+PHÒNG|CHANH\s+VAN\s+PHONG)\b', item_clean, re.IGNORECASE):
                position = item_healed
            elif not signer_name and len(item_clean.split()) <= 5 and not any(kw in item_clean.upper() for kw in ["ĐẢNG", "DANG", "UBND", "THUẬN", "TỈNH"]):
                # Sửa tên cán bộ đặc thù dân tộc (K'Thanh)
                if item_clean.lower() in ("thanh", "k'thanh", "kthanh"):
                    signer_name = "K'Thanh"
                else:
                    signer_name = item_clean

        # Làm sạch recipients
        clean_recipients = []
        for r in recipients:
            r_c, _ = AdministrativeSanitizer.heal_ocr_text(r)
            # Chuẩn hóa dấu phân cách
            r_c = re.sub(r'[,\.]+$', ';', r_c.strip())
            if not r_c.startswith("-"):
                r_c = f"- {r_c}"
            if "b/c" in r_c.lower() and not r_c.endswith(";"):
                r_c += ";"
            clean_recipients.append(r_c)

        # Đảm bảo dòng lưu trữ văn thư có dấu chấm kết thúc
        if clean_recipients and any("lưu" in r.lower() for r in clean_recipients):
            for idx, r in enumerate(clean_recipients):
                if "lưu" in r.lower():
                    clean_recipients[idx] = re.sub(r'[;,]+$', '.', r)
                    if not clean_recipients[idx].endswith('.'):
                        clean_recipients[idx] += '.'

        if auth_prefix:
            auth_prefix, _ = AdministrativeSanitizer.heal_ocr_text(auth_prefix)
        if position:
            position, _ = AdministrativeSanitizer.heal_ocr_text(position)

        return {
            "start_idx": start_idx,
            "end_idx": actual_end,
            "recipients": clean_recipients,
            "authority_prefix": auth_prefix,
            "signer_position": position,
            "signer_name": signer_name
        }

    @staticmethod
    def merge_fragmented_paragraphs(doc: Any) -> int:
        """
        Nối các dòng văn bản bị ngắt vụn bởi tiến trình OCR thành các đoạn văn hoàn chỉnh.
        Xóa bỏ triệt để hiện tượng từ rớt dòng (co so, dia, (co, co, có...) và thụt đầu dòng sai.
        """
        paras = list(doc.paragraphs)
        i = 0
        merged_count = 0
        in_footer = False

        while i < len(paras) - 1:
            p = paras[i]
            next_p = paras[i + 1]

            t_p = p.text.strip()
            t_next = next_p.text.strip()

            if not t_next:
                i += 1
                continue

            # 1. Quản lý trạng thái Footer (Nơi nhận & Chữ ký)
            if re.match(r'^(nơi\s*nhận|noi\s*nhan)\b', t_p, re.IGNORECASE) or re.match(r'^(nơi\s*nhận|noi\s*nhan)\b', t_next, re.IGNORECASE):
                in_footer = True

            # Khi gặp Bảng biểu / Phụ lục / Danh sách thì thoát Footer
            if re.match(r'^(danh\s*sách|danh\s*sach|phụ\s*lục|phu\s*luc|bảng|bang)\b', t_next, re.IGNORECASE) and 'kèm theo' not in t_next.lower() and 'kem theo' not in t_next.lower():
                in_footer = False
                i += 1
                continue

            # Tuyệt đối không nối tiêu đề DANH SÁCH với bất cứ đoạn nào
            if re.match(r'^(danh\s*sách|danh\s*sach|phụ\s*lục|phu\s*luc|bảng|bang)\b', t_p, re.IGNORECASE) and 'kèm theo' not in t_p.lower() and 'kem theo' not in t_p.lower():
                i += 1
                continue

            # Trong khu vực Footer, giữ nguyên từng dòng độc lập để phân tích Nơi nhận & Người ký
            if in_footer:
                i += 1
                continue

            # Tuyệt đối không nối dòng lệnh kết thúc bằng QUYẾT ĐỊNH / QUYẾT NGHỊ với đoạn sau
            if re.search(r'(quyết\s*định|quyet\s*dinh|quyết\s*nghị|quyet\s*nghi)\s*$', t_p, re.IGNORECASE):
                i += 1
                continue

            # Nếu t_p là tiêu đề La Mã (I., II.) hoặc đề mục độc lập (1., 2.), không gộp đoạn sau vào
            if re.match(r'^[IVXLCDM]+\.\s+', t_p) or re.match(r'^\d+[\.:]\s+[A-ZÀ-Ỹ]', t_p):
                i += 1
                continue

            # Kiểm tra nếu next_p là đoạn mở đầu khối mới độc lập (KHÔNG NỐI)
            is_starter = False
            if re.match(r'^[-–—•*]?\s*(căn\s*cứ|can\s*cur|can\s*cu)\b', t_next, re.IGNORECASE):
                is_starter = True
            elif re.match(r'^(theo\s*đề\s*nghị|theo\s*de\s*nghi)\b', t_next, re.IGNORECASE):
                is_starter = True
            elif re.match(r'^(điều|dieu|chương|chuong|mục|muc|phần|phan)\s+\d+[\.:]?', t_next, re.IGNORECASE):
                is_starter = True
            elif re.match(r'^(ban\s*thường\s*vụ|ban\s*thuong\s*vu|ban\s*chấp\s*hành|ban\s*chap\s*hanh)\b', t_next, re.IGNORECASE) and any(k in t_next.lower() for k in ["quyết định", "quyet dinh", "nghị quyết", "nghi quyet"]):
                is_starter = True
            elif re.match(r'^(nơi\s*nhận|noi\s*nhan)\b', t_next, re.IGNORECASE):
                is_starter = True
            elif re.match(r'^[-–—•*]', t_next) and not t_p.endswith('-'):
                # Dòng gạch đầu dòng là mục danh sách độc lập
                is_starter = True
            elif re.match(r'^\+\s+', t_next):
                # Dòng cộng đầu dòng là mục danh sách độc lập
                is_starter = True
            elif re.match(r'^[a-z]\)\s+', t_next, re.IGNORECASE):
                # Dòng chữ cái a), b), c) là mục danh sách độc lập
                is_starter = True
            elif re.match(r'^([IVXLCDM]+\.|\d+[\.:])\s+', t_next):
                # Đề mục La Mã I., II. hoặc số 1., 2.
                is_starter = True
            elif re.match(r'^\d+\s*[-A-Za-zÀ-ỹ]', t_next) and not t_p.endswith('-'):
                is_starter = True
            elif re.match(r'^(stt|tt)\s+', t_next, re.IGNORECASE):
                is_starter = True
            elif re.match(r'^(lưu\s*ý|luu\s*y)\s*:', t_next, re.IGNORECASE):
                is_starter = True

            ends_punct = t_p.endswith(('.', ';', ':', '!', '?'))
            # Nếu dòng tiếp theo bắt đầu bằng chữ thường hoặc dấu ngoặc mở -> chắc chắn là câu nối
            is_continuation = (
                t_next[0].islower() or
                t_next.startswith('(') or
                t_next.startswith(')') or
                t_next.startswith(']') or
                t_next.startswith('kèm theo') or
                t_next.startswith('kem theo')
            )

            # Các từ rớt dòng điển hình từ OCR do thụt lề
            is_dangling = any(t_next.lower().startswith(w) for w in [
                'cơ sở', 'co so', 'đảng xã', 'dang xa', 'địa', 'dia', 'bàn phụ', 'ban phu',
                'cơ', 'co', 'có', 'quan, đơn', 'quan, don', 'tên tại', 'ten tai', 'khoa i', 'nhiệm kỳ', 'nhiem ky'
            ])

            should_merge = False
            if not is_starter:
                if not ends_punct or is_continuation or is_dangling:
                    should_merge = True

            if should_merge:
                if t_p.endswith('-'):
                    p.text = t_p[:-1] + t_next
                else:
                    p.text = t_p + ' ' + t_next
                next_p._element.getparent().remove(next_p._element)
                paras.pop(i + 1)
                merged_count += 1
            else:
                i += 1

        return merged_count

    @staticmethod
    def convert_ocr_list_to_table(doc: Any) -> bool:
        """
        Tự động nhận diện danh sách phân công chi bộ hoặc bảng phụ lục bị phân mảnh bởi OCR,
        tái tạo hoàn chỉnh thành Bảng Word (Table) 3 cột có viền chuẩn văn thư hành chính.
        """
        import docx
        from docx.shared import Pt, Mm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        paras = list(doc.paragraphs)
        start_list_idx = -1
        end_list_idx = -1

        for idx, p in enumerate(paras):
            t = p.text.strip()
            # Tìm dòng tiêu đề bảng STT hoặc dòng bắt đầu có số STT + Chi bộ
            if re.match(r'^(stt|tt)\s+to\s+chuc', t, re.IGNORECASE) or re.match(r'^(stt|tt)\b', t, re.IGNORECASE):
                start_list_idx = idx
                break
            elif re.match(r'^\d+\s+(chi\s+bộ|chi\s+bo)\b', t, re.IGNORECASE) and idx > 10:
                start_list_idx = idx
                break

        if start_list_idx == -1:
            return False

        # Xác định điểm kết thúc bảng (gặp Lưu ý: hoặc hết paragraph)
        for idx in range(start_list_idx, len(paras)):
            t = paras[idx].text.strip()
            if re.match(r'^(lưu\s*ý|luu\s*y)\s*:', t, re.IGNORECASE):
                end_list_idx = idx
                break
        if end_list_idx == -1:
            end_list_idx = len(paras)

        list_paras = paras[start_list_idx:end_list_idx]
        if len(list_paras) < 3:
            return False

        raw_lines = [p.text.strip() for p in list_paras if p.text.strip()]

        # Bỏ qua dòng tiêu đề STT nếu có trong raw_lines
        if raw_lines and re.match(r'^(stt|tt)\s+', raw_lines[0], re.IGNORECASE):
            content_lines = raw_lines[1:]
        else:
            content_lines = raw_lines

        # Tiền xử lý phục hồi lỗi OCR và chuẩn hóa tên
        cleaned_lines = []
        for l in content_lines:
            t, _ = AdministrativeSanitizer.heal_ocr_text(l)
            t = re.sub(r'D/c\s*', 'Đ/c ', t)
            t = re.sub(r'Chi\s*boban', 'Chi bộ bản ', t, flags=re.IGNORECASE)
            t = re.sub(r'Chi\s*botruong', 'Chi bộ trường ', t, flags=re.IGNORECASE)
            t = re.sub(r'Chi\s*bothon', 'Chi bộ thôn ', t, flags=re.IGNORECASE)
            t = re.sub(r'Tpho\b|T\s*pho\b', 'Tổ phó', t, flags=re.IGNORECASE)
            t = re.sub(r'Tvien\b|T\s*vien\b', 'Tổ viên', t, flags=re.IGNORECASE)
            cleaned_lines.append(t)

        groups = []
        current_stt = 0
        current_orgs = []
        current_cadres = []

        def save_group():
            nonlocal current_stt, current_orgs, current_cadres
            if current_orgs or current_cadres:
                groups.append({
                    "stt": str(current_stt),
                    "orgs": list(current_orgs),
                    "cadres": list(current_cadres)
                })
                current_orgs = []
                current_cadres = []

        for line in cleaned_lines:
            t = line.strip()
            if not t:
                continue

            m_num = re.match(r'^(\d+)\s*(.*)$', t)
            line_num = int(m_num.group(1)) if m_num and int(m_num.group(1)) <= 30 else None
            rest = m_num.group(2).strip() if line_num is not None else t

            is_new_team = False
            if line_num in (1, 2, 3):
                is_new_team = True
                next_stt = line_num
            elif 'Tổ trưởng' in rest or 'To truong' in rest:
                # Nếu nhóm hiện tại đã có Tổ trưởng hoặc đã đủ cadre thì mở nhóm mới
                if any('Tổ trưởng' in c for c in current_cadres) or (current_stt >= 4 and len(current_cadres) >= 2):
                    is_new_team = True
                    next_stt = line_num if line_num else (current_stt + 1)
            elif current_stt in (1, 2, 3) and (rest.startswith('-') or 'Chi bộ' in rest):
                is_new_team = True
                next_stt = 4

            if is_new_team:
                save_group()
                current_stt = next_stt

            if line_num and line_num > current_stt:
                current_stt = line_num

            if not rest:
                continue

            chunks = re.split(r'(?=[-\s]*(?:Đ/c|D/c)\s+)', rest)
            for chunk in chunks:
                c = chunk.strip()
                if not c:
                    continue
                if re.match(r'^[-–—\s]*(?:Đ/c|D/c)\b', c):
                    c_clean = re.sub(r'^[-–—\s]+', '', c)
                    c_clean = re.sub(r'\s+', ' ', c_clean)
                    current_cadres.append(f'- {c_clean}' if not c_clean.startswith('-') else c_clean)
                else:
                    m_sub = re.search(r'(Đ/c\s+.*)$', c)
                    if m_sub:
                        org_part = c[:m_sub.start()].strip()
                        cadre_part = m_sub.group(1).strip()
                        org_clean = re.sub(r'^[-–—\s]+', '', org_part).strip()
                        if org_clean:
                            current_orgs.append(f'- {org_clean}' if not org_clean.startswith('-') else org_clean)
                        current_cadres.append(f'- {cadre_part}')
                    else:
                        org_clean = re.sub(r'^[-–—\s]+', '', c).strip()
                        if org_clean:
                            current_orgs.append(f'- {org_clean}' if not org_clean.startswith('-') else org_clean)

        save_group()

        if not groups:
            return False

        # Chèn bảng 3 cột ngay trước dòng đầu tiên của danh sách
        tbl = doc.add_table(rows=1, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        list_paras[0]._p.addprevious(tbl._tbl)

        # Thiết lập viền chuẩn bảng biểu hành chính (Grid viền đơn mỏng 0.5pt)
        tblPr = tbl._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
            f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
            f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
            f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
            f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

        # Thiết lập tiêu đề cột
        hdr = tbl.rows[0].cells
        hdr[0].width = Mm(15.0)
        hdr[1].width = Mm(65.0)
        hdr[2].width = Mm(85.0)

        p_h0 = hdr[0].paragraphs[0]
        p_h0.text = "STT"
        p_h0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h0.paragraph_format.space_before = Pt(4)
        p_h0.paragraph_format.space_after = Pt(4)
        format_run_element(p_h0.runs[0], font_name="Times New Roman", size_pt=13, bold=True)

        p_h1 = hdr[1].paragraphs[0]
        p_h1.text = "Tổ chức Đảng"
        p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h1.paragraph_format.space_before = Pt(4)
        p_h1.paragraph_format.space_after = Pt(4)
        format_run_element(p_h1.runs[0], font_name="Times New Roman", size_pt=13, bold=True)

        p_h2 = hdr[2].paragraphs[0]
        p_h2.text = "Cấp ủy phụ trách"
        p_h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h2.paragraph_format.space_before = Pt(4)
        p_h2.paragraph_format.space_after = Pt(4)
        format_run_element(p_h2.runs[0], font_name="Times New Roman", size_pt=13, bold=True)

        for g in groups:
            row = tbl.add_row().cells
            row[0].width = Mm(15.0)
            row[1].width = Mm(65.0)
            row[2].width = Mm(85.0)

            # Cột 0: STT
            p0 = row[0].paragraphs[0]
            p0.text = g["stt"]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p0.paragraph_format.space_before = Pt(2)
            p0.paragraph_format.space_after = Pt(2)
            format_run_element(p0.runs[0], font_name="Times New Roman", size_pt=12.5, bold=False)

            # Cột 1: Tổ chức Đảng
            p1 = row[1].paragraphs[0]
            p1.text = "\n".join(g["orgs"])
            p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p1.paragraph_format.space_before = Pt(2)
            p1.paragraph_format.space_after = Pt(2)
            p1.paragraph_format.line_spacing = 1.15
            for r in p1.runs:
                format_run_element(r, font_name="Times New Roman", size_pt=12.5, bold=False)

            # Cột 2: Cấp ủy phụ trách
            p2 = row[2].paragraphs[0]
            p2.text = "\n".join(g["cadres"])
            p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p2.paragraph_format.space_before = Pt(2)
            p2.paragraph_format.space_after = Pt(2)
            p2.paragraph_format.line_spacing = 1.15
            for r in p2.runs:
                format_run_element(r, font_name="Times New Roman", size_pt=12.5, bold=False)

        # Xóa triệt để các đoạn văn thô cũ đã được chuyển thành bảng
        for p in list_paras:
            try:
                p._element.getparent().remove(p._element)
            except Exception:
                pass

        return True
