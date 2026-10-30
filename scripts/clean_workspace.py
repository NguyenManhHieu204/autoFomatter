"""
scripts/clean_workspace.py: Dọn dẹp sạch sẽ toàn bộ rác phát sinh trong dự án và thư mục làm việc:
1. Dọn thư mục build/ của PyInstaller (tiết kiệm ~165MB)
2. Xóa tệp AutoFormatter_Standalone.exe dư thừa (tiết kiệm ~144MB)
3. Xóa các tệp đệm __pycache__, *.pyc, *.bak, *.tmp
4. Dọn sạch các tệp lặp đuôi _ChuanHoa_ChuanHoa... trong Downloads và thư mục test
"""

import os
import sys
import shutil
import glob
import re

def clean_all():
    sys.stdout.reconfigure(encoding='utf-8')
    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    print(f"=== BẮT ĐẦU DỌN DẸP SẠCH SẼ WORKSPACE: {root_dir} ===")

    # 1. Dọn thư mục build/
    build_dir = os.path.join(root_dir, 'build')
    if os.path.exists(build_dir):
        try:
            shutil.rmtree(build_dir, ignore_errors=True)
            print("✔ Đã xóa thư mục rác build/ (~165MB)")
        except Exception as e:
            print(f"⚠ Không thể xóa build/: {e}")

    # 2. Xóa tệp dist/AutoFormatter_Standalone.exe dư thừa
    standalone_exe = os.path.join(root_dir, 'dist', 'AutoFormatter_Standalone.exe')
    if os.path.exists(standalone_exe):
        try:
            os.remove(standalone_exe)
            print("✔ Đã xóa tệp AutoFormatter_Standalone.exe dư thừa (~144MB)")
        except Exception as e:
            print(f"⚠ Không thể xóa AutoFormatter_Standalone.exe: {e}")

    # 3. Dọn __pycache__ và *.pyc
    pycache_count = 0
    for current_dir, dirs, files in os.walk(root_dir):
        if '__pycache__' in dirs:
            p_dir = os.path.join(current_dir, '__pycache__')
            shutil.rmtree(p_dir, ignore_errors=True)
            pycache_count += 1
        for f in files:
            if f.endswith(('.pyc', '.pyo', '.bak', '.tmp')) or f.startswith('~$'):
                try:
                    os.remove(os.path.join(current_dir, f))
                except Exception:
                    pass
    print(f"✔ Đã dọn {pycache_count} thư mục đệm __pycache__ và tệp tạm (.pyc, .bak, .tmp, ~$*)")

    # 4. Dọn tệp test kết quả cũ trong Tai_Lieu_Mau_Test
    test_dir = os.path.join(root_dir, 'Tai_Lieu_Mau_Test')
    if os.path.exists(test_dir):
        for f in os.listdir(test_dir):
            if f.endswith('_ChuanHoa.docx') or f.endswith('_ChuanHoa.pdf') or '_ChuanHoa_' in f:
                try:
                    os.remove(os.path.join(test_dir, f))
                    print(f"✔ Đã xóa tệp kết quả test cũ: {f}")
                except Exception:
                    pass

    # 5. Dọn dẹp các tệp chuỗi lặp _ChuanHoa_ChuanHoa... trong Downloads của người dùng
    downloads_dir = os.path.expanduser('~/Downloads')
    if os.path.exists(downloads_dir):
        cleaned_downloads = 0
        pattern = re.compile(r'_ChuanHoa(_ChuanHoa)+', re.IGNORECASE)
        for f in os.listdir(downloads_dir):
            if pattern.search(f) and f.endswith(('.docx', '.pdf')):
                fpath = os.path.join(downloads_dir, f)
                try:
                    os.remove(fpath)
                    print(f"✔ Đã xóa tệp rác chuỗi lặp trong Downloads: {f}")
                    cleaned_downloads += 1
                except Exception:
                    pass
        if cleaned_downloads > 0:
            print(f"✔ Đã dọn dẹp sạch {cleaned_downloads} tệp rác chuỗi lặp thừa thãi trong Downloads!")

    print("=== DỌN DẸP HOÀN TẤT: KHÔNG GIAN ĐÃ SẠCH SẼ 100% ===")

if __name__ == '__main__':
    clean_all()
