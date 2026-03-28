# -*- coding: utf-8 -*-
"""
parse_poem.py
Đọc file .txt chứa bài thơ → trả về metadata và danh sách khổ thơ.

Định dạng file .txt chuẩn:
  Dòng 1  : Tên bài thơ (ALL CAPS hoặc thường)
  Dòng 2+ : Các dòng thơ, khổ cách nhau bằng 1 dòng trắng
  Dòng cuối (tùy chọn): Nguồn trích dẫn trong dấu ngoặc (...)
"""

import re
from pathlib import Path


def parse_poem(filepath: str) -> dict:
    """
    Đọc file .txt và trả về dict:
        {
          "title"  : str,          # tên bài thơ
          "stanzas": list[list[str]], # danh sách khổ, mỗi khổ là list dòng
          "source" : str,          # nguồn trích dẫn (nếu có)
          "all_lines": list[str],  # tất cả dòng thơ (không có dòng trắng/nguồn)
        }
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f'Không tìm thấy file: {filepath}')

    raw = path.read_text(encoding='utf-8').replace('\r\n', '\n').strip()
    lines = raw.split('\n')

    # Dòng đầu = tiêu đề
    title = lines[0].strip()
    rest  = lines[1:]

    # Tách nguồn trích dẫn (dòng bắt đầu bằng '(')
    source = ''
    body_lines = []
    for line in rest:
        stripped = line.strip()
        if stripped.startswith('(') and stripped.endswith(')'):
            source = stripped
        else:
            body_lines.append(stripped)

    # Tách thành khổ thơ (ngăn cách bằng dòng trống)
    stanzas = []
    current_stanza = []
    for line in body_lines:
        if line == '':
            if current_stanza:
                stanzas.append(current_stanza)
                current_stanza = []
        else:
            current_stanza.append(line)
    if current_stanza:
        stanzas.append(current_stanza)

    all_lines = [line for stanza in stanzas for line in stanza]

    return {
        'title'    : title,
        'stanzas'  : stanzas,
        'source'   : source,
        'all_lines': all_lines,
    }


def split_two_columns(stanzas: list, left_count: int = None) -> tuple:
    """
    Chia stanzas thành 2 cột để in 2 cột trên đề thi.
    left_count : số khổ cột trái (mặc định: ceil(n/2) làm tròn lên)
    Trả về (left_stanzas, right_stanzas)
    """
    import math
    n = len(stanzas)
    if left_count is None:
        left_count = math.ceil(n / 2)
    return stanzas[:left_count], stanzas[left_count:]


def validate_poem(poem: dict) -> list[str]:
    """Kiểm tra tính hợp lệ của bài thơ. Trả về list cảnh báo."""
    warnings = []
    if not poem['title']:
        warnings.append('⚠ Thiếu tiêu đề bài thơ.')
    if not poem['stanzas']:
        warnings.append('⚠ Không tìm thấy khổ thơ nào.')
    for i, s in enumerate(poem['stanzas'], 1):
        if len(s) not in (4, 5, 6):
            warnings.append(f'⚠ Khổ {i} có {len(s)} dòng (thường là 4 hoặc 5).')
    if not poem['source']:
        warnings.append('⚠ Thiếu nguồn trích dẫn.')
    return warnings


# ─── Demo khi chạy trực tiếp ─────────────────────────────────────────────────
if __name__ == '__main__':
    import sys
    txt = sys.argv[1] if len(sys.argv) > 1 else r'e:\Thu\nho_ba.txt'
    poem = parse_poem(txt)
    print(f"Tiêu đề : {poem['title']}")
    print(f"Số khổ  : {len(poem['stanzas'])}")
    print(f"Nguồn   : {poem['source']}")
    left, right = split_two_columns(poem['stanzas'])
    print(f"\nCột trái ({len(left)} khổ):")
    for s in left: print(' | '.join(s))
    print(f"\nCột phải ({len(right)} khổ):")
    for s in right: print(' | '.join(s))
    warns = validate_poem(poem)
    if warns:
        print('\n' + '\n'.join(warns))
    else:
        print('\n✅ Bài thơ hợp lệ.')
