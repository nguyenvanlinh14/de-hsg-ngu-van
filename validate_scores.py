# -*- coding: utf-8 -*-
"""
validate_scores.py
Kiểm tra tổng điểm của bảng thang điểm HDC.
"""

from fractions import Fraction


def _parse_score(s: str) -> float:
    """Chuyển '0,25' hoặc '4,0' → float."""
    return float(s.replace(',', '.')) if s.strip() else 0.0


def validate_c1_rows(rows: list, expected: float = 4.0) -> list[str]:
    """Kiểm tra tổng điểm C1_ROWS == expected."""
    errors = []
    total = sum(_parse_score(r[2]) for r in rows)
    total = round(total, 2)
    if total != expected:
        errors.append(
            f'❌ Câu 1: tổng điểm = {total} (kỳ vọng {expected})'
        )
    else:
        print(f'✅ Câu 1: tổng điểm = {total} ✓')
    return errors


def validate_c2_rows(rows: list, expected: float = 6.5) -> list[str]:
    """Kiểm tra tổng điểm C2_ROWS == expected."""
    errors = []
    total = sum(_parse_score(r[2]) for r in rows)
    total = round(total, 2)
    if total != expected:
        errors.append(
            f'❌ Câu 2: tổng điểm = {total} (kỳ vọng {expected})'
        )
    else:
        print(f'✅ Câu 2: tổng điểm = {total} ✓')
    return errors


def validate_all(c1_rows, c2_rows_list: list[tuple]) -> bool:
    """
    Kiểm tra toàn bộ.
    c2_rows_list: list của (label, rows) ví dụ [('Đề 1', C2_ROWS_DE1), ...]
    Trả về True nếu tất cả hợp lệ.
    """
    errors = validate_c1_rows(c1_rows)
    for label, rows in c2_rows_list:
        errs = validate_c2_rows(rows)
        for e in errs:
            errors.append(f'  [{label}] {e}')
    if errors:
        print('\n'.join(errors))
        return False
    return True


if __name__ == '__main__':
    # Import và kiểm tra ngay
    import sys
    sys.path.insert(0, r'e:\Thu')
    from tao_bo_de_final import C1_ROWS, C2_ROWS_DE1, C2_ROWS_DE2, C2_ROWS_DE3
    ok = validate_all(C1_ROWS, [
        ('Đề 1 – Đọc sách',        C2_ROWS_DE1),
        ('Đề 2 – Tình bạn',        C2_ROWS_DE2),
        ('Đề 3 – Lòng biết ơn',    C2_ROWS_DE3),
    ])
    if ok:
        print('\n✅ Tất cả thang điểm hợp lệ!')
    else:
        print('\n⚠ Vui lòng sửa các lỗi trên.')
