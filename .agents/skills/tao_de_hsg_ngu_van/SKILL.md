---
name: tao_de_hsg_ngu_van
description: >
  Tạo file Word (.doc/.docx) đề thi Học Sinh Giỏi Ngữ Văn THCS (lớp 6-9)
  bằng python-docx. Bao gồm phần Đề thi (1 trang) và Hướng dẫn chấm (có bảng
  điểm chi tiết), đúng format file mẫu Google Docs đề thi Việt Nam.
---

# Skill: Tạo Đề Thi HSG Ngữ Văn bằng python-docx

## Mô tả

Skill này dùng để tạo bộ đề thi Học Sinh Giỏi Ngữ Văn THCS (áp dụng lớp 7)
gồm 2 câu hỏi theo đúng cấu trúc chuẩn của Bộ GD-ĐT:

- **Câu 1 (4,0 điểm):** Đọc bài thơ → viết đoạn văn cảm xúc (~300 chữ)
- **Câu 2 (6,0 điểm):** Nghị luận xã hội trình bày ý kiến tán thành (~400-500 chữ)

Mỗi file Word gồm **2 phần trên 2 trang**:
1. Trang 1: Đề thi (thiết kế vừa khít 1 trang A4)
2. Trang 2+: Hướng dẫn chấm (bảng điểm 3 cột chi tiết)

---

## Yêu cầu môi trường

```bash
pip install python-docx
```

---

## Cấu trúc file Word chuẩn

### TRANG 1 – ĐỀ THI

```
[HEADER 2 CỘT - không viền]
Cột trái               | Cột phải
PHÒNG GD&ĐT ........   | KÌ THI CHỌN HỌC SINH GIỎI CẤP TRƯỜNG (in đậm)
TRƯỜNG THCS ....... (đậm) | NĂM HỌC 20XX – 20XX (đậm)
                       | MÔN: NGỮ VĂN – LỚP 7 (đậm)

            [In nghiêng, căn giữa]
    Thời gian làm bài: 90 phút, không kể thời gian giao đề
                (Đề thi có 01 trang)
─────────────────────────────────────────────────────── (kẻ ngang)

Câu 1. (4,0 điểm) Đọc bài thơ sau:
              [TÊN BÀI THƠ] (đậm, căn giữa)
     [Khổ 1-3 nghiêng]   |   [Khổ 4-5 nghiêng]   ← 2 CỘT không viền
                   (Nguồn trích dẫn - nghiêng, phải)
     Viết đoạn văn khoảng 300 chữ ghi lại cảm xúc...

Câu 2. (6,0 điểm)
     [Câu dẫn nhận định + trích dẫn in nghiêng]
     Hãy viết bài văn khoảng 400 đến 500 chữ...

                      --- Hết ---
```

### TRANG 2 – HƯỚNG DẪN CHẤM (format mới: 2 bảng riêng)

```
[HEADER 2 CỘT tương tự đề thi, cột phải = "HƯỚNG DẪN CHẤM..."]
─────────────────────────────

I. HƯỚNG DẪN CHUNG (3 điểm lưu ý đánh số 1. 2. 3.)

II. ĐÁP ÁN VÀ THANG ĐIỂM CHI TIẾT

Câu 1. (4,0 điểm)          ← in đậm, đứng ngoài bảng
┌─────────────────────┬──────────────────────────────┬────────┐
│ Phần                │ Nội dung cần đạt             │ Điểm   │
├─────────────────────┼──────────────────────────────┼────────┤
│ a. Hình thức (0,5)  │ Đúng hình thức đoạn văn...   │  0,5   │
│ b. Mở đoạn (0,5)   │ Giới thiệu tác giả...        │  0,5   │
│ c. Thân đoạn (2,5)  │ 1. Cảm xúc nội dung (1,5):   │  2,5   │
│                     │   - chi tiết...              │        │
│                     │ 2. Cảm xúc nghệ thuật (1,0): │        │
│                     │   - thể thơ, điệp ngữ...     │        │
│ d. Kết đoạn (0,25)  │ Khẳng định lại giá trị...    │  0,25  │
│ e. Sáng tạo (0,25)  │ Diễn đạt mới mẻ...           │  0,25  │
└─────────────────────┴──────────────────────────────┴────────┘
Lưu ý: ... (in nghiêng đậm, đứng ngoài bảng)

Câu 2. (6,0 điểm)          ← in đậm, đứng ngoài bảng
┌───────────────────────────┬──────────────────────────┬───────┐
│ Phần                      │ Nội dung cần đạt         │ Điểm  │
├───────────────────────────┼──────────────────────────┼───────┤
│ a. Bố cục (0,5)           │ Đủ 3 phần...             │  0,5  │
│ b. Xác định vấn đề (0,25) │ Đúng: [tên vấn đề]...   │  0,25 │
│ c. Mở bài (0,5)           │ Dẫn dắt, trích dẫn...   │  0,5  │
│ d. Thân bài (4,0)         │ 1. Giải thích (1,0):     │  4,0  │
│                           │   - bullet...            │       │
│                           │ 2. Bàn luận (2,0):       │       │
│                           │   - bullet...            │       │
│                           │ 3. Bài học (1,0):        │       │
│                           │   - bullet...            │       │
│ e. Kết bài (0,5)          │ Khẳng định lại...        │  0,5  │
│ f. Chính tả, ngữ pháp     │ Đảm bảo đúng...          │  0,25 │
│    (0,25)                 │                          │       │
│ g. Sáng tạo (0,5)         │ Lập luận độc đáo...      │  0,5  │
└───────────────────────────┴──────────────────────────┴───────┘
Lưu ý: ... (in nghiêng đậm, đứng ngoài bảng)

Ghi chú cách cho điểm chi tiết Câu 2: (đậm, gạch chân)
- 5,5 – 6,0:  Đầy đủ, lập luận sắc sảo, dẫn chứng sinh động, sáng tạo tốt.
- 4,5 – 5,25: Khá đầy đủ, lập luận rõ, dẫn chứng phù hợp, còn vài lỗi nhỏ.
- 3,5 – 4,25: Cơ bản, bố cục rõ nhưng nội dung chưa sâu, dẫn chứng chung.
- 2,5 – 3,25: Sơ sài, thiếu luận điểm, dẫn chứng nghèo nàn.
- Dưới 2,5:   Không đạt yêu cầu, ý rời rạc, nhiều lỗi.
```

---

## Key Technical Patterns (python-docx)

### 1. Setup cơ bản

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = 'Times New Roman'
SZ   = 11   # font size chuẩn cho đề thi VN

def setup(doc):
    s = doc.styles['Normal']
    s.font.name = FONT; s.font.size = Pt(SZ)
    s.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)  # QUAN TRỌNG cho tiếng Việt
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(3.0); sec.right_margin = Cm(2.0)
```

> **Lưu ý:** Luôn set `w:eastAsia` font khi dùng tiếng Việt, nếu không chữ sẽ bị sai font.

### 2. Thêm run có format

```python
def r(para, text, bold=False, italic=False, underline=False, sz=None):
    run = para.add_run(text)
    run.bold = bold; run.italic = italic; run.underline = underline
    run.font.name = FONT; run.font.size = Pt(sz or SZ)
    run._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), FONT)
    return run
```

### 3. Bảng KHÔNG viền (header 2 cột)

```python
def no_border(table):
    tbl  = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
    bd = OxmlElement('w:tblBorders')
    for e in ['top','left','bottom','right','insideH','insideV']:
        el = OxmlElement(f'w:{e}')
        el.set(qn('w:val'),'none'); el.set(qn('w:sz'),'0')
        el.set(qn('w:space'),'0'); el.set(qn('w:color'),'auto')
        bd.append(el)
    tblPr.append(bd)
```

### 4. Bảng CÓ viền đầy đủ (bảng điểm HDC)

```python
def full_border(table):
    tbl  = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
    bd = OxmlElement('w:tblBorders')
    for e in ['top','left','bottom','right','insideH','insideV']:
        el = OxmlElement(f'w:{e}')
        el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'4')
        el.set(qn('w:space'),'0'); el.set(qn('w:color'),'000000')
        bd.append(el)
    tblPr.append(bd)
```

### 5. Kẻ ngang (đường kẻ paragraph)

```python
def hrule(doc):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(2)
    bdr = OxmlElement('w:pBdr')
    btm = OxmlElement('w:bottom')
    btm.set(qn('w:val'),'single'); btm.set(qn('w:sz'),'6')
    btm.set(qn('w:space'),'1');   btm.set(qn('w:color'),'000000')
    bdr.append(btm)
    para._p.get_or_add_pPr().append(bdr)
```

### 6. Ngắt trang

```python
def pagebreak(doc):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    run = para.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)
```

### 7. Thơ 2 cột (tiết kiệm không gian)

```python
# Bài thơ chia: LEFT_STANZAS (khổ 1-3), RIGHT_STANZAS (khổ 4-5)
def fill_poem_col(cell, stanzas):
    # Xóa nội dung mặc định của cell
    for para in list(cell.paragraphs):
        para._element.getparent().remove(para._element)
    first = True
    for stanza in stanzas:
        if not first:
            blank = cell.add_paragraph()  # dòng trắng giữa khổ
            blank.alignment = WD_ALIGN_PARAGRAPH.CENTER
        first = False
        for line in stanza:
            para = cell.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(line)
            run.italic = True
            # set font...

# Tạo bảng 2 cột không viền
tp = doc.add_table(rows=1, cols=2)
no_border(tp)
tp.alignment = WD_TABLE_ALIGNMENT.CENTER
tp.cell(0,0).width = Cm(7.5)
tp.cell(0,1).width = Cm(7.5)
fill_poem_col(tp.cell(0,0), LEFT_STANZAS)
fill_poem_col(tp.cell(0,1), RIGHT_STANZAS)
```

### 8. Ô bảng nhiều đoạn văn (bảng HDC)

```python
# Xóa nội dung mặc định của cell trước khi điền
def clr(cell):
    for para in list(cell.paragraphs):
        para._element.getparent().remove(para._element)

# Thêm nhiều đoạn vào 1 ô - mỗi đoạn có format riêng
clr(cells[1])
for txt, bd, it in nd_items:  # nd_items = list[(text, bold, italic)]
    pr = cells[1].add_paragraph()
    pr.paragraph_format.space_before = Pt(0)
    pr.paragraph_format.space_after  = Pt(0)
    run = pr.add_run(txt)
    run.bold = bd; run.italic = it
    # set font...
```

> **QUAN TRỌNG:** Phải `clr(cell)` trước khi thêm nội dung vì python-docx
> tạo sẵn 1 paragraph rỗng trong mỗi cell khi tạo bảng mới.

### 9. Thụt đầu dòng (indent) cho yêu cầu đề

```python
para.paragraph_format.first_line_indent = Cm(1.0)
```

### 10. Tiết kiệm không gian để vừa 1 trang

```python
# Luôn đặt spacing tối thiểu
para.paragraph_format.space_before      = Pt(0)   # hoặc Pt(1), Pt(2) tùy chỗ
para.paragraph_format.space_after       = Pt(0)
para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
# Font size 11pt thay vì 12pt
# Lề dưới 1.5cm thay vì 2.0cm
```

---

## Workflow tạo 1 bộ đề đầy đủ

```
1. Đọc ngữ liệu từ file .txt
2. Chuẩn bị dữ liệu:
   - Bài thơ: chia LEFT_STANZAS (khổ 1-3) và RIGHT_STANZAS (khổ 4-5)
   - Câu 2: list (text, italic) cho nhận định + list (text, bold, italic) lệnh đề
   - HDC Câu 1: C1_ROWS cố định (chỉ thay tên bài thơ/tác giả)
   - HDC Câu 2: C2_ROWS_DEn riêng từng đề gồm 7 mục a–g
3. Gọi build_exam(doc, cau2_vd, cau2_yc) → trang 1
4. Gọi pagebreak(doc)
5. Gọi build_hdc(doc, c2_vande_label, c2_rows, c2_luu_y) → trang 2
6. doc.save(path)
```

### Helper: score_table() và luu_y()

```python
def score_table(doc, rows_data):
    """
    Tạo bảng thang điểm 3 cột: Phần | Nội dung cần đạt | Điểm
    rows_data: list of (phan_txt, nd_items, diem_txt)
      phan_txt : str              – cột Phần, ví dụ "a. Hình thức (0,5)"
      nd_items : list[(txt, bold, italic)] – nhiều đoạn trong 1 ô
      diem_txt : str              – cột Điểm
    Cột Phần in đậm. Mục số "1. Giải thích...", "2. Bàn luận..." in đậm.
    Bullets và nội dung thường để regular.
    """
    CW = [Cm(3.2), Cm(10.0), Cm(1.8)]   # chiều rộng 3 cột
    tbl = doc.add_table(rows=len(rows_data)+1, cols=3)
    full_border(tbl)
    # ... (xem tao_bo_de_final.py để biết chi tiết)

def luu_y(doc, lines):
    """Thêm đoạn Lưu ý sau bảng (in nghiêng đậm + nội dung nghiêng)."""
    para = doc.add_paragraph()
    run = para.add_run('Lưu ý: ')
    run.bold = True; run.italic = True
    run2 = para.add_run(lines[0])
    run2.italic = True
    for line in lines[1:]:
        p2 = doc.add_paragraph()
        p2.add_run(line).italic = True
```

---

## Thang điểm chuẩn (áp dụng cho mọi đề HSG Ngữ Văn 7)

### Câu 1 – Cảm xúc về bài thơ (4,0 điểm)

| Phần | Nội dung cần đạt | Điểm |
|------|-----------------|------|
| a. Hình thức (0,5) | Đủ mở-thân-kết, ~300 chữ | 0,5 |
| b. Mở đoạn (0,5) | Giới thiệu tác giả, tác phẩm; ấn tượng chung | 0,5 |
| c. Thân đoạn (2,5) | Nội dung (1,5đ): chi tiết hình ảnh bà, nỗi nhớ, tình cảm gia đình | 2,5 |
| | Nghệ thuật (1,0đ): thể thơ 5 chữ, điệp ngữ "ngỡ bà", hình ảnh gợi cảm | |
| d. Kết đoạn (0,25) | Khẳng định giá trị, liên hệ bản thân | 0,25 |
| e. Sáng tạo (0,25) | Diễn đạt mới mẻ, cảm xúc chân thành | 0,25 |

### Câu 2 – Nghị luận xã hội (6,0 điểm)

| Phần | Nội dung cần đạt | Điểm |
|------|-----------------|------|
| a. Bố cục (0,5) | Đủ mở-thân-kết bài | 0,5 |
| b. Xác định vấn đề (0,25) | Đúng trọng tâm nhận định | 0,25 |
| c. Mở bài (0,5) | Dẫn dắt, trích dẫn, nêu ý kiến | 0,5 |
| d. Thân bài (4,0) | Giải thích (1,0) + Bàn luận (2,0) + Bài học (1,0) | 4,0 |
| e. Kết bài (0,5) | Khẳng định, liên hệ bản thân | 0,5 |
| f. Chính tả, ngữ pháp (0,25) | Đảm bảo đúng chuẩn | 0,25 |
| g. Sáng tạo (0,5) | Lập luận độc đáo, dẫn chứng thuyết phục | 0,5 |

### Ghi chú điểm Câu 2 (5 mức)

| Mức điểm | Mô tả |
|----------|-------|
| 5,5 – 6,0 | Đầy đủ, lập luận sắc sảo, dẫn chứng sinh động, sáng tạo tốt |
| 4,5 – 5,25 | Khá đầy đủ, lập luận rõ, dẫn chứng phù hợp, còn vài lỗi nhỏ |
| 3,5 – 4,25 | Cơ bản, bố cục rõ nhưng nội dung chưa sâu, dẫn chứng còn chung |
| 2,5 – 3,25 | Sơ sài, thiếu luận điểm, dẫn chứng nghèo nàn |
| Dưới 2,5 | Không đạt yêu cầu, ý rời rạc, mắc nhiều lỗi |

---

## Script mẫu tham khảo

> File full: `e:\Thu\tao_bo_de_final.py`

Cấu trúc hàm chính:
```python
make(
    path             = r'e:\Thu\ĐỀ HSG NGỮ VĂN 7 2025-2026 (Đề N).doc',
    cau2_vd          = [(text, italic), ...],       # đoạn dẫn nhận định Câu 2
    cau2_yc          = [(text, bold, italic), ...], # lệnh đề Câu 2
    c2_vande_label   = 'tên vấn đề...',             # dùng ở mục b. HDC Câu 2
    c2_rows          = C2_ROWS_DEn,                 # list (phan, nd_items, diem)
    c2_luu_y         = ['Lưu ý dòng 1', ...],       # ghi chú sau bảng Câu 2
)
```

### Cấu trúc C2_ROWS (7 mục cố định)

```python
C2_ROWS_DEn = [
    ('a. Bố cục (0,5)',            [('Đủ 3 phần...', False, False)],   '0,5'),
    ('b. Xác định vấn đề (0,25)', [('Đúng: [tên vd]...', False, False)], '0,25'),
    ('c. Mở bài (0,5)',           [('Dẫn dắt...', False, False)],      '0,5'),
    ('d. Thân bài (4,0)',         [
        ('1. Giải thích (1,0):', True, False),
        ('- [bullet]', False, False),
        # ...
        ('2. Bàn luận (2,0):', True, False),
        ('- [bullet]', False, False),
        # ...
        ('3. Bài học (1,0):', True, False),
        ('- [bullet]', False, False),
    ], '4,0'),
    ('e. Kết bài (0,5)',           [('Khẳng định lại...', False, False)], '0,5'),
    ('f. Chính tả, ngữ pháp (0,25)', [('Đảm bảo...', False, False)],  '0,25'),
    ('g. Sáng tạo (0,5)',          [('Lập luận độc đáo...', False, False)], '0,5'),
]
```

---

## Lưu ý quan trọng

1. **Đặt tên file `.doc`** (không phải `.docx`) để tương thích khi gửi email/in
2. **Không dùng `tblPr = tbl.find(...) or OxmlElement(...)`** — lxml trả
   `None` nhưng phép `or` với Element sẽ gây `FutureWarning`/`True`. Dùng
   `if tblPr is None:` thay thế.
3. **Luôn `clr(cell)` trước khi thêm nội dung** vào cell bảng
4. **Thơ 5 chữ** thường có 5 khổ × 4 dòng = 20 dòng → chia 3+2 khổ cho 2 cột
5. **Font tiếng Việt**: Set cả `font.name` lẫn `w:eastAsia` attribute
6. Để in vừa 1 trang: dùng **font 11pt**, **lề dưới 1.5cm**, `line_spacing = SINGLE`, `space_before/after` tối thiểu
