# 📚 Bộ Đề HSG Ngữ Văn 7 – 2025-2026

Dự án Python tự động tạo bộ đề thi **Học Sinh Giỏi Ngữ Văn lớp 7** (năm học 2025–2026) dạng file Word (`.doc`), bao gồm **Đề thi** và **Hướng dẫn chấm** chi tiết.

---

## 📁 Cấu trúc dự án

```
de_hsg_ngu_van/
├── tao_bo_de_final.py    # Script chính – tạo file Word
├── parse_poem.py         # Đọc & phân tích bài thơ từ file .txt
├── validate_scores.py    # Kiểm tra tổng điểm thang điểm
├── nho_ba.txt            # Ngữ liệu bài thơ (Nhớ Bà – Trương Anh Tú)
├── output/               # Thư mục chứa file .doc được tạo ra (tự động tạo)
└── README.md
```

---

## ⚙️ Yêu cầu

- Python **3.10+**
- Thư viện `python-docx`

```bash
pip install python-docx
```

---

## 🚀 Cách dùng

### Tạo cả 3 bộ đề
```bash
python tao_bo_de_final.py
```

### Chỉ tạo một đề cụ thể
```bash
python tao_bo_de_final.py --de 1        # Đề 1: Đọc sách
python tao_bo_de_final.py --de 2        # Đề 2: Tình bạn chân chính
python tao_bo_de_final.py --de 3        # Đề 3: Lòng biết ơn
python tao_bo_de_final.py --de 1 3      # Tạo Đề 1 và Đề 3
```

### Chỉ kiểm tra thang điểm (không tạo file)
```bash
python tao_bo_de_final.py --validate
```

### Phân tích bài thơ nguồn
```bash
python parse_poem.py nho_ba.txt
```

---

## 📄 Cấu trúc file Word được tạo

Mỗi file `.doc` gồm **2 phần**:

### Trang 1 – Đề thi
- Header 2 cột (Phòng GD&ĐT | Thông tin kỳ thi)
- **Câu 1 (4,0 điểm):** Bài thơ *Nhớ Bà* (Trương Anh Tú) – in 2 cột để tiết kiệm trang → Viết đoạn văn cảm xúc ~300 chữ
- **Câu 2 (6,0 điểm):** Nhận định + lệnh đề nghị luận xã hội ~400–500 chữ

### Trang 2+ – Hướng dẫn chấm
- Header 2 cột (Phòng GD&ĐT | HƯỚNG DẪN CHẤM)
- I. Hướng dẫn chung
- II. Đáp án và thang điểm chi tiết (2 bảng: Câu 1 + Câu 2)
  - Cột: **Phần | Nội dung cần đạt | Điểm**
- Lưu ý chấm điểm + Ghi chú 5 mức điểm

---

## 🗂️ Danh sách 3 bộ đề

| Đề | Chủ đề Câu 2 | Nhận định |
|----|-------------|-----------|
| **Đề 1** | Thói quen đọc sách | *"Đọc sách không chỉ là để biết mà còn là để sống tốt hơn."* |
| **Đề 2** | Tình bạn chân chính | *"Tình bạn chân chính là ngọn lửa sưởi ấm tâm hồn..."* |
| **Đề 3** | Lòng biết ơn | *"Lòng biết ơn biến những gì ta có thành đủ..."* (Melody Beattie) |

---

## 🔧 Thêm đề mới

Để thêm Đề 4, mở `tao_bo_de_final.py` và thêm vào danh sách `ALL_EXAMS`:

```python
dict(
    so_de=4,
    filename='ĐỀ HSG NGỮ VĂN 7 2025-2026 (Đề 4).doc',
    label='Đề 4 – [Chủ đề mới]',
    cau2_vd=[
        ('Câu dẫn: ', False),
        ('"Trích dẫn nhận định."', True),
    ],
    cau2_yc=[
        ('Hãy viết bài văn khoảng ', False, False),
        ('400 đến 500 chữ', True, False),
        (' trình bày ý kiến tán thành của em.', False, False),
    ],
    c2_vande_label='[tên vấn đề]',
    c2_rows=C2_ROWS_DE4,   # định nghĩa ở trên
    c2_luu_y=C2_LUU_Y_DE4,
),
```

---

## 📝 Kỹ thuật

- **python-docx** để tạo file Word với format chuẩn (font, lề, bảng viền, in đậm/nghiêng)
- **Bài thơ 2 cột** (`parse_poem.py + split_two_columns()`) để đề thi vừa 1 trang A4
- **Bảng thang điểm** 3 cột: Phần | Nội dung cần đạt | Điểm
- **Validate điểm tự động** trước khi in: tổng Câu 1 = 4,0đ; Câu 2 = 6,0đ

---

## 👩‍🏫 Tác giả

Dự án phục vụ giáo viên Ngữ văn THCS trong việc soạn đề thi HSG hiệu quả, chuẩn format.
