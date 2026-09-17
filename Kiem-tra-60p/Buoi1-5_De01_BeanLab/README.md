# Bài kiểm tra thực hành — Lập trình phân tích dữ liệu · Buổi 1–5 (60 phút)

Bộ đề thực hành trên máy, 10 điểm, 11 câu trong 5 phần, bao trọn nội dung Buổi 1–5 và có thêm trực quan hoá:

| Phần | Nội dung | Buổi | Điểm |
|---|---|---|---|
| A | Python thuần: tập tin văn bản, regex, hàm tự viết, cấu trúc điều khiển, `dict`/`set`/`Counter` | 1 · 2 · 3 | 2,0 |
| B | NumPy: ma trận, `axis`, broadcasting, `linalg.solve` · lập trình hướng đối tượng, `@property`, kế thừa | 2 · 3 | 2,0 |
| C | Pandas: đọc CSV, Excel, SQLite · làm sạch dữ liệu · tích hợp dữ liệu | 3 · 4 | 2,5 |
| D | Biến đổi dữ liệu (mã hoá, log, one-hot) · thu gọn dữ liệu (tổng hợp, rời rạc hoá, tương quan, PCA) | 5 | 1,5 |
| E | Trực quan hoá: một hình 2×2 gồm biểu đồ đường, histogram, cột chồng, heatmap | 4 · EDA | 2,0 |

Dữ liệu là dữ liệu mô phỏng về một chuỗi cà phê hư cấu (BeanLab) có 4 chi nhánh, có **lỗi cài sẵn** để kiểm tra kỹ năng làm sạch và tích hợp.

## Cấu trúc thư mục

```
Buoi1-5_De01_BeanLab/
├── de_thi/          ← PHÁT CHO SINH VIÊN
│   ├── De_kiem_tra_LTPTDL.pdf
│   └── De_kiem_tra_LTPTDL.docx      (bản Word để sửa tên trường, học kỳ, quy chế tài liệu)
├── du_lieu/         ← PHÁT CHO SINH VIÊN
│   ├── nhat_ky_ca_lam.txt           nhật ký ca làm việc tháng 6/2025 (766 dòng dữ liệu)
│   ├── giao_dich.csv                2.430 dòng (đã gồm 30 dòng trùng)
│   ├── khach_hang.xlsx              sheet khach_hang (406, có 6 dòng trùng) · sheet hang_thanh_vien (3)
│   └── san_pham.db                  SQLite · bảng san_pham (16) · bảng chi_nhanh (4)
├── bai_lam/         ← PHÁT CHO SINH VIÊN
│   └── bai_lam_mau.ipynb            khung notebook, mỗi câu một ô
├── dap_an/          ← CHỈ GIẢNG VIÊN — xoá khỏi bản phát
│   ├── Dap_an_thang_diem.pdf        kết quả chuẩn, thang điểm chi tiết, lỗi thường gặp
│   ├── Dap_an_thang_diem.docx
│   ├── dap_an.ipynb                 lời giải đầy đủ, đã chạy sẵn kết quả
│   ├── ca_loi.txt                   file kết quả mẫu câu A1
│   └── bieu_do.png                  hình kết quả mẫu Phần E
└── cong_cu/
    ├── tao_du_lieu.py               sinh lại dữ liệu (seed cố định)
    └── dap_an_nguon.py              mã nguồn notebook đáp án (chạy được như script)
```

## Phát đề

Nén `de_thi/`, `du_lieu/`, `bai_lam/` thành một file zip. **Không** gửi kèm `dap_an/` và `cong_cu/` (script sinh dữ liệu để lộ các lỗi cài sẵn).

Sinh viên mở `bai_lam/bai_lam_mau.ipynb` — đường dẫn `DATA_DIR = Path('../du_lieu')` đã trỏ đúng. Nếu dùng Google Colab, tải 4 tập tin dữ liệu lên và đổi thành `Path('.')`.

Môi trường cần: Python ≥ 3.10, `numpy`, `pandas`, `matplotlib`, `seaborn`, `scipy`, `scikit-learn`, `openpyxl` (để đọc Excel). `sqlite3` có sẵn trong Python. Google Colab đã có sẵn tất cả.

## Tạo đề khác (đề B, đề thi lại…)

1. Mở `cong_cu/tao_du_lieu.py`, đổi `SEED = 2026` sang số khác, chạy `python3 tao_du_lieu.py`.
2. Mở `dap_an/dap_an.ipynb`, **Restart & Run All** để lấy kết quả chuẩn mới.
3. Cập nhật các con số trong `Dap_an_thang_diem.docx` theo notebook — nội dung đề thi không cần sửa, riêng Phần B không phụ thuộc seed.

Cấu trúc lỗi và các bẫy được giữ nguyên khi đổi seed, nhưng nên kiểm tra lại ba câu có đáp án phụ thuộc ngẫu nhiên: nhân viên có tỉ lệ đi trễ cao nhất (A2), cặp đặc trưng tương quan mạnh và số thành phần PCA (D2).
