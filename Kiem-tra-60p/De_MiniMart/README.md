# Bài kiểm tra thực hành — Lập trình phân tích dữ liệu (60 phút)

Bộ đề thực hành trên máy, 10 điểm, gồm 5 phần: Python thuần & regex · Pandas làm sạch/ghép bảng · phân tích tổng hợp · trực quan hoá · thống kê suy luận. Dữ liệu là dữ liệu mô phỏng về một chuỗi cửa hàng tiện lợi hư cấu (MiniMart), có **lỗi cài sẵn** để kiểm tra kỹ năng làm sạch.

## Cấu trúc thư mục

```
De_MiniMart/
├── de_thi/          ← PHÁT CHO SINH VIÊN
│   ├── De_kiem_tra_LTPTDL.pdf
│   └── De_kiem_tra_LTPTDL.docx      (bản Word để sửa tên trường, học kỳ, quy chế tài liệu)
├── du_lieu/         ← PHÁT CHO SINH VIÊN
│   ├── don_hang.csv                 2.436 dòng (đã gồm 36 dòng trùng)
│   ├── khach_hang.xlsx              sheet khach_hang (500) · sheet hang_thanh_vien (3)
│   └── nhat_ky_giao_hang.txt        nhật ký giao hàng tháng 6/2025
├── bai_lam/         ← PHÁT CHO SINH VIÊN
│   └── bai_lam_mau.ipynb            khung notebook, mỗi câu một ô
├── dap_an/          ← CHỈ GIẢNG VIÊN — xoá khỏi bản phát
│   ├── Dap_an_thang_diem.pdf        kết quả chuẩn, thang điểm chi tiết, lỗi thường gặp
│   ├── Dap_an_thang_diem.docx
│   ├── dap_an.ipynb                 lời giải đầy đủ, đã chạy sẵn kết quả
│   ├── dong_loi.txt                 file kết quả mẫu câu A1
│   └── bieu_do.png                  hình kết quả mẫu Phần D
└── cong_cu/
    └── tao_du_lieu.py               sinh lại dữ liệu (seed cố định)
```

## Phát đề

Nén `de_thi/`, `du_lieu/`, `bai_lam/` thành một file zip. **Không** gửi kèm `dap_an/` và `cong_cu/` (script sinh dữ liệu để lộ các lỗi cài sẵn).

Sinh viên mở `bai_lam/bai_lam_mau.ipynb` — đường dẫn `DATA_DIR = Path('../du_lieu')` đã trỏ đúng. Nếu dùng Google Colab, tải 3 tập tin dữ liệu lên và đổi thành `Path('.')`.

Môi trường cần: Python ≥ 3.10, `numpy`, `pandas`, `matplotlib`, `seaborn`, `scipy`, `openpyxl` (để đọc Excel). Google Colab đã có sẵn tất cả.

## Tạo đề khác (đề B, đề thi lại…)

1. Mở `cong_cu/tao_du_lieu.py`, đổi `SEED = 2025` sang số khác, chạy `python3 tao_du_lieu.py`.
2. Mở `dap_an/dap_an.ipynb`, **Restart & Run All** để lấy kết quả chuẩn mới.
3. Cập nhật các con số trong `Dap_an_thang_diem.docx` theo notebook — nội dung đề thi không cần sửa.

Cấu trúc lỗi và hiệu ứng thống kê được giữ nguyên khi đổi seed, nhưng nên kiểm tra lại hai câu có đáp án phụ thuộc ngẫu nhiên: danh mục dẫn đầu ở Đà Nẵng (C1) và tháng tăng trưởng mạnh nhất (C2).
