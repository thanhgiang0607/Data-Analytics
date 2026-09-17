<<<<<<< HEAD
# Data-Analytics
=======
# Bài kiểm tra thực hành — Lập trình phân tích dữ liệu · Buổi 1–5 (60 phút) · Đề số 02

Bộ đề thực hành trên máy, 10 điểm, 11 câu trong 5 phần. Cùng khung và thang điểm với Đề số 01 (BeanLab), nhưng **bối cảnh, dữ liệu, yêu cầu và các bẫy đều khác**.

| Phần | Nội dung | Buổi | Điểm |
|---|---|---|---|
| A | Python thuần: nhật ký khoá xe, regex khớp toàn dòng, hàm tự viết, `Counter`/`set` | 1 · 2 · 3 | 2,0 |
| B | NumPy: ma trận trạm × khung giờ, chọn cột không liền nhau, so sánh theo cột, `linalg.solve` (bảng giá cước) · OOP: thuộc tính lớp, `@property`, kế thừa, đa hình | 2 · 3 | 2,0 |
| C | Pandas: đọc CSV, Excel, SQLite · tính thời lượng từ hai mốc thời gian, làm sạch · tích hợp với khoá khác tên, chuẩn hoá số điện thoại | 3 · 4 | 2,5 |
| D | Mã hoá thứ bậc, log, rời rạc hoá khung giờ bằng `pd.cut` + one-hot · tổng hợp theo khách, `qcut`, tương quan, PCA | 5 | 1,5 |
| E | Trực quan hoá: hình 2×2 gồm biểu đồ đường, phân tán, cột ghép, heatmap | 4 · EDA | 2,0 |

Bối cảnh: **EcoRide**, dịch vụ cho thuê xe đạp, xe đạp điện và xe máy điện (hư cấu) có 6 trạm ở Đà Nẵng và Hội An, dữ liệu tháng 3–8/2025. Dữ liệu mô phỏng có **lỗi cài sẵn** để kiểm tra kỹ năng làm sạch và tích hợp.

## Khác gì so với Đề số 01

| | Đề 01 — BeanLab | Đề 02 — EcoRide |
|---|---|---|
| A — tập tin văn bản | Nhật ký ca làm, tính tỉ lệ đi trễ | Nhật ký khoá xe, chỉ xét sự kiện TRA, tỉ lệ trả xe pin yếu; bẫy thừa trường cuối dòng (`re.match` lọt), bẫy `<=` làm đổi xe cao nhất |
| B1 | So sánh với trung bình từng **dòng** | Chọn 2 cột không liền nhau, lọc bằng mặt nạ, so sánh với trung bình từng **cột** |
| B3 | `SanPham`, property `gia_ban` | `ChuyenDi`, thuộc tính lớp `PHI_MO_KHOA`, lớp con ghi đè `cuoc_phi` có `max(0, …)` |
| C2 | Cột số lượng và thời gian chờ có sẵn | Tự tính `phut` từ `bat_dau`/`ket_thuc` (bẫy `.dt.seconds`, bẫy lọc làm rơi NaN); “xe đạp” là tiền tố của “xe đạp điện” |
| C3 | Đổi tên khoá, chuẩn hoá giới tính | Ghép bằng `left_on`/`right_on`, chuẩn hoá số điện thoại bằng regex; bản ghi khách trùng khác cách viết nên `drop_duplicates()` không đủ |
| D | Rời rạc hoá bằng `qcut` | Thêm `pd.cut` với `right=False` và one-hot khung giờ; PCA không chuẩn hoá chỉ còn 1 thành phần |
| E | Đường · histogram · cột chồng · heatmap | Đường · phân tán · cột ghép theo thứ tự category · heatmap |

## Cấu trúc thư mục

```
Kiem-tra-LTPTDL-60p-Buoi1-5-De02/
├── de_thi/          ← PHÁT CHO SINH VIÊN
│   ├── De_kiem_tra_LTPTDL.pdf
│   └── De_kiem_tra_LTPTDL.docx      (bản Word để sửa tên trường, học kỳ, quy chế tài liệu)
├── du_lieu/         ← PHÁT CHO SINH VIÊN
│   ├── nhat_ky_khoa_xe.txt          nhật ký khoá xe điện tháng 7/2025 (802 dòng dữ liệu)
│   ├── chuyen_di.csv                2.635 dòng (đã gồm 35 dòng trùng)
│   ├── khach_hang.xlsx              sheet khach_hang (508, có 8 mã trùng) · sheet goi_cuoc (3)
│   └── he_thong.db                  SQLite · bảng tram (6) · bảng loai_xe (3)
├── bai_lam/         ← PHÁT CHO SINH VIÊN
│   └── bai_lam_mau.ipynb            khung notebook, mỗi câu một ô
├── dap_an/          ← CHỈ GIẢNG VIÊN — xoá khỏi bản phát
│   ├── Dap_an_thang_diem.pdf        kết quả chuẩn, thang điểm chi tiết, lỗi thường gặp
│   ├── Dap_an_thang_diem.docx
│   ├── dap_an.ipynb                 lời giải đầy đủ, đã chạy sẵn kết quả
│   ├── dong_loi.txt                 file kết quả mẫu câu A1
│   └── bieu_do.png                  hình kết quả mẫu Phần E
└── cong_cu/
    ├── tao_du_lieu.py               sinh lại dữ liệu (SEED = 2502)
    ├── dap_an_nguon.py              mã nguồn notebook đáp án (chạy được như script)
    └── tao_tai_lieu.py              dựng lại đề + đáp án ra PDF (Chrome headless) và DOCX
```

## Phát đề

Nén `de_thi/`, `du_lieu/`, `bai_lam/` thành một file zip. **Không** gửi kèm `dap_an/` và `cong_cu/` (script sinh dữ liệu để lộ các lỗi cài sẵn).

Sinh viên mở `bai_lam/bai_lam_mau.ipynb`, trong đó `DATA_DIR = Path('../du_lieu')` đã trỏ đúng. Nếu dùng Google Colab, tải 4 tập tin dữ liệu lên và đổi thành `Path('.')`.

Môi trường cần: Python ≥ 3.10, `numpy`, `pandas`, `matplotlib`, `seaborn`, `scipy`, `scikit-learn`, `openpyxl` (để đọc Excel). `sqlite3` có sẵn trong Python. Google Colab đã có sẵn tất cả.

## Dựng lại toàn bộ

```bash
cd cong_cu
python3 tao_du_lieu.py
cd ../dap_an
jupytext --to ipynb --output dap_an.ipynb ../cong_cu/dap_an_nguon.py
jupyter nbconvert --to notebook --execute --inplace dap_an.ipynb   # tạo dong_loi.txt, bieu_do.png
cd ../cong_cu
python3 tao_tai_lieu.py                                            # cần Google Chrome
```

Nếu đổi `SEED`, phải cập nhật các con số viết tay trong `tao_tai_lieu.py` theo notebook (riêng bảng A2 tự tính từ dữ liệu; Phần B không phụ thuộc seed). Các bẫy ở A2 (XM003 cao nhất, XD007 vượt lên khi dùng `<=`, XM010 bị loại, XM009 đúng 8 lần trả) được cài cố định, không phụ thuộc seed. Nên kiểm tra lại các cặp tương quan và số thành phần PCA ở D2.
>>>>>>> 09034b2 (First commit)
