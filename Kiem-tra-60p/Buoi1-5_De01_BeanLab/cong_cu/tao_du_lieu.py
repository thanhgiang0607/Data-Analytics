"""Sinh dữ liệu mô phỏng cho bài kiểm tra 60 phút — Lập trình phân tích dữ liệu (Buổi 1–5).

Chạy:  python3 tao_du_lieu.py
Seed cố định nên chạy lại bao nhiêu lần cũng ra đúng bộ dữ liệu này, và đáp án không đổi.
Muốn ra một ĐỀ KHÁC (vd. đề B): đổi SEED rồi chạy lại notebook đáp án để lấy số mới.

Tạo ra 4 tập tin trong ../du_lieu/:
  nhat_ky_ca_lam.txt   nhật ký ca làm việc tháng 6/2025, CÓ DÒNG SAI ĐỊNH DẠNG
  giao_dich.csv        ~2.400 giao dịch 01/2025–06/2025, CÓ LỖI CÀI SẴN
  khach_hang.xlsx      2 sheet: khach_hang (có dòng trùng, giới tính viết lộn xộn), hang_thanh_vien
  san_pham.db          SQLite, 2 bảng: san_pham (cột khoá tên product_id, có cột dư thừa), chi_nhanh
"""
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2026
rng = np.random.default_rng(SEED)
OUT = Path(__file__).resolve().parent.parent / 'du_lieu'
OUT.mkdir(parents=True, exist_ok=True)

CN = ['CN01', 'CN02', 'CN03', 'CN04']

# ================================================================ san_pham.db
SP = [
    ('SP01', 'Cà phê đen', 'Cà phê', 29000, 9000),
    ('SP02', 'Cà phê sữa', 'Cà phê', 35000, 11000),
    ('SP03', 'Bạc xỉu', 'Cà phê', 39000, 12000),
    ('SP04', 'Latte', 'Cà phê', 55000, 18000),
    ('SP05', 'Cold brew', 'Cà phê', 59000, 17000),
    ('SP06', 'Trà đào', 'Trà', 45000, 12000),
    ('SP07', 'Trà vải', 'Trà', 45000, 13000),
    ('SP08', 'Trà sen vàng', 'Trà', 49000, 14000),
    ('SP09', 'Trà sữa oolong', 'Trà', 52000, 16000),
    ('SP10', 'Matcha đá xay', 'Đá xay', 62000, 21000),
    ('SP11', 'Cookie đá xay', 'Đá xay', 65000, 22000),
    ('SP12', 'Chocolate đá xay', 'Đá xay', 60000, 20000),
    ('SP13', 'Croissant', 'Bánh', 35000, 15000),
    ('SP14', 'Tiramisu', 'Bánh', 45000, 20000),
    ('SP15', 'Bánh mì que', 'Bánh', 25000, 9000),
    ('SP16', 'Cheesecake', 'Bánh', 49000, 22000),
]
san_pham = pd.DataFrame(SP, columns=['product_id', 'ten_sp', 'nhom', 'gia_ban', 'gia_von'])
# Lỗi cài sẵn (tích hợp): cột dẫn xuất dư thừa gia_ban_vat = gia_ban × 1,08
san_pham.insert(4, 'gia_ban_vat', (san_pham['gia_ban'] * 1.08).round().astype(int))
chi_nhanh = pd.DataFrame({
    'ma_cn': CN,
    'ten_cn': ['Nguyễn Huệ', 'Thảo Điền', 'Hoàn Kiếm', 'Cầu Giấy'],
    'thanh_pho': ['Hồ Chí Minh', 'Hồ Chí Minh', 'Hà Nội', 'Hà Nội'],
})
db = OUT / 'san_pham.db'
db.unlink(missing_ok=True)
con = sqlite3.connect(db)
san_pham.to_sql('san_pham', con, index=False)
chi_nhanh.to_sql('chi_nhanh', con, index=False)
con.close()

# ================================================================ khach_hang.xlsx
N_KH = 400
HANG = ['Thường', 'Bạc', 'Vàng']
kh = pd.DataFrame({
    'ma_kh': [f'KH{i:04d}' for i in range(1, N_KH + 1)],
    'gioi_tinh': rng.choice(['Nam', 'Nữ'], N_KH, p=[0.42, 0.58]).astype(object),
    'nam_sinh': rng.integers(1972, 2008, N_KH),
    'hang_thanh_vien': rng.choice(HANG, N_KH, p=[0.58, 0.30, 0.12]),
    'ngay_dang_ky': (pd.Timestamp('2022-06-01')
                     + pd.to_timedelta(rng.integers(0, 900, N_KH), unit='D')).date,
})
hang_goc = kh['hang_thanh_vien'].copy()

# Lỗi cài sẵn (không nhất quán): giới tính viết nhiều kiểu
idx = rng.choice(N_KH, 44, replace=False)
bien_the = {'Nam': ['nam', 'NAM', 'M', ' Nam '], 'Nữ': ['nữ', 'NỮ', 'F', 'Nu']}
kh.loc[idx, 'gioi_tinh'] = [rng.choice(bien_the[g]) for g in kh.loc[idx, 'gioi_tinh']]

# Lỗi cài sẵn (tích hợp): 6 khách bị ghi trùng hoàn toàn → merge không khử trùng sẽ nhân dòng
trung = kh.loc[np.sort(rng.choice(N_KH, 6, replace=False))]
kh_xuat = pd.concat([kh, trung]).sort_index(kind='stable').reset_index(drop=True)
hang_tv = pd.DataFrame({'hang_thanh_vien': HANG, 'chiet_khau': [0.00, 0.05, 0.10]})

with pd.ExcelWriter(OUT / 'khach_hang.xlsx') as w:
    kh_xuat.to_excel(w, sheet_name='khach_hang', index=False)
    hang_tv.to_excel(w, sheet_name='hang_thanh_vien', index=False)

# ================================================================ giao_dich.csv
N = 2400
ngay = pd.date_range('2025-01-01', '2025-06-30', freq='D')
he_so_thang = {1: 1.15, 2: 0.80, 3: 0.95, 4: 1.00, 5: 1.05, 6: 1.20}
w_ngay = np.array([he_so_thang[d.month] * (1.3 if d.dayofweek >= 5 else 1.0) for d in ngay])
ngay_gd = pd.DatetimeIndex(np.sort(rng.choice(ngay.values, N, p=w_ngay / w_ngay.sum())))
gio = rng.choice(np.arange(7, 22), N, p=np.array([6, 9, 8, 6, 5, 7, 8, 6, 5, 5, 6, 7, 6, 4, 2]) / 90)
thoi_diem = ngay_gd + pd.to_timedelta(gio, unit='h') + pd.to_timedelta(rng.integers(0, 60, N), unit='min')
thoi_diem = thoi_diem.sort_values()
thang = thoi_diem.month.to_numpy()

ma_cn = rng.choice(CN, N, p=[0.32, 0.22, 0.26, 0.20])

# Kênh: tỉ trọng giao hàng tăng dần theo tháng → biểu đồ đường có xu hướng rõ
KENH = ['Tại quán', 'Mang đi', 'Giao hàng']
p_giao = 0.10 + 0.045 * (thang - 1)
u = rng.random(N)
kenh = np.where(u < p_giao, 'Giao hàng',
                np.where(u < p_giao + (1 - p_giao) * 0.45, 'Mang đi', 'Tại quán')).astype(object)

# Khách: 28% vãng lai (trống), 1,5% mã không có trong danh sách; khách hạng cao mua thường xuyên hơn
w_kh = hang_goc.map({'Thường': 1.0, 'Bạc': 1.4, 'Vàng': 2.0}).to_numpy()
ma_kh = rng.choice(kh['ma_kh'].to_numpy(), N, p=w_kh / w_kh.sum()).astype(object)
r = rng.random(N)
ma_kh[r < 0.28] = ''
la = (r >= 0.28) & (r < 0.295)
ma_kh[la] = [f'KH{i:04d}' for i in rng.integers(401, 460, la.sum())]

# Sản phẩm: 1% mã không có trong danh mục
ma_sp = rng.choice([s[0] for s in SP], N,
                   p=np.array([10, 12, 9, 7, 5, 9, 6, 5, 6, 5, 4, 4, 6, 4, 5, 3]) / 100).astype(object)
ma_sp[rng.random(N) < 0.01] = rng.choice(['SP17', 'SP18'])

hang_gd = pd.Series(ma_kh).map(kh.set_index('ma_kh')['hang_thanh_vien']).fillna('')
lam = (0.45 + hang_gd.map({'Thường': 0.15, 'Bạc': 0.35, 'Vàng': 0.75}).fillna(0).to_numpy()
       + np.where(kenh == 'Giao hàng', 0.5, 0.0))
so_luong = np.minimum(1 + rng.poisson(lam), 6).astype(float)

# Thời gian chờ (phút): lệch phải, giao hàng lâu hơn hẳn
cho = np.where(kenh == 'Giao hàng', rng.lognormal(3.15, 0.35, N),
               np.where(kenh == 'Mang đi', rng.lognormal(1.45, 0.45, N), rng.lognormal(1.65, 0.50, N)))
# Điểm đánh giá giảm khi chờ lâu → heatmap có tương quan âm
diem = np.clip(np.rint(4.35 - 0.045 * cho + rng.normal(0, 0.8, N)), 1, 5).astype(int)
NHAN = {1: 'Rất tệ', 2: 'Tệ', 3: 'Bình thường', 4: 'Tốt', 5: 'Rất tốt'}
danh_gia = pd.Series(diem).map(NHAN).astype(object)
danh_gia[rng.random(N) < 0.20] = np.nan

gd = pd.DataFrame({
    'ma_gd': [f'GD{i:05d}' for i in range(1, N + 1)],
    'thoi_gian': thoi_diem.strftime('%d/%m/%Y %H:%M'),
    'ma_cn': ma_cn,
    'ma_kh': ma_kh,
    'ma_sp': ma_sp,
    'so_luong': so_luong,
    'kenh': kenh,
    'thoi_gian_cho': np.round(cho, 1),
    'danh_gia': danh_gia,
})

# --- Lỗi 1: kênh viết lộn xộn (18%), có cả bản không dấu
idx = rng.choice(N, int(N * 0.18), replace=False)
bien_the = {'Tại quán': ['tại quán', 'TẠI QUÁN', ' Tại quán', 'Tai quan'],
            'Mang đi': ['mang đi', 'MANG ĐI', 'Mang đi  ', 'Mang di'],
            'Giao hàng': ['giao hàng', 'GIAO HÀNG', ' Giao hàng', 'Giao hang']}
gd.loc[idx, 'kenh'] = [rng.choice(bien_the[k]) for k in gd.loc[idx, 'kenh']]

# --- Lỗi 2: so_luong trống
gd.loc[rng.choice(N, 60, replace=False), 'so_luong'] = np.nan
# --- Lỗi 3: so_luong nhập nhầm
con = gd.index[gd['so_luong'].notna()]
gd.loc[rng.choice(con, 8, replace=False), 'so_luong'] = [40, 60, 80, 100, 120, 200, 999, 999]
# --- Lỗi 4: thời gian chờ âm (lỗi thiết bị)
gd.loc[rng.choice(N, 6, replace=False), 'thoi_gian_cho'] = [-1.0, -2.5, -3.0, -4.0, -5.0, -12.0]

gd['so_luong'] = gd['so_luong'].astype('Int64')
# --- Lỗi 5: giao dịch ghi trùng hoàn toàn (làm SAU CÙNG)
dup = gd.loc[np.sort(rng.choice(N, 30, replace=False))]
gd = pd.concat([gd, dup]).sort_index(kind='stable').reset_index(drop=True)
gd.to_csv(OUT / 'giao_dich.csv', index=False, encoding='utf-8')

# ================================================================ nhat_ky_ca_lam.txt
NV = [f'NV{i:03d}' for i in range(1, 15)]
ti_le_tre = dict.fromkeys(NV, 0.08)
ti_le_tre.update(NV004=0.16, NV009=0.30, NV012=0.22)
CA = {'SANG': (7 * 60, 11 * 60), 'CHIEU': (12 * 60, 17 * 60), 'TOI': (17 * 60, 22 * 60)}
# NV001–NV003 xoay ca ở cả 4 chi nhánh; những người khác gắn với 1–2 chi nhánh
nha = {nv: CN for nv in NV[:3]}
for i, nv in enumerate(NV[3:]):
    nha[nv] = [CN[i % 4]] if i % 3 else [CN[i % 4], CN[(i + 1) % 4]]


def hhmm(phut):
    return f'{phut // 60:02d}:{phut % 60:02d}'


ban_ghi = []
for ngay_lam in pd.date_range('2025-06-01', '2025-06-30'):
    for ca, (bd, kt) in CA.items():
        for cn in CN:
            ung_vien = [nv for nv in NV if cn in nha[nv]]
            for nv in rng.choice(ung_vien, 2, replace=False):         # mỗi ca 2 nhân viên
                tre = rng.random() < ti_le_tre[nv]
                lech = int(rng.integers(6, 26)) if tre else int(rng.integers(-15, 6))   # +5 phút KHÔNG tính trễ
                ban_ghi.append((ngay_lam, str(nv), cn, ca, bd + lech, kt + int(rng.integers(-5, 21))))

# Bẫy: NV015 chỉ có 4 ca, trễ 3 — tỉ lệ cao nhất nhưng không đủ 10 ca để xét
for k, (d, ca, lech) in enumerate([(3, 'SANG', 12), (11, 'TOI', 9), (19, 'CHIEU', 3), (26, 'SANG', 18)]):
    bd, kt = CA[ca]
    ban_ghi.append((pd.Timestamp(f'2025-06-{d:02d}'), 'NV015', CN[k % 4], ca, bd + lech, kt))
ban_ghi.sort(key=lambda b: (b[0], b[4]))

dong = [f'{d:%Y-%m-%d} | {nv} | {cn} | ca={ca} | vao={hhmm(v)} | ra={hhmm(ra)}'
        for d, nv, cn, ca, v, ra in ban_ghi]


def lam_hong(d, nv, cn, ca, v, ra, kieu):
    if kieu == 0:
        nv = nv[:2] + nv[3:]                                        # NV07 — thiếu 1 chữ số
    elif kieu == 1:
        ca = ca.lower()                                             # ca viết thường
    elif kieu == 2:
        return f'{d} | {nv} | {cn} | ca={ca} | vao={v}'             # thiếu trường ra
    elif kieu == 3:
        v = v.replace(':', 'h')                                     # vao=07h02 — sai dấu phân cách giờ
    elif kieu == 4:
        return f'{d}|{nv}|{cn}|ca={ca}|vao={v}|ra={ra}'             # thiếu khoảng trắng quanh '|'
    elif kieu == 5:
        ca = 'DEM'                                                  # loại ca ngoài danh sách
    return f'{d} | {nv} | {cn} | ca={ca} | vao={v} | ra={ra}'


chen = []
for kieu in range(6):
    for _ in range(7):
        d = f'2025-06-{int(rng.integers(1, 31)):02d}'
        ca = str(rng.choice(list(CA)))
        v = hhmm(CA[ca][0] + int(rng.integers(-10, 15)))
        chen.append(lam_hong(d, str(rng.choice(NV)), str(rng.choice(CN)), ca, v, hhmm(CA[ca][1]), kieu))
chen += [''] * 10 + ['# --- dong bo lai tu may cham cong ---'] * 3
rng.shuffle(chen)
for x in chen:
    dong.insert(int(rng.integers(0, len(dong) + 1)), x)

tieu_de = ['# NHAT KY CA LAM VIEC - THANG 06/2025',
           '# Dinh dang: YYYY-MM-DD | NVxxx | CNxx | ca=CA | vao=HH:MM | ra=HH:MM',
           '# CA: SANG (07:00-11:00), CHIEU (12:00-17:00), TOI (17:00-22:00)']
(OUT / 'nhat_ky_ca_lam.txt').write_text('\n'.join(tieu_de + dong) + '\n', encoding='utf-8')

print('Đã tạo:', *sorted(p.name for p in OUT.iterdir()), sep='\n  ')
