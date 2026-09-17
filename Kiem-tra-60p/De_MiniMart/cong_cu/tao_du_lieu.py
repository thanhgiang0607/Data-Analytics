"""Sinh dữ liệu mô phỏng cho bài kiểm tra 60 phút — Lập trình phân tích dữ liệu.

Chạy:  python3 tao_du_lieu.py
Seed cố định nên chạy lại bao nhiêu lần cũng ra đúng bộ dữ liệu này, và đáp án không đổi.
Muốn ra một ĐỀ KHÁC (vd. đề B): đổi SEED rồi chạy lại notebook đáp án để lấy số mới.

Tạo ra 3 tập tin trong ../du_lieu/:
  khach_hang.xlsx         2 sheet: khach_hang (500 dòng), hang_thanh_vien (3 dòng)
  don_hang.csv            ~2.400 đơn, 01/2025–06/2025, CÓ LỖI CÀI SẴN
  nhat_ky_giao_hang.txt   nhật ký giao hàng tháng 6/2025, CÓ DÒNG SAI ĐỊNH DẠNG
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2025
rng = np.random.default_rng(SEED)
OUT = Path(__file__).resolve().parent.parent / 'du_lieu'
OUT.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------- khach_hang.xlsx
N_KH = 500
HANG = ['Thường', 'Bạc', 'Vàng']
kh = pd.DataFrame({
    'ma_kh': [f'KH{i:04d}' for i in range(1, N_KH + 1)],
    'gioi_tinh': rng.choice(['Nam', 'Nữ'], N_KH, p=[0.45, 0.55]),
    'nam_sinh': rng.integers(1965, 2007, N_KH),
    'hang_thanh_vien': rng.choice(HANG, N_KH, p=[0.60, 0.28, 0.12]),
    'ngay_dang_ky': (pd.Timestamp('2022-01-01')
                     + pd.to_timedelta(rng.integers(0, 1095, N_KH), unit='D')).date,
})
hang_tv = pd.DataFrame({'hang_thanh_vien': HANG, 'chiet_khau': [0.00, 0.05, 0.10]})

with pd.ExcelWriter(OUT / 'khach_hang.xlsx') as w:
    kh.to_excel(w, sheet_name='khach_hang', index=False)
    hang_tv.to_excel(w, sheet_name='hang_thanh_vien', index=False)

# ---------------------------------------------------------------- don_hang.csv
N_DON = 2400
ngay = pd.date_range('2025-01-01', '2025-06-30', freq='D')
# Tết (tháng 1) cao, tháng 2 thấp, tháng 6 tăng vọt → câu hỏi tăng trưởng có một đáp án rõ
he_so_thang = {1: 1.20, 2: 0.85, 3: 0.90, 4: 0.95, 5: 1.00, 6: 1.40}
w_ngay = np.array([he_so_thang[d.month] * (1.25 if d.dayofweek >= 5 else 1.0) for d in ngay])
ngay_dat = pd.DatetimeIndex(np.sort(rng.choice(ngay.values, N_DON, p=w_ngay / w_ngay.sum())))

# Khách hạng cao mua thường xuyên hơn; 6% vãng lai (trống), 1,5% mã không có trong danh sách
w_kh = kh['hang_thanh_vien'].map({'Thường': 1.0, 'Bạc': 1.3, 'Vàng': 1.8}).to_numpy()
ma_kh = rng.choice(kh['ma_kh'].to_numpy(), N_DON, p=w_kh / w_kh.sum()).astype(object)
r = rng.random(N_DON)
ma_kh[r < 0.06] = ''
la_ma_la = (r >= 0.06) & (r < 0.075)
ma_kh[la_ma_la] = [f'KH{i:04d}' for i in rng.integers(501, 560, la_ma_la.sum())]

TP = ['Hồ Chí Minh', 'Hà Nội', 'Đà Nẵng', 'Cần Thơ', 'Huế']
thanh_pho = rng.choice(TP, N_DON, p=[0.32, 0.30, 0.15, 0.12, 0.11]).astype(object)

# Hiệu ứng cài sẵn cho câu chi-bình phương: 2 thành phố lớn chuộng ví điện tử
PT = ['Tiền mặt', 'Chuyển khoản', 'Ví điện tử']
lon = np.isin(thanh_pho, TP[:2])
phuong_thuc = np.where(lon,
                       rng.choice(PT, N_DON, p=[0.20, 0.35, 0.45]),
                       rng.choice(PT, N_DON, p=[0.42, 0.33, 0.25]))

DM_GIA = {  # danh mục: (giá thấp nhất, giá cao nhất) nghìn đồng
    'Đồ uống': (12, 45), 'Thực phẩm': (20, 150), 'Gia dụng': (80, 600),
    'Mỹ phẩm': (90, 450), 'Văn phòng phẩm': (5, 60),
}
# Cơ cấu danh mục khác nhau giữa các thành phố → câu "danh mục dẫn đầu" không cùng một đáp án
DM_THEO_TP = {  # Đồ uống, Thực phẩm, Gia dụng, Mỹ phẩm, Văn phòng phẩm
    'Hồ Chí Minh': [0.26, 0.22, 0.10, 0.25, 0.17],
    'Hà Nội':      [0.28, 0.24, 0.10, 0.20, 0.18],
    'Đà Nẵng':     [0.28, 0.30, 0.16, 0.10, 0.16],
    'Cần Thơ':     [0.30, 0.42, 0.08, 0.07, 0.13],
    'Huế':         [0.30, 0.40, 0.08, 0.08, 0.14],
}
danh_muc = np.empty(N_DON, dtype=object)
for tp, p in DM_THEO_TP.items():
    chon = thanh_pho == tp
    danh_muc[chon] = rng.choice(list(DM_GIA), chon.sum(), p=p)
lo = np.array([DM_GIA[d][0] for d in danh_muc])
hi = np.array([DM_GIA[d][1] for d in danh_muc])
don_gia = rng.integers(lo, hi + 1) * 1000

# Hiệu ứng cài sẵn cho câu t-test: khách Vàng mua số lượng lớn hơn.
# Số lượng cũng khác nhau theo danh mục → điền trung vị THEO NHÓM khác với trung vị chung.
hang_don = pd.Series(ma_kh).map(kh.set_index('ma_kh')['hang_thanh_vien'])
lam = (hang_don.map({'Thường': 1.6, 'Bạc': 2.0, 'Vàng': 3.0}).fillna(1.4).to_numpy()
       + pd.Series(danh_muc).map({'Đồ uống': 1.5, 'Văn phòng phẩm': 1.0, 'Thực phẩm': 0.5,
                                  'Mỹ phẩm': 0.0, 'Gia dụng': -0.6}).to_numpy())
so_luong = np.minimum(1 + rng.poisson(lam), 15).astype(float)

danh_gia = rng.choice([1, 2, 3, 4, 5], N_DON, p=[0.05, 0.10, 0.25, 0.35, 0.25]).astype(float)
danh_gia[rng.random(N_DON) < 0.20] = np.nan

dh = pd.DataFrame({
    'ma_don': [f'DH{i:05d}' for i in range(1, N_DON + 1)],
    'ngay_dat': ngay_dat.strftime('%d/%m/%Y'),
    'ma_kh': ma_kh,
    'thanh_pho': thanh_pho,
    'danh_muc': danh_muc,
    'don_gia': don_gia,
    'so_luong': so_luong,
    'phuong_thuc_tt': phuong_thuc,
    'danh_gia': danh_gia,
})

# --- Lỗi cài sẵn 1: tên thành phố viết lộn xộn hoa/thường, thừa khoảng trắng (18% dòng)
idx = rng.choice(N_DON, int(N_DON * 0.18), replace=False)
bien_the = [str.lower, str.upper, lambda s: '  ' + s, lambda s: s + ' ',
            lambda s: ' ' + s.lower() + '  ']
kieu = rng.integers(0, len(bien_the), len(idx))
dh.loc[idx, 'thanh_pho'] = [bien_the[k](dh.at[i, 'thanh_pho']) for i, k in zip(idx, kieu)]

# --- Lỗi cài sẵn 2: so_luong bị trống
dh.loc[rng.choice(N_DON, 70, replace=False), 'so_luong'] = np.nan

# --- Lỗi cài sẵn 3: lỗi nhập liệu so_luong vô lý
con_so = dh.index[dh['so_luong'].notna()]
dh.loc[rng.choice(con_so, 9, replace=False), 'so_luong'] = [120, 150, 200, 250, 300, 500, 999, 999, 999]

dh['so_luong'] = dh['so_luong'].astype('Int64')   # ghi ra CSV là "3" chứ không phải "3.0"
dh['danh_gia'] = dh['danh_gia'].astype('Int64')

# --- Lỗi cài sẵn 4: đơn bị ghi trùng hoàn toàn (làm SAU CÙNG để bản sao giống hệt bản gốc)
dup = dh.loc[np.sort(rng.choice(N_DON, 36, replace=False))]
dh = pd.concat([dh, dup]).sort_index(kind='stable').reset_index(drop=True)

dh.to_csv(OUT / 'don_hang.csv', index=False, encoding='utf-8')

# ---------------------------------------------------------------- nhat_ky_giao_hang.txt
SP = [f'SP{i:02d}' for i in range(1, 11)]
ti_le = dict.fromkeys(SP, 0.90)
ti_le.update(SP03=0.84, SP07=0.72, SP09=0.86)

N_LOG = 880
thoi_diem = (pd.Timestamp('2025-06-01')
             + pd.to_timedelta(rng.integers(0, 30, N_LOG), unit='D')
             + pd.to_timedelta(rng.integers(7 * 60, 21 * 60, N_LOG), unit='min')).sort_values()
shipper = rng.choice(SP, N_LOG).astype(object)
# Bẫy: SP11 chỉ chạy 6 đơn, tỉ lệ thấp nhất nhưng KHÔNG đủ 20 đơn để xét
vi_tri_sp11 = np.sort(rng.choice(N_LOG, 6, replace=False))
shipper[vi_tri_sp11] = 'SP11'
trang_thai_sp11 = iter(['GIAO_THANH_CONG', 'THAT_BAI', 'GIAO_THANH_CONG',
                        'THAT_BAI', 'HOAN_TRA', 'THAT_BAI'])

dong = []
for t, sp in zip(thoi_diem, shipper):
    ma = f'DH{rng.integers(1, N_DON + 1):05d}'
    if sp == 'SP11':
        tt = next(trang_thai_sp11)
    elif rng.random() < ti_le[sp]:
        tt = 'GIAO_THANH_CONG'
    else:
        tt = 'THAT_BAI' if rng.random() < 0.7 else 'HOAN_TRA'
    dong.append(f'{t:%Y-%m-%d %H:%M} | {ma} | {tt} | shipper={sp}')


def lam_hong(t, ma, tt, sp, kieu):
    if kieu == 0:
        ma = ma[:2] + ma[3:]                        # DH1234 — thiếu 1 chữ số
    elif kieu == 1:
        tt = tt.lower()                             # trạng thái viết thường
    elif kieu == 2:
        return f'{t} | {ma} | {tt} | shipper:{sp}'  # sai dấu ':'
    elif kieu == 3:
        return f'{t} | {ma} | {tt}'                 # thiếu trường shipper
    elif kieu == 4:
        tt = 'DANG_GIAO'                            # trạng thái ngoài danh sách
    elif kieu == 5:
        sp = 'SP' + sp[-1]                          # SP7 — thiếu chữ số
    elif kieu == 6:
        return f'{t}|{ma}|{tt}|shipper={sp}'        # thiếu khoảng trắng quanh '|'
    return f'{t} | {ma} | {tt} | shipper={sp}'


chen = []
for kieu in range(7):
    for _ in range(6):
        t = pd.Timestamp('2025-06-01') + pd.Timedelta(days=int(rng.integers(0, 30)),
                                                       minutes=int(rng.integers(420, 1260)))
        chen.append(lam_hong(f'{t:%Y-%m-%d %H:%M}', f'DH{rng.integers(1, N_DON + 1):05d}',
                             rng.choice(['GIAO_THANH_CONG', 'THAT_BAI']), rng.choice(SP), kieu))
chen += [''] * 12 + ['# --- đồng bộ lại từ máy quét ---'] * 2
rng.shuffle(chen)
for x in chen:
    dong.insert(int(rng.integers(0, len(dong) + 1)), x)

tieu_de = ['# NHAT KY GIAO HANG - THANG 06/2025',
           '# Dinh dang: YYYY-MM-DD HH:MM | DHxxxxx | TRANG_THAI | shipper=SPxx',
           '# TRANG_THAI: GIAO_THANH_CONG, THAT_BAI, HOAN_TRA']
(OUT / 'nhat_ky_giao_hang.txt').write_text('\n'.join(tieu_de + dong) + '\n', encoding='utf-8')

print('Đã tạo:', *sorted(p.name for p in OUT.iterdir()), sep='\n  ')
