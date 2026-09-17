"""Sinh dữ liệu mô phỏng cho bài kiểm tra 60 phút — Lập trình phân tích dữ liệu (Buổi 1–5) · ĐỀ SỐ 02.

Bối cảnh: EcoRide — dịch vụ cho thuê xe đạp và xe điện công cộng tại Đà Nẵng và Hội An.

Chạy:  python3 tao_du_lieu.py
Seed cố định nên chạy lại bao nhiêu lần cũng ra đúng bộ dữ liệu này, và đáp án không đổi.

Tạo ra 4 tập tin trong ../du_lieu/:
  nhat_ky_khoa_xe.txt  nhật ký khoá thông minh của xe điện tháng 7/2025, CÓ DÒNG SAI ĐỊNH DẠNG
  chuyen_di.csv        ~2.600 chuyến đi 03/2025–08/2025, CÓ LỖI CÀI SẴN
  khach_hang.xlsx      2 sheet: khach_hang (trùng mã, số điện thoại nhiều định dạng), goi_cuoc
  he_thong.db          SQLite, 2 bảng: tram (khoá tên id), loai_xe (có cột dẫn xuất dư thừa)
"""
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2502
rng = np.random.default_rng(SEED)
OUT = Path(__file__).resolve().parent.parent / 'du_lieu'
OUT.mkdir(parents=True, exist_ok=True)

# ================================================================ he_thong.db
TRAM = [
    ('DN01', 'Cầu Rồng', 'Đà Nẵng', 30),
    ('DN02', 'Mỹ Khê', 'Đà Nẵng', 24),
    ('DN03', 'Chợ Hàn', 'Đà Nẵng', 20),
    ('DN04', 'Sơn Trà', 'Đà Nẵng', 16),
    ('HA01', 'Phố cổ', 'Hội An', 28),
    ('HA02', 'An Bàng', 'Hội An', 18),
]
tram = pd.DataFrame(TRAM, columns=['id', 'ten_tram', 'thanh_pho', 'suc_chua'])
LOAI = [('XC', 'Xe đạp', 3000, 200), ('XD', 'Xe đạp điện', 5000, 500), ('XM', 'Xe máy điện', 8000, 1000)]
loai_xe = pd.DataFrame(LOAI, columns=['ma_loai', 'ten_loai', 'phi_mo_khoa', 'gia_phut'])
# Lỗi cài sẵn (tích hợp): cột dẫn xuất dư thừa gia_30_phut = phi_mo_khoa + 30 × gia_phut
loai_xe['gia_30_phut'] = loai_xe['phi_mo_khoa'] + 30 * loai_xe['gia_phut']
db = OUT / 'he_thong.db'
db.unlink(missing_ok=True)
con = sqlite3.connect(db)
tram.to_sql('tram', con, index=False)
loai_xe.to_sql('loai_xe', con, index=False)
con.close()

# ================================================================ khach_hang.xlsx
N_KH = 500
GOI = ['Tiêu chuẩn', 'Sinh viên', 'Thân thiết']
goi = rng.choice(GOI, N_KH, p=[0.55, 0.28, 0.17])
nam_sinh = np.where(goi == 'Sinh viên', rng.integers(2003, 2008, N_KH), rng.integers(1968, 2003, N_KH))
dau_so = rng.choice(['090', '093', '035', '038', '070', '077', '098'], N_KH)
sdt = np.array([d + ''.join(map(str, rng.integers(0, 10, 7))) for d in dau_so], dtype=object)
kh = pd.DataFrame({
    'ma_kh': [f'KH{i:04d}' for i in range(1, N_KH + 1)],
    'nam_sinh': nam_sinh,
    'so_dien_thoai': sdt,
    'goi_cuoc': goi,
    'ngay_dang_ky': (pd.Timestamp('2023-01-01')
                     + pd.to_timedelta(rng.integers(0, 780, N_KH), unit='D')).date,
})
goi_goc = kh['goi_cuoc'].copy()


def dinh_dang_lai(so, kieu):
    """Viết lại một số điện thoại 10 chữ số theo kiểu khác (cùng một số)."""
    if kieu == 0:
        return f'{so[:4]} {so[4:7]} {so[7:]}'          # 0905 123 456
    if kieu == 1:
        return '+84' + so[1:]                           # +84905123456
    if kieu == 2:
        return f'84-{so[1:4]}-{so[4:7]}-{so[7:]}'       # 84-905-123-456
    return f'{so[:4]}.{so[4:7]}.{so[7:]}'               # 0905.123.456


# Lỗi cài sẵn (không nhất quán): số điện thoại viết nhiều kiểu
idx = rng.choice(N_KH, 110, replace=False)
kh.loc[idx, 'so_dien_thoai'] = [dinh_dang_lai(s, int(rng.integers(0, 4))) for s in kh.loc[idx, 'so_dien_thoai']]
# Lỗi cài sẵn: 9 số điện thoại thiếu 1 chữ số (không sửa được — chỉ đếm)
con_lai = np.setdiff1d(np.arange(N_KH), idx)
thieu = rng.choice(con_lai, 9, replace=False)
kh.loc[thieu, 'so_dien_thoai'] = [s[:-1] for s in kh.loc[thieu, 'so_dien_thoai']]

# Lỗi cài sẵn (tích hợp): 8 khách bị ghi 2 lần — 3 bản trùng hoàn toàn, 5 bản khác cách viết số điện thoại
# → drop_duplicates() không có subset chỉ bỏ được 3, merge vẫn nhân dòng
trung = np.sort(rng.choice(np.setdiff1d(con_lai, thieu), 8, replace=False))
ban_sao = kh.loc[trung].copy()
ban_sao.iloc[3:, ban_sao.columns.get_loc('so_dien_thoai')] = [
    dinh_dang_lai(s, k) for s, k in zip(ban_sao['so_dien_thoai'].iloc[3:], [0, 1, 2, 3, 1])]
kh_xuat = pd.concat([kh, ban_sao]).sort_index(kind='stable').reset_index(drop=True)
goi_cuoc = pd.DataFrame({'goi_cuoc': GOI, 'giam_gia': [0.00, 0.20, 0.10]})

with pd.ExcelWriter(OUT / 'khach_hang.xlsx') as w:
    kh_xuat.to_excel(w, sheet_name='khach_hang', index=False)
    goi_cuoc.to_excel(w, sheet_name='goi_cuoc', index=False)

# ================================================================ chuyen_di.csv
N = 2600
ngay = pd.date_range('2025-03-01', '2025-08-31', freq='D')
he_so_thang = {3: 0.80, 4: 0.90, 5: 1.00, 6: 1.15, 7: 1.30, 8: 1.25}
w_ngay = np.array([he_so_thang[d.month] * (1.3 if d.dayofweek >= 5 else 1.0) for d in ngay])
ngay_cd = pd.DatetimeIndex(np.sort(rng.choice(ngay.values, N, p=w_ngay / w_ngay.sum())))
w_gio = np.array([3, 9, 12, 7, 5, 4, 4, 4, 4, 5, 6, 8, 12, 11, 7, 5, 3, 1], dtype=float)   # 5h → 22h
gio = rng.choice(np.arange(5, 23), N, p=w_gio / w_gio.sum())
bat_dau = ngay_cd + pd.to_timedelta(gio, unit='h') + pd.to_timedelta(rng.integers(0, 60, N), unit='min')
thu_tu = np.argsort(bat_dau.values, kind='stable')
bat_dau = bat_dau[thu_tu]
gio = gio[thu_tu]
thang = bat_dau.month.to_numpy()

MA_TRAM = [t[0] for t in TRAM]
tram_muon = rng.choice(MA_TRAM, N, p=[0.21, 0.19, 0.16, 0.10, 0.21, 0.13]).astype(object)
thanh_pho = np.where(pd.Series(tram_muon).str.startswith('DN'), 'DN', 'HA')
tram_tra = np.where(thanh_pho == 'DN', rng.choice(MA_TRAM[:4], N, p=[0.3, 0.3, 0.25, 0.15]),
                    rng.choice(MA_TRAM[4:], N, p=[0.6, 0.4])).astype(object)

# Khách: 34% khách lẻ (trống), 1,5% mã không có trong danh sách; khách gói cao đi nhiều hơn
w_kh = goi_goc.map({'Tiêu chuẩn': 1.0, 'Sinh viên': 1.6, 'Thân thiết': 2.4}).to_numpy()
ma_kh = rng.choice(kh['ma_kh'].to_numpy(), N, p=w_kh / w_kh.sum()).astype(object)
r = rng.random(N)
ma_kh[r < 0.34] = ''
la = (r >= 0.34) & (r < 0.355)
ma_kh[la] = [f'KH{i:04d}' for i in rng.integers(501, 560, la.sum())]
goi_cd = pd.Series(ma_kh).map(kh.set_index('ma_kh')['goi_cuoc']).fillna('').to_numpy()

# Loại xe: xe máy điện tăng dần theo tháng; sinh viên chuộng xe đạp, khách thân thiết chuộng xe máy điện
LOAI_TEN = ['Xe đạp', 'Xe đạp điện', 'Xe máy điện']
p_xm = 0.10 + 0.06 * (thang - 3)
p_xc = 0.46 - 0.05 * (thang - 3)
W = np.column_stack([p_xc, 1 - p_xc - p_xm, p_xm])
W[goi_cd == 'Sinh viên'] *= [2.2, 1.0, 0.5]
W[goi_cd == 'Thân thiết'] *= [0.5, 1.0, 1.8]
W /= W.sum(axis=1, keepdims=True)
k = (rng.random(N)[:, None] > W.cumsum(axis=1)).sum(axis=1)
loai = np.array(LOAI_TEN, dtype=object)[k]

# Thời lượng (phút): lệch phải, xe máy điện đi lâu hơn, Hội An (khách du lịch) lâu hơn
mu = np.array([np.log(16), np.log(22), np.log(34)])[k] + np.where(thanh_pho == 'HA', 0.25, 0.0)
phut = np.maximum(np.rint(rng.lognormal(mu, 0.5)), 3).astype(int)
toc_do = np.clip(rng.normal(np.array([11.0, 16.0, 23.0])[k], np.array([2.0, 2.5, 3.5])[k]), 5, None)
km = np.round(toc_do * phut / 60 * rng.normal(1, 0.08, N), 1)

# Đánh giá: xe máy điện được thích hơn; giờ tan tầm ở Đà Nẵng bị chấm thấp (kẹt xe); đi quá lâu mệt
tan_tam = (gio >= 16) & (gio < 20) & (thanh_pho == 'DN')
diem = np.clip(np.rint(4.1 + np.array([-0.25, 0.0, 0.35])[k] - 0.7 * tan_tam
                       - 0.012 * np.maximum(phut - 40, 0) + rng.normal(0, 0.75, N)), 1, 5).astype(int)
NHAN = {1: 'Rất không hài lòng', 2: 'Không hài lòng', 3: 'Bình thường', 4: 'Hài lòng', 5: 'Rất hài lòng'}
danh_gia = pd.Series(diem).map(NHAN).astype(object)
danh_gia[rng.random(N) < 0.22] = np.nan

ket_thuc = bat_dau + pd.to_timedelta(phut, unit='min')
cd = pd.DataFrame({
    'ma_cd': [f'CD{i:05d}' for i in range(1, N + 1)],
    'bat_dau': bat_dau.strftime('%d/%m/%Y %H:%M'),
    'ket_thuc': ket_thuc.strftime('%d/%m/%Y %H:%M'),
    'tram_muon': tram_muon,
    'tram_tra': tram_tra,
    'ma_kh': ma_kh,
    'loai_xe': loai,
    'quang_duong_km': km,
    'danh_gia': danh_gia,
})

# --- Lỗi 1: loại xe viết lộn xộn (20%), có bản không dấu; "xe đạp" là tiền tố của "xe đạp điện"
idx = rng.choice(N, int(N * 0.20), replace=False)
bien_the = {'Xe đạp': ['xe đạp', 'XE ĐẠP', ' Xe đạp', 'Xe dap'],
            'Xe đạp điện': ['xe đạp điện', 'XE ĐẠP ĐIỆN', 'Xe đạp điện  ', 'Xe dap dien'],
            'Xe máy điện': ['xe máy điện', 'XE MÁY ĐIỆN', ' Xe máy điện', 'Xe may dien']}
cd.loc[idx, 'loai_xe'] = [rng.choice(bien_the[x]) for x in cd.loc[idx, 'loai_xe']]

tu_do = np.arange(N)
# --- Lỗi 2: mất kết nối — không ghi được giờ trả xe
mat = rng.choice(tu_do, 48, replace=False)
cd.loc[mat, 'ket_thuc'] = np.nan
tu_do = np.setdiff1d(tu_do, mat)
# --- Lỗi 3: lỗi đồng hồ — giờ trả xe sớm hơn giờ mượn
lech = rng.choice(tu_do, 9, replace=False)
cd.loc[lech, 'ket_thuc'] = (bat_dau[lech] - pd.to_timedelta(phut[lech], unit='min')).strftime('%d/%m/%Y %H:%M')
tu_do = np.setdiff1d(tu_do, lech)
# --- Lỗi 4: mở khoá thử — trả xe sau 1 phút ở ngay trạm mượn, quãng đường ~0
thu = rng.choice(tu_do, 27, replace=False)
cd.loc[thu, 'ket_thuc'] = (bat_dau[thu] + pd.to_timedelta(1, unit='min')).strftime('%d/%m/%Y %H:%M')
cd.loc[thu, 'tram_tra'] = cd.loc[thu, 'tram_muon']
cd.loc[thu, 'quang_duong_km'] = rng.choice([0.0, 0.0, 0.1], 27)
tu_do = np.setdiff1d(tu_do, thu)
# --- Lỗi 5: quãng đường trống / GPS lỗi (quá lớn hoặc âm)
cd.loc[rng.choice(tu_do, 40, replace=False), 'quang_duong_km'] = np.nan
gps = rng.choice(np.setdiff1d(tu_do, cd.index[cd['quang_duong_km'].isna()]), 11, replace=False)
cd.loc[gps, 'quang_duong_km'] = [125.4, 188.0, 240.7, 312.5, 999.9, 999.9, -0.8, -1.5, -2.3, -3.1, -12.0]
# --- Lỗi 6: trạm DN07 (trạm thử nghiệm đã gỡ) không có trong danh mục
cd.loc[rng.choice(N, 24, replace=False), 'tram_muon'] = 'DN07'
# --- Lỗi 7: chuyến ghi trùng hoàn toàn (làm SAU CÙNG)
dup = cd.loc[np.sort(rng.choice(N, 35, replace=False))]
cd = pd.concat([cd, dup]).sort_index(kind='stable').reset_index(drop=True)
cd.to_csv(OUT / 'chuyen_di.csv', index=False, encoding='utf-8')

# ================================================================ nhat_ky_khoa_xe.txt
# Mỗi chuyến của một xe điện sinh 2 dòng: MUON (pin lúc mở khoá) và TRA (pin lúc khoá lại)
XE_DN = [f'XD{i:03d}' for i in range(1, 10)] + [f'XM{i:03d}' for i in range(1, 7)]
XE_HA = [f'XD{i:03d}' for i in range(10, 15)] + [f'XM{i:03d}' for i in range(7, 10)]
TRAM_TP = {'DN': MA_TRAM[:4], 'HA': MA_TRAM[4:]}
# (số lần trả, số lần trả pin < 20%, số lần trả pin đúng 20%)
KICH_BAN = {nv: (int(rng.integers(12, 21)), None, int(rng.integers(0, 2))) for nv in XE_DN + XE_HA}
KICH_BAN.update(XM003=(20, 7, 1),      # cao nhất khi so sánh đúng  pin < 20
                XD007=(22, 7, 4),      # vượt lên khi dùng nhầm   pin <= 20
                XM009=(8, 2, 0),       # đúng bằng toi_thieu = 8 → vẫn được xét
                XM010=(5, 4, 0))       # bẫy: tỉ lệ cao nhất nhưng chỉ 5 lần trả
DI_CHUYEN = {'XD004': 3, 'XD012': 2, 'XM008': 3, 'XM010': 2}    # số chuyến ở thành phố còn lại


def hhmm(phut_):
    return f'{phut_ // 60:02d}:{phut_ % 60:02d}'


ban_ghi = []
for xe, (n, n_yeu, n_20) in KICH_BAN.items():
    if n_yeu is None:
        n_yeu = int(rng.binomial(n, 0.10))
    pin_tra = np.concatenate([rng.integers(4, 20, n_yeu), np.full(n_20, 20), rng.integers(21, 96, n - n_yeu - n_20)])
    rng.shuffle(pin_tra)
    tp_nha = 'DN' if (xe in XE_DN or xe == 'XM010') else 'HA'
    tp_khac = 'HA' if tp_nha == 'DN' else 'DN'
    so_khac = DI_CHUYEN.get(xe, 0)
    tp = [tp_khac] * so_khac + [tp_nha] * (n - so_khac)
    rng.shuffle(tp)
    ngay_ = np.sort(rng.integers(1, 32, n))
    for j in range(n):
        tram_ = TRAM_TP[tp[j]]
        m = int(rng.integers(6 * 60, 21 * 60))
        t = m + int(rng.integers(8, 90))
        p_tra = int(pin_tra[j])
        p_muon = min(100, p_tra + int(rng.integers(4, 36)))
        d = f'2025-07-{ngay_[j]:02d}'
        ban_ghi.append((d, m, str(rng.choice(tram_)), xe, 'MUON', p_muon))
        ban_ghi.append((d, min(t, 23 * 60 + 59), str(rng.choice(tram_)), xe, 'TRA', p_tra))
ban_ghi.sort(key=lambda b: (b[0], b[1], b[4] == 'TRA'))
dong = [f'{d} {hhmm(m)} | {t} | {xe} | {sk} | pin={p}%' for d, m, t, xe, sk, p in ban_ghi]


def lam_hong(d, m, t, xe, sk, p, kieu):
    if kieu == 0:
        xe = xe[:2] + xe[3:]                                        # XD07 — thiếu 1 chữ số
    elif kieu == 1:
        sk = sk.lower()                                             # tra — sự kiện viết thường
    elif kieu == 2:
        return f'{d} {m} | {t} | {xe} | {sk}'                       # thiếu trường pin
    elif kieu == 3:
        m = m.replace(':', 'h')                                     # 07h15 — sai dấu phân cách giờ
    elif kieu == 4:
        t = 'QN' + t[2:]                                            # QN01 — trạm ngoài hệ thống
    elif kieu == 5:
        return f'{d} {m} | {t} | {xe} | {sk} | pin={p}% | gps=off'  # thừa trường ở cuối dòng
    return f'{d} {m} | {t} | {xe} | {sk} | pin={p}%'


chen = []
for kieu in range(6):
    for _ in range(8):
        chen.append(lam_hong(f'2025-07-{int(rng.integers(1, 32)):02d}', hhmm(int(rng.integers(6 * 60, 22 * 60))),
                             str(rng.choice(MA_TRAM)), str(rng.choice(XE_DN + XE_HA)),
                             str(rng.choice(['MUON', 'TRA'])), int(rng.integers(5, 100)), kieu))
chen += [''] * 12 + ['# --- mat ket noi may chu, dong bo lai ---'] * 4
rng.shuffle(chen)
for x in chen:
    dong.insert(int(rng.integers(0, len(dong) + 1)), x)

tieu_de = ['# NHAT KY KHOA THONG MINH - XE DIEN ECORIDE - THANG 07/2025',
           '# Dinh dang: YYYY-MM-DD HH:MM | TRAM | MA_XE | SU_KIEN | pin=N%',
           '# SU_KIEN: MUON (mo khoa), TRA (khoa lai) - pin la % pin tai thoi diem ghi']
(OUT / 'nhat_ky_khoa_xe.txt').write_text('\n'.join(tieu_de + dong) + '\n', encoding='utf-8')

print('Đã tạo:', *sorted(p.name for p in OUT.iterdir()), sep='\n  ')
