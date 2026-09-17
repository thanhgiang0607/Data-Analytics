# %% [markdown]
# # ĐÁP ÁN — Bài kiểm tra Lập trình phân tích dữ liệu · Buổi 1–5 (60 phút)
# Notebook tham chiếu cho giảng viên. Các con số in ra ở mỗi ô là **kết quả chuẩn** để đối chiếu bài làm.

# %%
import re
import sqlite3
from collections import Counter
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA

DATA_DIR = Path('../du_lieu')

# %% [markdown]
# ## Phần A — Python thuần: tập tin văn bản, regex, hàm (2,0 điểm)
# ### A1 (1,0 đ) Lọc dòng hợp lệ bằng regex, ghi dòng lỗi ra file

# %%
MAU = re.compile(r'(\d{4}-\d{2}-\d{2}) \| (NV\d{3}) \| (CN\d{2}) \| ca=(SANG|CHIEU|TOI) '
                 r'\| vao=(\d{2}:\d{2}) \| ra=(\d{2}:\d{2})')

hop_le, loi = [], []
with open(DATA_DIR / 'nhat_ky_ca_lam.txt', encoding='utf-8') as f:
    for dong in f:
        dong = dong.rstrip('\n')
        if dong.strip() == '' or dong.startswith('#'):
            continue
        (hop_le if MAU.fullmatch(dong) else loi).append(dong)

with open('ca_loi.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(loi) + '\n')

print('Số dòng hợp lệ      :', len(hop_le))
print('Số dòng không hợp lệ:', len(loi))
print('Số ca theo loại     :', Counter(MAU.fullmatch(d).group(4) for d in hop_le))

# %% [markdown]
# ### A2 (1,0 đ) Hàm thống kê đi trễ & tập chi nhánh

# %%
GIO_BAT_DAU = {'SANG': '07:00', 'CHIEU': '12:00', 'TOI': '17:00'}


def doi_phut(hhmm):
    gio, phut = hhmm.split(':')
    return int(gio) * 60 + int(phut)


def thong_ke_di_tre(cac_dong, tre_cho_phep=5, toi_thieu=10):
    so_ca, so_tre = Counter(), Counter()
    for d in cac_dong:
        _, nv, _, ca, vao, _ = MAU.fullmatch(d).groups()
        so_ca[nv] += 1
        if doi_phut(vao) - doi_phut(GIO_BAT_DAU[ca]) > tre_cho_phep:
            so_tre[nv] += 1
    return {nv: so_tre[nv] / n for nv, n in so_ca.items() if n >= toi_thieu}


kq = thong_ke_di_tre(hop_le)
for nv in sorted(kq):
    print(f'{nv}: {kq[nv]:.3f}')
nv_max = max(kq, key=kq.get)
print(f'\nTỉ lệ đi trễ cao nhất (>= 10 ca): {nv_max} — {kq[nv_max]:.1%}')

chi_nhanh = {}
for d in hop_le:
    _, nv, cn, *_ = MAU.fullmatch(d).groups()
    chi_nhanh.setdefault(nv, set()).add(cn)
print('Làm ở đủ 4 chi nhánh:', sorted(nv for nv, s in chi_nhanh.items() if len(s) == 4))

# Kiểm tra bẫy
tat_ca = thong_ke_di_tre(hop_le, toi_thieu=0)
print('Bị loại vì < 10 ca:', {nv: f'{v:.1%}' for nv, v in tat_ca.items() if nv not in kq})
print('Nếu tính trễ khi >= 5 phút:', max(thong_ke_di_tre(hop_le, tre_cho_phep=4), key=thong_ke_di_tre(hop_le, tre_cho_phep=4).get))

# %% [markdown]
# ## Phần B — NumPy & lập trình hướng đối tượng (2,0 điểm)
# ### B1 (0,5 đ) Ma trận doanh thu chi nhánh × ngày

# %%
M = np.array([[18.2, 16.5, 17.1, 19.4, 24.8, 31.2, 29.5],
              [11.4, 10.9, 12.2, 12.8, 15.1, 21.7, 20.3],
              [15.6, 14.8, 15.9, 16.4, 20.2, 26.9, 25.1],
              [ 9.8, 10.2,  9.5, 11.1, 13.6, 18.4, 17.2]])
THU = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN']

tong_cn = M.sum(axis=1)
print('Tổng tuần theo chi nhánh:', tong_cn.round(1), '→ cao nhất: CN0%d' % (tong_cn.argmax() + 1))
tb_ngay = M.mean(axis=0)
print('TB theo thứ:', tb_ngay.round(2), '→ cao nhất:', THU[tb_ngay.argmax()])
ty_trong = M / M.sum(axis=1, keepdims=True)
print('Tỉ trọng T7 + CN mỗi chi nhánh:', ty_trong[:, 5:].sum(axis=1).round(3))
print('Số ô cao hơn TB của chính chi nhánh đó:', (M > M.mean(axis=1, keepdims=True)).sum())

# %% [markdown]
# ### B2 (0,5 đ) Giải hệ phương trình tìm đơn giá nguyên liệu

# %%
A = np.array([[5, 10, 2],
              [3, 12, 1],
              [8, 6, 4]])
b = np.array([1470, 1069, 2052])
x = np.linalg.solve(A, b)
print('Đơn giá (nghìn đồng): cà phê/kg, sữa/lít, đường/kg =', x)
print('det(A) =', round(np.linalg.det(A), 2), '· kiểm tra A @ x == b:', np.allclose(A @ x, b))

# %% [markdown]
# ### B3 (1,0 đ) Lớp SanPham và SanPhamKhuyenMai

# %%
class SanPham:
    def __init__(self, ma, ten, gia_ban, gia_von):
        self.ma, self.ten, self.gia_von = ma, ten, gia_von
        self.gia_ban = gia_ban

    @property
    def gia_ban(self):
        return self._gia_ban

    @gia_ban.setter
    def gia_ban(self, gia):
        if gia < self.gia_von:
            raise ValueError(f'{self.ma}: giá bán {gia} thấp hơn giá vốn {self.gia_von}')
        self._gia_ban = gia

    def gia_thuc_te(self):
        return self.gia_ban

    def bien_loi_nhuan(self):
        return (self.gia_thuc_te() - self.gia_von) / self.gia_thuc_te()

    def __str__(self):
        return f'{self.ma} {self.ten}: {self.gia_thuc_te():,.0f}đ · biên LN {self.bien_loi_nhuan():.1%}'


class SanPhamKhuyenMai(SanPham):
    def __init__(self, ma, ten, gia_ban, gia_von, giam):
        super().__init__(ma, ten, gia_ban, gia_von)
        self.giam = giam

    def gia_thuc_te(self):
        return self.gia_ban * (1 - self.giam)


latte = SanPham('SP04', 'Latte', 55000, 18000)
cookie = SanPhamKhuyenMai('SP11', 'Cookie đá xay', 65000, 22000, 0.20)
print(latte)
print(cookie)
try:
    latte.gia_ban = 15000
except ValueError as e:
    print('Lỗi:', e)

# %% [markdown]
# ## Phần C — Pandas: đọc, làm sạch, tích hợp (2,5 điểm)
# ### C1 (0,5 đ) Đọc dữ liệu từ CSV, Excel và SQLite

# %%
gd = pd.read_csv(DATA_DIR / 'giao_dich.csv')
gd['thoi_gian'] = pd.to_datetime(gd['thoi_gian'], format='%d/%m/%Y %H:%M')
kh = pd.read_excel(DATA_DIR / 'khach_hang.xlsx', sheet_name='khach_hang')
hang = pd.read_excel(DATA_DIR / 'khach_hang.xlsx', sheet_name='hang_thanh_vien')
with sqlite3.connect(DATA_DIR / 'san_pham.db') as con:
    sp = pd.read_sql('SELECT * FROM san_pham', con)
    cn = pd.read_sql('SELECT * FROM chi_nhanh', con)

for ten, bang in [('giao_dich', gd), ('khach_hang', kh), ('san_pham', sp), ('chi_nhanh', cn)]:
    print(f'{ten:11}: {bang.shape}')
print('\nGiá trị thiếu trong giao_dich:')
print(gd.isna().sum()[lambda s: s > 0])

# %% [markdown]
# ### C2 (1,0 đ) Làm sạch bảng giao dịch theo đúng 5 bước

# %%
n0 = len(gd)
gd = gd.drop_duplicates()                                                    # 1
n1 = len(gd)

CHUAN_KENH = {'tại quán': 'Tại quán', 'tai quan': 'Tại quán', 'mang đi': 'Mang đi',
              'mang di': 'Mang đi', 'giao hàng': 'Giao hàng', 'giao hang': 'Giao hàng'}
gd['kenh'] = gd['kenh'].str.strip().str.lower().map(CHUAN_KENH)             # 2

gd = gd[~(gd['so_luong'] > 20)]                                              # 3
n3 = len(gd)
gd['so_luong'] = gd['so_luong'].fillna(gd.groupby('kenh')['so_luong'].transform('median')).astype(int)

gd.loc[gd['thoi_gian_cho'] < 0, 'thoi_gian_cho'] = np.nan                    # 4
so_am = gd['thoi_gian_cho'].isna().sum()

q1 = gd.groupby('kenh')['thoi_gian_cho'].transform(lambda s: s.quantile(0.25))   # 5
q3 = gd.groupby('kenh')['thoi_gian_cho'].transform(lambda s: s.quantile(0.75))
tren = q3 + 1.5 * (q3 - q1)
ngoai_lai = gd['thoi_gian_cho'] > tren
print('Ngoại lai theo kênh:', gd.loc[ngoai_lai, 'kenh'].value_counts().to_dict())
q1c, q3c = gd['thoi_gian_cho'].quantile([0.25, 0.75])                        # bẫy: IQR chung cả cột
print('Nếu dùng IQR chung (ngưỡng %.1f phút):' % (q3c + 1.5 * (q3c - q1c)),
      gd.loc[gd['thoi_gian_cho'] > q3c + 1.5 * (q3c - q1c), 'kenh'].value_counts().to_dict())
gd['thoi_gian_cho'] = gd['thoi_gian_cho'].clip(upper=tren)
gd['thoi_gian_cho'] = gd['thoi_gian_cho'].fillna(gd.groupby('kenh')['thoi_gian_cho'].transform('median'))

print('Ban đầu:', n0, '· sau bước 1:', n1, f'(bỏ {n0 - n1})', '· sau bước 3:', n3, f'(bỏ {n1 - n3})')
print('Số đơn theo kênh:', gd['kenh'].value_counts().to_dict())
print('Trung vị so_luong theo kênh:', gd.groupby('kenh')['so_luong'].median().to_dict())
print('Thời gian chờ âm → NaN:', so_am)
print('Còn thiếu:', gd[['so_luong', 'thoi_gian_cho', 'kenh']].isna().sum().to_dict())

# %% [markdown]
# ### C3 (1,0 đ) Tích hợp bốn nguồn

# %%
sp = sp.rename(columns={'product_id': 'ma_sp'})
print('gia_ban_vat là cột dẫn xuất:', np.allclose(sp['gia_ban_vat'], sp['gia_ban'] * 1.08, atol=1))
sp = sp.drop(columns='gia_ban_vat')

print('Mã khách trùng:', kh['ma_kh'].duplicated().sum())
print('Bẫy — merge khi CHƯA khử trùng khách:', len(gd.merge(kh, on='ma_kh', how='left')), 'dòng thay vì', len(gd))
kh = kh.drop_duplicates(subset='ma_kh')
kh['gioi_tinh'] = kh['gioi_tinh'].str.strip().str.lower().map(
    {'nam': 'Nam', 'm': 'Nam', 'nữ': 'Nữ', 'nu': 'Nữ', 'f': 'Nữ'})
print('Giới tính:', kh['gioi_tinh'].value_counts().to_dict())
kh = kh.merge(hang, on='hang_thanh_vien', how='left')

khong_co_sp = ~gd['ma_sp'].isin(sp['ma_sp'])
print('Giao dịch có ma_sp ngoài danh mục:', khong_co_sp.sum())
df = (gd[~khong_co_sp]
      .merge(sp, on='ma_sp', how='left', validate='many_to_one')
      .merge(kh[['ma_kh', 'gioi_tinh', 'hang_thanh_vien', 'chiet_khau']], on='ma_kh', how='left',
             validate='many_to_one')
      .merge(cn, on='ma_cn', how='left'))
df['hang_thanh_vien'] = df['hang_thanh_vien'].fillna('Vãng lai')
df['chiet_khau'] = df['chiet_khau'].fillna(0)
df['doanh_thu'] = df['so_luong'] * df['gia_ban'] * (1 - df['chiet_khau'])
df['loi_nhuan'] = df['doanh_thu'] - df['so_luong'] * df['gia_von']

print('\nSố dòng df:', len(df), '· vãng lai:', (df['hang_thanh_vien'] == 'Vãng lai').sum())
print(f'Tổng doanh thu: {df["doanh_thu"].sum():,.0f} · lợi nhuận: {df["loi_nhuan"].sum():,.0f}')
print((df.groupby('ten_cn')['doanh_thu'].sum() / 1e6).round(1).sort_values(ascending=False))

# %% [markdown]
# ## Phần D — Biến đổi & thu gọn dữ liệu (1,5 điểm)
# ### D1 (0,75 đ) Biến đổi

# %%
THU_TU = {'Rất tệ': 1, 'Tệ': 2, 'Bình thường': 3, 'Tốt': 4, 'Rất tốt': 5}
df['diem_dg'] = df['danh_gia'].map(THU_TU)
mode_dg = df['diem_dg'].mode()[0]
df['diem_dg'] = df['diem_dg'].fillna(mode_dg).astype(int)
print('Mode:', mode_dg, '· phân bố:', df['diem_dg'].value_counts().sort_index().to_dict())

import scipy.stats as st
df['log_cho'] = np.log1p(df['thoi_gian_cho'])
print('Skew thoi_gian_cho:', round(st.skew(df['thoi_gian_cho']), 3), '→ sau log1p:', round(st.skew(df['log_cho']), 3))

dummy = pd.get_dummies(df['kenh'], prefix='kenh', dtype=int)
print('One-hot:', dummy.columns.tolist(), dummy.sum().to_dict())

# %% [markdown]
# ### D2 (0,75 đ) Tổng hợp theo khách hàng & giảm chiều

# %%
thanh_vien = df[df['hang_thanh_vien'] != 'Vãng lai']
khach = thanh_vien.groupby('ma_kh').agg(
    so_gd=('ma_gd', 'count'),
    tong_chi=('doanh_thu', 'sum'),
    sl_tb=('so_luong', 'mean'),
    cho_tb=('thoi_gian_cho', 'mean'),
    diem_tb=('diem_dg', 'mean'),
    ti_le_giao=('kenh', lambda s: (s == 'Giao hàng').mean()),
)
khach['tong_chi'] = khach['tong_chi'] / 1e6
print('Bảng khách:', khach.shape)
khach['nhom_chi'] = pd.qcut(khach['tong_chi'], q=3, labels=['Thấp', 'Vừa', 'Cao'])
print('Nhóm chi tiêu:', khach['nhom_chi'].value_counts(sort=False).to_dict())

so = khach.drop(columns='nhom_chi')
tq = so.corr()
print(tq.round(2))
cap = [(a, b, round(tq.loc[a, b], 2)) for i, a in enumerate(tq.columns) for b in tq.columns[i + 1:] if abs(tq.loc[a, b]) > 0.7]
print('Cặp |r| > 0.7:', cap)

Z = StandardScaler().fit_transform(so)
pca = PCA(n_components=0.9).fit(Z)
print('PCA giữ >= 90%:', pca.n_components_, 'thành phần ·', PCA().fit(Z).explained_variance_ratio_.cumsum().round(3))

# %% [markdown]
# ## Phần E — Trực quan hoá (2,0 điểm)

# %%
fig, axes = plt.subplots(2, 2, figsize=(14, 10))

theo_thang = (df.groupby([df['thoi_gian'].dt.month, 'kenh'])['doanh_thu'].sum() / 1e6).unstack()
theo_thang.plot(ax=axes[0, 0], marker='o')
axes[0, 0].set(title='Doanh thu theo tháng và kênh', xlabel='Tháng', ylabel='Doanh thu (triệu đồng)')

sns.histplot(data=df, x='log_cho', hue='kenh', bins=30, ax=axes[0, 1])
axes[0, 1].set(title='Phân phối thời gian chờ sau log1p', xlabel='log(1 + thời gian chờ, phút)', ylabel='Số giao dịch')

pv = df.pivot_table(index='ten_cn', columns='nhom', values='doanh_thu', aggfunc='sum') / 1e6
pv.plot(kind='bar', stacked=True, ax=axes[1, 0], rot=0)
axes[1, 0].set(title='Doanh thu theo chi nhánh và nhóm sản phẩm', xlabel='Chi nhánh', ylabel='Doanh thu (triệu đồng)')

sns.heatmap(tq, annot=True, fmt='.2f', cmap='coolwarm', vmin=-1, vmax=1, ax=axes[1, 1])
axes[1, 1].set_title('Tương quan các đặc trưng khách hàng')

fig.tight_layout()
fig.savefig('bieu_do.png', dpi=150)
plt.show()
print(theo_thang.round(1))
print(pv.round(1))
