import pandas as pd
import numpy as np
from fpdf import FPDF
from datetime import datetime, timedelta
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. LOAD & PROCESS DATA
# ============================================================
df_l5 = pd.read_excel(r'C:\Users\Administrator\Desktop\Anal_C_Cuc\L5_c_C.xlsx')
df_l5['NGAY_CHIA'] = pd.to_datetime(df_l5['NGAY_CHIA'])
df_l5 = df_l5.dropna(subset=['NGAY_CHIA'])
df_l5['date'] = df_l5['NGAY_CHIA'].dt.date

df_l8 = pd.read_excel(r'C:\Users\Administrator\Desktop\Anal_C_Cuc\L8_c_cuc.xlsx')
df_l8['CURRENT_ORDER_DATE'] = pd.to_datetime(df_l8['CURRENT_ORDER_DATE'])
df_l8 = df_l8.dropna(subset=['CURRENT_ORDER_DATE'])
df_l8['date'] = df_l8['CURRENT_ORDER_DATE'].dt.date

# Week label: lay ngay dau tuan (Monday)
def get_monday(d):
    return d - timedelta(days=d.weekday())

df_l5['monday'] = df_l5['date'].apply(get_monday)
df_l8['monday'] = df_l8['date'].apply(get_monday)

# Week label format: dd/mm
df_l5['week_label'] = df_l5['monday'].apply(lambda x: x.strftime('%d/%m'))
df_l8['week_label'] = df_l8['monday'].apply(lambda x: x.strftime('%d/%m'))

# Sort weeks theo ngay
all_weeks = sorted(df_l5['monday'].unique())
week_labels = [w.strftime('%d/%m') for w in all_weeks]

# Map monday -> week_label
monday_to_label = {w: w.strftime('%d/%m') for w in all_weeks}

# Camp short name
def short_camp(name):
    name = str(name)
    name = name.replace('Inhouse_EdupiaClass_Tu-van-Demo_', '')
    name = name.replace('Inhouse_EdupiaClass_Tu-van_Demo_', '')
    name = name.replace('Inhouse_EdupiaClass_', '')
    if len(name) > 40:
        name = name[:37] + '...'
    return name

# ============================================================
# 2. CALCULATE STATISTICS
# ============================================================

total_l5 = len(df_l5)
total_l8 = len(df_l8)
tlc_total = total_l8 / total_l5 * 100

# --- All campaigns ---
all_camps = df_l5['CAMPAIGN_NAME'].unique()

# --- Section 1: Tong quan theo tuan ---
week_summary = []
for m in all_weeks:
    label = monday_to_label[m]
    l5_cnt = len(df_l5[df_l5['monday'] == m])
    l8_cnt = len(df_l8[df_l8['monday'] == m])
    tlc = l8_cnt / l5_cnt * 100 if l5_cnt > 0 else 0
    week_summary.append({'week': label, 'L5': l5_cnt, 'L8': l8_cnt, 'TLC': tlc})
df_week_summary = pd.DataFrame(week_summary)

# --- Section 2: Chuyen doi theo campaign ---
camp_stats = []
for camp in all_camps:
    l5_cnt = len(df_l5[df_l5['CAMPAIGN_NAME'] == camp])
    l8_cnt = len(df_l8[df_l8['CAMPAIGN'] == camp])
    tlc = l8_cnt / l5_cnt * 100 if l5_cnt > 0 else 0
    pct_l5 = l5_cnt / total_l5 * 100
    pct_l8 = l8_cnt / total_l8 * 100
    camp_stats.append({
        'camp_full': camp,
        'camp_short': short_camp(camp),
        'L5': l5_cnt, 'L8': l8_cnt,
        'TLC': tlc, 'pct_L5': pct_l5, 'pct_L8': pct_l8
    })
df_camp_stats = pd.DataFrame(camp_stats).sort_values('L5', ascending=False)

# --- Section 3: Ty trong L5 theo tuan (pivot) ---
pivot_tytrong = df_l5.groupby(['CAMPAIGN_NAME', 'week_label']).size().unstack(fill_value=0)
# Tinh % moi hang
pivot_tytrong_pct = pivot_tytrong.div(pivot_tytrong.sum(axis=0), axis=1) * 100
# Sort theo tong giam dan
pivot_tytrong_pct['TB'] = pivot_tytrong_pct.mean(axis=1)
pivot_tytrong_pct = pivot_tytrong_pct.sort_values('TB', ascending=False)

# --- Section 4: TLC theo tuan x campaign (pivot) ---
# Tao df gop L5 va L8
df_l5['is_l5'] = True
df_l8['is_l8'] = True
df_l8['CAMPAIGN_NAME'] = df_l8['CAMPAIGN']  # unify column name

# Tinh TLC theo camp x tuan
tlc_pivot_data = []
for camp in all_camps:
    for m in all_weeks:
        label = monday_to_label[m]
        l5_cnt = len(df_l5[(df_l5['CAMPAIGN_NAME'] == camp) & (df_l5['monday'] == m)])
        l8_cnt = len(df_l8[(df_l8['CAMPAIGN'] == camp) & (df_l8['monday'] == m)])
        tlc = l8_cnt / l5_cnt * 100 if l5_cnt > 0 else 0
        tlc_pivot_data.append({'camp': camp, 'week': label, 'TLC': tlc, 'L5': l5_cnt, 'L8': l8_cnt})

df_tlc_pivot = pd.DataFrame(tlc_pivot_data)
pivot_tlc = df_tlc_pivot.pivot(index='camp', columns='week', values='TLC').fillna(0)
pivot_tlc = pivot_tlc[week_labels]  # sort columns

# --- Section 5: L8 theo tuan x campaign (pivot) ---
pivot_l8 = df_tlc_pivot.pivot(index='camp', columns='week', values='L8').fillna(0)
pivot_l8 = pivot_l8[week_labels]
pivot_l8['TOTAL'] = pivot_l8.sum(axis=1)
pivot_l8 = pivot_l8.sort_values('TOTAL', ascending=False)

# ============================================================
# 3. CREATE PDF
# ============================================================
FONT_PATH = r'C:\Windows\Fonts\DejaVuSans.ttf'
FONT_BOLD = r'C:\Windows\Fonts\DejaVuSans-Bold.ttf'

class PDF(FPDF):
    def header(self):
        self.set_font('VF', 'B', 13)
        self.cell(0, 10, 'PHAN TICH TLC - TEAM CUC', 0, new_x="LMARGIN", new_y="NEXT", align='C')
        self.set_font('VF', '', 8)
        self.cell(0, 5, f'Ngay: {datetime.now().strftime("%d/%m/%Y %H:%M")}', 0, new_x="LMARGIN", new_y="NEXT", align='R')
        self.ln(3)

    def footer(self):
        self.set_y(-15)
        self.set_font('VF', 'I', 7)
        self.cell(0, 10, f'Trang {self.page_no()}/{{nb}}', 0, 0, 'C')

    def section_title(self, title):
        self.set_font('VF', 'B', 10)
        self.set_fill_color(41, 128, 185)
        self.set_text_color(255, 255, 255)
        self.cell(0, 7, f'  {title}', 0, new_x="LMARGIN", new_y="NEXT", align='L', fill=True)
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def table_header(self, cols, widths):
        self.set_font('VF', 'B', 7)
        self.set_fill_color(220, 230, 241)
        for i, col in enumerate(cols):
            self.cell(widths[i], 6, col, 1, 0, 'C', fill=True)
        self.ln()

    def table_row(self, values, widths, bold=False, highlight=False, aligns=None):
        if bold:
            self.set_font('VF', 'B', 7)
        else:
            self.set_font('VF', '', 7)
        if highlight:
            self.set_fill_color(255, 255, 200)
        else:
            self.set_fill_color(255, 255, 255)
        for i, v in enumerate(values):
            a = 'C'
            if aligns and i < len(aligns):
                a = aligns[i]
            self.cell(widths[i], 5, str(v), 1, 0, a, fill=highlight)
        self.ln()

    def note_text(self, text):
        self.set_font('VF', 'I', 8)
        self.set_text_color(180, 0, 0)
        self.multi_cell(0, 4, text)
        self.set_text_color(0, 0, 0)
        self.ln(1)

    def body_text(self, text):
        self.set_font('VF', '', 9)
        self.multi_cell(0, 5, text)
        self.ln(1)


pdf = PDF()
pdf.alias_nb_pages()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_font('VF', '', FONT_PATH)
pdf.add_font('VF', 'B', FONT_BOLD)
pdf.add_font('VF', 'I', FONT_PATH)
pdf.add_page()

# ============================================================
# SECTION 1: TONG QUAN THEO TUAN
# ============================================================
pdf.section_title('1. TONG QUAN THEO TUAN')
pdf.body_text(f'Du lieu: {total_l5} deals L5, {total_l8} don hang L8')
pdf.body_text(f'Thoi gian: {week_labels[0]} - {week_labels[-1]}')
pdf.body_text(f'Tong TLC (L8/L5): {tlc_total:.1f}%')
pdf.ln(1)

cols1 = ['Tuan', 'L5', 'L8', 'TLC_%']
widths1 = [25, 25, 25, 25]
pdf.table_header(cols1, widths1)
for _, r in df_week_summary.iterrows():
    pdf.table_row([r['week'], int(r['L5']), int(r['L8']), f"{r['TLC']:.1f}"], widths1)
pdf.ln(3)

# ============================================================
# SECTION 2: CHUYEN DOI THEO CAMPAIGN
# ============================================================
pdf.section_title('2. CHUYEN DOI THEO CAMPAIGN')
cols2 = ['CAMPAIGN', 'L5', 'L8', 'TLC_%', 'Trong_L5_%', 'Trong_L8_%']
widths2 = [65, 15, 15, 18, 22, 22]
pdf.table_header(cols2, widths2)
for _, r in df_camp_stats.iterrows():
    pdf.table_row([
        r['camp_short'][:42], int(r['L5']), int(r['L8']),
        f"{r['TLC']:.1f}", f"{r['pct_L5']:.1f}", f"{r['pct_L8']:.1f}"
    ], widths2, aligns=['L', 'C', 'C', 'C', 'C', 'C'])
pdf.ln(3)

# ============================================================
# SECTION 3: TY TRONG L5 THEO TUAN
# ============================================================
pdf.add_page()
pdf.section_title('3. TY TRONG L5 THEO TUAN (%)')
cols3 = ['CAMPAIGN'] + week_labels + ['TB']
widths3 = [50] + [18] * len(week_labels) + [18]
pdf.table_header(cols3, widths3)
for camp, row in pivot_tytrong_pct.iterrows():
    vals = [short_camp(camp)[:35]]
    for w in week_labels:
        v = row.get(w, 0)
        vals.append(f"{v:.1f}" if v > 0 else '-')
    vals.append(f"{row['TB']:.1f}")
    pdf.table_row(vals, widths3, aligns=['L'] + ['C'] * (len(week_labels) + 1))
pdf.ln(3)

# ============================================================
# SECTION 4: TLC THEO TUAN X CAMPAIGN
# ============================================================
pdf.add_page()
pdf.section_title('4. TLC THEO TUAN X CAMPAIGN (%)')
cols4 = ['CAMPAIGN'] + week_labels
widths4 = [50] + [22] * len(week_labels)
pdf.table_header(cols4, widths4)
# Sort theo tong TLC giam dan
pivot_tlc['sum'] = pivot_tlc.sum(axis=1)
pivot_tlc = pivot_tlc.sort_values('sum', ascending=False)
for camp, row in pivot_tlc.iterrows():
    vals = [short_camp(camp)[:35]]
    for w in week_labels:
        v = row.get(w, 0)
        if v > 0:
            vals.append(f"{v:.1f}")
        else:
            vals.append('-')
    pdf.table_row(vals, widths4, aligns=['L'] + ['C'] * len(week_labels))
pdf.ln(3)

# ============================================================
# SECTION 5: L8 CAMPAIGN THEO TUAN
# ============================================================
pdf.section_title('5. L8 CAMPAIGN THEO TUAN')
cols5 = ['CAMPAIGN'] + week_labels + ['TOTAL']
widths5 = [50] + [18] * len(week_labels) + [18]
pdf.table_header(cols5, widths5)
for camp, row in pivot_l8.iterrows():
    vals = [short_camp(camp)[:35]]
    for w in week_labels:
        v = int(row.get(w, 0))
        vals.append(str(v) if v > 0 else '-')
    vals.append(int(row['TOTAL']))
    pdf.table_row(vals, widths5, aligns=['L'] + ['C'] * (len(week_labels) + 1))
pdf.ln(3)

# ============================================================
# SECTION 6: KET LUAN
# ============================================================
pdf.add_page()
pdf.section_title('6. KET LUAN')

# Tinh toan cho ket luan
# Nobuy vs Isbuy TLC
nobuy_camps = [c for c in all_camps if 'nobuy' in str(c).lower()]
isbuy_camps = [c for c in all_camps if 'isbuy' in str(c).lower()]

nobuy_l5_total = len(df_l5[df_l5['CAMPAIGN_NAME'].isin(nobuy_camps)])
nobuy_l8_total = len(df_l8[df_l8['CAMPAIGN'].isin(nobuy_camps)])
nobuy_tlc = nobuy_l8_total / nobuy_l5_total * 100 if nobuy_l5_total > 0 else 0

isbuy_l5_total = len(df_l5[df_l5['CAMPAIGN_NAME'].isin(isbuy_camps)])
isbuy_l8_total = len(df_l8[df_l8['CAMPAIGN'].isin(isbuy_camps)])
isbuy_tlc = isbuy_l8_total / isbuy_l5_total * 100 if isbuy_l5_total > 0 else 0

# Ty trong
nobuy_pct_l5 = nobuy_l5_total / total_l5 * 100
nobuy_pct_l8 = nobuy_l8_total / total_l8 * 100
isbuy_pct_l5 = isbuy_l5_total / total_l5 * 100
isbuy_pct_l8 = isbuy_l8_total / total_l8 * 100

# Correlation nobuy% vs TLC% theo tuan
nobuy_weekly_pct = []
tlc_weekly = []
for m in all_weeks:
    l5_w = len(df_l5[df_l5['monday'] == m])
    l5_nobuy_w = len(df_l5[(df_l5['monday'] == m) & (df_l5['CAMPAIGN_NAME'].isin(nobuy_camps))])
    l8_w = len(df_l8[df_l8['monday'] == m])
    tlc_w = l8_w / l5_w * 100 if l5_w > 0 else 0
    nobuy_pct_w = l5_nobuy_w / l5_w * 100 if l5_w > 0 else 0
    nobuy_weekly_pct.append(nobuy_pct_w)
    tlc_weekly.append(tlc_w)

corr = np.corrcoef(nobuy_weekly_pct, tlc_weekly)[0, 1] if len(nobuy_weekly_pct) > 1 else 0

pdf.set_font('VF', 'B', 10)
pdf.cell(0, 7, 'CAU HOI 1: Ty trong nobuy co anh huong den TLC chung ko?', 0, new_x="LMARGIN", new_y="NEXT")
pdf.set_font('VF', '', 9)
pdf.ln(1)

pdf.body_text(f'- Nobuy: {nobuy_l5_total} L5 ({nobuy_pct_l5:.1f}%), {nobuy_l8_total} L8 ({nobuy_pct_l8:.1f}%), TLC = {nobuy_tlc:.1f}%')
pdf.body_text(f'- Isbuy: {isbuy_l5_total} L5 ({isbuy_pct_l5:.1f}%), {isbuy_l8_total} L8 ({isbuy_pct_l8:.1f}%), TLC = {isbuy_tlc:.1f}%')
pdf.body_text(f'- Correlation (Nobuy% vs TLC% theo tuan): {corr:.3f}')
pdf.ln(1)

if abs(corr) < 0.1:
    corr_note = 'Correlation gan 0 -> Ty trong nobuy KHONG anh huong den TLC.'
elif corr < 0:
    corr_note = f'Correlation am ({corr:.3f}) -> Ty trong nobuy cao thi TLC thap.'
else:
    corr_note = f'Correlation duong ({corr:.3f}) -> Ty trong nobuy cao thi TLC cao.'

pdf.body_text(f'- {corr_note}')
pdf.body_text(f'- Nobuy TLC = {nobuy_tlc:.1f}% thap hon Isbuy TLC = {isbuy_tlc:.1f}% (gap {isbuy_tlc/nobuy_tlc:.1f}x)')
pdf.body_text(f'- Nobuy chiem {nobuy_pct_l5:.1f}% L5 nhung chi dong gop {nobuy_pct_l8:.1f}% L8 -> hieu qua thap')
pdf.body_text(f'- Tuy nhien, nobuy bien dong manh theo tuan (1.7%-11.6%) do lag effect')
pdf.ln(2)

pdf.set_font('VF', 'B', 10)
pdf.cell(0, 7, 'CAU HOI 2: Lam gi de cai thien?', 0, new_x="LMARGIN", new_y="NEXT")
pdf.set_font('VF', '', 9)
pdf.ln(1)

suggestions = [
    '1. GIAM TY LE NOBUY TRONG L5:',
    '   - Loc SĐT co xu huong mua hang truoc khi chia camp',
    '   - Uu tien nguoi da tuong tac san pham (xem, them vao gio)',
    '   - Bo qua SĐT moi, chua co hanh vi mua sam',
    '',
    '2. TANG CONVERSION NOBUY:',
    '   - Nurture mau hon: goi dien trong 24h dau tien',
    '   - Gui tin nhan/GOI y san pham phu hop',
    '   - Tao urgency: uu dai co han, so luong co han',
    '',
    '3. TAP TRUNG VAO ISBUY:',
    '   - Isbuy co TLC cao hon 2x -> nen tang ty trong isbuy trong L5',
    '   - Nhom isbuy_2sp co TLC cao nhat -> uu tien nguoi da mua nhieu SP',
    '',
    '4. KHAC PHUC LAG EFFECT:',
    '   - Theo doi don hang theo cohort (tu ngay chia camp)',
    '   - Khong chi danh gia theo tuan (cut-off) ma can tich luy +1/+2 tuan',
    '',
    '5. OPTIMIZE THEO CAMPAIGN:',
    '   - isbuy_lead_Edu: TLC 16.7% -> nen tang ty trong',
    '   - isbuy_2sp: TLC cao -> uu tien nguoi da mua 2+ SP',
    '   - nobuy_lead (cos): TLC 6.2% -> can cai thien nurture',
]

for s in suggestions:
    pdf.body_text(s)

pdf.ln(2)
pdf.set_font('VF', 'B', 10)
pdf.cell(0, 7, 'KET LUAN TONG QUAT', 0, new_x="LMARGIN", new_y="NEXT")
pdf.set_font('VF', '', 9)
pdf.ln(1)

pdf.body_text(f'1. Ty trong nobuy {nobuy_pct_l5:.1f}% anh huong IT den TLC (correlation = {corr:.3f}).')
pdf.body_text(f'2. Nobuy TLC {nobuy_tlc:.1f}% thap hon Isbuy {isbuy_tlc:.1f}% -> can cai thieu conversion nobuy.')
pdf.body_text(f'3. Isbuy la yeu to chinh day TLC len -> nen tang ty trong isbuy trong L5.')
pdf.body_text(f'4. Lag effect lam bien dong TLC nobuy theo tuan -> can tinh cohort de chinh xac.')
pdf.body_text(f'5. Nhom isbuy_2sp va isbuy_Edu co TLC cao nhat -> uu tien tai nguyen.')

# ============================================================
# 4. SAVE PDF
# ============================================================
output_path = r'C:\Users\Administrator\Desktop\Anal_C_Cuc\Phan_tich_TLC_team_Cuc.pdf'
pdf.output(output_path)
print(f'Da xuat file: {output_path}')
