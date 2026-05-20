import streamlit as st
import pandas as pd
import numpy as np
import io, os, glob
from datetime import timedelta

st.set_page_config(page_title="NXA-X12M-1 시뮬레이션", layout="wide")
st.title("NXA-X12M-1 광미래 납품 시뮬레이션")
st.caption("SAP 수불 파일 → 라인별 생산 / 납품 / 가용재고 자동 계산")

LEAD_TIME  = 4
PROD_LOC   = '3000'
SHIP_LOC   = '5700'
LOT_LN_IDX = 6
MAX_LOT_KG = 3000

LINE_MAP = {
    '03C':'G3C','04C':'G4C','05C':'G5C',
    '06C':'G6C','08C':'G8C',
    'C01':'C1','C02':'C2','C03':'C3',
}
LINE_ORDER = ['G3C','G4C','G5C','G6C','G8C','C1','C2','C3']

def fmt(v):
    """숫자 표시: 0이면 빈칸, 소수점 불필요하면 정수로"""
    if isinstance(v, (int, float)):
        if v == 0: return ''
        return int(v) if v == int(v) else round(v, 2)
    return v

with st.sidebar:
    st.header("⚙️ 설정")
    lead_time = st.number_input("품질합격 반영 리드타임 (D+N)", 1, 10, LEAD_TIME)
    st.divider()
    st.header("📂 파일 선택")
    base_dir  = os.path.dirname(os.path.abspath(__file__))
    file_list = glob.glob(os.path.join(base_dir,'*.xlsx')) + \
                glob.glob(os.path.join(base_dir,'*.csv'))
    # 엑셀 열려있을 때 생기는 임시파일 (~$로 시작) 제외
    fnames    = [os.path.basename(f) for f in file_list
                 if not os.path.basename(f).startswith('~$')]
    if not fnames:
        st.warning("같은 폴더에 xlsx/csv 없음")
        selected_file = None
    else:
        sel = st.selectbox("수불 파일", fnames)
        selected_file = os.path.join(base_dir, sel)

    # 품질판정 파일 (선택사항)
    st.divider()
    st.header("📋 품질판정 파일 (선택)")
    qc_fnames = [f for f in fnames if f != sel] if fnames else []
    use_qc = st.checkbox("품질판정 파일 사용", value=False)
    qc_file = None
    if use_qc and qc_fnames:
        qc_sel  = st.selectbox("품질판정 파일", qc_fnames)
        qc_file = os.path.join(base_dir, qc_sel)

@st.cache_data
def process(filepath, lead_time, qc_filepath=None):
    # ── 파일 읽기
    def read_file(path):
        if path.endswith('.csv'):
            for enc in ['utf-8-sig','cp949','utf-8']:
                try: return pd.read_csv(path, dtype=str, encoding=enc)
                except: continue
        return pd.read_excel(path, dtype=str)

    df = read_file(filepath)
    df.columns = df.columns.str.strip()

    def find(kws, cols):
        for kw in kws:
            for c in cols:
                if kw in c: return c
        return None

    loc_col  = find(['저장 위치','저장위치'], df.columns)
    lot_col  = find(['LOT NO','LOT번호','배치'], df.columns)
    date_col = find(['입고일'], df.columns)
    in_col   = find(['입고소계'], df.columns)

    miss = [n for n,c in [('저장위치',loc_col),('LOT NO',lot_col),
                           ('입고일',date_col),('입고소계',in_col)] if c is None]
    if miss:
        return None, f"컬럼 없음: {miss} / 실제컬럼: {list(df.columns)}"

    df = df.rename(columns={loc_col:'loc',lot_col:'lot',
                             date_col:'date',in_col:'qty_in'})
    df['qty_in'] = pd.to_numeric(
        df['qty_in'].astype(str).str.replace(',','').str.strip(), errors='coerce'
    ).fillna(0)
    df['date'] = pd.to_datetime(
        df['date'].astype(str).str.extract(r'(\d{4}-\d{2}-\d{2})')[0], errors='coerce'
    )
    df = df.dropna(subset=['date'])
    df['loc'] = df['loc'].astype(str).str.strip()

    def get_line(lot):
        if not isinstance(lot, str): return 'ETC'
        p = lot.strip().split('-')
        return LINE_MAP.get(p[LOT_LN_IDX].upper(), p[LOT_LN_IDX].upper()) \
               if len(p) > LOT_LN_IDX else 'ETC'
    df['line'] = df['lot'].apply(get_line)

    # ── 생산: 3000창고 LOT별 최초 1회, 3000KG 캡
    prod_raw   = df[df['loc'] == PROD_LOC].copy()
    prod_raw['qty_in'] = prod_raw['qty_in'].clip(upper=MAX_LOT_KG)
    prod_dedup = (prod_raw.sort_values('date')
                  .groupby('lot', as_index=False).first())

    # ── 납품: 5700창고
    ship_raw = df[df['loc'] == SHIP_LOC].copy()

    # 날짜 범위
    all_d = pd.concat([prod_dedup['date'], ship_raw['date']]).dropna()
    if len(all_d) == 0:
        return None, "유효한 날짜 없음"
    dr = pd.date_range(all_d.min(), all_d.max())

    # KG→ton 자동판단
    med = prod_dedup['qty_in'][prod_dedup['qty_in']>0].median() if len(prod_dedup)>0 else 0
    div = 1000 if med > 500 else 1

    # 라인별 일별 생산 (ton)
    prod_line_pivot = (
        prod_dedup.groupby(['date','line'])['qty_in'].sum() / div
    ).reset_index().pivot_table(
        index='date', columns='line', values='qty_in', aggfunc='sum'
    ).reindex(dr).fillna(0)

    # 일별 생산 소계
    prod_total = (prod_dedup.groupby('date')['qty_in'].sum() / div).reindex(dr).fillna(0)

    # 일별 납품
    ship_total = (ship_raw.groupby('date')['qty_in'].sum() / div).reindex(dr).fillna(0)

    # ── 품질합격 반영
    # 품질판정 파일 있으면 그 날짜 기준, 없으면 생산일 + lead_time
    if qc_filepath:
        try:
            qc_df = read_file(qc_filepath)
            qc_df.columns = qc_df.columns.str.strip()
            # 판정일 / LOT / 합격여부 컬럼 자동탐지
            qc_date_col   = find(['판정일','합격일','검사일','입고일'], qc_df.columns)
            qc_lot_col    = find(['LOT NO','LOT번호','배치'], qc_df.columns)
            qc_result_col = find(['품질판정','판정','합격'], qc_df.columns)
            if qc_date_col and qc_lot_col:
                qc_df = qc_df.rename(columns={qc_date_col:'qc_date', qc_lot_col:'lot'})
                if qc_result_col:
                    qc_df = qc_df.rename(columns={qc_result_col:'qc_result'})
                    qc_df = qc_df[qc_df['qc_result'].astype(str).str.upper().isin(['Y','합격','OK','PASS'])]
                qc_df['qc_date'] = pd.to_datetime(
                    qc_df['qc_date'].astype(str).str.extract(r'(\d{4}-\d{2}-\d{2})')[0], errors='coerce'
                )
                # 각 LOT의 합격일 → 생산수량 매핑
                lot_qty = prod_dedup.set_index('lot')['qty_in']
                qc_merged = qc_df.dropna(subset=['qc_date']).copy()
                qc_merged['qty'] = qc_merged['lot'].map(lot_qty).fillna(0) / div
                approved_total = qc_merged.groupby('qc_date')['qty'].sum().reindex(dr).fillna(0)
            else:
                approved_total = None
        except:
            approved_total = None
    else:
        approved_total = None

    # 품질판정 데이터 없으면 생산일 + lead_time 으로 대체
    if approved_total is None:
        approved_total = prod_total.copy()
        approved_total.index = approved_total.index + pd.Timedelta(days=lead_time)
        approved_total = approved_total.reindex(dr).fillna(0)

    # ── 가용재고 = 합격재고 누계 - 납품 누계
    approved_cum = approved_total.cumsum()
    ship_cum     = ship_total.cumsum()
    avail_stock  = (approved_cum - ship_cum).round(1)

    # ── 분석대기 = 생산 누계 - 합격 누계 (아직 판정 안 된 물량)
    prod_cum     = prod_total.cumsum()
    pending      = (prod_cum - approved_cum).clip(lower=0).round(1)

    r = {
        'dr': dr, 'div': div,
        'prod_line':   prod_line_pivot,
        'prod_total':  prod_total,
        'approved':    approved_total,
        'ship_total':  ship_total,
        'avail_stock': avail_stock,
        'pending':     pending,
        'total_prod':  round(float(prod_total.sum()),1),
        'total_ship':  round(float(ship_total.sum()),1),
        'cur_stock':   round(float(avail_stock.iloc[-1]),1),
        'neg_days':    int((avail_stock < 0).sum()),
        'prod_rows':   len(prod_dedup),
        'ship_rows':   len(ship_raw),
        'lead_time':   lead_time,
        'has_qc':      qc_filepath is not None and approved_total is not None,
    }
    return r, None

# ── 메인
if not selected_file:
    st.info("← 왼쪽에서 파일을 선택하세요")
    st.stop()

with st.spinner("처리 중..."):
    res, err = process(selected_file, lead_time, qc_file if use_qc else None)

if err:
    st.error(err); st.stop()

dr           = res['dr']
prod_line    = res['prod_line']
prod_total   = res['prod_total']
approved     = res['approved']
ship_total   = res['ship_total']
avail_stock  = res['avail_stock']
pending      = res['pending']

# ── 요약 카드
c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("기간", f"{dr[0].strftime('%m/%d')} ~ {dr[-1].strftime('%m/%d')}")
c2.metric("총 생산", f"{res['total_prod']:,.1f} t")
c3.metric("총 납품", f"{res['total_ship']:,.1f} t")
c4.metric("현재 가용재고", f"{res['cur_stock']:,.1f} t")
c5.metric("재고 부족일", f"{res['neg_days']} 일")
st.caption(
    f"단위: {'KG→ton(÷1000)' if res['div']==1000 else 'ton'} | "
    f"생산 LOT: {res['prod_rows']}건 | 출하: {res['ship_rows']}건 | "
    f"합격반영: {'QC파일' if res['has_qc'] else 'D+' + str(res['lead_time']) + ' 자동반영'}"
)

st.divider()

# ── 날짜 필터
st.subheader("일별 시뮬레이션")
cf1, cf2 = st.columns(2)
with cf1: s_date = st.date_input("시작일", dr[0].date())
with cf2: e_date = st.date_input("종료일", min(dr[-1].date(),(dr[0]+timedelta(days=59)).date()))

mask = (dr >= pd.Timestamp(s_date)) & (dr <= pd.Timestamp(e_date))
fdr  = dr[mask]
fidx = np.where(mask)[0]

if len(fdr) == 0:
    st.warning("선택 기간에 데이터 없음"); st.stop()

days_kr = ['월','화','수','목','금','토','일']

# ── 테이블 구성
exist_lines = [l for l in LINE_ORDER if l in prod_line.columns]
other_lines = [l for l in prod_line.columns if l not in LINE_ORDER and l != 'ETC']
all_lines   = exist_lines + other_lines

def make_row(구분, 항목, series, fdr):
    row = {'구분': 구분, '항목': 항목}
    for d in fdr:
        v = float(series.loc[d]) if d in series.index else 0
        key = f"{d.month}/{d.day}"
        row[key] = fmt(v)
    return row

rows = []

# 라인별 생산
for ln in all_lines:
    row = {'구분':'생산', '항목': ln}
    for d in fdr:
        v = float(prod_line.loc[d, ln]) if d in prod_line.index and ln in prod_line.columns else 0
        row[f"{d.month}/{d.day}"] = fmt(v)
    rows.append(row)

# 생산 소계
rows.append(make_row('생산', '소계', prod_total, fdr))

# 합격재고 반영 (분석 완료)
rows.append(make_row('합격반영', f'D+{lead_time}', approved, fdr))

# 납품
rows.append(make_row('납품', '남경', ship_total, fdr))

# 가용재고
rows.append(make_row('가용재고', '', avail_stock, fdr))

# 분석대기
rows.append(make_row('분석대기', '(미판정)', pending, fdr))

df_tbl = pd.DataFrame(rows)
num_cols = [c for c in df_tbl.columns if c not in ['구분','항목']]

# ── 스타일: 검은글씨 + 가운데정렬 기본, 연회색/진회색 구분
BASE  = 'color:#111111; text-align:center;'
LIGHT = BASE + 'background-color:#eeeeee;'          # 연한회색 - 일반행
DARK  = BASE + 'background-color:#999999; font-weight:bold;'  # 진한회색 - 가용재고
RED   = BASE + 'background-color:#f5c6cb; font-weight:bold;'  # 음수

def style_fn(df):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for i, r in df.iterrows():
        구분 = r.get('구분','')
        styles.loc[i] = DARK if 구분 == '가용재고' else LIGHT
    for c in num_cols:
        for i, v in df[c].items():
            if isinstance(v, (int, float)) and v < 0:
                styles.loc[i, c] = RED
    return styles

st.dataframe(
    df_tbl.style.apply(style_fn, axis=None),
    use_container_width=True,
    height=min(80 + len(rows)*38, 650),
    hide_index=True,
    column_config={
        '구분': st.column_config.TextColumn(width=80),
        '항목': st.column_config.TextColumn(width=70),
    }
)

st.caption("※ 합격반영: QC파일 없을 시 생산일 기준 D+N 자동반영 | 분석대기: 생산 누계 - 합격 누계")

# ── 재고 추이 차트
st.divider()
st.subheader("재고 추이")
chart_df = pd.DataFrame({
    '가용재고(t)': avail_stock.values,
    '분석대기(t)': pending.values,
    '생산(t)':    prod_total.values,
    '납품(t)':    ship_total.values,
}, index=dr)
chart_df = chart_df[
    (chart_df.index >= pd.Timestamp(s_date)) &
    (chart_df.index <= pd.Timestamp(e_date))
]
t1, t2 = st.tabs(["재고 추이","생산/납품"])
with t1: st.line_chart(chart_df[['가용재고(t)','분석대기(t)']])
with t2: st.bar_chart(chart_df[['생산(t)','납품(t)']])

# ── 엑셀 다운로드
st.divider()
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as w:
    df_tbl.to_excel(w, index=False, sheet_name='시뮬레이션')
st.download_button(
    "📥 엑셀 다운로드", buf.getvalue(),
    file_name=f"시뮬레이션_{dr[0].strftime('%Y%m%d')}_{dr[-1].strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)