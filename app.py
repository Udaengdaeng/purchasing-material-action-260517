import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import date
import plotly.express as px

# ---------------------------
# 페이지 설정 + UI FIX
# ---------------------------
st.set_page_config(layout="wide")

st.markdown("""
<style>
.block-container {
    padding-top: 3.5rem !important;
    padding-bottom: 2rem;
}

header[data-testid="stHeader"] {
    height: 0px;
}

.big-title {
    font-size:2rem;
    font-weight:800;
    margin-bottom:0.5rem;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="big-title">📦 Material Action Dashboard</div>', unsafe_allow_html=True)

# ---------------------------
# 업로드
# ---------------------------
uploaded = st.sidebar.file_uploader("엑셀 업로드", type=["xlsx"])

loss_rate = st.sidebar.number_input("Loss 반영률", value=0.03)

if not uploaded:
    st.info("엑셀을 업로드하세요.")
    st.stop()

# ---------------------------
# 데이터 로드 (안터지게)
# ---------------------------
try:
    xls = pd.ExcelFile(uploaded)
    sheets = {s: pd.read_excel(uploaded, sheet_name=s, header=None) for s in xls.sheet_names}
except:
    st.error("엑셀 로드 실패")
    st.stop()

# ---------------------------
# 생산계획 간단 파싱
# ---------------------------
def parse_production(df):
    rows = []
    for i in range(5, len(df)):
        try:
            buyer = str(df.iloc[i,1])
            style = str(df.iloc[i,8])
            color = str(df.iloc[i,9])
            qty = float(df.iloc[i,10])
            
            if not style or qty == 0:
                continue
            
            rows.append({
                "PO": str(df.iloc[i,6]),
                "style": style,
                "color": color,
                "qty": qty,
                "date": df.iloc[i,18]
            })
        except:
            continue
    return pd.DataFrame(rows)

# ---------------------------
# BOM 간단 파싱
# ---------------------------
def parse_bom(df):
    rows = []
    for i in range(12, len(df)):
        try:
            item = str(df.iloc[i,7])
            usage = float(df.iloc[i,18])
            if usage == 0:
                continue
            
            rows.append({
                "style": str(df.iloc[1,18]),
                "item": item,
                "usage": usage
            })
        except:
            continue
    return pd.DataFrame(rows)

# ---------------------------
# 재고 파싱
# ---------------------------
def parse_stock(df):
    rows = []
    for i in range(5, len(df)):
        try:
            item = str(df.iloc[i,7])
            qty = float(df.iloc[i,13])
            if not item:
                continue
            
            rows.append({
                "item": item,
                "stock": qty
            })
        except:
            continue
    return pd.DataFrame(rows)

# ---------------------------
# 실행
# ---------------------------
prod_sheet = [s for s in sheets if "생산" in s][0]
bom_sheet = [s for s in sheets if "품목" in s][0]
stock_sheet = [s for s in sheets if "재고" in s][0]

prod = parse_production(sheets[prod_sheet])
bom = parse_bom(sheets[bom_sheet])
stock = parse_stock(sheets[stock_sheet])

# ---------------------------
# 계산
# ---------------------------
data = []

for _, p in prod.iterrows():
    matched = bom[bom["style"] == p["style"]]
    
    for _, m in matched.iterrows():
        need = p["qty"] * m["usage"] * (1+loss_rate)
        stock_qty = stock[stock["item"] == m["item"]]["stock"].sum()
        
        shortage = max(need - stock_qty, 0)
        
        data.append({
            "PO": p["PO"],
            "item": m["item"],
            "need": need,
            "stock": stock_qty,
            "shortage": shortage
        })

df = pd.DataFrame(data)

# ---------------------------
# 대시보드
# ---------------------------
st.subheader("부족 자재")

if df.empty:
    st.warning("데이터 없음")
else:
    fig = px.bar(df.groupby("item")["shortage"].sum().reset_index(),
                 x="item", y="shortage")
    st.plotly_chart(fig, use_container_width=True)

# ---------------------------
# 상세
# ---------------------------
po_list = df["PO"].dropna().unique()

selected = st.selectbox("PO 선택", po_list)

detail = df[df["PO"] == selected]

st.subheader("상세")

st.dataframe(detail)

# ---------------------------
# 액션
# ---------------------------
st.subheader("구매 액션")

for _, r in detail.iterrows():
    if r["shortage"] > 0:
        st.write(f"{r['item']} {round(r['shortage'],1)}개 발주 필요")
