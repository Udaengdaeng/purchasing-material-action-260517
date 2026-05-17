
# app.py
# Material Action Dashboard - Safe Version
# 실행: streamlit run app.py

from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# 0. Page Config / CSS
# =========================================================

st.set_page_config(
    page_title="Material Action Dashboard",
    page_icon="📦",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 2rem !important;
    }
    header[data-testid="stHeader"] {
        height: 0px;
    }
    .big-title {
        font-size: 2rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    .section-sub {
        color: #6b7280;
        font-size: 0.95rem;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.55rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 1. Utility
# =========================================================

def clean(x: Any) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def norm(x: Any) -> str:
    s = clean(x).lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^0-9a-z가-힣#./+\- ]", "", s)
    return s.strip()


def to_num(x: Any, default: float = 0.0) -> float:
    if x is None:
        return default
    try:
        if pd.isna(x):
            return default
    except Exception:
        pass
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("$", "").replace("%", "")
    if s == "" or s.lower() in ["nan", "none"]:
        return default
    try:
        return float(s)
    except Exception:
        return default


def parse_date(x: Any) -> Optional[pd.Timestamp]:
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    try:
        dt = pd.to_datetime(x, errors="coerce")
        if pd.isna(dt):
            return None
        return pd.Timestamp(dt).normalize()
    except Exception:
        return None


def fmt_qty(x: Any) -> str:
    n = to_num(x, 0)
    if abs(n - round(n)) < 1e-9:
        return f"{n:,.0f}"
    return f"{n:,.2f}"


def find_sheet(sheet_names: List[str], candidates: List[str]) -> Optional[str]:
    for c in candidates:
        for s in sheet_names:
            if c in s:
                return s
    return None


def risk_label(r: str) -> str:
    if r == "긴급":
        return "🔴 긴급"
    if r == "주의":
        return "🟠 주의"
    if r == "정상":
        return "🟢 정상"
    return "⚪ 확인 필요"


# =========================================================
# 2. Load Excel
# =========================================================

@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> Tuple[List[str], Dict[str, pd.DataFrame]]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}
    for s in xls.sheet_names:
        try:
            sheets[s] = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=None)
        except Exception:
            sheets[s] = pd.DataFrame()
    return xls.sheet_names, sheets


# =========================================================
# 3. Parsers
# =========================================================

def parse_production(raw: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "prod_key", "source_row", "buyer", "prod_year", "prod_month", "ship_month",
        "pr_no", "buyer_po", "item_code", "style", "color", "prod_qty",
        "ref_qty", "sewing_from", "sewing_to", "ship_date", "completed_qty",
        "balance", "material_status", "remark"
    ]

    if raw is None or raw.empty or raw.shape[1] < 12:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(0, len(raw)):
        r = raw.iloc[i]

        buyer = clean(r.iloc[1] if len(r) > 1 else "")
        buyer_po = clean(r.iloc[6] if len(r) > 6 else "")
        style = clean(r.iloc[8] if len(r) > 8 else "")
        color = clean(r.iloc[9] if len(r) > 9 else "")
        prod_qty = to_num(r.iloc[10] if len(r) > 10 else 0)

        # 실제 생산계획 행만 남김
        if not buyer or not style or prod_qty <= 0:
            continue
        if norm(buyer) in ["buyer", "ppm"] or norm(style) in ["style"]:
            continue
        if "total" in norm(style) or "city" in norm(style):
            continue

        sewing_from = parse_date(r.iloc[18] if len(r) > 18 else None)
        sewing_to = parse_date(r.iloc[19] if len(r) > 19 else None)
        ship_date = parse_date(r.iloc[20] if len(r) > 20 else None)

        if sewing_from is None and ship_date is not None:
            sewing_from = ship_date

        prod_key = f"{buyer_po}|{style}|{color}|{prod_qty}|row{i+1}"

        rows.append({
            "prod_key": prod_key,
            "source_row": i + 1,
            "buyer": buyer,
            "prod_year": clean(r.iloc[2] if len(r) > 2 else ""),
            "prod_month": clean(r.iloc[3] if len(r) > 3 else ""),
            "ship_month": clean(r.iloc[4] if len(r) > 4 else ""),
            "pr_no": clean(r.iloc[5] if len(r) > 5 else ""),
            "buyer_po": buyer_po,
            "item_code": clean(r.iloc[7] if len(r) > 7 else ""),
            "style": style,
            "color": color,
            "prod_qty": prod_qty,
            "ref_qty": to_num(r.iloc[11] if len(r) > 11 else 0),
            "sewing_from": sewing_from,
            "sewing_to": sewing_to,
            "ship_date": ship_date,
            "completed_qty": to_num(r.iloc[21] if len(r) > 21 else 0),
            "balance": to_num(r.iloc[22] if len(r) > 22 else 0),
            "material_status": clean(r.iloc[24] if len(r) > 24 else ""),
            "remark": clean(r.iloc[25] if len(r) > 25 else ""),
        })

    return pd.DataFrame(rows, columns=cols)


def parse_bom(raw: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "source_row", "style", "style_color", "style_item_code", "supplier",
        "lead_time", "leadtime_days", "item", "spec", "material_color",
        "unit", "buyer_price", "aj_price", "moq", "usage_per_unit"
    ]

    if raw is None or raw.empty or raw.shape[0] < 13 or raw.shape[1] < 19:
        return pd.DataFrame(columns=cols)

    style_cols = []
    max_col = min(raw.shape[1], 28)
    for c in range(18, max_col):
        style = clean(raw.iat[1, c] if raw.shape[0] > 1 else "")
        item_code = clean(raw.iat[2, c] if raw.shape[0] > 2 else "")
        style_color = clean(raw.iat[3, c] if raw.shape[0] > 3 else "")
        if style:
            style_cols.append((c, style, item_code, style_color))

    rows = []
    for i in range(12, len(raw)):
        r = raw.iloc[i]
        item = clean(r.iloc[7] if len(r) > 7 else "")
        if not item:
            continue

        supplier = clean(r.iloc[4] if len(r) > 4 else "")
        lead_time = clean(r.iloc[6] if len(r) > 6 else "")
        lead_nums = re.findall(r"\d+", lead_time)
        lead_days = int(lead_nums[0]) * 7 if lead_nums and "week" in lead_time.lower() else int(lead_nums[0]) if lead_nums else 30

        for c, style, item_code, style_color in style_cols:
            usage = to_num(r.iloc[c] if len(r) > c else 0)
            if usage <= 0:
                continue
            rows.append({
                "source_row": i + 1,
                "style": style,
                "style_color": style_color,
                "style_item_code": item_code,
                "supplier": supplier,
                "lead_time": lead_time,
                "leadtime_days": lead_days,
                "item": item,
                "spec": clean(r.iloc[8] if len(r) > 8 else ""),
                "material_color": clean(r.iloc[9] if len(r) > 9 else ""),
                "unit": clean(r.iloc[11] if len(r) > 11 else ""),
                "buyer_price": to_num(r.iloc[12] if len(r) > 12 else 0),
                "aj_price": to_num(r.iloc[13] if len(r) > 13 else 0),
                "moq": to_num(r.iloc[15] if len(r) > 15 else 0),
                "usage_per_unit": usage,
            })

    return pd.DataFrame(rows, columns=cols)


def parse_inventory(raw: pd.DataFrame) -> pd.DataFrame:
    cols = ["source_row", "supplier", "item", "material_color", "unit", "stock_qty", "unit_price"]

    if raw is None or raw.empty or raw.shape[1] < 14:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(0, len(raw)):
        r = raw.iloc[i]
        item = clean(r.iloc[7] if len(r) > 7 else "")
        stock_qty = to_num(r.iloc[13] if len(r) > 13 else 0)
        if not item or stock_qty == 0:
            continue
        rows.append({
            "source_row": i + 1,
            "supplier": clean(r.iloc[6] if len(r) > 6 else ""),
            "item": item,
            "material_color": clean(r.iloc[9] if len(r) > 9 else ""),
            "unit": clean(r.iloc[11] if len(r) > 11 else ""),
            "stock_qty": stock_qty,
            "unit_price": to_num(r.iloc[12] if len(r) > 12 else 0),
        })

    return pd.DataFrame(rows, columns=cols)


def parse_simple_flow(raw: pd.DataFrame, kind: str) -> pd.DataFrame:
    cols = ["kind", "source_row", "supplier", "item", "color", "unit", "qty"]
    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)

    rows = []
    for i in range(0, len(raw)):
        r = raw.iloc[i]
        vals = [clean(v) for v in r.tolist()]
        # item 후보: 가장 그럴듯한 긴 텍스트
        item = ""
        for v in vals:
            if len(v) >= 4 and not re.fullmatch(r"[0-9.,/$\-]+", v):
                if any(k in norm(v) for k in ["ny", "poly", "zipper", "web", "fabric", "chain", "guard", "non woven", "elastic", "slider"]):
                    item = v
                    break
        if not item:
            continue
        qty_candidates = [to_num(v, np.nan) for v in vals]
        qty_candidates = [q for q in qty_candidates if not pd.isna(q) and q > 0]
        qty = qty_candidates[-1] if qty_candidates else 0
        rows.append({
            "kind": kind,
            "source_row": i + 1,
            "supplier": "",
            "item": item,
            "color": "",
            "unit": "",
            "qty": qty,
        })

    return pd.DataFrame(rows, columns=cols)


# =========================================================
# 4. Calculation
# =========================================================

def build_analysis(prod: pd.DataFrame, bom: pd.DataFrame, inv: pd.DataFrame, loss_rate: float) -> Tuple[pd.DataFrame, pd.DataFrame]:
    detail_cols = [
        "prod_key", "buyer_po", "buyer", "style", "style_color", "prod_qty",
        "need_date", "sewing_to", "ship_date", "supplier", "item",
        "material_color", "unit", "usage_per_unit", "required_qty",
        "required_with_loss", "stock_qty", "shortage_qty", "moq",
        "recommended_order_qty", "leadtime_days", "risk", "action",
        "source_row_prod", "source_row_bom"
    ]
    summary_cols = [
        "prod_key", "buyer_po", "buyer", "style", "style_color", "prod_qty",
        "need_date", "sewing_to", "ship_date", "shortage_items", "urgent_items",
        "risk", "main_action"
    ]

    if prod.empty:
        return pd.DataFrame(columns=detail_cols), pd.DataFrame(columns=summary_cols)

    if bom.empty:
        # 생산계획만이라도 보여주기
        summary = prod.copy()
        summary["shortage_items"] = 0
        summary["urgent_items"] = 0
        summary["risk"] = "확인필요"
        summary["main_action"] = "BOM 데이터 매칭 필요"
        return pd.DataFrame(columns=detail_cols), summary[summary_cols]

    inv_group = pd.DataFrame(columns=["item_key", "stock_qty"])
    if not inv.empty:
        inv2 = inv.copy()
        inv2["item_key"] = inv2["item"].apply(norm)
        inv_group = inv2.groupby("item_key", as_index=False)["stock_qty"].sum()

    rows = []
    today = pd.Timestamp.today().normalize()

    for _, p in prod.iterrows():
        p_style = norm(p["style"])
        p_color = norm(p["color"])

        matched = bom[(bom["style"].apply(norm) == p_style) & (bom["style_color"].apply(norm) == p_color)].copy()
        if matched.empty:
            matched = bom[bom["style"].apply(norm) == p_style].copy()

        if matched.empty:
            continue

        for _, b in matched.iterrows():
            required = float(p["prod_qty"]) * float(b["usage_per_unit"])
            required_loss = required * (1 + loss_rate)

            item_key = norm(b["item"])
            stock_qty = 0.0
            if not inv_group.empty:
                hit = inv_group[inv_group["item_key"] == item_key]
                if not hit.empty:
                    stock_qty = float(hit["stock_qty"].iloc[0])

            shortage = max(required_loss - stock_qty, 0)
            moq = float(b["moq"] or 0)
            rec = 0 if shortage <= 0 else max(shortage, moq)

            need_date = p["sewing_from"]
            days_left = None
            if need_date is not None and not pd.isna(need_date):
                days_left = int((pd.Timestamp(need_date).normalize() - today).days)

            lead = int(b["leadtime_days"] or 30)
            if shortage <= 0:
                risk = "정상"
                action = "현재 재고로 커버 가능"
            elif days_left is not None and days_left < lead:
                risk = "긴급"
                action = f"{b['supplier'] or '공급업체 확인'}에 {b['item']} {fmt_qty(rec)}{b['unit']}를 생산 전까지 입고 가능한지 즉시 확인"
            else:
                risk = "주의"
                action = f"{b['supplier'] or '공급업체 확인'}에 {b['item']} {fmt_qty(rec)}{b['unit']} 발주 검토"

            rows.append({
                "prod_key": p["prod_key"],
                "buyer_po": p["buyer_po"],
                "buyer": p["buyer"],
                "style": p["style"],
                "style_color": p["color"],
                "prod_qty": p["prod_qty"],
                "need_date": p["sewing_from"],
                "sewing_to": p["sewing_to"],
                "ship_date": p["ship_date"],
                "supplier": b["supplier"],
                "item": b["item"],
                "material_color": b["material_color"],
                "unit": b["unit"],
                "usage_per_unit": b["usage_per_unit"],
                "required_qty": required,
                "required_with_loss": required_loss,
                "stock_qty": stock_qty,
                "shortage_qty": shortage,
                "moq": moq,
                "recommended_order_qty": rec,
                "leadtime_days": lead,
                "risk": risk,
                "action": action,
                "source_row_prod": p["source_row"],
                "source_row_bom": b["source_row"],
            })

    detail = pd.DataFrame(rows, columns=detail_cols)

    if detail.empty:
        summary = prod.copy()
        summary["shortage_items"] = 0
        summary["urgent_items"] = 0
        summary["risk"] = "확인필요"
        summary["main_action"] = "BOM과 생산계획 매칭 실패"
        return detail, summary[summary_cols]

    summary = detail.groupby("prod_key", as_index=False).agg(
        buyer_po=("buyer_po", "first"),
        buyer=("buyer", "first"),
        style=("style", "first"),
        style_color=("style_color", "first"),
        prod_qty=("prod_qty", "first"),
        need_date=("need_date", "first"),
        sewing_to=("sewing_to", "first"),
        ship_date=("ship_date", "first"),
        shortage_items=("shortage_qty", lambda x: int((x > 0).sum())),
        urgent_items=("risk", lambda x: int((x == "긴급").sum())),
    )

    def summary_risk(r):
        if r["urgent_items"] > 0:
            return "긴급"
        if r["shortage_items"] > 0:
            return "주의"
        return "정상"

    summary["risk"] = summary.apply(summary_risk, axis=1)

    actions = []
    for key, g in detail.groupby("prod_key"):
        short = g[g["shortage_qty"] > 0].copy()
        if short.empty:
            actions.append({"prod_key": key, "main_action": "현재 재고로 커버 가능"})
        else:
            rank = {"긴급": 0, "주의": 1, "정상": 2}
            short["rank"] = short["risk"].map(rank).fillna(9)
            top = short.sort_values(["rank", "shortage_qty"], ascending=[True, False]).iloc[0]
            actions.append({"prod_key": key, "main_action": top["action"]})

    summary = summary.merge(pd.DataFrame(actions), on="prod_key", how="left")
    summary["main_action"] = summary["main_action"].fillna("확인 필요")
    rank2 = {"긴급": 0, "주의": 1, "정상": 2, "확인필요": 3}
    summary["rank"] = summary["risk"].map(rank2).fillna(9)
    summary = summary.sort_values(["rank", "need_date"], ascending=[True, True])

    return detail, summary[summary_cols]


# =========================================================
# 5. UI
# =========================================================

def show_intro():
    st.markdown('<div class="big-title">📦 Material Action Dashboard</div>', unsafe_allow_html=True)
    st.info("왼쪽 사이드바에서 엑셀 파일을 업로드하면 분석이 시작됩니다.")
    st.markdown(
        """
        ### 이 앱이 보여주는 것
        - 어떤 PO가 자재 부족 위험이 있는지
        - 어떤 자재가 얼마나 부족한지
        - 어느 공급업체에 몇 개를 언제까지 요청해야 하는지
        - 상세 화면에서 BOM/재고/입고/출고 원본 데이터 확인
        """
    )


def show_dashboard(summary: pd.DataFrame, detail: pd.DataFrame):
    st.markdown('<div class="big-title">📦 Material Action Dashboard</div>', unsafe_allow_html=True)

    if summary.empty:
        st.warning("분석 가능한 생산계획 데이터가 없습니다. 데이터 점검 탭에서 시트 로드 상태를 확인해 주세요.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 생산계획", f"{len(summary):,}")
    c2.metric("긴급 PO", f"{(summary['risk'] == '긴급').sum():,}")
    c3.metric("주의 PO", f"{(summary['risk'] == '주의').sum():,}")
    c4.metric("부족 자재 항목", f"{(detail['shortage_qty'] > 0).sum() if not detail.empty else 0:,}")

    left, right = st.columns([0.9, 1.1])

    with left:
        st.subheader("위험도 분포")
        rc = summary["risk"].value_counts().reset_index()
        rc.columns = ["risk", "count"]
        fig = px.pie(
            rc,
            names="risk",
            values="count",
            hole=0.45,
            color="risk",
            color_discrete_map={"긴급": "#ef4444", "주의": "#f59e0b", "정상": "#10b981", "확인필요": "#9ca3af"},
        )
        fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("부족 자재 TOP")
        short = detail[detail["shortage_qty"] > 0].copy() if not detail.empty else pd.DataFrame()
        if short.empty:
            st.success("부족 자재가 없습니다.")
        else:
            top = short.groupby(["item", "unit"], as_index=False)["shortage_qty"].sum()
            top = top.sort_values("shortage_qty", ascending=False).head(10)
            fig = px.bar(top, x="shortage_qty", y="item", orientation="h", text="shortage_qty")
            fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10), yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("PO별 요약")
    view = summary.copy()
    view["상태"] = view["risk"].apply(risk_label)
    view["생산시작"] = pd.to_datetime(view["need_date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    view["생산종료"] = pd.to_datetime(view["sewing_to"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    st.dataframe(
        view[[
            "상태", "buyer_po", "style", "style_color", "prod_qty",
            "생산시작", "생산종료", "shortage_items", "urgent_items", "main_action"
        ]].rename(columns={
            "buyer_po": "PO No.",
            "style": "Style",
            "style_color": "Color",
            "prod_qty": "생산수량",
            "shortage_items": "부족 자재 수",
            "urgent_items": "긴급 자재 수",
            "main_action": "우선 액션",
        }),
        use_container_width=True,
        hide_index=True,
    )


def show_detail(summary: pd.DataFrame, detail: pd.DataFrame, parsed: Dict[str, pd.DataFrame]):
    st.header("PO 상세 조회")

    if summary.empty:
        st.warning("상세 조회 가능한 PO가 없습니다.")
        return

    opts = summary.copy()
    opts["label"] = (
        opts["risk"].apply(risk_label)
        + " | "
        + opts["buyer_po"].astype(str)
        + " | "
        + opts["style"].astype(str)
        + " / "
        + opts["style_color"].astype(str)
        + " | "
        + opts["prod_qty"].apply(lambda x: f"{x:,.0f}PCS")
    )

    selected = st.selectbox("PO 선택", opts["label"].tolist())
    key = opts.loc[opts["label"] == selected, "prod_key"].iloc[0]
    srow = summary[summary["prod_key"] == key].iloc[0]
    d = detail[detail["prod_key"] == key].copy()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("상태", risk_label(srow["risk"]))
    c2.metric("생산수량", f"{srow['prod_qty']:,.0f} PCS")
    c3.metric("부족 자재", f"{int(srow['shortage_items']):,}")
    c4.metric("긴급 자재", f"{int(srow['urgent_items']):,}")

    st.subheader("1. 생산 정보")
    info = pd.DataFrame([
        ["PO No.", srow["buyer_po"]],
        ["Style / Color", f"{srow['style']} / {srow['style_color']}"],
        ["생산수량", f"{srow['prod_qty']:,.0f} PCS"],
        ["생산시작", pd.to_datetime(srow["need_date"]).strftime("%Y-%m-%d") if pd.notna(srow["need_date"]) else "-"],
        ["생산종료", pd.to_datetime(srow["sewing_to"]).strftime("%Y-%m-%d") if pd.notna(srow["sewing_to"]) else "-"],
        ["선적일", pd.to_datetime(srow["ship_date"]).strftime("%Y-%m-%d") if pd.notna(srow["ship_date"]) else "-"],
    ], columns=["항목", "내용"])
    st.dataframe(info, use_container_width=True, hide_index=True)

    st.subheader("2. 자재 필요량 · 재고 · 발주 액션")
    if d.empty:
        st.info("이 PO는 BOM 매칭 결과가 없습니다.")
    else:
        dv = d.copy()
        dv["상태"] = dv["risk"].apply(risk_label)
        st.dataframe(
            dv[[
                "상태", "supplier", "item", "material_color", "unit",
                "usage_per_unit", "prod_qty", "required_with_loss",
                "stock_qty", "shortage_qty", "moq", "recommended_order_qty",
                "leadtime_days", "action"
            ]].rename(columns={
                "supplier": "Supplier",
                "item": "자재명",
                "material_color": "자재 Color",
                "unit": "단위",
                "usage_per_unit": "1개당 소요량",
                "prod_qty": "생산수량",
                "required_with_loss": "Loss 포함 필요량",
                "stock_qty": "현재 재고",
                "shortage_qty": "부족량",
                "moq": "MOQ",
                "recommended_order_qty": "권장 발주량",
                "leadtime_days": "리드타임(일)",
                "action": "요청 액션",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("3. 업체별 요청 문장")
        need = dv[dv["recommended_order_qty"] > 0]
        if need.empty:
            st.success("현재 재고 기준 추가 발주가 필요하지 않습니다.")
        else:
            for _, r in need.iterrows():
                st.markdown(
                    f"- **{r['supplier'] or '공급업체 확인'}**에 "
                    f"**{r['item']} / {r['material_color']}** "
                    f"**{fmt_qty(r['recommended_order_qty'])}{r['unit']}**를 "
                    f"생산 전까지 입고 가능한지 확인 필요."
                )

    st.subheader("4. 원본 데이터 확인")
    with st.expander("파싱된 생산계획 원본"):
        st.dataframe(parsed["production"].head(200), use_container_width=True)
    with st.expander("파싱된 BOM 원본"):
        st.dataframe(parsed["bom"].head(200), use_container_width=True)
    with st.expander("파싱된 재고 원본"):
        st.dataframe(parsed["inventory"].head(200), use_container_width=True)


def show_quality(parsed: Dict[str, pd.DataFrame]):
    st.header("데이터 점검")
    q = pd.DataFrame([
        {"시트/데이터": "생산계획", "행 수": len(parsed["production"]), "상태": "OK" if len(parsed["production"]) else "확인 필요"},
        {"시트/데이터": "자재품목리스트(BOM)", "행 수": len(parsed["bom"]), "상태": "OK" if len(parsed["bom"]) else "확인 필요"},
        {"시트/데이터": "자재 재고 내역", "행 수": len(parsed["inventory"]), "상태": "OK" if len(parsed["inventory"]) else "확인 필요"},
        {"시트/데이터": "입고내역", "행 수": len(parsed["inbound"]), "상태": "OK" if len(parsed["inbound"]) else "선택"},
        {"시트/데이터": "출고내역", "행 수": len(parsed["outbound"]), "상태": "OK" if len(parsed["outbound"]) else "선택"},
    ])
    st.dataframe(q, use_container_width=True, hide_index=True)


# =========================================================
# 6. Main
# =========================================================

def main():
    st.sidebar.title("설정")
    uploaded = st.sidebar.file_uploader("엑셀 업로드", type=["xlsx", "xlsm"])
    loss_rate = st.sidebar.number_input("Loss 반영률", min_value=0.0, max_value=0.2, value=0.03, step=0.005, format="%.3f")

    if uploaded is None:
        show_intro()
        return

    try:
        sheet_names, sheets = load_excel(uploaded.getvalue())

        sh_prod = find_sheet(sheet_names, ["생산계획"])
        sh_bom = find_sheet(sheet_names, ["자재품목"])
        sh_inv = find_sheet(sheet_names, ["자재 재고"])
        sh_in = find_sheet(sheet_names, ["자재입고"])
        sh_out = find_sheet(sheet_names, ["자재출고"])

        prod = parse_production(sheets.get(sh_prod, pd.DataFrame()))
        bom = parse_bom(sheets.get(sh_bom, pd.DataFrame()))
        inv = parse_inventory(sheets.get(sh_inv, pd.DataFrame()))
        inbound = parse_simple_flow(sheets.get(sh_in, pd.DataFrame()), "입고")
        outbound = parse_simple_flow(sheets.get(sh_out, pd.DataFrame()), "출고")

        parsed = {
            "production": prod,
            "bom": bom,
            "inventory": inv,
            "inbound": inbound,
            "outbound": outbound,
        }

        detail, summary = build_analysis(prod, bom, inv, float(loss_rate))

    except Exception as e:
        st.error("엑셀 분석 중 오류가 발생했습니다. 데이터 점검이 필요합니다.")
        st.exception(e)
        return

    tab1, tab2, tab3 = st.tabs(["대시보드", "PO 상세", "데이터 점검"])

    with tab1:
        show_dashboard(summary, detail)
    with tab2:
        show_detail(summary, detail, parsed)
    with tab3:
        show_quality(parsed)


if __name__ == "__main__":
    main()
