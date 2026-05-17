
# app.py
# 생산계획 기반 자재 부족/발주 액션 대시보드
# 실행: streamlit run app.py

from __future__ import annotations

import io
import re
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# 0. 페이지 설정
# =========================================================

st.set_page_config(
    page_title="Material Action Dashboard",
    page_icon="📦",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    .small-note {color:#6b7280; font-size:0.9rem;}
    .big-title {font-size:2rem; font-weight:800; margin-bottom:0.2rem;}
    div[data-testid="stMetricValue"] {font-size: 1.65rem;}
    .risk-high {color:#b91c1c; font-weight:700;}
    .risk-mid {color:#b45309; font-weight:700;}
    .risk-low {color:#047857; font-weight:700;}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 1. 안전 유틸 함수
# =========================================================

def clean_text(x: Any) -> str:
    if pd.isna(x):
        return ""
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def norm(x: Any) -> str:
    s = clean_text(x).lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^0-9a-z가-힣#./+\- ]", "", s)
    return s.strip()


def to_num(x: Any, default: float = 0.0) -> float:
    if pd.isna(x):
        return default
    if isinstance(x, (int, float, np.number)):
        try:
            if pd.isna(x):
                return default
            return float(x)
        except Exception:
            return default
    s = str(x).strip().replace(",", "").replace("$", "").replace("%", "")
    if s == "":
        return default
    try:
        return float(s)
    except Exception:
        return default


def to_date_safe(x: Any) -> Optional[pd.Timestamp]:
    if pd.isna(x) or x == "":
        return None
    try:
        dt = pd.to_datetime(x, errors="coerce")
        if pd.isna(dt):
            return None
        return pd.Timestamp(dt).normalize()
    except Exception:
        return None


def parse_leadtime_days(x: Any, default: int = 30) -> int:
    s = norm(x)
    if not s:
        return default
    nums = re.findall(r"\d+", s)
    if not nums:
        return default
    n = int(nums[0])
    if "week" in s or "weeks" in s or "주" in s:
        return n * 7
    return n


def safe_div(a: float, b: float) -> float:
    try:
        if b == 0:
            return 0.0
        return a / b
    except Exception:
        return 0.0


def find_sheet_name(sheet_names: List[str], keywords: List[str]) -> Optional[str]:
    for kw in keywords:
        for s in sheet_names:
            if kw in s:
                return s
    return None


def make_material_key(item: Any, color: Any, unit: Any = "") -> str:
    return f"{norm(item)}|{norm(color)}|{norm(unit)}"


def make_item_key(item: Any) -> str:
    return norm(item)


def format_qty(x: Any) -> str:
    n = to_num(x, np.nan)
    if pd.isna(n):
        return "-"
    if abs(n) >= 1000:
        return f"{n:,.0f}"
    if abs(n - round(n)) < 1e-6:
        return f"{n:,.0f}"
    return f"{n:,.2f}"


# =========================================================
# 2. 엑셀 로드
# =========================================================

@st.cache_data(show_spinner=False)
def load_workbook_bytes(file_bytes: bytes) -> Tuple[List[str], Dict[str, pd.DataFrame]]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}
    for sh in xls.sheet_names:
        try:
            sheets[sh] = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, header=None)
        except Exception:
            # 한 시트가 깨져도 전체 앱이 죽지 않게 처리
            sheets[sh] = pd.DataFrame()
    return xls.sheet_names, sheets


# =========================================================
# 3. 시트별 파서
# =========================================================

def parse_po(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty:
        return pd.DataFrame()

    # STYLE CODE가 있는 행을 헤더로 추정
    header_row = None
    for i in range(min(len(raw), 30)):
        vals = [norm(v) for v in raw.iloc[i].tolist()]
        if any("style code" in v for v in vals) and any("po qty" in v for v in vals):
            header_row = i
            break

    if header_row is None:
        return pd.DataFrame()

    headers = [clean_text(x) or f"col_{j}" for j, x in enumerate(raw.iloc[header_row].tolist())]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers[: len(df.columns)]

    rename = {}
    for c in df.columns:
        nc = norm(c)
        if "style code" in nc:
            rename[c] = "item_code"
        elif "style" in nc and "description" in nc:
            rename[c] = "style"
        elif "color" in nc:
            rename[c] = "color"
        elif "po qty" in nc:
            rename[c] = "po_qty"
        elif nc == "mtl":
            rename[c] = "mtl"
        elif nc == "cmt":
            rename[c] = "cmt"
        elif nc == "fob":
            rename[c] = "fob"
        elif "amount" in nc:
            rename[c] = "amount"
        elif "buyer po" in nc:
            rename[c] = "buyer_po_short"

    df = df.rename(columns=rename)
    keep = [c for c in ["item_code", "style", "color", "po_qty", "mtl", "cmt", "fob", "amount", "buyer_po_short"] if c in df.columns]
    df = df[keep].dropna(how="all")
    if "po_qty" in df.columns:
        df["po_qty"] = df["po_qty"].apply(to_num)
    return df


def parse_production(raw: pd.DataFrame) -> pd.DataFrame:
    """
    생산계획 시트는 고정 위치가 비교적 명확하므로 컬럼 위치 기반으로 읽는다.
    화면 기준:
    BUYER=1, PROD YEAR=2, PROD MONTH=3, SHIP MONTH=4, MTL PR NO=5,
    BUYER PO=6, ITEM CODE=7, STYLE=8, COLOR=9, PROD QTY=10,
    REF=11, FROM=18, TO=19, SHIP DATE=20, COMPLETED=21, BAL=22, MATERIAL=24, REMARK=25
    """
    if raw.empty or raw.shape[1] < 25:
        return pd.DataFrame()

    rows = []
    for idx in range(5, len(raw)):
        r = raw.iloc[idx]
        buyer = clean_text(r.iloc[1] if len(r) > 1 else "")
        buyer_po = clean_text(r.iloc[6] if len(r) > 6 else "")
        style = clean_text(r.iloc[8] if len(r) > 8 else "")
        color = clean_text(r.iloc[9] if len(r) > 9 else "")
        prod_qty = to_num(r.iloc[10] if len(r) > 10 else 0)

        # 실제 오더 행만 남김
        if not buyer or not style or prod_qty <= 0:
            continue
        if buyer in [".", "BUYER"] or "2026" == style:
            continue

        sewing_from = to_date_safe(r.iloc[18] if len(r) > 18 else None)
        sewing_to = to_date_safe(r.iloc[19] if len(r) > 19 else None)
        ship_date = to_date_safe(r.iloc[20] if len(r) > 20 else None)

        if sewing_from is None and ship_date is not None:
            # 생산 시작일이 비어 있으면 최소한 선적일 기준으로 표시
            sewing_from = ship_date

        rows.append(
            {
                "source_row": idx + 1,
                "buyer": buyer,
                "prod_year": clean_text(r.iloc[2] if len(r) > 2 else ""),
                "prod_month": clean_text(r.iloc[3] if len(r) > 3 else ""),
                "ship_month": clean_text(r.iloc[4] if len(r) > 4 else ""),
                "pr_no": clean_text(r.iloc[5] if len(r) > 5 else ""),
                "buyer_po": buyer_po,
                "item_code": clean_text(r.iloc[7] if len(r) > 7 else ""),
                "style": style,
                "color": color,
                "prod_qty": prod_qty,
                "ref_qty": to_num(r.iloc[11] if len(r) > 11 else 0),
                "sewing_from": sewing_from,
                "sewing_to": sewing_to,
                "ship_date": ship_date,
                "completed_qty": to_num(r.iloc[21] if len(r) > 21 else 0),
                "balance": to_num(r.iloc[22] if len(r) > 22 else 0),
                "progress": to_num(r.iloc[23] if len(r) > 23 else 0),
                "material_status": clean_text(r.iloc[24] if len(r) > 24 else ""),
                "remark": clean_text(r.iloc[25] if len(r) > 25 else ""),
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # 동일 PO가 분할 생산된 경우를 구분하기 위한 행 키
    df["prod_key"] = (
        df["buyer_po"].astype(str) + " | " +
        df["style"].astype(str) + " | " +
        df["color"].astype(str) + " | " +
        df["prod_qty"].round(4).astype(str) + " | " +
        df["source_row"].astype(str)
    )
    return df


def parse_bom(raw: pd.DataFrame) -> pd.DataFrame:
    """
    자재품목리스트는 18~25열에 스타일별 소요량이 있고,
    1행 STYLE, 2행 ITEM#, 3행 COLOR가 스타일 메타데이터다.
    10행이 자재 기본 헤더, 12행부터 자재 데이터.
    """
    if raw.empty or raw.shape[1] < 20:
        return pd.DataFrame()

    style_cols = []
    max_col = min(raw.shape[1], 28)
    for c in range(18, max_col):
        style = clean_text(raw.iat[1, c] if raw.shape[0] > 1 else "")
        item_code = clean_text(raw.iat[2, c] if raw.shape[0] > 2 else "")
        style_color = clean_text(raw.iat[3, c] if raw.shape[0] > 3 else "")
        if style:
            style_cols.append((c, style, item_code, style_color))

    rows = []
    for idx in range(12, len(raw)):
        r = raw.iloc[idx]
        item = clean_text(r.iloc[7] if len(r) > 7 else "")
        if not item:
            continue

        base = {
            "source_row": idx + 1,
            "key": clean_text(r.iloc[0] if len(r) > 0 else ""),
            "ajs_code": clean_text(r.iloc[1] if len(r) > 1 else ""),
            "mtl_type": clean_text(r.iloc[3] if len(r) > 3 else ""),
            "supplier": clean_text(r.iloc[4] if len(r) > 4 else ""),
            "co": clean_text(r.iloc[5] if len(r) > 5 else ""),
            "lead_time_raw": clean_text(r.iloc[6] if len(r) > 6 else ""),
            "leadtime_days": parse_leadtime_days(r.iloc[6] if len(r) > 6 else ""),
            "item": item,
            "spec": clean_text(r.iloc[8] if len(r) > 8 else ""),
            "material_color": clean_text(r.iloc[9] if len(r) > 9 else ""),
            "color2": clean_text(r.iloc[10] if len(r) > 10 else ""),
            "unit": clean_text(r.iloc[11] if len(r) > 11 else ""),
            "buyer_price": to_num(r.iloc[12] if len(r) > 12 else 0),
            "aj_price": to_num(r.iloc[13] if len(r) > 13 else 0),
            "remark": clean_text(r.iloc[14] if len(r) > 14 else ""),
            "moq": to_num(r.iloc[15] if len(r) > 15 else 0),
            "surcharge": to_num(r.iloc[16] if len(r) > 16 else 0),
            "leadtime_extra": clean_text(r.iloc[17] if len(r) > 17 else ""),
            "ttl_cons": to_num(r.iloc[26] if len(r) > 26 else 0),
            "loss_qty": to_num(r.iloc[27] if len(r) > 27 else 0),
        }

        for c, style, item_code, style_color in style_cols:
            usage = to_num(r.iloc[c] if len(r) > c else 0)
            if usage <= 0:
                continue
            row = base.copy()
            row.update(
                {
                    "style": style,
                    "style_item_code": item_code,
                    "style_color": style_color,
                    "usage_per_unit": usage,
                    "material_key": make_material_key(item, base["material_color"], base["unit"]),
                    "item_key": make_item_key(item),
                }
            )
            rows.append(row)

    return pd.DataFrame(rows)


def parse_inventory(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty or raw.shape[1] < 14:
        return pd.DataFrame()

    rows = []
    for idx in range(5, len(raw)):
        r = raw.iloc[idx]
        item = clean_text(r.iloc[7] if len(r) > 7 else "")
        if not item:
            continue

        qty = to_num(r.iloc[13] if len(r) > 13 else 0)
        rows.append(
            {
                "source_row": idx + 1,
                "rack": clean_text(r.iloc[2] if len(r) > 2 else ""),
                "status": clean_text(r.iloc[3] if len(r) > 3 else ""),
                "buyer": clean_text(r.iloc[4] if len(r) > 4 else ""),
                "pr_no": clean_text(r.iloc[5] if len(r) > 5 else ""),
                "supplier": clean_text(r.iloc[6] if len(r) > 6 else ""),
                "item": item,
                "spec": clean_text(r.iloc[8] if len(r) > 8 else ""),
                "material_color": clean_text(r.iloc[9] if len(r) > 9 else ""),
                "color2": clean_text(r.iloc[10] if len(r) > 10 else ""),
                "unit": clean_text(r.iloc[11] if len(r) > 11 else ""),
                "unit_price": to_num(r.iloc[12] if len(r) > 12 else 0),
                "stock_qty": qty,
                "amount": to_num(r.iloc[14] if len(r) > 14 else 0),
                "material_key": make_material_key(item, clean_text(r.iloc[9] if len(r) > 9 else ""), clean_text(r.iloc[11] if len(r) > 11 else "")),
                "item_key": make_item_key(item),
            }
        )

    return pd.DataFrame(rows)


def parse_inbound_outbound(raw: pd.DataFrame, kind: str) -> pd.DataFrame:
    if raw.empty or raw.shape[1] < 13:
        return pd.DataFrame()

    rows = []
    for idx in range(5, len(raw)):
        r = raw.iloc[idx]
        buyer = clean_text(r.iloc[1] if kind == "inbound" and len(r) > 1 else r.iloc[0] if len(r) > 0 else "")
        po = clean_text(r.iloc[2] if kind == "inbound" and len(r) > 2 else r.iloc[1] if len(r) > 1 else "")
        supplier = clean_text(r.iloc[4] if kind == "inbound" and len(r) > 4 else r.iloc[3] if len(r) > 3 else "")
        item = clean_text(r.iloc[6] if kind == "inbound" and len(r) > 6 else r.iloc[5] if len(r) > 5 else "")
        color = clean_text(r.iloc[8] if kind == "inbound" and len(r) > 8 else r.iloc[7] if len(r) > 7 else "")
        unit = clean_text(r.iloc[10] if kind == "inbound" and len(r) > 10 else r.iloc[9] if len(r) > 9 else "")
        qty = to_num(r.iloc[12] if kind == "inbound" and len(r) > 12 else r.iloc[11] if len(r) > 11 else 0)
        if not item or qty == 0:
            continue
        rows.append(
            {
                "source_row": idx + 1,
                "kind": kind,
                "buyer": buyer,
                "po": po,
                "supplier": supplier,
                "item": item,
                "color": color,
                "unit": unit,
                "qty": qty,
                "item_key": make_item_key(item),
                "material_key": make_material_key(item, color, unit),
            }
        )
    return pd.DataFrame(rows)


def parse_purchase(raw: pd.DataFrame) -> pd.DataFrame:
    """
    구매데이터는 자재품목리스트와 유사하나 MOQ/SURCHARGE가 명시되어 있어 보조 정보로 사용.
    """
    if raw.empty or raw.shape[1] < 16:
        return pd.DataFrame()

    rows = []
    for idx in range(12, len(raw)):
        r = raw.iloc[idx]
        item = clean_text(r.iloc[5] if len(r) > 5 else "")
        if not item:
            continue
        color = clean_text(r.iloc[7] if len(r) > 7 else "")
        unit = clean_text(r.iloc[9] if len(r) > 9 else "")
        rows.append(
            {
                "source_row": idx + 1,
                "mtl_type": clean_text(r.iloc[0] if len(r) > 0 else ""),
                "ajs_code": clean_text(r.iloc[2] if len(r) > 2 else ""),
                "supplier": clean_text(r.iloc[3] if len(r) > 3 else ""),
                "co": clean_text(r.iloc[4] if len(r) > 4 else ""),
                "item": item,
                "spec": clean_text(r.iloc[6] if len(r) > 6 else ""),
                "material_color": color,
                "color2": clean_text(r.iloc[8] if len(r) > 8 else ""),
                "unit": unit,
                "aj_price": to_num(r.iloc[10] if len(r) > 10 else 0),
                "buyer_price": to_num(r.iloc[11] if len(r) > 11 else 0),
                "remark": clean_text(r.iloc[12] if len(r) > 12 else ""),
                "moq": to_num(r.iloc[13] if len(r) > 13 else 0),
                "surcharge": to_num(r.iloc[14] if len(r) > 14 else 0),
                "material_key": make_material_key(item, color, unit),
                "item_key": make_item_key(item),
            }
        )
    return pd.DataFrame(rows)


def parse_all(sheets: Dict[str, pd.DataFrame], sheet_names: List[str]) -> Dict[str, pd.DataFrame]:
    sh_po = find_sheet_name(sheet_names, ["바이어 PO"])
    sh_prod = find_sheet_name(sheet_names, ["생산계획"])
    sh_bom = find_sheet_name(sheet_names, ["자재품목"])
    sh_stock = find_sheet_name(sheet_names, ["자재 재고"])
    sh_in = find_sheet_name(sheet_names, ["자재입고"])
    sh_out = find_sheet_name(sheet_names, ["자재출고"])
    sh_purchase = find_sheet_name(sheet_names, ["구매데이터"])

    parsed = {
        "po": parse_po(sheets.get(sh_po, pd.DataFrame())) if sh_po else pd.DataFrame(),
        "production": parse_production(sheets.get(sh_prod, pd.DataFrame())) if sh_prod else pd.DataFrame(),
        "bom": parse_bom(sheets.get(sh_bom, pd.DataFrame())) if sh_bom else pd.DataFrame(),
        "inventory": parse_inventory(sheets.get(sh_stock, pd.DataFrame())) if sh_stock else pd.DataFrame(),
        "inbound": parse_inbound_outbound(sheets.get(sh_in, pd.DataFrame()), "inbound") if sh_in else pd.DataFrame(),
        "outbound": parse_inbound_outbound(sheets.get(sh_out, pd.DataFrame()), "outbound") if sh_out else pd.DataFrame(),
        "purchase": parse_purchase(sheets.get(sh_purchase, pd.DataFrame())) if sh_purchase else pd.DataFrame(),
    }
    return parsed


# =========================================================
# 4. 자재 필요량/부족량 계산
# =========================================================

def enrich_bom_with_purchase(bom: pd.DataFrame, purchase: pd.DataFrame) -> pd.DataFrame:
    if bom.empty:
        return bom
    out = bom.copy()
    if purchase.empty:
        return out

    purch_cols = ["material_key", "moq", "surcharge", "aj_price", "buyer_price", "supplier"]
    p = purchase[purch_cols].copy()
    p = p.groupby("material_key", as_index=False).agg(
        {
            "moq": "max",
            "surcharge": "max",
            "aj_price": "max",
            "buyer_price": "max",
            "supplier": "first",
        }
    )
    out = out.merge(p, on="material_key", how="left", suffixes=("", "_purchase"))
    for col in ["moq", "surcharge", "aj_price", "buyer_price"]:
        if f"{col}_purchase" in out.columns:
            out[col] = np.where(out[col].fillna(0).eq(0), out[f"{col}_purchase"].fillna(0), out[col].fillna(0))
    if "supplier_purchase" in out.columns:
        out["supplier"] = np.where(out["supplier"].eq(""), out["supplier_purchase"].fillna(""), out["supplier"])
    return out


def build_stock_maps(inventory: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if inventory.empty:
        return pd.DataFrame(), pd.DataFrame()

    by_key = inventory.groupby("material_key", as_index=False).agg(
        stock_qty=("stock_qty", "sum"),
        stock_amount=("amount", "sum"),
        stock_rows=("source_row", lambda x: ", ".join(map(str, x.head(5)))),
    )
    by_item = inventory.groupby("item_key", as_index=False).agg(
        stock_qty_item=("stock_qty", "sum"),
    )
    return by_key, by_item


def build_requirements(
    production: pd.DataFrame,
    bom: pd.DataFrame,
    inventory: pd.DataFrame,
    loss_rate: float = 0.03,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if production.empty or bom.empty:
        return pd.DataFrame(), pd.DataFrame()

    stock_by_key, stock_by_item = build_stock_maps(inventory)

    rows = []
    for _, prod in production.iterrows():
        prod_style = norm(prod.get("style"))
        prod_color = norm(prod.get("color"))

        # 스타일 + 색상 기준 매칭
        matched = bom[
            (bom["style"].apply(norm) == prod_style)
            & (bom["style_color"].apply(norm) == prod_color)
        ].copy()

        # 색상이 약간 다를 경우 스타일만이라도 매칭
        if matched.empty:
            matched = bom[bom["style"].apply(norm) == prod_style].copy()

        if matched.empty:
            rows.append(
                {
                    "prod_key": prod.get("prod_key", ""),
                    "buyer_po": prod.get("buyer_po", ""),
                    "style": prod.get("style", ""),
                    "style_color": prod.get("color", ""),
                    "prod_qty": prod.get("prod_qty", 0),
                    "need_date": prod.get("sewing_from", None),
                    "item": "(BOM 매칭 실패)",
                    "supplier": "",
                    "unit": "",
                    "usage_per_unit": 0,
                    "required_qty": 0,
                    "required_with_loss": 0,
                    "stock_qty": 0,
                    "shortage_qty": 0,
                    "recommended_order_qty": 0,
                    "leadtime_days": 0,
                    "risk": "확인필요",
                    "action": "BOM에서 해당 STYLE/COLOR 매칭 필요",
                    "match_status": "BOM 매칭 실패",
                }
            )
            continue

        for _, m in matched.iterrows():
            required = float(prod.get("prod_qty", 0)) * float(m.get("usage_per_unit", 0))
            required_loss = required * (1 + loss_rate)

            stock_qty = 0.0
            match_status = "자재명+색상+단위 매칭"
            if not stock_by_key.empty:
                stock_hit = stock_by_key[stock_by_key["material_key"] == m.get("material_key", "")]
                if not stock_hit.empty:
                    stock_qty = float(stock_hit["stock_qty"].iloc[0])

            if stock_qty == 0 and not stock_by_item.empty:
                item_hit = stock_by_item[stock_by_item["item_key"] == m.get("item_key", "")]
                if not item_hit.empty:
                    stock_qty = float(item_hit["stock_qty_item"].iloc[0])
                    match_status = "자재명 기준 매칭"

            shortage = max(required_loss - stock_qty, 0)
            moq = float(m.get("moq", 0) or 0)
            rec_order = 0 if shortage <= 0 else max(shortage, moq)

            need_date = prod.get("sewing_from", None)
            today = pd.Timestamp(date.today()).normalize()
            days_left = None
            if need_date is not None and not pd.isna(need_date):
                days_left = int((pd.Timestamp(need_date).normalize() - today).days)

            lead_days = int(m.get("leadtime_days", 30) or 30)

            if shortage <= 0:
                risk = "정상"
                action = "현재 재고로 커버 가능"
            elif days_left is not None and days_left < lead_days:
                risk = "긴급"
                action = f"{clean_text(m.get('supplier')) or '공급업체 확인'}에 {clean_text(m.get('item'))} {format_qty(rec_order)}{clean_text(m.get('unit'))} 발주/납기 가능 여부 즉시 확인"
            else:
                risk = "주의"
                action = f"{clean_text(m.get('supplier')) or '공급업체 확인'}에 {clean_text(m.get('item'))} {format_qty(rec_order)}{clean_text(m.get('unit'))} 발주 검토"

            rows.append(
                {
                    "prod_key": prod.get("prod_key", ""),
                    "source_row_prod": prod.get("source_row", ""),
                    "buyer": prod.get("buyer", ""),
                    "pr_no": prod.get("pr_no", ""),
                    "buyer_po": prod.get("buyer_po", ""),
                    "item_code": prod.get("item_code", ""),
                    "style": prod.get("style", ""),
                    "style_color": prod.get("color", ""),
                    "prod_qty": float(prod.get("prod_qty", 0)),
                    "need_date": need_date,
                    "sewing_to": prod.get("sewing_to", None),
                    "ship_date": prod.get("ship_date", None),
                    "material_status": prod.get("material_status", ""),
                    "supplier": clean_text(m.get("supplier", "")),
                    "mtl_type": clean_text(m.get("mtl_type", "")),
                    "item": clean_text(m.get("item", "")),
                    "spec": clean_text(m.get("spec", "")),
                    "material_color": clean_text(m.get("material_color", "")),
                    "unit": clean_text(m.get("unit", "")),
                    "usage_per_unit": float(m.get("usage_per_unit", 0)),
                    "required_qty": required,
                    "loss_rate": loss_rate,
                    "required_with_loss": required_loss,
                    "stock_qty": stock_qty,
                    "shortage_qty": shortage,
                    "moq": moq,
                    "recommended_order_qty": rec_order,
                    "leadtime_days": lead_days,
                    "buyer_price": float(m.get("buyer_price", 0) or 0),
                    "aj_price": float(m.get("aj_price", 0) or 0),
                    "estimated_purchase_amount": rec_order * float(m.get("aj_price", 0) or 0),
                    "risk": risk,
                    "action": action,
                    "match_status": match_status,
                    "source_row_bom": m.get("source_row", ""),
                }
            )

    detail = pd.DataFrame(rows)

    if detail.empty:
        return detail, pd.DataFrame()

    def agg_action(g: pd.DataFrame) -> str:
        shortages = g[g["shortage_qty"] > 0]
        if shortages.empty:
            return "현재 재고로 커버 가능"
        top = shortages.sort_values(["risk", "shortage_qty"], ascending=[True, False]).head(1).iloc[0]
        return top["action"]

    po_summary = detail.groupby("prod_key", as_index=False).agg(
        buyer=("buyer", "first"),
        pr_no=("pr_no", "first"),
        buyer_po=("buyer_po", "first"),
        item_code=("item_code", "first"),
        style=("style", "first"),
        style_color=("style_color", "first"),
        prod_qty=("prod_qty", "first"),
        need_date=("need_date", "first"),
        sewing_to=("sewing_to", "first"),
        ship_date=("ship_date", "first"),
        material_status=("material_status", "first"),
        total_required_items=("item", "count"),
        shortage_items=("shortage_qty", lambda x: int((x > 0).sum())),
        urgent_items=("risk", lambda x: int((x == "긴급").sum())),
        estimated_purchase_amount=("estimated_purchase_amount", "sum"),
        main_action=("action", agg_action),
    )

    def po_risk(row):
        if row["urgent_items"] > 0:
            return "긴급"
        if row["shortage_items"] > 0:
            return "주의"
        return "정상"

    po_summary["risk"] = po_summary.apply(po_risk, axis=1)
    po_summary["risk_order"] = po_summary["risk"].map({"긴급": 0, "주의": 1, "정상": 2}).fillna(3)
    po_summary = po_summary.sort_values(["risk_order", "need_date", "shortage_items"], ascending=[True, True, False])

    return detail, po_summary


# =========================================================
# 5. UI 컴포넌트
# =========================================================

def risk_badge(risk: str) -> str:
    if risk == "긴급":
        return "🔴 긴급"
    if risk == "주의":
        return "🟠 주의"
    if risk == "정상":
        return "🟢 정상"
    return "⚪ 확인필요"


def display_dashboard(po_summary: pd.DataFrame, detail: pd.DataFrame) -> None:
    st.markdown('<div class="big-title">📦 생산계획 기반 자재 액션 대시보드</div>', unsafe_allow_html=True)
    st.caption("현재 업로드된 엑셀 기준으로 생산 예정 PO별 자재 필요량, 재고 부족, 발주 필요 항목을 계산합니다.")

    if po_summary.empty:
        st.warning("분석 가능한 PO 요약 데이터가 없습니다. 시트명과 생산계획/BOM/재고 구조를 확인해 주세요.")
        return

    total_po = len(po_summary)
    urgent_po = int((po_summary["risk"] == "긴급").sum())
    caution_po = int((po_summary["risk"] == "주의").sum())
    shortage_materials = int((detail["shortage_qty"] > 0).sum()) if not detail.empty else 0
    purchase_amount = float(detail["estimated_purchase_amount"].sum()) if not detail.empty else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("총 생산계획 건수", f"{total_po:,}")
    k2.metric("긴급 PO", f"{urgent_po:,}")
    k3.metric("부족 자재 항목", f"{shortage_materials:,}")
    k4.metric("추정 추가 발주 금액", f"${purchase_amount:,.0f}")

    c1, c2 = st.columns([0.9, 1.1])

    with c1:
        st.subheader("위험도 분포")
        risk_count = po_summary["risk"].value_counts().reset_index()
        risk_count.columns = ["risk", "count"]
        if not risk_count.empty:
            fig = px.pie(
                risk_count,
                names="risk",
                values="count",
                hole=0.45,
                color="risk",
                color_discrete_map={"긴급": "#ef4444", "주의": "#f59e0b", "정상": "#10b981"},
            )
            fig.update_traces(textposition="inside", textinfo="label+percent")
            fig.update_layout(height=330, margin=dict(l=10, r=10, t=30, b=10), showlegend=True)
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("생산일 기준 위험 PO")
        timeline = po_summary.copy()
        timeline["start"] = pd.to_datetime(timeline["need_date"], errors="coerce")
        timeline["end"] = pd.to_datetime(timeline["sewing_to"], errors="coerce")
        timeline["end"] = timeline["end"].fillna(timeline["start"] + pd.Timedelta(days=1))
        timeline = timeline.dropna(subset=["start"])
        if timeline.empty:
            st.info("생산일 정보가 없어 타임라인을 표시할 수 없습니다.")
        else:
            timeline["label"] = timeline["buyer_po"].astype(str) + " / " + timeline["style"].astype(str)
            fig = px.timeline(
                timeline.head(30),
                x_start="start",
                x_end="end",
                y="label",
                color="risk",
                color_discrete_map={"긴급": "#ef4444", "주의": "#f59e0b", "정상": "#10b981"},
                hover_data=["prod_qty", "shortage_items", "urgent_items", "main_action"],
            )
            fig.update_yaxes(autorange="reversed")
            fig.update_layout(height=330, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns([1, 1])
    with c3:
        st.subheader("부족 자재 TOP 10")
        short = detail[detail["shortage_qty"] > 0].copy() if not detail.empty else pd.DataFrame()
        if short.empty:
            st.success("현재 기준 부족 자재가 없습니다.")
        else:
            top = short.groupby(["item", "unit"], as_index=False).agg(shortage_qty=("shortage_qty", "sum"))
            top = top.sort_values("shortage_qty", ascending=False).head(10)
            fig = px.bar(top, x="shortage_qty", y="item", orientation="h", text="shortage_qty")
            fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

    with c4:
        st.subheader("공급업체별 발주 검토 금액")
        short = detail[detail["recommended_order_qty"] > 0].copy() if not detail.empty else pd.DataFrame()
        if short.empty:
            st.success("추가 발주 검토 항목이 없습니다.")
        else:
            sup = short.groupby("supplier", as_index=False).agg(
                estimated_purchase_amount=("estimated_purchase_amount", "sum"),
                item_count=("item", "nunique"),
            ).sort_values("estimated_purchase_amount", ascending=False).head(10)
            fig = px.bar(sup, x="supplier", y="estimated_purchase_amount", text="estimated_purchase_amount")
            fig.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), xaxis_tickangle=-30)
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("PO별 요약")
    view = po_summary.copy()
    view["상태"] = view["risk"].apply(risk_badge)
    view["생산일"] = view["need_date"].dt.strftime("%Y-%m-%d").fillna("-")
    view["생산종료"] = view["sewing_to"].dt.strftime("%Y-%m-%d").fillna("-")
    show_cols = [
        "상태", "buyer_po", "style", "style_color", "prod_qty", "생산일", "생산종료",
        "shortage_items", "urgent_items", "estimated_purchase_amount", "main_action", "prod_key"
    ]
    rename = {
        "buyer_po": "PO No.",
        "style": "Style",
        "style_color": "Color",
        "prod_qty": "생산수량",
        "shortage_items": "부족 자재 수",
        "urgent_items": "긴급 자재 수",
        "estimated_purchase_amount": "추정 발주 금액",
        "main_action": "우선 액션",
        "prod_key": "상세 선택 키",
    }
    st.dataframe(
        view[show_cols].rename(columns=rename),
        use_container_width=True,
        hide_index=True,
    )


def display_po_detail(selected_key: str, po_summary: pd.DataFrame, detail: pd.DataFrame, parsed: Dict[str, pd.DataFrame]) -> None:
    po_row = po_summary[po_summary["prod_key"] == selected_key]
    if po_row.empty:
        st.info("선택한 PO 정보를 찾을 수 없습니다.")
        return

    po = po_row.iloc[0]
    po_detail = detail[detail["prod_key"] == selected_key].copy()

    st.markdown("---")
    st.header(f"🔎 PO 상세: {po['buyer_po']} / {po['style']} / {po['style_color']}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("상태", risk_badge(po["risk"]))
    c2.metric("생산수량", f"{po['prod_qty']:,.0f} PCS")
    c3.metric("부족 자재", f"{int(po['shortage_items']):,}개")
    c4.metric("긴급 자재", f"{int(po['urgent_items']):,}개")

    st.subheader("1. 생산 정보")
    prod_info = pd.DataFrame(
        [
            ["Buyer", po.get("buyer", "")],
            ["PR No.", po.get("pr_no", "")],
            ["PO No.", po.get("buyer_po", "")],
            ["Item Code", po.get("item_code", "")],
            ["Style / Color", f"{po.get('style', '')} / {po.get('style_color', '')}"],
            ["생산수량", f"{po.get('prod_qty', 0):,.0f} PCS"],
            ["생산기간", f"{po.get('need_date').strftime('%Y-%m-%d') if pd.notna(po.get('need_date')) else '-'} ~ {po.get('sewing_to').strftime('%Y-%m-%d') if pd.notna(po.get('sewing_to')) else '-'}"],
            ["선적일", po.get("ship_date").strftime("%Y-%m-%d") if pd.notna(po.get("ship_date")) else "-"],
            ["Material 상태", po.get("material_status", "")],
        ],
        columns=["항목", "내용"],
    )
    st.dataframe(prod_info, use_container_width=True, hide_index=True)

    st.subheader("2. 자재별 필요량 · 재고 · 발주 액션")
    if po_detail.empty:
        st.warning("해당 PO에 연결된 자재 데이터가 없습니다.")
        return

    action_df = po_detail.copy()
    action_df["필요일"] = pd.to_datetime(action_df["need_date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
    action_df["상태"] = action_df["risk"].apply(risk_badge)
    action_df["필요량"] = action_df["required_with_loss"]
    action_df["현재재고"] = action_df["stock_qty"]
    action_df["부족량"] = action_df["shortage_qty"]
    action_df["권장발주량"] = action_df["recommended_order_qty"]
    action_df["예상금액"] = action_df["estimated_purchase_amount"]

    cols = [
        "상태", "supplier", "item", "material_color", "unit",
        "usage_per_unit", "prod_qty", "필요량", "현재재고", "부족량",
        "moq", "권장발주량", "필요일", "leadtime_days", "예상금액", "action"
    ]
    rename = {
        "supplier": "Supplier",
        "item": "자재명",
        "material_color": "자재 Color",
        "unit": "단위",
        "usage_per_unit": "1개당 소요량",
        "prod_qty": "생산수량",
        "moq": "MOQ",
        "leadtime_days": "리드타임(일)",
        "action": "요청 액션",
    }
    st.dataframe(
        action_df[cols].rename(columns=rename),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("3. 업체별 요청 문장")
    order_needed = action_df[action_df["권장발주량"] > 0].copy()
    if order_needed.empty:
        st.success("이 PO는 현재 재고 기준 추가 발주가 필요하지 않습니다.")
    else:
        for _, r in order_needed.sort_values(["risk", "supplier", "item"]).iterrows():
            st.markdown(
                f"- **{r['supplier'] or '공급업체 확인'}**에 "
                f"**{r['item']} / {r['material_color']}** "
                f"**{format_qty(r['권장발주량'])}{r['unit']}**를 "
                f"**{r['필요일']} 생산 전까지** 입고 가능한지 확인 필요. "
                f"(부족 {format_qty(r['부족량'])}{r['unit']}, MOQ {format_qty(r['moq'])}{r['unit']}, 리드타임 {int(r['leadtime_days'])}일)"
            )

    st.subheader("4. 원본 데이터 더보기")
    with st.expander("BOM 원본 연결 정보"):
        st.dataframe(
            po_detail[
                [
                    "source_row_bom", "supplier", "mtl_type", "item", "spec", "material_color",
                    "unit", "usage_per_unit", "buyer_price", "aj_price", "moq", "leadtime_days", "match_status"
                ]
            ].rename(
                columns={
                    "source_row_bom": "BOM 원본 행",
                    "supplier": "Supplier",
                    "mtl_type": "자재분류",
                    "item": "자재명",
                    "spec": "Spec",
                    "material_color": "Color",
                    "unit": "Unit",
                    "usage_per_unit": "소요량",
                    "buyer_price": "Buyer 단가",
                    "aj_price": "AJ 단가",
                    "moq": "MOQ",
                    "leadtime_days": "리드타임(일)",
                    "match_status": "재고 매칭 방식",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("현재 재고 원본"):
        inv = parsed.get("inventory", pd.DataFrame())
        if inv.empty:
            st.info("자재 재고 내역을 찾을 수 없습니다.")
        else:
            item_keys = set(po_detail["item_key"].dropna().tolist()) if "item_key" in po_detail.columns else set()
            inv_view = inv[inv["item_key"].isin(item_keys)].copy() if item_keys else pd.DataFrame()
            if inv_view.empty:
                st.info("선택 PO의 자재명과 직접 매칭되는 재고 행이 없습니다.")
            else:
                st.dataframe(inv_view, use_container_width=True, hide_index=True)

    with st.expander("최근 입고/출고 원본"):
        inb = parsed.get("inbound", pd.DataFrame())
        outb = parsed.get("outbound", pd.DataFrame())
        item_keys = set(po_detail["item_key"].dropna().tolist()) if "item_key" in po_detail.columns else set()

        col1, col2 = st.columns(2)
        with col1:
            st.caption("입고")
            if inb.empty:
                st.info("입고 데이터 없음")
            else:
                st.dataframe(inb[inb["item_key"].isin(item_keys)].head(50), use_container_width=True, hide_index=True)
        with col2:
            st.caption("출고")
            if outb.empty:
                st.info("출고 데이터 없음")
            else:
                st.dataframe(outb[outb["item_key"].isin(item_keys)].head(50), use_container_width=True, hide_index=True)


def display_data_quality(parsed: Dict[str, pd.DataFrame]) -> None:
    st.subheader("데이터 로드 상태")
    rows = []
    for name, df in parsed.items():
        rows.append({"데이터": name, "행 수": len(df), "상태": "OK" if not df.empty else "확인 필요"})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    st.caption("엑셀 양식이 바뀐 경우에도 앱이 중단되지 않도록, 비어 있는 데이터는 '확인 필요'로 표시합니다.")


# =========================================================
# 6. 메인 앱
# =========================================================

def main() -> None:
    st.sidebar.title("설정")

    uploaded = st.sidebar.file_uploader(
        "엑셀 파일 업로드",
        type=["xlsx", "xlsm"],
        help="산학협력자료(자재관련).xlsx 형식의 파일을 업로드하세요.",
    )

    loss_rate = st.sidebar.number_input(
        "Loss 반영률",
        min_value=0.0,
        max_value=0.20,
        value=0.03,
        step=0.005,
        format="%.3f",
    )

    st.sidebar.caption("현재 버전은 히스토리 기능 없이, 업로드된 엑셀의 현재 생산계획 기준으로 자재 부족과 발주 액션을 계산합니다.")

    if not uploaded:
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
        return

    try:
        file_bytes = uploaded.getvalue()
        sheet_names, sheets = load_workbook_bytes(file_bytes)
        parsed = parse_all(sheets, sheet_names)
        parsed["bom"] = enrich_bom_with_purchase(parsed.get("bom", pd.DataFrame()), parsed.get("purchase", pd.DataFrame()))

        detail, po_summary = build_requirements(
            production=parsed.get("production", pd.DataFrame()),
            bom=parsed.get("bom", pd.DataFrame()),
            inventory=parsed.get("inventory", pd.DataFrame()),
            loss_rate=float(loss_rate),
        )

    except Exception as e:
        st.error("엑셀을 분석하는 중 오류가 발생했습니다. 파일 구조가 크게 달라졌을 수 있습니다.")
        st.exception(e)
        return

    tab1, tab2, tab3 = st.tabs(["대시보드", "PO 상세", "데이터 점검"])

    with tab1:
        display_dashboard(po_summary, detail)

    with tab2:
        if po_summary.empty:
            st.warning("상세 조회 가능한 PO가 없습니다.")
        else:
            display_options = po_summary.copy()
            display_options["label"] = (
                display_options["risk"].apply(risk_badge)
                + " | "
                + display_options["buyer_po"].astype(str)
                + " | "
                + display_options["style"].astype(str)
                + " / "
                + display_options["style_color"].astype(str)
                + " | "
                + display_options["prod_qty"].apply(lambda x: f"{x:,.0f}PCS")
            )

            selected_label = st.selectbox("상세 조회할 PO 선택", display_options["label"].tolist())
            selected_key = display_options.loc[display_options["label"] == selected_label, "prod_key"].iloc[0]
            display_po_detail(selected_key, po_summary, detail, parsed)

    with tab3:
        display_data_quality(parsed)
        with st.expander("파싱된 생산계획 데이터 미리보기"):
            st.dataframe(parsed.get("production", pd.DataFrame()).head(100), use_container_width=True)
        with st.expander("파싱된 BOM 데이터 미리보기"):
            st.dataframe(parsed.get("bom", pd.DataFrame()).head(100), use_container_width=True)
        with st.expander("파싱된 재고 데이터 미리보기"):
            st.dataframe(parsed.get("inventory", pd.DataFrame()).head(100), use_container_width=True)


if __name__ == "__main__":
    main()
