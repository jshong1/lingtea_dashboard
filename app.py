import io
import os
from datetime import datetime

import gspread
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials
import plotly.graph_objects as go
import plotly.express as px

# -----------------------------------
# 기본 설정
# -----------------------------------

st.set_page_config(
    page_title="Lingtea Dashboard v5.2",
    layout="wide"
)

st.title("📊 Lingtea Dashboard v5.2")
st.caption("월별 채널/제품 분석 + 매출/출고량/공헌이익 통합 대시보드")

SHEET_ID = "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"

# -----------------------------------
# 공통 유틸
# -----------------------------------

def get_client():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]

    if os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file(
            "service_account.json",
            scopes=scope
        )
    else:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=scope
        )

    return gspread.authorize(creds)

def load_cost_input(sh):

    ws = sh.worksheet("COST_INPUT")
    data = ws.get_all_values()

    if len(data) < 3:
        return {}, {}

    header = data[0]  # 월 헤더 (B~M)
    months = header[1:]  # 첫 컬럼 제외

    # -----------------------------
    # 1. 물류비
    # -----------------------------
    logistics_row = data[1]

    logistics_dict = {}
    for i, m in enumerate(months):
        val = logistics_row[i + 1]
        if val not in ["", None]:
            val = str(val).replace(",", "").strip()
            logistics_dict[m] = float(val) if val != "" else 0
        else:
            logistics_dict[m] = 0

    # -----------------------------
    # 2. 광고비
    # -----------------------------
    ad_dict = {}

    for row in data[4:]:  # 상품 영역 시작
        category = str(row[0]).strip()

        if category == "":
            continue

        for i, m in enumerate(months):
            val = row[i + 1]

            try:
                val_float = float(val)
            except:
                val_float = 0

            val = str(val).replace(",", "").strip()

            if val == "" or val.lower() == "nan":
                val_float = 0
            else:
                try:
                    val_float = float(val)
                except:
                    val_float = 0

            ad_dict[(category, m)] = val_float

    return logistics_dict, ad_dict


def load_channel_cost(sh):
    """
    CHANNEL_COST 시트 로드
    A열: 년월 (예: 2026-01)
    B열: 거래처명
    C열: 품목군
    D열: 비용항목
    E열: 금액(VAT-)
    반환: dict { (년월, 거래처명, 품목군): 금액합계 }
    """
    try:
        ws = sh.worksheet("CHANNEL_COST")
        data = ws.get_all_values()
    except Exception:
        return {}

    if len(data) < 2:
        return {}

    channel_cost_dict = {}
    for row in data[1:]:
        if len(row) < 5:
            continue
        year_month = str(row[0]).strip()
        channel    = str(row[1]).strip()
        item_group = str(row[2]).strip()
        amount_str = str(row[4]).replace(",", "").strip()

        if year_month == "" or channel == "" or item_group == "":
            continue

        try:
            amount = float(amount_str)
        except:
            amount = 0

        key = (year_month, channel, item_group)
        channel_cost_dict[key] = channel_cost_dict.get(key, 0) + amount

    return channel_cost_dict


def clean_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False).str.strip(),
        errors="coerce"
    )


def sort_month_cols(cols):
    return sorted(cols, key=lambda x: pd.to_datetime(f"{x}-01", errors="coerce"))


def format_won(value):
    return f"{value:,.0f} 원"


def format_pct(value):
    return f"{value:.2f}%"


def safe_divide(a, b):
    return np.where(b != 0, a / b, 0)


def make_excel_file(sheet_dict: dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_dict.items():
            df.to_excel(writer, sheet_name=sheet_name[:31])
    return output.getvalue()

# -----------------------------------
# VIEW_TABLE 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_view_table():
    client = get_client()

    ws = client.open_by_key(SHEET_ID).worksheet("VIEW_TABLE")
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0]).copy()

    use_cols = [
        "출고일자",
        "출고년월",
        "거래처코드",
        "내품상품명",
        "총내품출고수량",
        "품목별매출(VAT제외)"
    ]
    df = df[use_cols].copy()

    df["출고일자"] = pd.to_datetime(df["출고일자"], errors="coerce")
    df["출고년월"] = df["출고년월"].astype(str).str.strip()
    df["총내품출고수량"] = clean_numeric(df["총내품출고수량"])
    df["품목별매출(VAT제외)"] = clean_numeric(df["품목별매출(VAT제외)"])

    # 최근 12개월만
    today = datetime.today()
    one_year_ago = today - relativedelta(months=12)
    df = df[df["출고일자"] >= one_year_ago].copy()

    return df


@st.cache_data(ttl=600)
def load_cost_master():
    """
    COST_MASTER 시트 로드
    A열: 년월 (예: 2025-12)
    B열: 상품코드
    C열: 상품명
    D열: 제품원가
    반환: dict { (년월, 상품명): 제품원가 }
    공헌이익 원가 시차 규칙: 출고년월 N월 → 전월(N-1월) 원가 사용
    """
    client = get_client()
    ws = client.open_by_key(SHEET_ID).worksheet("COST_MASTER")
    data = ws.get_all_values()

    cost_dict = {}
    for row in data[1:]:
        if len(row) < 4:
            continue
        year_month = str(row[0]).strip()
        item_name  = str(row[2]).strip()
        cost_str   = str(row[3]).replace(",", "").strip()
        if year_month == "" or item_name == "":
            continue
        try:
            cost = float(cost_str)
        except:
            cost = 0
        cost_dict[(year_month, item_name)] = cost

    return cost_dict


# -----------------------------------
# MASTER 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_master():
    client = get_client()

    # ITEM_MASTER (품목군 분류만 사용, 원가는 COST_MASTER에서 가져옴)
    item_ws = client.open_by_key(SHEET_ID).worksheet("ITEM_MASTER")
    item_data = item_ws.get_all_values()
    item_df = pd.DataFrame(item_data[1:], columns=item_data[0]).copy()

    # CUSTOMER_MASTER
    cust_ws = client.open_by_key(SHEET_ID).worksheet("CUSTOMER_MASTER")
    cust_data = cust_ws.get_all_values()
    cust_df = pd.DataFrame(cust_data[1:], columns=cust_data[0]).copy()

    if "거래처분류" not in cust_df.columns:
        cust_df["거래처분류"] = np.nan

    if "수수료율" not in cust_df.columns:
        cust_df["수수료율"] = 0

    # D열: 국내/해외 구분 컬럼 (4번째 컬럼)
    col_list = cust_df.columns.tolist()
    if len(col_list) >= 4:
        domestic_col = col_list[3]  # D열 (0-indexed: 3)
        cust_df["국내여부"] = cust_df[domestic_col].astype(str).str.strip()
    else:
        cust_df["국내여부"] = "국내"  # 컬럼 없으면 기본값 국내

    cust_df["수수료율"] = clean_numeric(cust_df["수수료율"]).fillna(0)
    cust_df["수수료율"] = np.where(cust_df["수수료율"] > 1, cust_df["수수료율"] / 100, cust_df["수수료율"])
    cust_df["거래처명"] = cust_df["거래처명"].astype(str).str.strip()
    cust_df["거래처분류"] = cust_df["거래처분류"].astype(str).str.strip()

    return item_df, cust_df


# -----------------------------------
# 통합 데이터셋 구성
# -----------------------------------

@st.cache_data(ttl=600)
def build_dataset():
    df = load_view_table()
    item_df, cust_df = load_master()
    cost_dict = load_cost_master()

    # 품목군만 조인 (원가는 COST_MASTER에서 별도 매칭)
    merged = df.merge(
        item_df[["상품명", "품목군"]],
        left_on="내품상품명",
        right_on="상품명",
        how="left"
    )

    merged["거래처분류"] = merged["거래처코드"]

    merged = merged.merge(
        cust_df[["거래처분류", "수수료율", "국내여부"]].drop_duplicates(),
        on="거래처분류",
        how="left"
    )

    merged["수수료율"] = merged["수수료율"].fillna(0)
    merged["국내여부"] = merged["국내여부"].fillna("국내")

    # 거래처분류 없는 경우 거래처코드로 대체
    merged["거래처분류"] = merged["거래처분류"].replace("", np.nan)
    merged["거래처분류"] = merged["거래처분류"].fillna(merged["거래처코드"])

    # -----------------------------
    # 원가 시차 반영: 출고년월 N월 → 전월(N-1월) COST_MASTER 원가 사용
    # 예) 2026-01 출고 → 2025-12 원가 적용
    # -----------------------------
    def get_prev_month_cost(row):
        try:
            ship_dt = pd.to_datetime(row["출고년월"] + "-01")
            prev_dt = ship_dt - relativedelta(months=1)
            prev_ym = prev_dt.strftime("%Y-%m")
        except:
            return 0
        return cost_dict.get((prev_ym, row["내품상품명"]), 0)

    merged["제품원가"] = merged.apply(get_prev_month_cost, axis=1)

    merged["원가총액"] = merged["총내품출고수량"] * merged["제품원가"]
    merged["채널수수료"] = merged["품목별매출(VAT제외)"] * merged["수수료율"]
    merged["마진"] = (
        merged["품목별매출(VAT제외)"]
        - merged["원가총액"]
        - merged["채널수수료"]
    )
    merged["마진율"] = safe_divide(merged["마진"], merged["품목별매출(VAT제외)"])

    return merged


df = build_dataset()
# -----------------------------------
# COST INPUT 초기화 (비용 계산 전에 먼저 실행)
# -----------------------------------
client = get_client()
sh = client.open_by_key(SHEET_ID)

if "logistics_table" not in st.session_state:
    logistics_dict, ad_dict = load_cost_input(sh)
    st.session_state["logistics_table"] = logistics_dict
    st.session_state["ad_cost_monthly"] = ad_dict

if "channel_cost" not in st.session_state:
    st.session_state["channel_cost"] = load_channel_cost(sh)

# -----------------------------------
# 사이드바 필터
# -----------------------------------

st.sidebar.header("📌 필터")

all_months = sort_month_cols(df["출고년월"].dropna().unique().tolist())
all_channel_groups = sorted(df["거래처분류"].dropna().unique().tolist())
all_items = sorted(df["내품상품명"].dropna().unique().tolist())

selected_months = st.sidebar.multiselect(
    "출고년월",
    options=all_months,
    default=all_months
)

selected_channel_groups = st.sidebar.multiselect(
    "채널",
    options=all_channel_groups,
    default=all_channel_groups
)

selected_items = st.sidebar.multiselect(
    "품목",
    options=all_items,
    default=all_items
)

logistics_cost_input = {}

filtered_df = df[
    (df["출고년월"].isin(selected_months)) &
    (df["거래처분류"].isin(selected_channel_groups)) &
    (df["내품상품명"].isin(selected_items))
].copy()

# -----------------------------------
# 비용 배분 계산 (필터 영향 없이 계산) [v4.1 추가]
# -----------------------------------

filtered_df["물류비"] = 0
filtered_df["광고비"] = 0

for m in selected_months:

    month_mask = filtered_df["출고년월"] == m

    # 국내 거래처의 전체 매출합 기준으로 안분
    domestic_mask_all = (df["출고년월"] == m) & (df["국내여부"] == "국내")
    month_domestic_sales = df.loc[domestic_mask_all, "품목별매출(VAT제외)"].sum()

    if month_domestic_sales > 0:
        # filtered_df 중 국내인 행만 물류비 배분
        domestic_month_mask = month_mask & (filtered_df["국내여부"] == "국내")
        ratio = (
            filtered_df.loc[domestic_month_mask, "품목별매출(VAT제외)"]
            / month_domestic_sales
        )
        filtered_df.loc[domestic_month_mask, "물류비"] = (
            ratio * st.session_state["logistics_table"].get(m, 0)
        )

for (category, month), ad_cost in st.session_state["ad_cost_monthly"].items():

    mask = (
        (filtered_df["품목군"] == category) &
        (filtered_df["출고년월"] == month)
    )

    month_sales = filtered_df.loc[mask, "품목별매출(VAT제외)"].sum()

    if month_sales > 0:

        ratio = (
            filtered_df.loc[mask, "품목별매출(VAT제외)"]
            / month_sales
        )

        filtered_df.loc[mask, "광고비"] = ratio * ad_cost

if filtered_df.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# -----------------------------------
# KPI
# -----------------------------------
st.markdown("""
<style>
[data-testid="stMetricValue"] {
    font-size: 28px;
}
</style>
""", unsafe_allow_html=True)

total_sales = filtered_df["품목별매출(VAT제외)"].sum()
total_qty = filtered_df["총내품출고수량"].sum()
total_margin = filtered_df["마진"].sum()
margin_rate = (total_margin / total_sales * 100) if total_sales != 0 else 0
# -----------------------------------
# 공헌이익 계산 (v3.4 추가 기능)
# -----------------------------------

filtered_df["공헌이익"] = (
    filtered_df["마진"]
    - filtered_df["물류비"]
    - filtered_df["광고비"]
)

filtered_df["공헌이익률"] = safe_divide(
    filtered_df["공헌이익"],
    filtered_df["품목별매출(VAT제외)"]
)

monthly_kpi = (
    filtered_df.groupby("출고년월", as_index=False)[["품목별매출(VAT제외)", "총내품출고수량", "마진"]]
    .sum()
)
monthly_kpi["출고년월_dt"] = pd.to_datetime(monthly_kpi["출고년월"] + "-01", errors="coerce")
monthly_kpi = monthly_kpi.sort_values("출고년월_dt")

if len(monthly_kpi) >= 2:
    current_sales = monthly_kpi.iloc[-1]["품목별매출(VAT제외)"]
    prev_sales = monthly_kpi.iloc[-2]["품목별매출(VAT제외)"]
    sales_mom = ((current_sales - prev_sales) / prev_sales * 100) if prev_sales != 0 else 0
else:
    sales_mom = 0

top_channel = (
    filtered_df.groupby("거래처분류")["품목별매출(VAT제외)"]
    .sum()
    .sort_values(ascending=False)
    .index[0]
)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("총 매출", f"{total_sales:,.0f} 원")
c2.metric("총 출고량", f"{total_qty:,.0f}")
c3.metric("총 마진(매출액 - 원가 - 채널별 수수료)", f"{total_margin:,.0f} 원")
c4.metric("마진율", f"{margin_rate:.2f}%")
c5.metric("Top 채널", top_channel, delta=f"{sales_mom:.2f}% MoM")

st.divider()


# -----------------------------------
# 탭 구성
# -----------------------------------

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 월별 추이",
    "🏪 채널 분석",
    "📦 제품 분석",
    "📊 공헌이익 분석",
    "📥 다운로드"
])

# -----------------------------------
# 월별 추이
# -----------------------------------
with tab1:
    st.subheader("📈 월별 추이")

    # ---------------------
    # 1. 데이터 집계
    # ---------------------
    monthly = (
        filtered_df.groupby("출고년월", as_index=False)[
            ["품목별매출(VAT제외)", "마진", "총내품출고수량"]
        ].sum()
    )

    monthly["출고년월_dt"] = pd.to_datetime(monthly["출고년월"] + "-01", errors="coerce")
    monthly = monthly.sort_values("출고년월_dt")

    # 👉 컬럼명 통일 (가독성)
    monthly = monthly.rename(columns={
        "품목별매출(VAT제외)": "매출액",
        "총내품출고수량": "출고량"
    })

    # ---------------------
    # 2. 라벨 표시 옵션
    # ---------------------
    show_label = st.checkbox("📊 라벨 표시", value=True)

    # ---------------------
    # 3. 매출 + 마진 그래프
    # ---------------------
    fig = go.Figure()

    # 매출 (Bar)
    fig.add_trace(go.Bar(
        x=monthly["출고년월"],
        y=monthly["매출액"],
        name="매출액",
        text=monthly["매출액"] if show_label else None,
        texttemplate='%{text:,.0f}',
        textposition='outside',
        cliponaxis=False
    ))

    # 마진 (Line)
    fig.add_trace(go.Scatter(
        x=monthly["출고년월"],
        y=monthly["마진"],
        name="마진",
        mode="lines+markers+text" if show_label else "lines+markers",
        line=dict(width=4),
        text=monthly["마진"] if show_label else None,
        texttemplate='%{text:,.0f}',
        textposition="top center"
    ))

    # Y축 여유 (잘림 방지)
    y_max = max(
        monthly["매출액"].max(),
        monthly["마진"].max()
    ) * 1.25

    fig.update_layout(
        yaxis=dict(range=[0, y_max]),
        legend=dict(orientation="h"),
        margin=dict(t=40)
    )

    # 라벨 스타일
    fig.update_traces(
        textfont=dict(size=12, color="black")
    )

    st.plotly_chart(fig, use_container_width=True)

    # ---------------------
    # 4. 출고량 그래프
    # ---------------------
    st.markdown("### 📦 월별 출고량")

    fig_qty = go.Figure()

    fig_qty.add_trace(go.Bar(
        x=monthly["출고년월"],
        y=monthly["출고량"],
        name="출고량",
        text=monthly["출고량"] if show_label else None,
        texttemplate='%{text:,.0f}',
        textposition='outside',
        cliponaxis=False
    ))

    y_max_qty = monthly["출고량"].max() * 1.25

    fig_qty.update_layout(
        yaxis=dict(range=[0, y_max_qty]),
        margin=dict(t=40)
    )

    fig_qty.update_traces(
        textfont=dict(size=12, color="black")
    )

    st.plotly_chart(fig_qty, use_container_width=True)

# -----------------------------------
# 채널 분석
# -----------------------------------

with tab2:
    st.subheader("🏪 채널 분석")

    channel_summary = (
        filtered_df.groupby("거래처분류", as_index=False)[
            [
            "총내품출고수량",
            "품목별매출(VAT제외)",
            "원가총액",
            "마진",
            "채널수수료",
            "물류비",
            "광고비",
            "공헌이익"
            ]
        ]
        .sum()
    )
    channel_summary["마진율"] = safe_divide(
        channel_summary["마진"],
        channel_summary["품목별매출(VAT제외)"]
    )
    channel_summary["수수료율(실적)"] = safe_divide(
        channel_summary["채널수수료"],
        channel_summary["품목별매출(VAT제외)"]
    )
    channel_summary["공헌이익률"] = safe_divide(
        channel_summary["공헌이익"],
        channel_summary["품목별매출(VAT제외)"]
    )
    channel_summary = channel_summary.sort_values("품목별매출(VAT제외)", ascending=False)

    ch1, ch2 = st.columns(2)

    with ch1:
        fig_channel_sales = px.bar(
            channel_summary.sort_values("품목별매출(VAT제외)", ascending=True),
            x="품목별매출(VAT제외)",
            y="거래처분류",
            orientation="h",
            title="채널별 매출"
        )
        fig_channel_sales.update_layout(xaxis_title="매출액", yaxis_title="채널")
        fig_channel_sales.update_xaxes(tickformat=",")
        st.plotly_chart(fig_channel_sales, use_container_width=True)

    with ch2:
        fig_channel_margin = px.bar(
            channel_summary.sort_values("마진", ascending=True),
            x="마진",
            y="거래처분류",
            orientation="h",
            title="채널별 마진"
        )
        fig_channel_margin.update_layout(xaxis_title="마진액", yaxis_title="채널")
        fig_channel_margin.update_xaxes(tickformat=",")
        st.plotly_chart(fig_channel_margin, use_container_width=True)

    st.subheader("📦 월별 채널별 출고량")
    channel_qty_pivot = pd.pivot_table(
        filtered_df,
        values="총내품출고수량",
        index="거래처분류",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    channel_qty_pivot = channel_qty_pivot.reindex(columns=sort_month_cols(channel_qty_pivot.columns.tolist()))
    st.dataframe(channel_qty_pivot.style.format("{:,.0f}"), use_container_width=True)

    st.subheader("💰 월별 채널별 매출액")
    channel_sales_pivot = pd.pivot_table(
        filtered_df,
        values="품목별매출(VAT제외)",
        index="거래처분류",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    channel_sales_pivot = channel_sales_pivot.reindex(columns=sort_month_cols(channel_sales_pivot.columns.tolist()))
    st.dataframe(channel_sales_pivot.style.format("{:,.0f}"), use_container_width=True)

# -----------------------------------
# 제품 분석
# -----------------------------------

with tab3:
    st.subheader("📦 제품 분석")

    # 🔥 반드시 맨 위에서 초기화
    if "ad_cost_map" not in st.session_state:
        st.session_state["ad_cost_map"] = {}

    product_summary = (
        filtered_df.groupby("내품상품명", as_index=False)[[
        "총내품출고수량",
        "품목별매출(VAT제외)",
        "마진",
        "물류비",
        "광고비",
        "공헌이익"
        ]].sum()
    )    
    product_summary["마진율"] = safe_divide(
        product_summary["마진"],
        product_summary["품목별매출(VAT제외)"]
    )
    product_summary["공헌이익률"] = safe_divide(
        product_summary["공헌이익"],
        product_summary["품목별매출(VAT제외)"]
    )
    product_summary = product_summary.sort_values("품목별매출(VAT제외)", ascending=False)

    top_n = st.selectbox("Top 제품 기준", [10, 20, 30, 50], index=1)

    top_products = product_summary.head(top_n).copy()

    pr1, pr2 = st.columns(2)

    with pr1:
        fig_product_sales = px.bar(
            top_products.sort_values("품목별매출(VAT제외)", ascending=True),
            x="품목별매출(VAT제외)",
            y="내품상품명",
            orientation="h",
            title=f"Top {top_n} 제품 매출"
        )
        fig_product_sales.update_layout(xaxis_title="매출액", yaxis_title="제품명")
        fig_product_sales.update_xaxes(tickformat=",")
        st.plotly_chart(fig_product_sales, use_container_width=True)

    with pr2:
        fig_product_margin = px.bar(
            top_products.sort_values("마진", ascending=True),
            x="마진",
            y="내품상품명",
            orientation="h",
            title=f"Top {top_n} 제품 마진"
        )
        fig_product_margin.update_layout(xaxis_title="마진액", yaxis_title="제품명")
        fig_product_margin.update_xaxes(tickformat=",")
        st.plotly_chart(fig_product_margin, use_container_width=True)

    st.subheader("📦 월별 제품별 출고량")
    product_qty_pivot = pd.pivot_table(
        filtered_df[filtered_df["내품상품명"].isin(top_products["내품상품명"])],
        values="총내품출고수량",
        index="내품상품명",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    product_qty_pivot = product_qty_pivot.reindex(columns=sort_month_cols(product_qty_pivot.columns.tolist()))
    st.dataframe(product_qty_pivot.style.format("{:,.0f}"), use_container_width=True)

    st.subheader("💰 월별 제품별 매출액")
    product_sales_pivot = pd.pivot_table(
        filtered_df[filtered_df["내품상품명"].isin(top_products["내품상품명"])],
        values="품목별매출(VAT제외)",
        index="내품상품명",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    product_sales_pivot = product_sales_pivot.reindex(columns=sort_month_cols(product_sales_pivot.columns.tolist()))
    st.dataframe(product_sales_pivot.style.format("{:,.0f}"), use_container_width=True)

# -----------------------------------
# 공헌이익 분석
# -----------------------------------
with tab4:
    st.subheader("📊 공헌이익 분석")
    st.caption("※ 물류비 / 광고비는 COST_INPUT 시트에서 수정됩니다")

    # -----------------------------
    # 1. 월별 물류비 테이블
    # -----------------------------
    st.markdown("### 🚚 월별 물류비")

    # -----------------------------
    # (선택) 화면 표시용 테이블
    # -----------------------------
    logistics_df = pd.DataFrame([{
        m: st.session_state["logistics_table"].get(m, 0)
        for m in all_months
    }], index=["월별 물류비"])

    st.dataframe(logistics_df.style.format("{:,.0f}"), use_container_width=True)

    # -----------------------------
    # 2. 제품 광고비 (월별 테이블)
    # -----------------------------
    st.markdown("### 📢 제품 광고비")

    # 광고비 dict → DataFrame 변환
    ad_data = []

    products = sorted(set([k[0] for k in st.session_state["ad_cost_monthly"].keys()]))

    for p in products:
        row = {"제품명": p}
        for m in all_months:
            row[m] = st.session_state["ad_cost_monthly"].get((p, m), 0)
        ad_data.append(row)

    ad_df = pd.DataFrame(ad_data)

    # 보기용 테이블 (수정 불가)
    month_cols = [c for c in ad_df.columns if c != "제품명"]
    fmt = {m: "{:,.0f}" for m in month_cols}
    st.dataframe(ad_df.style.format(fmt), use_container_width=True)

    # -----------------------------
    # 3-1. 채널별 후정산 비용 현황
    # -----------------------------
    st.markdown("### 💸 채널별 후정산 비용")

    if st.session_state["channel_cost"]:
        cc_rows = []
        for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
            cc_rows.append({"년월": ym, "거래처명": ch, "품목군": ig, "비용(VAT-)": amt})
        cc_df = pd.DataFrame(cc_rows).sort_values(["년월", "거래처명", "품목군"])
        st.dataframe(
            cc_df.style.format({"비용(VAT-)": "{:,.0f}"}),
            use_container_width=True
        )
    else:
        st.info("CHANNEL_COST 시트에 데이터가 없습니다.")
    st.markdown("### 📦 품목군별 공헌이익")

    temp_df = df.copy()  # 🔥 핵심: filtered_df ❌ → df ✅

    # -----------------------------
    # 1. 물류비 (전체 기준 배분)
    # -----------------------------
    temp_df["물류비"] = 0

    for m in all_months:

        mask = temp_df["출고년월"] == m

        # 국내 거래처 매출합 기준 안분
        domestic_mask = mask & (temp_df["국내여부"] == "국내")
        month_domestic_total = temp_df.loc[domestic_mask, "품목별매출(VAT제외)"].sum()

        if month_domestic_total > 0:
            ratio = temp_df.loc[domestic_mask, "품목별매출(VAT제외)"] / month_domestic_total
            temp_df.loc[domestic_mask, "물류비"] = (
                ratio * st.session_state["logistics_table"].get(m, 0)
            )

    # -----------------------------
    # 2. 광고비 (그대로 매핑, 분배 X)
    # -----------------------------
    temp_df["광고비"] = 0

    for (category, month), ad_cost in st.session_state["ad_cost_monthly"].items():

        mask = (
            (temp_df["품목군"] == category) &
            (temp_df["출고년월"] == month)
        )

        month_sales = temp_df.loc[mask, "품목별매출(VAT제외)"].sum()

        if month_sales > 0:

            ratio = (
                temp_df.loc[mask, "품목별매출(VAT제외)"]
                / month_sales
            )

            temp_df.loc[mask, "광고비"] = ratio * ad_cost

    # -----------------------------
    # 3. 공헌이익 계산 (비용 제외 중간값)
    # -----------------------------
    temp_df["공헌이익"] = (
        temp_df["마진"]
        - temp_df["물류비"]
        - temp_df["광고비"]
    )

    temp_df["공헌이익률"] = safe_divide(
        temp_df["공헌이익"],
        temp_df["품목별매출(VAT제외)"]
    )

    # -----------------------------
    # 🔥 필터는 여기서 적용
    # -----------------------------
    temp_df = temp_df[
        (temp_df["출고년월"].isin(selected_months)) &
        (temp_df["거래처분류"].isin(selected_channel_groups)) &
        (temp_df["내품상품명"].isin(selected_items))
    ].copy()


    product_contrib = (
        temp_df.groupby("품목군", as_index=False)[[
            "총내품출고수량",
            "품목별매출(VAT제외)",
            "원가총액",
            "채널수수료",
            "마진",
            "물류비",
            "광고비",
        ]].sum()
    )

    # -----------------------------
    # 채널별×품목군별 후정산 비용 → 품목군별 직접 매핑
    # CHANNEL_COST 키: (년월, 거래처명, 품목군) → 해당 품목군에 바로 합산
    # -----------------------------
    product_contrib["비용"] = 0.0

    for (year_month, channel_name, item_group), cost_amount in st.session_state["channel_cost"].items():
        if year_month not in selected_months:
            continue
        # 선택된 채널 필터 적용
        if channel_name not in selected_channel_groups:
            continue
        row_mask = product_contrib["품목군"] == item_group
        if row_mask.any():
            product_contrib.loc[row_mask, "비용"] += cost_amount

    # 최종 공헌이익 = 마진 - 물류비 - 광고비 - 비용
    product_contrib["공헌이익"] = (
        product_contrib["마진"]
        - product_contrib["물류비"]
        - product_contrib["광고비"]
        - product_contrib["비용"]
    )

    product_contrib["마진율"] = safe_divide(
        product_contrib["마진"],
        product_contrib["품목별매출(VAT제외)"]
    )

    product_contrib["공헌이익률"] = safe_divide(
        product_contrib["공헌이익"],
        product_contrib["품목별매출(VAT제외)"]
    )

    st.dataframe(
        product_contrib.style.format({
            "총내품출고수량": "{:,.0f}",
            "품목별매출(VAT제외)": "{:,.0f}",
            "원가총액": "{:,.0f}",
            "채널수수료": "{:,.0f}",
            "물류비": "{:,.0f}",
            "광고비": "{:,.0f}",
            "비용": "{:,.0f}",
            "마진": "{:,.0f}",
            "마진율": "{:.2%}",
            "공헌이익": "{:,.0f}",
            "공헌이익률": "{:.2%}",
        }),
        use_container_width=True
    )

    # -----------------------------
    # 4. 채널별 공헌이익
    # -----------------------------
    st.markdown("### 🏪 채널별 공헌이익")

    channel_contrib = (
        temp_df.groupby("거래처분류", as_index=False)[[
            "총내품출고수량",
            "품목별매출(VAT제외)",
            "원가총액",
            "채널수수료",
            "마진",
            "물류비",
            "광고비",
            "공헌이익"
        ]].sum()
    )

    # CHANNEL_COST 후정산 비용 반영 (년월, 거래처명, 품목군) 기준 합산
    channel_contrib["비용"] = 0.0

    for (year_month, channel_name, item_group), cost_amount in st.session_state["channel_cost"].items():
        if year_month not in selected_months:
            continue
        row_mask = channel_contrib["거래처분류"] == channel_name
        if row_mask.any():
            channel_contrib.loc[row_mask, "비용"] += cost_amount

    # 최종 공헌이익 = 마진 - 물류비 - 광고비 - 비용
    channel_contrib["공헌이익"] = (
        channel_contrib["마진"]
        - channel_contrib["물류비"]
        - channel_contrib["광고비"]
        - channel_contrib["비용"]
    )

    channel_contrib["수수료율"] = safe_divide(
        channel_contrib["채널수수료"],
        channel_contrib["품목별매출(VAT제외)"]
    )

    channel_contrib["마진율"] = safe_divide(
        channel_contrib["마진"],
        channel_contrib["품목별매출(VAT제외)"]
    )

    channel_contrib["공헌이익률"] = safe_divide(
        channel_contrib["공헌이익"],
        channel_contrib["품목별매출(VAT제외)"]
    )

    st.dataframe(
        channel_contrib.style.format({
            "총내품출고수량": "{:,.0f}",
            "품목별매출(VAT제외)": "{:,.0f}",
            "원가총액": "{:,.0f}",
            "물류비": "{:,.0f}",
            "광고비": "{:,.0f}",
            "비용": "{:,.0f}",
            "채널수수료": "{:,.0f}",
            "수수료율": "{:.2%}",
            "마진": "{:,.0f}",
            "마진율": "{:.2%}",
            "공헌이익": "{:,.0f}",
            "공헌이익률": "{:.2%}",
        }),
        use_container_width=True
    )
# -----------------------------------
# 다운로드
# -----------------------------------

with tab5:
    st.subheader("📥 다운로드")

    channel_qty_pivot = pd.pivot_table(
        filtered_df,
        values="총내품출고수량",
        index="거래처분류",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    channel_qty_pivot = channel_qty_pivot.reindex(columns=sort_month_cols(channel_qty_pivot.columns.tolist()))

    channel_sales_pivot = pd.pivot_table(
        filtered_df,
        values="품목별매출(VAT제외)",
        index="거래처분류",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    channel_sales_pivot = channel_sales_pivot.reindex(columns=sort_month_cols(channel_sales_pivot.columns.tolist()))

    product_qty_pivot = pd.pivot_table(
        filtered_df,
        values="총내품출고수량",
        index="내품상품명",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    product_qty_pivot = product_qty_pivot.reindex(columns=sort_month_cols(product_qty_pivot.columns.tolist()))

    product_sales_pivot = pd.pivot_table(
        filtered_df,
        values="품목별매출(VAT제외)",
        index="내품상품명",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )
    product_sales_pivot = product_sales_pivot.reindex(columns=sort_month_cols(product_sales_pivot.columns.tolist()))

    download_file = make_excel_file({
        "월별채널출고량": channel_qty_pivot,
        "월별채널매출": channel_sales_pivot,
        "월별제품출고량": product_qty_pivot,
        "월별제품매출": product_sales_pivot
    })

    st.download_button(
        label="📥 분석 결과 통합 엑셀 다운로드",
        data=download_file,
        file_name="Lingtea_Dashboard_v3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("### 포함 시트")
    st.write("- 월별 채널 출고량")
    st.write("- 월별 채널 매출액")
    st.write("- 월별 제품 출고량")
    st.write("- 월별 제품 매출액")

st.success("🚀 Lingtea Dashboard v5.2 Ready")
