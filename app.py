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

# -----------------------------------
# 기본 설정
# -----------------------------------

st.set_page_config(
    page_title="Lingtea Dashboard v3.3",
    layout="wide"
)

st.title("📊 Lingtea Dashboard v3.3")
st.caption("월별 채널/제품 분석 + 매출/출고량/마진 통합 대시보드")

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


# -----------------------------------
# MASTER 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_master():
    client = get_client()

    # ITEM_MASTER
    item_ws = client.open_by_key(SHEET_ID).worksheet("ITEM_MASTER")
    item_data = item_ws.get_all_values()
    item_df = pd.DataFrame(item_data[1:], columns=item_data[0]).copy()

    if "제품원가" not in item_df.columns:
        item_df["제품원가"] = 0

    item_df["제품원가"] = clean_numeric(item_df["제품원가"]).fillna(0)

    # CUSTOMER_MASTER
    cust_ws = client.open_by_key(SHEET_ID).worksheet("CUSTOMER_MASTER")
    cust_data = cust_ws.get_all_values()
    cust_df = pd.DataFrame(cust_data[1:], columns=cust_data[0]).copy()

    if "거래처분류" not in cust_df.columns:
        cust_df["거래처분류"] = np.nan

    if "수수료율" not in cust_df.columns:
        cust_df["수수료율"] = 0

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

    merged = df.merge(
        item_df[["상품명", "제품원가"]],
        left_on="내품상품명",
        right_on="상품명",
        how="left"
    )

    merged["거래처분류"] = merged["거래처코드"]

    merged = merged.merge(
        cust_df[["거래처분류", "수수료율"]].drop_duplicates(),
        on="거래처분류",
        how="left"
    )

    merged["제품원가"] = merged["제품원가"].fillna(0)
    merged["수수료율"] = merged["수수료율"].fillna(0)

    # 거래처분류 없는 경우 거래처코드로 대체
    merged["거래처분류"] = merged["거래처분류"].replace("", np.nan)
    merged["거래처분류"] = merged["거래처분류"].fillna(merged["거래처코드"])

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

# -----------------------------------
# 비용 입력 (v3.4 추가 기능)
# -----------------------------------

st.sidebar.markdown("---")
st.sidebar.header("💸 비용 입력")

# 월별 물류비 입력
st.sidebar.markdown("### 🚚 월별 물류비")

logistics_cost_input = {}

for m in all_months:
    logistics_cost_input[m] = st.sidebar.number_input(
        f"{m} 물류비",
        min_value=0,
        value=0,
        step=10000,
        key=f"logistics_{m}"
    )


filtered_df = df[
    (df["출고년월"].isin(selected_months)) &
    (df["거래처분류"].isin(selected_channel_groups)) &
    (df["내품상품명"].isin(selected_items))
].copy()

# -----------------------------------
# 비용 배분 계산 (v3.4 추가 기능)
# -----------------------------------

filtered_df["물류비"] = 0

for m in selected_months:

    month_mask = filtered_df["출고년월"] == m

    month_sales = filtered_df.loc[
        month_mask,
        "품목별매출(VAT제외)"
    ].sum()

    # 물류비 매출 비중 배분
    if month_sales > 0:

        ratio = (
            filtered_df.loc[month_mask, "품목별매출(VAT제외)"]
            / month_sales
        )

        filtered_df.loc[month_mask, "물류비"] = (
            ratio * logistics_cost_input.get(m, 0)
        )

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
c3.metric("총 마진", f"{total_margin:,.0f} 원")
c4.metric("마진율", f"{margin_rate:.2f}%")
c5.metric("Top 채널", top_channel, delta=f"{sales_mom:.2f}% MoM")

st.divider()

# -----------------------------------
# 월별 추이
# -----------------------------------

st.subheader("📈 월별 추이")

# ---------------------
# 월별 매출
# ---------------------

monthly_sales = (
    filtered_df.groupby("출고년월", as_index=False)["품목별매출(VAT제외)"]
    .sum()
)
monthly_sales["출고년월_dt"] = pd.to_datetime(monthly_sales["출고년월"] + "-01", errors="coerce")
monthly_sales = monthly_sales.sort_values("출고년월_dt")

fig_sales = px.bar(
    monthly_sales,
    x="출고년월",
    y="품목별매출(VAT제외)",
    title="월별 매출"
)

fig_sales.update_xaxes(type="category")
fig_sales.update_layout(xaxis_title="출고년월", yaxis_title="매출액")
fig_sales.update_yaxes(tickformat=",", ticksuffix=" 원")

st.plotly_chart(fig_sales, use_container_width=True)

# ---------------------
# 월별 출고량
# ---------------------

monthly_qty = (
    filtered_df.groupby("출고년월", as_index=False)["총내품출고수량"]
    .sum()
)
monthly_qty["출고년월_dt"] = pd.to_datetime(monthly_qty["출고년월"] + "-01", errors="coerce")
monthly_qty = monthly_qty.sort_values("출고년월_dt")

fig_qty = px.bar(
    monthly_qty,
    x="출고년월",
    y="총내품출고수량",
    title="월별 출고량"
)

fig_qty.update_xaxes(type="category")
fig_qty.update_layout(xaxis_title="출고년월", yaxis_title="출고량")
fig_qty.update_yaxes(tickformat=",")

st.plotly_chart(fig_qty, use_container_width=True)

# ---------------------
# 월별 마진
# ---------------------

monthly_margin = (
    filtered_df.groupby("출고년월", as_index=False)["마진"]
    .sum()
)
monthly_margin["출고년월_dt"] = pd.to_datetime(monthly_margin["출고년월"] + "-01", errors="coerce")
monthly_margin = monthly_margin.sort_values("출고년월_dt")

fig_margin = px.bar(
    monthly_margin,
    x="출고년월",
    y="마진",
    title="월별 마진"
)

fig_margin.update_xaxes(type="category")
fig_margin.update_layout(xaxis_title="출고년월", yaxis_title="마진액")
fig_margin.update_yaxes(tickformat=",", ticksuffix=" 원")

st.plotly_chart(fig_margin, use_container_width=True)

st.divider()

# -----------------------------------
# 탭 구성
# -----------------------------------

tab1, tab2, tab3 = st.tabs(["🏪 채널 분석", "📦 제품 분석", "📥 다운로드"])

# -----------------------------------
# 채널 분석
# -----------------------------------

with tab1:
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

    st.subheader("📋 채널 수익성 요약")
    channel_view = channel_summary.rename(columns={
        "거래처분류": "채널",
        "총내품출고수량": "출고량",
        "품목별매출(VAT제외)": "매출액",
        "원가총액": "원가",
        "채널수수료": "수수료액"
    })

    channel_view = channel_view[
        [
            "채널",
            "출고량",
            "매출액",
            "원가",
            "마진",
            "수수료액",
            "마진율",
            "수수료율(실적)",
            "공헌이익률",
        ]
    ]

    st.dataframe(
        channel_view.style.format({
            "출고량": "{:,.0f}",
            "매출액": "{:,.0f}",
            "원가": "{:,.0f}",
            "마진": "{:,.0f}",
            "수수료액": "{:,.0f}",
            "물류비": "{:,.0f}",
            "광고비": "{:,.0f}",
            "공헌이익": "{:,.0f}",
            "마진율": "{:.2%}",
            "수수료율(실적)": "{:.2%}",
            "공헌이익률": "{:.2%}"
        }),
        use_container_width=True
    )

# -----------------------------------
# 제품 분석
# -----------------------------------

with tab2:
    st.subheader("📦 제품 분석")

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

    st.subheader("📋 제품 광고비 입력")

    # 제품명만 추출 (중복 제거)
    product_view = top_products[["내품상품명"]].drop_duplicates().copy()
    
    product_view = product_view.rename(columns={
        "내품상품명": "제품명"
    })
    
    # 광고비 컬럼 생성
    product_view["광고비"] = 0
    
    # 광고비 입력 테이블
    edited_product = st.data_editor(
        product_view,
        use_container_width=True,
        column_config={
            "광고비": st.column_config.NumberColumn(
                "광고비",
                step=10000,
                format="%,d"
            ),
        },
        disabled=["제품명"]
    )
    # -----------------------------------
    # 광고비를 filtered_df에 반영 (수정)
    # -----------------------------------
    
    ad_map = edited_product[["제품명", "광고비"]].copy()
    ad_map["광고비"] = ad_map["광고비"].fillna(0)
    
    filtered_df = filtered_df.drop(columns=["광고비"], errors="ignore")
    
    final_product = (
        filtered_df.groupby("내품상품명", as_index=False)[
            ["총내품출고수량","품목별매출(VAT제외)","마진","물류비"]
        ]
        .sum()
    )
    
    # 광고비 붙이기
    final_product = final_product.merge(
        edited_product,
        left_on="내품상품명",
        right_on="제품명",
        how="left"
    )
    
    final_product["광고비"] = final_product["광고비"].fillna(0)
    
    # 공헌이익 계산
    final_product["공헌이익"] = (
        final_product["마진"]
        - final_product["물류비"]
        - final_product["광고비"]
    )
    
    final_product["공헌이익률"] = safe_divide(
        final_product["공헌이익"],
        final_product["품목별매출(VAT제외)"]
    )
    
    filtered_df = filtered_df.drop(columns=["제품명"], errors="ignore")
    
    st.subheader("📊 공헌이익 반영")

    # -----------------------------------
    # 공헌이익 반영 테이블 (수정)
    # -----------------------------------
    
    final_product = (
        filtered_df.groupby("내품상품명", as_index=False)[
            ["총내품출고수량", "품목별매출(VAT제외)", "마진", "물류비", "광고비", "공헌이익"]
        ]
        .sum()
    )
    
    final_product["마진율"] = safe_divide(
        final_product["마진"],
        final_product["품목별매출(VAT제외)"]
    )
    
    final_product["공헌이익률"] = safe_divide(
        final_product["공헌이익"],
        final_product["품목별매출(VAT제외)"]
    )
    
    final_product = final_product.rename(columns={
        "내품상품명": "제품명",
        "총내품출고수량": "출고량",
        "품목별매출(VAT제외)": "매출액"
    })
    
    # 숫자 정리
    for col in ["출고량","매출액","마진","물류비","광고비","공헌이익"]:
        final_product[col] = final_product[col].round(0).astype(int)
    
    st.dataframe(
        final_product.style.format({
            "출고량": "{:,.0f}",
            "매출액": "{:,.0f}",
            "마진": "{:,.0f}",
            "물류비": "{:,.0f}",
            "광고비": "{:,.0f}",
            "공헌이익": "{:,.0f}",
            "마진율": "{:.2%}",
            "공헌이익률": "{:.2%}"
        }),
        use_container_width=True
    )

    # -----------------------------------
    # 공헌이익 재계산 (광고비 반영 후)
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

# -----------------------------------
# 다운로드
# -----------------------------------

with tab3:
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

st.success("🚀 Lingtea Dashboard v3.3 Ready")
