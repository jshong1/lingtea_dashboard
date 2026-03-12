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
    page_title="Lingtea Dashboard",
    layout="wide"
)

st.title("📊 Lingtea Dashboard")
st.caption("매출 · 출고 · 채널 · 제품 수익성 대시보드")

SHEET_ID = "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"

# -----------------------------------
# 유틸 함수
# -----------------------------------

def get_gspread_client():
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


def to_numeric_clean(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False).str.strip(),
        errors="coerce"
    )


def format_won(value: float) -> str:
    if pd.isna(value):
        return "-"
    return f"{value:,.0f} 원"


def format_uk(value: float) -> str:
    if pd.isna(value):
        return "-"
    return f"{value / 100000000:.1f} 억"


def format_pct(value: float) -> str:
    if pd.isna(value):
        return "-"
    return f"{value:.1%}"


# -----------------------------------
# VIEW_TABLE 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_view_data():
    client = get_gspread_client()

    ws = client.open_by_key(SHEET_ID).worksheet("VIEW_TABLE")
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])

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
    df["출고년월"] = df["출고년월"].astype(str)

    df["총내품출고수량"] = to_numeric_clean(df["총내품출고수량"])
    df["품목별매출(VAT제외)"] = to_numeric_clean(df["품목별매출(VAT제외)"])

    # 최근 12개월 컷
    today = datetime.today()
    one_year_ago = today - relativedelta(months=12)
    df = df[df["출고일자"] >= one_year_ago].copy()

    return df


# -----------------------------------
# MASTER 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_master_data():
    client = get_gspread_client()

    # ITEM_MASTER
    item_ws = client.open_by_key(SHEET_ID).worksheet("ITEM_MASTER")
    item_data = item_ws.get_all_values()
    item_df = pd.DataFrame(item_data[1:], columns=item_data[0]).copy()

    if "제품원가" not in item_df.columns:
        item_df["제품원가"] = np.nan

    item_df["제품원가"] = to_numeric_clean(item_df["제품원가"])

    # CUSTOMER_MASTER
    cust_ws = client.open_by_key(SHEET_ID).worksheet("CUSTOMER_MASTER")
    cust_data = cust_ws.get_all_values()
    cust_df = pd.DataFrame(cust_data[1:], columns=cust_data[0]).copy()

    # 컬럼 없을 때 방어
    if "거래처분류" not in cust_df.columns:
        cust_df["거래처분류"] = "미분류"
    if "수수료율" not in cust_df.columns:
        cust_df["수수료율"] = 0

    cust_df["수수료율"] = to_numeric_clean(cust_df["수수료율"])

    # 12, 3.5 같이 입력된 경우 %로 간주하여 0.12, 0.035 변환
    cust_df["수수료율"] = np.where(
        cust_df["수수료율"] > 1,
        cust_df["수수료율"] / 100,
        cust_df["수수료율"]
    )

    cust_df["거래처분류"] = cust_df["거래처분류"].replace("", "미분류").fillna("미분류")

    return item_df, cust_df


# -----------------------------------
# 데이터 결합
# -----------------------------------

@st.cache_data(ttl=600)
def build_dataset():
    df = load_view_data()
    item_df, cust_df = load_master_data()

    merged = df.merge(
        item_df[["상품명", "제품원가"]],
        left_on="내품상품명",
        right_on="상품명",
        how="left"
    )

    merged = merged.merge(
        cust_df[["거래처명", "거래처분류", "수수료율"]],
        left_on="거래처코드",
        right_on="거래처명",
        how="left"
    )

    merged["거래처분류"] = merged["거래처분류"].fillna("미분류")
    merged["수수료율"] = merged["수수료율"].fillna(0)
    merged["제품원가"] = merged["제품원가"].fillna(0)

    merged["원가총액"] = merged["총내품출고수량"] * merged["제품원가"]
    merged["채널수수료"] = merged["품목별매출(VAT제외)"] * merged["수수료율"]
    merged["마진"] = (
        merged["품목별매출(VAT제외)"]
        - merged["원가총액"]
        - merged["채널수수료"]
    )
    merged["마진율"] = np.where(
        merged["품목별매출(VAT제외)"] != 0,
        merged["마진"] / merged["품목별매출(VAT제외)"],
        0
    )

    return merged


df = build_dataset()

# -----------------------------------
# 사이드바 필터
# -----------------------------------

st.sidebar.header("📌 필터")

all_months = sorted(df["출고년월"].dropna().unique())
all_channel_groups = sorted(df["거래처분류"].dropna().unique())
all_customers = sorted(df["거래처코드"].dropna().unique())
all_items = sorted(df["내품상품명"].dropna().unique())

selected_months = st.sidebar.multiselect(
    "출고년월",
    options=all_months,
    default=all_months
)

selected_channel_groups = st.sidebar.multiselect(
    "채널구분",
    options=all_channel_groups,
    default=all_channel_groups
)

selected_customers = st.sidebar.multiselect(
    "거래처명",
    options=all_customers,
    default=all_customers
)

selected_items = st.sidebar.multiselect(
    "품목",
    options=all_items,
    default=all_items
)

filtered_df = df[
    (df["출고년월"].isin(selected_months)) &
    (df["거래처분류"].isin(selected_channel_groups)) &
    (df["거래처코드"].isin(selected_customers)) &
    (df["내품상품명"].isin(selected_items))
].copy()

# -----------------------------------
# 빈 데이터 방어
# -----------------------------------

if filtered_df.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# -----------------------------------
# KPI
# -----------------------------------

total_qty = filtered_df["총내품출고수량"].sum()
total_sales = filtered_df["품목별매출(VAT제외)"].sum()
total_margin = filtered_df["마진"].sum()
margin_rate = total_margin / total_sales if total_sales != 0 else 0

monthly_summary = (
    filtered_df.groupby("출고년월", as_index=False)[["품목별매출(VAT제외)", "마진", "총내품출고수량"]]
    .sum()
)

monthly_summary["출고년월_dt"] = pd.to_datetime(monthly_summary["출고년월"] + "-01", errors="coerce")
monthly_summary = monthly_summary.sort_values("출고년월_dt")

if len(monthly_summary) >= 2:
    current_sales = monthly_summary.iloc[-1]["품목별매출(VAT제외)"]
    previous_sales = monthly_summary.iloc[-2]["품목별매출(VAT제외)"]
    mom_sales = ((current_sales - previous_sales) / previous_sales) if previous_sales != 0 else 0
else:
    mom_sales = 0

top_channel_group = (
    filtered_df.groupby("거래처분류")["품목별매출(VAT제외)"].sum().sort_values(ascending=False).index[0]
    if not filtered_df.empty else "-"
)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("총 매출", format_uk(total_sales))
k2.metric("총 출고량", f"{total_qty:,.0f}")
k3.metric("총 마진", format_uk(total_margin))
k4.metric("마진율", format_pct(margin_rate))
k5.metric("Top 채널구분", top_channel_group, delta=f"{mom_sales:.1%} MoM")

st.divider()

# -----------------------------------
# 탭
# -----------------------------------

tab1, tab2, tab3 = st.tabs(["📈 종합", "📦 제품 분석", "🏪 채널 분석"])

# -----------------------------------
# 종합 탭
# -----------------------------------

with tab1:
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("📈 월별 매출 추이")

        monthly_sales = monthly_summary.copy()
        monthly_sales["매출(억원)"] = monthly_sales["품목별매출(VAT제외)"] / 100000000

        fig_sales = px.line(
            monthly_sales,
            x="출고년월_dt",
            y="매출(억원)",
            markers=True
        )
        fig_sales.update_layout(
            xaxis_title="출고년월",
            yaxis_title="매출(억원)"
        )
        fig_sales.update_xaxes(tickformat="%Y-%m")
        st.plotly_chart(fig_sales, use_container_width=True)

    with c2:
        st.subheader("💰 월별 마진 추이")

        monthly_margin = monthly_summary.copy()
        monthly_margin["마진(억원)"] = monthly_margin["마진"] / 100000000

        fig_margin_month = px.line(
            monthly_margin,
            x="출고년월_dt",
            y="마진(억원)",
            markers=True
        )
        fig_margin_month.update_layout(
            xaxis_title="출고년월",
            yaxis_title="마진(억원)"
        )
        fig_margin_month.update_xaxes(tickformat="%Y-%m")
        st.plotly_chart(fig_margin_month, use_container_width=True)

    st.subheader("📦 월별 출고량")

    monthly_qty = monthly_summary.copy()

    fig_qty = px.bar(
        monthly_qty,
        x="출고년월_dt",
        y="총내품출고수량"
    )
    fig_qty.update_layout(
        xaxis_title="출고년월",
        yaxis_title="출고수량"
    )
    fig_qty.update_xaxes(tickformat="%Y-%m")
    st.plotly_chart(fig_qty, use_container_width=True)

    st.subheader("📋 월별 요약 테이블")

    monthly_table = monthly_summary[[
        "출고년월", "총내품출고수량", "품목별매출(VAT제외)", "마진"
    ]].copy()

    monthly_table["마진율"] = np.where(
        monthly_table["품목별매출(VAT제외)"] != 0,
        monthly_table["마진"] / monthly_table["품목별매출(VAT제외)"],
        0
    )

    monthly_table = monthly_table.rename(columns={
        "총내품출고수량": "출고수량",
        "품목별매출(VAT제외)": "매출액"
    })

    st.dataframe(
        monthly_table.style.format({
            "출고수량": "{:,.0f}",
            "매출액": "{:,.0f}",
            "마진": "{:,.0f}",
            "마진율": "{:.1%}"
        }),
        use_container_width=True
    )

# -----------------------------------
# 제품 분석 탭
# -----------------------------------

with tab2:
    product_summary = (
        filtered_df.groupby("내품상품명", as_index=False)[["총내품출고수량", "품목별매출(VAT제외)", "마진"]]
        .sum()
    )
    product_summary["마진율"] = np.where(
        product_summary["품목별매출(VAT제외)"] != 0,
        product_summary["마진"] / product_summary["품목별매출(VAT제외)"],
        0
    )
    product_summary = product_summary.sort_values("품목별매출(VAT제외)", ascending=False)

    top10_products = product_summary.head(10).copy()

    p1, p2 = st.columns(2)

    with p1:
        st.subheader("🏆 Top10 제품 매출")

        fig_top_product = px.bar(
            top10_products.sort_values("품목별매출(VAT제외)", ascending=True),
            x="품목별매출(VAT제외)",
            y="내품상품명",
            orientation="h"
        )
        fig_top_product.update_layout(
            xaxis_title="매출액",
            yaxis_title="제품명"
        )
        st.plotly_chart(fig_top_product, use_container_width=True)

    with p2:
        st.subheader("💰 Top10 제품 마진")

        fig_top_product_margin = px.bar(
            top10_products.sort_values("마진", ascending=True),
            x="마진",
            y="내품상품명",
            orientation="h"
        )
        fig_top_product_margin.update_layout(
            xaxis_title="마진액",
            yaxis_title="제품명"
        )
        st.plotly_chart(fig_top_product_margin, use_container_width=True)

    st.subheader("🔥 Top10 제품 × 월 매출 히트맵")

    top10_names = top10_products["내품상품명"].tolist()
    heatmap_source = filtered_df[filtered_df["내품상품명"].isin(top10_names)].copy()

    heatmap_df = pd.pivot_table(
        heatmap_source,
        values="품목별매출(VAT제외)",
        index="내품상품명",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )

    # 월 컬럼 정렬
    sorted_heatmap_cols = sorted(heatmap_df.columns.tolist())
    heatmap_df = heatmap_df[sorted_heatmap_cols]

    # 억원 단위로 변환
    heatmap_df_uk = heatmap_df / 100000000

    fig_heatmap = px.imshow(
        heatmap_df_uk,
        text_auto=".1f",
        aspect="auto",
        labels=dict(color="매출(억원)")
    )
    fig_heatmap.update_layout(
        xaxis_title="출고년월",
        yaxis_title="제품명"
    )
    st.plotly_chart(fig_heatmap, use_container_width=True)

    st.subheader("📋 제품 수익성 요약")

    product_view = top10_products.rename(columns={
        "내품상품명": "제품명",
        "총내품출고수량": "출고수량",
        "품목별매출(VAT제외)": "매출액"
    })

    st.dataframe(
        product_view.style.format({
            "출고수량": "{:,.0f}",
            "매출액": "{:,.0f}",
            "마진": "{:,.0f}",
            "마진율": "{:.1%}"
        }),
        use_container_width=True
    )

# -----------------------------------
# 채널 분석 탭
# -----------------------------------

with tab3:
    channel_summary = (
        filtered_df.groupby("거래처분류", as_index=False)[["품목별매출(VAT제외)", "마진", "채널수수료"]]
        .sum()
    )
    channel_summary["마진율"] = np.where(
        channel_summary["품목별매출(VAT제외)"] != 0,
        channel_summary["마진"] / channel_summary["품목별매출(VAT제외)"],
        0
    )
    channel_summary["수수료율_실적"] = np.where(
        channel_summary["품목별매출(VAT제외)"] != 0,
        channel_summary["채널수수료"] / channel_summary["품목별매출(VAT제외)"],
        0
    )
    channel_summary = channel_summary.sort_values("품목별매출(VAT제외)", ascending=False)

    ch1, ch2 = st.columns(2)

    with ch1:
        st.subheader("📊 채널별 매출")

        fig_channel_sales = px.bar(
            channel_summary.sort_values("품목별매출(VAT제외)", ascending=True),
            x="품목별매출(VAT제외)",
            y="거래처분류",
            orientation="h"
        )
        fig_channel_sales.update_layout(
            xaxis_title="매출액",
            yaxis_title="채널구분"
        )
        st.plotly_chart(fig_channel_sales, use_container_width=True)

    with ch2:
        st.subheader("💹 채널별 마진율")

        fig_channel_margin_rate = px.bar(
            channel_summary.sort_values("마진율", ascending=True),
            x="마진율",
            y="거래처분류",
            orientation="h",
            text_auto=".1%"
        )
        fig_channel_margin_rate.update_layout(
            xaxis_title="마진율",
            yaxis_title="채널구분"
        )
        st.plotly_chart(fig_channel_margin_rate, use_container_width=True)

    st.subheader("📋 채널 수익성 요약")

    channel_view = channel_summary.rename(columns={
        "거래처분류": "채널구분",
        "품목별매출(VAT제외)": "매출액",
        "채널수수료": "수수료액"
    })

    st.dataframe(
        channel_view.style.format({
            "매출액": "{:,.0f}",
            "마진": "{:,.0f}",
            "수수료액": "{:,.0f}",
            "마진율": "{:.1%}",
            "수수료율_실적": "{:.1%}"
        }),
        use_container_width=True
    )

    st.subheader("📊 거래처 × 월 매출 피벗")

    customer_pivot = pd.pivot_table(
        filtered_df,
        values="품목별매출(VAT제외)",
        index="거래처코드",
        columns="출고년월",
        aggfunc="sum",
        fill_value=0
    )

    sorted_cols = sorted(customer_pivot.columns.tolist())
    customer_pivot = customer_pivot[sorted_cols]

    if len(sorted_cols) > 0:
        latest_month = sorted_cols[-1]
        customer_pivot = customer_pivot.sort_values(by=latest_month, ascending=False)

    st.dataframe(
        customer_pivot.style.format("{:,.0f}"),
        use_container_width=True
    )

    def convert_df_to_excel(df_to_save):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_to_save.to_excel(writer, sheet_name="거래처월피벗")
        return output.getvalue()

    excel_data = convert_df_to_excel(customer_pivot)

    st.download_button(
        label="📥 거래처 × 월 피벗 다운로드",
        data=excel_data,
        file_name="거래처_월_피벗.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.success("🚀 Lingtea Dashboard Ready")
