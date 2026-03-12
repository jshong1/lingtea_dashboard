import streamlit as st
import pandas as pd
import numpy as np
import gspread
import io
from google.oauth2.service_account import Credentials
import plotly.express as px
from datetime import datetime
from dateutil.relativedelta import relativedelta

# -----------------------------------
# 기본 설정
# -----------------------------------

st.set_page_config(
    page_title="Lingtea Dashboard",
    layout="wide"
)

st.title("📊 Lingtea Dashboard")

# -----------------------------------
# 🔐 화이트리스트
# -----------------------------------

ALLOWED_USERS = [
    "js.hong1@lingtea.co.kr",
    "finance@company.com",
    "marketing@company.com"
]

user_email = "public_user"

# -----------------------------------
# 📥 Google Sheets 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_data():

    scope = [
        "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]

    import os

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

    client = gspread.authorize(creds)

    sheet = client.open_by_key(
        "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"
    ).worksheet("VIEW_TABLE")

    data = sheet.get_all_values()

    df = pd.DataFrame(data[1:], columns=data[0])

    df = df[[
        "출고일자",
        "출고년월",
        "거래처코드",
        "내품상품명",
        "총내품출고수량",
        "품목별매출(VAT제외)"
    ]]

    df["출고일자"] = pd.to_datetime(df["출고일자"], errors="coerce")

    df["총내품출고수량"] = (
        df["총내품출고수량"]
        .astype(str)
        .str.replace(",", "", regex=False)
    )
    df["총내품출고수량"] = pd.to_numeric(
        df["총내품출고수량"],
        errors="coerce"
    )

    df["품목별매출(VAT제외)"] = (
        df["품목별매출(VAT제외)"]
        .astype(str)
        .str.replace(",", "", regex=False)
    )
    df["품목별매출(VAT제외)"] = pd.to_numeric(
        df["품목별매출(VAT제외)"],
        errors="coerce"
    )

    today = datetime.today()
    one_year_ago = today - relativedelta(months=12)
    df = df[df["출고일자"] >= one_year_ago]

    df["출고년월"] = df["출고년월"].astype(str)

    return df


df = load_data()

# -----------------------------------
# 🎛 필터
# -----------------------------------

st.sidebar.header("📌 필터")

all_months = sorted(df["출고년월"].unique())
select_all_months = st.sidebar.checkbox("전체 출고년월 선택", value=True)

if select_all_months:
    selected_months = all_months
else:
    selected_months = st.sidebar.multiselect(
        "출고년월 선택",
        options=all_months
    )

all_channels = sorted(df["거래처코드"].unique())
select_all_channels = st.sidebar.checkbox("전체 채널 선택", value=True)

if select_all_channels:
    selected_channels = all_channels
else:
    selected_channels = st.sidebar.multiselect(
        "채널 선택",
        options=all_channels
    )

all_items = sorted(df["내품상품명"].unique())
select_all_items = st.sidebar.checkbox("전체 품목 선택", value=True)

if select_all_items:
    selected_items = all_items
else:
    selected_items = st.sidebar.multiselect(
        "품목 선택",
        options=all_items
    )

# -----------------------------------
# 📊 필터 적용
# -----------------------------------

filtered_df = df[
    (df["출고년월"].isin(selected_months)) &
    (df["거래처코드"].isin(selected_channels)) &
    (df["내품상품명"].isin(selected_items))
]

# -----------------------------------
# KPI
# -----------------------------------

total_qty = filtered_df["총내품출고수량"].sum()
total_sales = filtered_df["품목별매출(VAT제외)"].sum()

monthly_sales = (
    filtered_df.groupby("출고년월")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
)

monthly_sales["출고년월_dt"] = pd.to_datetime(monthly_sales["출고년월"] + "-01")
monthly_sales = monthly_sales.sort_values("출고년월_dt")

if len(monthly_sales) >= 2:
    current_month_sales = monthly_sales.iloc[-1]["품목별매출(VAT제외)"]
    previous_month_sales = monthly_sales.iloc[-2]["품목별매출(VAT제외)"]

    mom = ((current_month_sales - previous_month_sales) / previous_month_sales * 100) if previous_month_sales != 0 else 0
else:
    mom = 0

top_channel = (
    filtered_df.groupby("거래처코드")["품목별매출(VAT제외)"]
    .sum()
    .idxmax()
    if not filtered_df.empty else "-"
)

col1, col2, col3, col4 = st.columns(4)

col1.metric("총 출고수량", f"{total_qty:,.0f}")
col2.metric("총 매출액(VAT제외)", f"{total_sales:,.0f} 원")
col3.metric("전월 대비", f"{mom:.2f} %")
col4.metric("Top 채널", top_channel)

st.divider()

# -----------------------------------
# 📈 월별 추이 (행 분리)
# -----------------------------------

st.header("📈 월별 추이")

# 월별 매출
monthly_sales_chart = (
    filtered_df.groupby("출고년월")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
)

fig_sales = px.line(
    monthly_sales_chart,
    x="출고년월",
    y="품목별매출(VAT제외)",
    markers=True
)

st.subheader("월별 매출")
st.plotly_chart(fig_sales, use_container_width=True)


# 월별 출고량
monthly_qty_chart = (
    filtered_df.groupby("출고년월")["총내품출고수량"]
    .sum()
    .reset_index()
)

fig_qty = px.bar(
    monthly_qty_chart,
    x="출고년월",
    y="총내품출고수량"
)

st.subheader("월별 출고량")
st.plotly_chart(fig_qty, use_container_width=True)

# -----------------------------------
# 📊 채널 분석 (전체 채널)
# -----------------------------------

st.header("📊 채널 분석")

channel_sales = (
    filtered_df.groupby("거래처코드")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
    .sort_values(by="품목별매출(VAT제외)", ascending=False)
)

fig_channel = px.bar(
    channel_sales,
    x="거래처코드",
    y="품목별매출(VAT제외)"
)

st.subheader("채널별 매출")
st.plotly_chart(fig_channel, use_container_width=True)

# -----------------------------------
# 🏆 품목별 랭킹
# -----------------------------------

st.header("🏆 품목별 랭킹")

item_rank = (
    filtered_df.groupby("내품상품명")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
    .sort_values(by="품목별매출(VAT제외)", ascending=False)
)

item_rank = item_rank.rename(
    columns={"품목별매출(VAT제외)": "품목별매출액(VAT제외)"}
)

def highlight_top3(s):
    top3 = s.nlargest(3).values
    return [
        "background-color: #FFD700; font-weight: bold" if v in top3 else ""
        for v in s
    ]

styled_rank = (
    item_rank
    .style
    .format({"품목별매출액(VAT제외)": "{:,.0f}"})
    .apply(highlight_top3, subset=["품목별매출액(VAT제외)"])
)

st.dataframe(styled_rank, use_container_width=True)

# -----------------------------------
# 📊 거래처 × 월 피벗
# -----------------------------------

st.header("📊 거래처 × 월 피벗")

pivot_table = pd.pivot_table(
    filtered_df,
    values="품목별매출(VAT제외)",
    index="거래처코드",
    columns="출고년월",
    aggfunc="sum",
    fill_value=0
)

latest_month = sorted(selected_months)[-1]

if latest_month in pivot_table.columns:
    pivot_table = pivot_table.sort_values(
        by=latest_month,
        ascending=False
    )

st.dataframe(
    pivot_table.style.format("{:,.0f}"),
    use_container_width=True
)

# -----------------------------------
# 엑셀 다운로드
# -----------------------------------

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Pivot")
    return output.getvalue()

excel_data = convert_df_to_excel(pivot_table)

st.download_button(
    label="📥 피벗 엑셀 다운로드",
    data=excel_data,
    file_name="거래처_월_피벗.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("🚀 Lingtea Dashboard v3.2 Ready")
