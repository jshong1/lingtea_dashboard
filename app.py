import streamlit as st
import pandas as pd
import gspread
import io
from google.oauth2.service_account import Credentials
import plotly.express as px
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os

# -----------------------------------
# 기본 설정
# -----------------------------------

st.set_page_config(
    page_title="Lingtea Dashboard",
    layout="wide"
)

st.title("📊 Lingtea Dashboard")

# -----------------------------------
# Google Sheets 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_data():

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

    client = gspread.authorize(creds)

    sheet = client.open_by_key(
        "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"
    ).worksheet("VIEW_TABLE")

    data = sheet.get_all_values()

    df = pd.DataFrame(data[1:], columns=data[0])

    df = df[
        [
            "출고일자",
            "출고년월",
            "거래처코드",
            "내품상품명",
            "총내품출고수량",
            "품목별매출(VAT제외)"
        ]
    ]

    df["출고일자"] = pd.to_datetime(df["출고일자"], errors="coerce")

    df["총내품출고수량"] = (
        df["총내품출고수량"]
        .astype(str)
        .str.replace(",", "")
        .astype(float)
    )

    df["품목별매출(VAT제외)"] = (
        df["품목별매출(VAT제외)"]
        .astype(str)
        .str.replace(",", "")
        .astype(float)
    )

    # 최근 12개월
    today = datetime.today()
    one_year_ago = today - relativedelta(months=12)

    df = df[df["출고일자"] >= one_year_ago]

    return df


df = load_data()

# -----------------------------------
# MASTER DATA
# -----------------------------------

@st.cache_data(ttl=600)
def load_master():

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

    client = gspread.authorize(creds)

    item_ws = client.open_by_key(
        "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"
    ).worksheet("ITEM_MASTER")

    item_data = item_ws.get_all_values()
    item_df = pd.DataFrame(item_data[1:], columns=item_data[0])

    item_df["제품원가"] = pd.to_numeric(item_df["제품원가"], errors="coerce")

    cust_ws = client.open_by_key(
        "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"
    ).worksheet("CUSTOMER_MASTER")

    cust_data = cust_ws.get_all_values()
    cust_df = pd.DataFrame(cust_data[1:], columns=cust_data[0])

    cust_df["수수료율"] = pd.to_numeric(cust_df["수수료율"], errors="coerce")

    return item_df, cust_df


item_df, cust_df = load_master()

# -----------------------------------
# MASTER MERGE
# -----------------------------------

df = df.merge(
    item_df[["상품명", "제품원가"]],
    left_on="내품상품명",
    right_on="상품명",
    how="left"
)

df = df.merge(
    cust_df[["거래처명", "거래처분류", "수수료율"]],
    left_on="거래처코드",
    right_on="거래처명",
    how="left"
)

# -----------------------------------
# 마진 계산
# -----------------------------------

df["원가총액"] = df["총내품출고수량"] * df["제품원가"]

df["채널수수료"] = df["품목별매출(VAT제외)"] * df["수수료율"]

df["마진"] = (
    df["품목별매출(VAT제외)"]
    - df["원가총액"]
    - df["채널수수료"]
)

# -----------------------------------
# 필터
# -----------------------------------

st.sidebar.header("📌 필터")

all_months = sorted(df["출고년월"].unique())
selected_months = st.sidebar.multiselect(
    "출고년월",
    all_months,
    default=all_months
)

all_channels = sorted(df["거래처분류"].dropna().unique())
selected_channels = st.sidebar.multiselect(
    "채널",
    all_channels,
    default=all_channels
)

all_items = sorted(df["내품상품명"].unique())
selected_items = st.sidebar.multiselect(
    "품목",
    all_items,
    default=all_items
)

filtered_df = df[
    (df["출고년월"].isin(selected_months)) &
    (df["거래처분류"].isin(selected_channels)) &
    (df["내품상품명"].isin(selected_items))
]

# -----------------------------------
# KPI
# -----------------------------------

total_sales = filtered_df["품목별매출(VAT제외)"].sum()
total_qty = filtered_df["총내품출고수량"].sum()
total_margin = filtered_df["마진"].sum()

margin_rate = (
    total_margin / total_sales * 100
    if total_sales != 0 else 0
)

col1, col2, col3, col4 = st.columns(4)

col1.metric("총 매출", f"{total_sales:,.0f} 원")
col2.metric("총 출고량", f"{total_qty:,.0f}")
col3.metric("총 마진", f"{total_margin:,.0f} 원")
col4.metric("마진율", f"{margin_rate:.2f}%")

st.divider()

# -----------------------------------
# 월별 매출 추이
# -----------------------------------

st.subheader("📈 월별 매출 추이")

monthly_sales = (
    filtered_df
    .groupby("출고년월")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
)

fig = px.line(
    monthly_sales,
    x="출고년월",
    y="품목별매출(VAT제외)",
    markers=True
)

st.plotly_chart(fig, use_container_width=True)

# -----------------------------------
# 월별 채널 출고량
# -----------------------------------

st.subheader("📦 월별 채널 출고량")

channel_qty = pd.pivot_table(
    filtered_df,
    values="총내품출고수량",
    index="거래처분류",
    columns="출고년월",
    aggfunc="sum",
    fill_value=0
)

st.dataframe(channel_qty.style.format("{:,.0f}"))

# -----------------------------------
# 월별 채널 매출
# -----------------------------------

st.subheader("💰 월별 채널 매출")

channel_sales = pd.pivot_table(
    filtered_df,
    values="품목별매출(VAT제외)",
    index="거래처분류",
    columns="출고년월",
    aggfunc="sum",
    fill_value=0
)

st.dataframe(channel_sales.style.format("{:,.0f}"))

# -----------------------------------
# 월별 제품 출고량
# -----------------------------------

st.subheader("📦 월별 제품 출고량")

product_qty = pd.pivot_table(
    filtered_df,
    values="총내품출고수량",
    index="내품상품명",
    columns="출고년월",
    aggfunc="sum",
    fill_value=0
)

st.dataframe(product_qty.style.format("{:,.0f}"))

# -----------------------------------
# 월별 제품 매출
# -----------------------------------

st.subheader("💰 월별 제품 매출")

product_sales = pd.pivot_table(
    filtered_df,
    values="품목별매출(VAT제외)",
    index="내품상품명",
    columns="출고년월",
    aggfunc="sum",
    fill_value=0
)

st.dataframe(product_sales.style.format("{:,.0f}"))

# -----------------------------------
# 엑셀 다운로드
# -----------------------------------

def convert_excel(df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer)

    return output.getvalue()

excel_data = convert_excel(product_sales)

st.download_button(
    label="📥 제품 매출 엑셀 다운로드",
    data=excel_data,
    file_name="product_sales.xlsx"
)

st.success("🚀 Lingtea Dashboard Ready")
