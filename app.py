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
# 🔐 화이트리스트 설정
# -----------------------------------

ALLOWED_USERS = [
    "js.hong1@lingtea.co.kr",
    "finance@company.com",
    "marketing@company.com"
]

# -----------------------------------
# 🔐 로그인 처리 (로컬 + 클라우드 호환)
# -----------------------------------

is_cloud = hasattr(st, "user") and hasattr(st.user, "email")

if is_cloud:
    # Streamlit Cloud 환경
    try:
        if not st.user.email:
            st.login()

        user_email = st.user.email

        if user_email not in ALLOWED_USERS:
            st.error("접근 권한이 없습니다.")
            st.stop()

        st.success(f"로그인 사용자: {user_email}")

    except:
        st.error("로그인이 필요합니다.")
        st.stop()

else:
    # 로컬 실행 환경
    st.warning("로컬 실행 모드 (로그인 비활성화)")
    user_email = "local_test_user"

# -----------------------------------
# 📥 Google Sheets 로드 (Cloud + 로컬 겸용)
# -----------------------------------

@st.cache_data(ttl=600)
def load_data():

    scope = [
        "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]

    import os

    # 🔥 로컬 실행이면 JSON 파일 사용
    if os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file(
            "service_account.json",
            scopes=scope
        )

    # 🔥 Streamlit Cloud 실행이면 Secrets 사용
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

    # 최근 12개월 컷
    today = datetime.today()
    one_year_ago = today - relativedelta(months=12)
    df = df[df["출고일자"] >= one_year_ago]

    df["출고년월"] = df["출고년월"].astype(str)

    return df


df = load_data()

# -----------------------------------
# 🎛 필터 영역
# -----------------------------------

st.sidebar.header("📌 필터")

# -------------------------
# 출고년월
# -------------------------

all_months = sorted(df["출고년월"].unique())
select_all_months = st.sidebar.checkbox("전체 출고년월 선택", value=True)

if select_all_months:
    selected_months = all_months
else:
    selected_months = st.sidebar.multiselect(
        "출고년월 선택",
        options=all_months
    )

# -------------------------
# 채널
# -------------------------

all_channels = sorted(df["거래처코드"].unique())
select_all_channels = st.sidebar.checkbox("전체 채널 선택", value=True)

if select_all_channels:
    selected_channels = all_channels
else:
    selected_channels = st.sidebar.multiselect(
        "채널 선택",
        options=all_channels
    )

# -------------------------
# 품목
# -------------------------

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
# 📊 필터 적용 데이터
# -----------------------------------

filtered_df = df[
    (df["출고년월"].isin(selected_months)) &
    (df["거래처코드"].isin(selected_channels)) &
    (df["내품상품명"].isin(selected_items))
]

# -----------------------------------
# 💎 KPI 계산 (최근 선택 월 기준 전월 대비)
# -----------------------------------

total_qty = filtered_df["총내품출고수량"].sum()
total_sales = filtered_df["품목별매출(VAT제외)"].sum()

monthly_sales = (
    filtered_df.groupby("출고년월")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
)

# 날짜 정렬 안정화
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

# -----------------------------------
# 📊 KPI 카드
# -----------------------------------

col1, col2, col3, col4 = st.columns(4)

col1.metric("총 출고수량", f"{total_qty:,.0f}")
col2.metric("총 매출액(VAT제외)", f"{total_sales:,.0f} 원")
col3.metric("전월 대비 (선택기준)", f"{mom:.2f} %")
col4.metric("Top 채널", top_channel)

st.divider()

# -----------------------------------
# 📈 월별 추이
# -----------------------------------
st.subheader("📈 월별 매출 추이")

monthly_trend = (
    filtered_df
    .groupby("출고년월")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
)

monthly_trend["출고년월_dt"] = pd.to_datetime(
    monthly_trend["출고년월"] + "-01"
)

monthly_trend = monthly_trend.sort_values("출고년월_dt")

# 🔥 억 단위 변환
monthly_trend["매출(억원)"] = monthly_trend["품목별매출(VAT제외)"] / 100000000

fig_trend = px.line(
    monthly_trend,
    x="출고년월_dt",
    y="매출(억원)",
    markers=True
)

fig_trend.update_layout(
    xaxis_title="출고년월",
    yaxis_title="매출액 (억원)"
)

fig_trend.update_xaxes(tickformat="%Y-%m")

st.plotly_chart(fig_trend, use_container_width=True)

# -----------------------------------
# 📊 채널별 매출
# -----------------------------------

st.subheader("📊 채널별 매출")

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

st.plotly_chart(fig_channel, use_container_width=True)

# -----------------------------------
# 🏆 품목별 랭킹
# -----------------------------------
st.subheader("🏆 품목별 랭킹")

item_rank = (
    filtered_df
    .groupby("내품상품명")["품목별매출(VAT제외)"]
    .sum()
    .reset_index()
    .sort_values(by="품목별매출(VAT제외)", ascending=False)
)

item_rank = item_rank.rename(
    columns={"품목별매출(VAT제외)": "품목별매출액(VAT제외)"}
)

# 🔥 상위 3개 강조 함수
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
st.subheader("📊 거래처 × 월 피벗")

pivot_table = pd.pivot_table(
    filtered_df,
    values="품목별매출(VAT제외)",
    index="거래처코드",
    columns="출고년월",
    aggfunc="sum",
    fill_value=0
)

# 🔥 최신 선택 월 기준 정렬
latest_month = sorted(selected_months)[-1]

if latest_month in pivot_table.columns:
    pivot_table = pivot_table.sort_values(
        by=latest_month,
        ascending=False
    )

# 🔥 상위 3개 강조
def highlight_top3_pivot(df):
    if latest_month in df.columns:
        top3_vals = df[latest_month].nlargest(3).values
        return df.style.apply(
            lambda col: [
                "background-color: #C6F6D5; font-weight: bold"
                if (col.name == latest_month and v in top3_vals)
                else ""
                for v in col
            ]
        )
    return df.style

styled_pivot = (
    highlight_top3_pivot(pivot_table)
    .format("{:,.0f}")
)

st.dataframe(styled_pivot, use_container_width=True)

# -----------------------------------
# 엑셀 다운로드 버튼 추가
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


st.success("🚀 Lingtea Dashboard Ready")
