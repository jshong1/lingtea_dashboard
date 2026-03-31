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

# -----------------------------------
# 기본 설정
# -----------------------------------

st.set_page_config(page_title="Lingtea Dashboard", layout="wide")
st.title("📊 Lingtea Dashboard")
st.caption("월별 채널/제품 분석 + 매출/출고량/공헌이익 통합 대시보드")

SHEET_ID = "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"

# -----------------------------------
# 공통 유틸
# -----------------------------------

def get_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    if os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=scope)
    else:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)


def load_cost_input(sh):
    ws   = sh.worksheet("COST_INPUT")
    data = ws.get_all_values()
    if len(data) < 3:
        return {}, {}

    months = data[0][1:]

    # 물류비
    logistics_dict = {}
    for i, m in enumerate(months):
        val = str(data[1][i + 1]).replace(",", "").strip()
        try:
            logistics_dict[m] = float(val) if val else 0
        except:
            logistics_dict[m] = 0

    # 광고비 (품목군 × 월)
    ad_dict = {}
    for row in data[4:]:
        category = str(row[0]).strip()
        if not category:
            continue
        for i, m in enumerate(months):
            val = str(row[i + 1]).replace(",", "").strip()
            try:
                ad_dict[(category, m)] = float(val) if val and val.lower() != "nan" else 0
            except:
                ad_dict[(category, m)] = 0

    return logistics_dict, ad_dict


def load_channel_cost(sh):
    """
    CHANNEL_COST: A=년월, B=거래처명, C=품목군, D=비용항목, E=금액(VAT-)
    반환: dict { (년월, 거래처명, 품목군): 금액합계 }
    """
    try:
        ws   = sh.worksheet("CHANNEL_COST")
        data = ws.get_all_values()
    except Exception:
        return {}
    if len(data) < 2:
        return {}

    result = {}
    for row in data[1:]:
        if len(row) < 5:
            continue
        ym, ch, ig = str(row[0]).strip(), str(row[1]).strip(), str(row[2]).strip()
        if not (ym and ch and ig):
            continue
        try:
            amt = float(str(row[4]).replace(",", "").strip())
        except:
            amt = 0
        key = (ym, ch, ig)
        result[key] = result.get(key, 0) + amt
    return result


def clean_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False)
              .str.replace("%", "", regex=False).str.strip(),
        errors="coerce"
    )


def sort_month_cols(cols):
    return sorted(cols, key=lambda x: pd.to_datetime(f"{x}-01", errors="coerce"))


def safe_divide(a, b):
    return np.where(b != 0, a / b, 0)


def make_excel_file(sheet_dict: dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheet_dict.items():
            df.to_excel(writer, sheet_name=name[:31])
    return output.getvalue()


def add_total_row(pivot: pd.DataFrame) -> pd.DataFrame:
    """최상단에 합계 행 추가"""
    total     = pivot.sum(numeric_only=True)
    total_row = pd.DataFrame([total], index=["[합계]"])
    return pd.concat([total_row, pivot])


def sort_pivot_by_last_month(pivot: pd.DataFrame, month_cols: list) -> pd.DataFrame:
    """가장 마지막 월 기준 내림차순 정렬 (합계 행 보존)"""
    if not month_cols or month_cols[-1] not in pivot.columns:
        return pivot
    last_m    = month_cols[-1]
    total_part = pivot[pivot.index == "[합계]"]
    data_part  = pivot[pivot.index != "[합계]"].sort_values(last_m, ascending=False)
    return pd.concat([total_part, data_part])


# -----------------------------------
# 데이터 로드
# -----------------------------------

@st.cache_data(ttl=600)
def load_view_table():
    client = get_client()
    ws     = client.open_by_key(SHEET_ID).worksheet("VIEW_TABLE")
    data   = ws.get_all_values()
    df     = pd.DataFrame(data[1:], columns=data[0]).copy()

    use_cols = ["출고일자", "출고년월", "거래처코드", "내품상품명",
                "총내품출고수량", "품목별매출(VAT제외)"]
    df = df[use_cols].copy()
    df["출고일자"]          = pd.to_datetime(df["출고일자"], errors="coerce")
    df["출고년월"]          = df["출고년월"].astype(str).str.strip()
    df["총내품출고수량"]      = clean_numeric(df["총내품출고수량"])
    df["품목별매출(VAT제외)"] = clean_numeric(df["품목별매출(VAT제외)"])

    cutoff = datetime.today() - relativedelta(months=12)
    return df[df["출고일자"] >= cutoff].copy()


@st.cache_data(ttl=600)
def load_cost_master():
    client = get_client()
    ws     = client.open_by_key(SHEET_ID).worksheet("COST_MASTER")
    data   = ws.get_all_values()
    result = {}
    for row in data[1:]:
        if len(row) < 4:
            continue
        ym, name = str(row[0]).strip(), str(row[2]).strip()
        if not (ym and name):
            continue
        try:
            result[(ym, name)] = float(str(row[3]).replace(",", "").strip())
        except:
            result[(ym, name)] = 0
    return result


@st.cache_data(ttl=600)
def load_master():
    client = get_client()

    item_ws  = client.open_by_key(SHEET_ID).worksheet("ITEM_MASTER")
    item_df  = pd.DataFrame(item_ws.get_all_values()[1:], columns=item_ws.get_all_values()[0]).copy()

    cust_ws  = client.open_by_key(SHEET_ID).worksheet("CUSTOMER_MASTER")
    cust_raw = cust_ws.get_all_values()
    cust_df  = pd.DataFrame(cust_raw[1:], columns=cust_raw[0]).copy()

    if "거래처분류" not in cust_df.columns:
        cust_df["거래처분류"] = np.nan
    if "수수료율" not in cust_df.columns:
        cust_df["수수료율"] = 0

    cols = cust_df.columns.tolist()
    cust_df["국내여부"] = cust_df[cols[3]].astype(str).str.strip() if len(cols) >= 4 else "국내"

    cust_df["수수료율"]  = clean_numeric(cust_df["수수료율"]).fillna(0)
    cust_df["수수료율"]  = np.where(cust_df["수수료율"] > 1, cust_df["수수료율"] / 100, cust_df["수수료율"])
    cust_df["거래처명"]  = cust_df["거래처명"].astype(str).str.strip()
    cust_df["거래처분류"] = cust_df["거래처분류"].astype(str).str.strip()
    return item_df, cust_df


@st.cache_data(ttl=600)
def build_dataset():
    df        = load_view_table()
    item_df, cust_df = load_master()
    cost_dict = load_cost_master()

    merged = df.merge(item_df[["상품명", "품목군"]], left_on="내품상품명", right_on="상품명", how="left")
    merged["거래처분류"] = merged["거래처코드"]
    merged = merged.merge(cust_df[["거래처분류", "수수료율", "국내여부"]].drop_duplicates(),
                          on="거래처분류", how="left")
    merged["수수료율"] = merged["수수료율"].fillna(0)
    merged["국내여부"] = merged["국내여부"].fillna("국내")
    merged["거래처분류"] = merged["거래처분류"].replace("", np.nan).fillna(merged["거래처코드"])

    # 원가 시차 반영 (N월 출고 → N-1월 원가)
    def get_cost(row):
        try:
            prev_ym = (pd.to_datetime(row["출고년월"] + "-01") - relativedelta(months=1)).strftime("%Y-%m")
        except:
            return 0
        return cost_dict.get((prev_ym, row["내품상품명"]), 0)

    merged["제품원가"]  = merged.apply(get_cost, axis=1)
    merged["원가총액"]  = merged["총내품출고수량"] * merged["제품원가"]
    merged["채널수수료"] = merged["품목별매출(VAT제외)"] * merged["수수료율"]

    # ★ 매출총이익 = 매출액 - 원가 (채널수수료 미포함)
    merged["매출총이익"]  = merged["품목별매출(VAT제외)"] - merged["원가총액"]
    merged["매출총이익률"] = safe_divide(merged["매출총이익"], merged["품목별매출(VAT제외)"])

    return merged


# -----------------------------------
# 초기화
# -----------------------------------
df     = build_dataset()
client = get_client()
sh     = client.open_by_key(SHEET_ID)

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
all_months         = sort_month_cols(df["출고년월"].dropna().unique().tolist())
all_channel_groups = sorted(df["거래처분류"].dropna().unique().tolist())
all_items          = sorted(df["내품상품명"].dropna().unique().tolist())

selected_months         = st.sidebar.multiselect("출고년월", all_months,         default=all_months)
selected_channel_groups = st.sidebar.multiselect("채널",    all_channel_groups, default=all_channel_groups)
selected_items          = st.sidebar.multiselect("품목",    all_items,          default=all_items)

# -----------------------------------
# filtered_df
# -----------------------------------
filtered_df = df[
    (df["출고년월"].isin(selected_months)) &
    (df["거래처분류"].isin(selected_channel_groups)) &
    (df["내품상품명"].isin(selected_items))
].copy()

# -----------------------------------
# 비용 배분: 물류비 + 광고비
# 광고비 변경: 전체 채널 매출 비중으로 배분 (채널 기준)
# -----------------------------------
filtered_df["물류비"] = 0.0
filtered_df["광고비"] = 0.0

for m in selected_months:
    month_mask = filtered_df["출고년월"] == m

    # 물류비: 국내 거래처 매출 비중 안분
    dom_mask_all = (df["출고년월"] == m) & (df["국내여부"] == "국내")
    dom_total    = df.loc[dom_mask_all, "품목별매출(VAT제외)"].sum()
    if dom_total > 0:
        dom_month_mask = month_mask & (filtered_df["국내여부"] == "국내")
        ratio = filtered_df.loc[dom_month_mask, "품목별매출(VAT제외)"] / dom_total
        filtered_df.loc[dom_month_mask, "물류비"] = ratio * st.session_state["logistics_table"].get(m, 0)

    # 광고비: 해당 월 전체 품목군 광고비 합산 → 전체 채널 매출 비중으로 배분
    month_total_ad = sum(v for (cat, mon), v in st.session_state["ad_cost_monthly"].items() if mon == m)
    if month_total_ad > 0:
        month_total_sales = filtered_df.loc[month_mask, "품목별매출(VAT제외)"].sum()
        if month_total_sales > 0:
            ratio = filtered_df.loc[month_mask, "품목별매출(VAT제외)"] / month_total_sales
            filtered_df.loc[month_mask, "광고비"] = ratio * month_total_ad

if filtered_df.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# -----------------------------------
# 공헌이익 = 매출액 - 원가 - 채널수수료 - 물류비 - 광고비
# -----------------------------------
filtered_df["공헌이익"] = (
    filtered_df["품목별매출(VAT제외)"]
    - filtered_df["원가총액"]
    - filtered_df["채널수수료"]
    - filtered_df["물류비"]
    - filtered_df["광고비"]
)
filtered_df["공헌이익률"] = safe_divide(filtered_df["공헌이익"], filtered_df["품목별매출(VAT제외)"])

# -----------------------------------
# KPI
# -----------------------------------
st.markdown("""
<style>[data-testid="stMetricValue"] { font-size: 28px; }</style>
""", unsafe_allow_html=True)

total_sales        = filtered_df["품목별매출(VAT제외)"].sum()
total_qty          = filtered_df["총내품출고수량"].sum()
total_gross_profit = filtered_df["매출총이익"].sum()   # 매출액 - 원가
gross_profit_rate  = (total_gross_profit / total_sales * 100) if total_sales else 0

monthly_kpi = filtered_df.groupby("출고년월", as_index=False)[["품목별매출(VAT제외)"]].sum()
monthly_kpi["dt"] = pd.to_datetime(monthly_kpi["출고년월"] + "-01", errors="coerce")
monthly_kpi = monthly_kpi.sort_values("dt")
if len(monthly_kpi) >= 2:
    cur, prv = monthly_kpi.iloc[-1]["품목별매출(VAT제외)"], monthly_kpi.iloc[-2]["품목별매출(VAT제외)"]
    sales_mom = ((cur - prv) / prv * 100) if prv else 0
else:
    sales_mom = 0

top_channel = (
    filtered_df.groupby("거래처분류")["품목별매출(VAT제외)"]
    .sum().sort_values(ascending=False).index[0]
)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("누적 매출",                    f"{total_sales:,.0f} 원")
c2.metric("누적 출고량",                   f"{total_qty:,.0f}")
c3.metric("매출 총 이익 (매출액 - 원가)",   f"{total_gross_profit:,.0f} 원")
c4.metric("매출 총 이익률",                f"{gross_profit_rate:.2f}%")
c5.metric("Top 채널", top_channel,         delta=f"{sales_mom:.2f}% MoM")

st.divider()

# -----------------------------------
# 탭
# -----------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 월별 추이", "🏪 채널 분석", "📦 제품 분석", "📊 공헌이익 분석", "📥 다운로드"
])

# ===================================
# TAB 1: 월별 추이
# ===================================
with tab1:
    st.subheader("📈 월별 추이")

    # -----------------------------------
    # 전년(2025) 월별 매출 하드코딩
    # -----------------------------------
    PREV_YEAR_SALES = {
        "2025-01": 1991931168,
        "2025-02": 2176591336,
        "2025-03": 3651018838,
        "2025-04": 3872623112,
        "2025-05": 4175587255,
        "2025-06": 5201516827,
        "2025-07": 7693114287,
        "2025-08": 7275245463,
        "2025-09": 4754097322,
        "2025-10": 3676493223,
        "2025-11": 2421817295,
        "2025-12": 3094450532,
    }

    monthly = (
        filtered_df.groupby("출고년월", as_index=False)[
            ["품목별매출(VAT제외)", "매출총이익", "총내품출고수량"]
        ].sum()
    )
    monthly["dt"] = pd.to_datetime(monthly["출고년월"] + "-01", errors="coerce")
    monthly = monthly.sort_values("dt").rename(columns={
        "품목별매출(VAT제외)": "매출액", "총내품출고수량": "출고량"
    })

    # 26년 데이터가 있는 월 추출 → 해당 월의 전년 동월 데이터 매핑
    current_months_26 = [m for m in monthly["출고년월"].tolist() if m.startswith("2026")]
    has_26_data = len(current_months_26) > 0

    # 전년 동월 매출 컬럼 추가
    def get_prev_year_sales(ym):
        try:
            dt = pd.to_datetime(ym + "-01")
            prev_ym = (dt - relativedelta(years=1)).strftime("%Y-%m")
            return PREV_YEAR_SALES.get(prev_ym, None)
        except:
            return None

    monthly["전년동월매출"] = monthly["출고년월"].apply(get_prev_year_sales)

    # 컨트롤 영역
    ctrl1, ctrl2 = st.columns([1, 1])
    with ctrl1:
        show_label = st.checkbox("📊 라벨 표시", value=True)
    with ctrl2:
        # 26년 데이터가 있는 경우에만 토글 표시, 기본값 ON
        if has_26_data:
            show_prev_year = st.checkbox("📅 전년 동월 비교 표시", value=True)
        else:
            show_prev_year = False

    # -----------------------------------
    # 매출액 + 매출총이익 + 전년동월 그래프
    # -----------------------------------
    fig = go.Figure()

    # 전년 동월 Bar (26년 데이터가 있고 토글 ON일 때만)
    if has_26_data and show_prev_year:
        prev_data = monthly.dropna(subset=["전년동월매출"])
        if not prev_data.empty:
            fig.add_trace(go.Bar(
                x=prev_data["출고년월"],
                y=prev_data["전년동월매출"],
                name="매출액 (전년 동월)",
                marker_color="rgba(214, 39, 40, 0.4)",
                text=prev_data["전년동월매출"] if show_label else None,
                texttemplate='%{text:,.0f}', textposition='outside', cliponaxis=False
            ))
            
    # 26년 매출 Bar
    fig.add_trace(go.Bar(
        x=monthly["출고년월"],
        y=monthly["매출액"],
        name="매출액 (26년)",
        marker_color="#1f77b4",
        text=monthly["매출액"] if show_label else None,
        texttemplate='%{text:,.0f}', textposition='outside', cliponaxis=False
    ))

    # 매출총이익 Line
    fig.add_trace(go.Scatter(
        x=monthly["출고년월"],
        y=monthly["매출총이익"],
        name="매출총이익",
        mode="lines+markers+text" if show_label else "lines+markers",
        line=dict(width=4, color="#e73535"),
        text=monthly["매출총이익"] if show_label else None,
        texttemplate='%{text:,.0f}', textposition="top center"
    ))


    # Y축 범위: 전년 데이터 포함해서 계산
    y_vals = [monthly["매출액"].max(), monthly["매출총이익"].max()]
    if has_26_data and show_prev_year and not monthly["전년동월매출"].dropna().empty:
        y_vals.append(monthly["전년동월매출"].dropna().max())
    y_max = max(y_vals) * 1.25

    fig.update_layout(
        height=400,
        barmode="group",
        yaxis=dict(range=[0, y_max]),
        legend=dict(orientation="h"),
        margin=dict(t=40)
    )
    fig.update_traces(textfont=dict(size=12, color="black"))
    st.plotly_chart(fig, use_container_width=True)

    # 전년 동월 비교 표 (26년 데이터가 있고 토글 ON일 때만)
    if has_26_data and show_prev_year:
        st.markdown("##### 📊 전년 동월 대비 매출 비교")
        compare_df = monthly[monthly["출고년월"].isin(current_months_26)][
            ["출고년월", "매출액", "전년동월매출"]
        ].copy()
        compare_df = compare_df.rename(columns={
            "출고년월": "월",
            "매출액": "26년 매출액",
            "전년동월매출": "25년 동월 매출액"
        })
        compare_df["증감액"] = compare_df["26년 매출액"] - compare_df["25년 동월 매출액"].fillna(0)
        compare_df["증감률"] = np.where(
            compare_df["25년 동월 매출액"] > 0,
            (compare_df["증감액"] / compare_df["25년 동월 매출액"]) * 100,
            np.nan
        )
        st.dataframe(
            compare_df.style.format({
                "26년 매출액":    "{:,.0f}",
                "25년 동월 매출액": "{:,.0f}",
                "증감액":         "{:+,.0f}",
                "증감률":         "{:+.1f}%",
            }).applymap(
                lambda v: "color: #d62728" if isinstance(v, str) and v.startswith("-") else
                          ("color: #2ca02c" if isinstance(v, str) and v.startswith("+") else ""),
                subset=["증감액", "증감률"]
            ),
            use_container_width=True
        )

    # -----------------------------------
    # 출고량 그래프
    # -----------------------------------
    st.markdown("### 📦 월별 출고량")
    fig_qty = go.Figure()
    fig_qty.add_trace(go.Bar(
        x=monthly["출고년월"], y=monthly["출고량"], name="출고량",
        text=monthly["출고량"] if show_label else None,
        texttemplate='%{text:,.0f}', textposition='outside', cliponaxis=False
    ))
    fig_qty.update_layout(
        height=400,
        yaxis=dict(range=[0, monthly["출고량"].max() * 1.25]),
        margin=dict(t=40)
    )
    fig_qty.update_traces(textfont=dict(size=12, color="black"))
    st.plotly_chart(fig_qty, use_container_width=True)

# ===================================
# TAB 2: 채널 분석
# ===================================
with tab2:
    st.subheader("🏪 채널 분석")

    channel_summary = (
        filtered_df.groupby("거래처분류", as_index=False)[[
            "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
            "매출총이익", "채널수수료", "물류비", "광고비", "공헌이익"
        ]].sum()
    )
    channel_summary["매출총이익률"] = safe_divide(channel_summary["매출총이익"], channel_summary["품목별매출(VAT제외)"])
    channel_summary["공헌이익률"]   = safe_divide(channel_summary["공헌이익"],   channel_summary["품목별매출(VAT제외)"])
    channel_summary = channel_summary.sort_values("품목별매출(VAT제외)", ascending=False)

    # Top N 필터
    top_n_ch   = st.selectbox("Top 채널 기준", [5, 10, 20, 30], index=1, key="top_n_ch")
    top_channels = channel_summary.head(top_n_ch).copy()

    ch1, ch2 = st.columns(2)
    with ch1:
        fig_ch_s = px.bar(
            top_channels.sort_values("품목별매출(VAT제외)", ascending=True),
            x="품목별매출(VAT제외)", y="거래처분류", orientation="h", title="채널별 매출"
        )
        fig_ch_s.update_layout(xaxis_title="매출액", yaxis_title="채널")
        fig_ch_s.update_xaxes(tickformat=",")
        st.plotly_chart(fig_ch_s, use_container_width=True)
    with ch2:
        fig_ch_g = px.bar(
            top_channels.sort_values("매출총이익", ascending=True),
            x="매출총이익", y="거래처분류", orientation="h", title="채널별 매출 총 이익"
        )
        fig_ch_g.update_layout(xaxis_title="매출 총 이익", yaxis_title="채널")
        fig_ch_g.update_xaxes(tickformat=",")
        st.plotly_chart(fig_ch_g, use_container_width=True)

    # ── 월별 채널별 매출액 (위) ──
    st.subheader("💰 월별 채널별 매출액 및 구성비")
    ch_sales_pivot = pd.pivot_table(
        filtered_df, values="품목별매출(VAT제외)", index="거래처분류",
        columns="출고년월", aggfunc="sum", fill_value=0
    )
    ch_sales_mcols = sort_month_cols(ch_sales_pivot.columns.tolist())
    ch_sales_pivot = ch_sales_pivot.reindex(columns=ch_sales_mcols)

    # 월별 전체 합계 (구성비 분모)
    ch_month_totals = ch_sales_pivot.sum()

    # 합계 행 추가 후 정렬
    ch_sales_pivot = add_total_row(ch_sales_pivot)
    ch_sales_pivot = sort_pivot_by_last_month(ch_sales_pivot, ch_sales_mcols)

    # 월별 [매출액 / 구성비] 컬럼 쌍으로 재구성
    ch_display_cols = []
    ch_display_fmt  = {}
    for m in ch_sales_mcols:
        ratio_col = f"{m}_구성비"
        ch_sales_pivot[ratio_col] = ch_sales_pivot[m] / ch_month_totals[m] if ch_month_totals[m] > 0 else 0
        ch_display_cols += [m, ratio_col]
        ch_display_fmt[m]          = "{:,.0f}"
        ch_display_fmt[ratio_col]  = "{:.1%}"

    # 컬럼명 보기 좋게 rename
    rename_map = {f"{m}_구성비": f"{m}_구성비" for m in ch_sales_mcols}
    ch_display = ch_sales_pivot[ch_display_cols].copy()

    # MultiIndex 헤더 구성
    multi_cols = pd.MultiIndex.from_tuples(
        [(m, "매출액") if "_구성비" not in c else (m.replace("_구성비", ""), "구성비")
         for c in ch_display_cols
         for m in ([c] if "_구성비" not in c else [c.replace("_구성비", "")])],
        names=["월", "구분"]
    )
    # MultiIndex 직접 구성
    tuples = []
    for c in ch_display_cols:
        if c in ch_sales_mcols:
            tuples.append((c, "매출액"))
        else:
            base_m = c.replace("_구성비", "")
            tuples.append((base_m, "구성비"))
    ch_display.columns = pd.MultiIndex.from_tuples(tuples)

    st.dataframe(ch_display.style.format(ch_display_fmt), use_container_width=True)

    # ── 월별 채널별 출고량 (아래) ──
    st.subheader("📦 월별 채널별 출고량")
    ch_qty_pivot = pd.pivot_table(
        filtered_df, values="총내품출고수량", index="거래처분류",
        columns="출고년월", aggfunc="sum", fill_value=0
    )
    ch_qty_mcols = sort_month_cols(ch_qty_pivot.columns.tolist())
    ch_qty_pivot = ch_qty_pivot.reindex(columns=ch_qty_mcols)
    ch_qty_pivot = add_total_row(ch_qty_pivot)
    ch_qty_pivot = sort_pivot_by_last_month(ch_qty_pivot, ch_qty_mcols)
    st.dataframe(ch_qty_pivot.style.format("{:,.0f}"), use_container_width=True)

# ===================================
# TAB 3: 제품 분석
# ===================================
with tab3:
    st.subheader("📦 제품 분석")

    if "ad_cost_map" not in st.session_state:
        st.session_state["ad_cost_map"] = {}

    product_summary = (
        filtered_df.groupby("내품상품명", as_index=False)[[
            "총내품출고수량", "품목별매출(VAT제외)", "매출총이익",
            "물류비", "광고비", "공헌이익"
        ]].sum()
    )
    product_summary["매출총이익률"] = safe_divide(product_summary["매출총이익"], product_summary["품목별매출(VAT제외)"])
    product_summary["공헌이익률"]   = safe_divide(product_summary["공헌이익"],   product_summary["품목별매출(VAT제외)"])
    product_summary["개당 단가"]    = safe_divide(product_summary["품목별매출(VAT제외)"], product_summary["총내품출고수량"])
    total_prod = product_summary["품목별매출(VAT제외)"].sum()
    product_summary["제품 구성비"]  = safe_divide(product_summary["품목별매출(VAT제외)"], total_prod)
    product_summary = product_summary.sort_values("품목별매출(VAT제외)", ascending=False)

    top_n = st.selectbox("Top 제품 기준", [10, 20, 30, 50], index=0)
    top_products = product_summary.head(top_n).copy()

    pr1, pr2 = st.columns(2)
    with pr1:
        fig_pr_s = px.bar(
            top_products.sort_values("품목별매출(VAT제외)", ascending=True),
            x="품목별매출(VAT제외)", y="내품상품명", orientation="h",
            title=f"Top {top_n} 제품 매출"
        )
        fig_pr_s.update_layout(xaxis_title="매출액", yaxis_title="제품명")
        fig_pr_s.update_xaxes(tickformat=",")
        st.plotly_chart(fig_pr_s, use_container_width=True)
    with pr2:
        fig_pr_g = px.bar(
            top_products.sort_values("매출총이익", ascending=True),
            x="매출총이익", y="내품상품명", orientation="h",
            title=f"Top {top_n} 제품 매출 총 이익"
        )
        fig_pr_g.update_layout(xaxis_title="매출 총 이익", yaxis_title="제품명")
        fig_pr_g.update_xaxes(tickformat=",")
        st.plotly_chart(fig_pr_g, use_container_width=True)

    # ── 월별 제품별 매출액 (위) ──
    st.subheader("💰 월별 제품별 매출액 및 구성비 / 개당 단가")

    filt_top = filtered_df[filtered_df["내품상품명"].isin(top_products["내품상품명"])]

    prod_s_pivot = pd.pivot_table(
        filt_top, values="품목별매출(VAT제외)", index="내품상품명",
        columns="출고년월", aggfunc="sum", fill_value=0
    )
    prod_s_mcols = sort_month_cols(prod_s_pivot.columns.tolist())
    prod_s_pivot = prod_s_pivot.reindex(columns=prod_s_mcols)

    prod_q_by_month = pd.pivot_table(
        filt_top, values="총내품출고수량", index="내품상품명",
        columns="출고년월", aggfunc="sum", fill_value=0
    ).reindex(columns=prod_s_mcols, fill_value=0)

    # 월별 합계
    prod_month_totals = prod_s_pivot.sum()

    # 합계 행 추가 후 정렬
    prod_s_pivot  = add_total_row(prod_s_pivot)
    prod_q_by_month = add_total_row(prod_q_by_month)
    prod_s_pivot  = sort_pivot_by_last_month(prod_s_pivot, prod_s_mcols)
    prod_q_by_month = prod_q_by_month.reindex(prod_s_pivot.index, fill_value=0)

    # 월별 [매출액 / 구성비 / 개당단가] 컬럼 쌍으로 재구성
    prod_display_cols = []
    prod_display_fmt  = {}
    for m in prod_s_mcols:
        ratio_col = f"{m}_구성비"
        price_col = f"{m}_개당단가"

        month_total = prod_month_totals[m] if prod_month_totals[m] > 0 else 1
        prod_s_pivot[ratio_col] = prod_s_pivot[m] / month_total

        qty_series = prod_q_by_month[m].reindex(prod_s_pivot.index).fillna(0)
        prod_s_pivot[price_col] = np.where(
            qty_series != 0,
            prod_s_pivot[m] / qty_series,
            0
        )

        prod_display_cols += [m, ratio_col, price_col]
        prod_display_fmt[m]          = "{:,.0f}"
        prod_display_fmt[ratio_col]  = "{:.1%}"
        prod_display_fmt[price_col]  = "{:,.0f}"

    prod_display = prod_s_pivot[prod_display_cols].copy()

    # MultiIndex 헤더
    tuples_prod = []
    for c in prod_display_cols:
        if c in prod_s_mcols:
            tuples_prod.append((c, "매출액"))
        elif "_구성비" in c:
            tuples_prod.append((c.replace("_구성비", ""), "구성비"))
        else:
            tuples_prod.append((c.replace("_개당단가", ""), "개당단가"))
    prod_display.columns = pd.MultiIndex.from_tuples(tuples_prod)

    st.dataframe(prod_display.style.format(prod_display_fmt), use_container_width=True)

    # ── 월별 제품별 출고량 (아래) ──
    st.subheader("📦 월별 제품별 출고량")
    prod_q_pivot = pd.pivot_table(
        filt_top, values="총내품출고수량", index="내품상품명",
        columns="출고년월", aggfunc="sum", fill_value=0
    )
    prod_q_mcols = sort_month_cols(prod_q_pivot.columns.tolist())
    prod_q_pivot = prod_q_pivot.reindex(columns=prod_q_mcols)
    prod_q_pivot = add_total_row(prod_q_pivot)
    prod_q_pivot = sort_pivot_by_last_month(prod_q_pivot, prod_q_mcols)
    st.dataframe(prod_q_pivot.style.format("{:,.0f}"), use_container_width=True)

# ===================================
# TAB 4: 공헌이익 분석
# ===================================
with tab4:
    st.subheader("📊 공헌이익 분석")
    st.caption("※ 물류비 / 광고비는 COST_INPUT 시트에서 수정됩니다")

    # 1. 월별 물류비
    st.markdown("### 🚚 월별 물류비")
    logistics_df = pd.DataFrame([{m: st.session_state["logistics_table"].get(m, 0) for m in all_months}],
                                 index=["월별 물류비"])
    st.dataframe(logistics_df.style.format("{:,.0f}"), use_container_width=True)

    # 2. 제품 광고비
    st.markdown("### 📢 제품 광고비")
    ad_data  = []
    products = sorted(set(k[0] for k in st.session_state["ad_cost_monthly"].keys()))
    for p in products:
        row = {"제품명": p}
        for m in all_months:
            row[m] = st.session_state["ad_cost_monthly"].get((p, m), 0)
        ad_data.append(row)
    ad_df = pd.DataFrame(ad_data)
    ad_mcols = [c for c in ad_df.columns if c != "제품명"]
    st.dataframe(ad_df.style.format({m: "{:,.0f}" for m in ad_mcols}), use_container_width=True)

    # 3. 채널별 후정산 비용
    st.markdown("### 💸 채널별 후정산 비용")
    if st.session_state["channel_cost"]:
        cc_rows = [{"년월": ym, "거래처명": ch, "품목군": ig, "비용(VAT-)": amt}
                   for (ym, ch, ig), amt in st.session_state["channel_cost"].items()]
        cc_df = pd.DataFrame(cc_rows).sort_values(["년월", "거래처명", "품목군"])
        st.dataframe(cc_df.style.format({"비용(VAT-)": "{:,.0f}"}), use_container_width=True)
    else:
        st.info("CHANNEL_COST 시트에 데이터가 없습니다.")

    # ── 품목군별 공헌이익 ──
    st.markdown("### 📦 품목군별 공헌이익")

    temp_df = df.copy()

    # 물류비 (국내 기준 안분)
    temp_df["물류비"] = 0.0
    for m in all_months:
        mask = temp_df["출고년월"] == m
        dom  = mask & (temp_df["국내여부"] == "국내")
        dom_total = temp_df.loc[dom, "품목별매출(VAT제외)"].sum()
        if dom_total > 0:
            ratio = temp_df.loc[dom, "품목별매출(VAT제외)"] / dom_total
            temp_df.loc[dom, "물류비"] = ratio * st.session_state["logistics_table"].get(m, 0)

    # 광고비 (전체 채널 매출 비중 기준 배분)
    temp_df["광고비"] = 0.0
    for m in all_months:
        mask = temp_df["출고년월"] == m
        month_total_ad = sum(v for (cat, mon), v in st.session_state["ad_cost_monthly"].items() if mon == m)
        if month_total_ad > 0:
            month_total_sales = temp_df.loc[mask, "품목별매출(VAT제외)"].sum()
            if month_total_sales > 0:
                ratio = temp_df.loc[mask, "품목별매출(VAT제외)"] / month_total_sales
                temp_df.loc[mask, "광고비"] = ratio * month_total_ad

    # 필터 적용
    temp_df = temp_df[
        (temp_df["출고년월"].isin(selected_months)) &
        (temp_df["거래처분류"].isin(selected_channel_groups)) &
        (temp_df["내품상품명"].isin(selected_items))
    ].copy()

    # 매출총이익 재계산
    temp_df["매출총이익"] = temp_df["품목별매출(VAT제외)"] - temp_df["원가총액"]

    product_contrib = (
        temp_df.groupby("품목군", as_index=False)[[
            "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
            "매출총이익", "채널수수료", "물류비", "광고비"
        ]].sum()
    )

    # CHANNEL_COST → 품목군별 직접 매핑
    product_contrib["비용"] = 0.0
    for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
        if ym not in selected_months or ch not in selected_channel_groups:
            continue
        mask = product_contrib["품목군"] == ig
        if mask.any():
            product_contrib.loc[mask, "비용"] += amt

    # 공헌이익 = 매출액 - 원가 - 채널수수료 - 물류비 - 광고비 - 비용
    product_contrib["공헌이익"] = (
        product_contrib["품목별매출(VAT제외)"]
        - product_contrib["원가총액"]
        - product_contrib["채널수수료"]
        - product_contrib["물류비"]
        - product_contrib["광고비"]
        - product_contrib["비용"]
    )
    product_contrib["공헌이익률"] = safe_divide(product_contrib["공헌이익"], product_contrib["품목별매출(VAT제외)"])

    # 컬럼 순서 정렬
    product_contrib = product_contrib[[
        "품목군", "총내품출고수량", "품목별매출(VAT제외)",
        "원가총액", "매출총이익", "채널수수료",
        "물류비", "광고비", "비용", "공헌이익", "공헌이익률"
    ]]

    st.dataframe(
        product_contrib.style.format({
            "총내품출고수량":      "{:,.0f}",
            "품목별매출(VAT제외)": "{:,.0f}",
            "원가총액":           "{:,.0f}",
            "매출총이익":          "{:,.0f}",
            "채널수수료":          "{:,.0f}",
            "물류비":             "{:,.0f}",
            "광고비":             "{:,.0f}",
            "비용":               "{:,.0f}",
            "공헌이익":            "{:,.0f}",
            "공헌이익률":          "{:.2%}",
        }),
        use_container_width=True
    )

    # ── 채널별 공헌이익 (토글) ──
    with st.expander("🏪 채널별 공헌이익", expanded=False):
        channel_contrib = (
            temp_df.groupby("거래처분류", as_index=False)[[
                "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                "매출총이익", "채널수수료", "물류비", "광고비"
            ]].sum()
        )

        # CHANNEL_COST 후정산 비용
        channel_contrib["비용"] = 0.0
        for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
            if ym not in selected_months:
                continue
            mask = channel_contrib["거래처분류"] == ch
            if mask.any():
                channel_contrib.loc[mask, "비용"] += amt

        # 공헌이익 = 매출액 - 원가 - 채널수수료 - 물류비 - 광고비 - 비용
        channel_contrib["공헌이익"] = (
            channel_contrib["품목별매출(VAT제외)"]
            - channel_contrib["원가총액"]
            - channel_contrib["채널수수료"]
            - channel_contrib["물류비"]
            - channel_contrib["광고비"]
            - channel_contrib["비용"]
        )
        channel_contrib["공헌이익률"] = safe_divide(channel_contrib["공헌이익"], channel_contrib["품목별매출(VAT제외)"])

        # 컬럼 순서 정렬
        channel_contrib = channel_contrib[[
            "거래처분류", "총내품출고수량", "품목별매출(VAT제외)",
            "원가총액", "매출총이익", "채널수수료",
            "물류비", "광고비", "비용", "공헌이익", "공헌이익률"
        ]]

        st.dataframe(
            channel_contrib.style.format({
                "총내품출고수량":      "{:,.0f}",
                "품목별매출(VAT제외)": "{:,.0f}",
                "원가총액":           "{:,.0f}",
                "매출총이익":          "{:,.0f}",
                "채널수수료":          "{:,.0f}",
                "물류비":             "{:,.0f}",
                "광고비":             "{:,.0f}",
                "비용":               "{:,.0f}",
                "공헌이익":            "{:,.0f}",
                "공헌이익률":          "{:.2%}",
            }),
            use_container_width=True
        )

# ===================================
# TAB 5: 다운로드
# ===================================
with tab5:
    st.subheader("📥 다운로드")

    dl_ch_sales = pd.pivot_table(filtered_df, values="품목별매출(VAT제외)", index="거래처분류",
                                  columns="출고년월", aggfunc="sum", fill_value=0)
    dl_ch_sales = dl_ch_sales.reindex(columns=sort_month_cols(dl_ch_sales.columns.tolist()))

    dl_ch_qty = pd.pivot_table(filtered_df, values="총내품출고수량", index="거래처분류",
                                columns="출고년월", aggfunc="sum", fill_value=0)
    dl_ch_qty = dl_ch_qty.reindex(columns=sort_month_cols(dl_ch_qty.columns.tolist()))

    dl_prod_sales = pd.pivot_table(filtered_df, values="품목별매출(VAT제외)", index="내품상품명",
                                    columns="출고년월", aggfunc="sum", fill_value=0)
    dl_prod_sales = dl_prod_sales.reindex(columns=sort_month_cols(dl_prod_sales.columns.tolist()))

    dl_prod_qty = pd.pivot_table(filtered_df, values="총내품출고수량", index="내품상품명",
                                  columns="출고년월", aggfunc="sum", fill_value=0)
    dl_prod_qty = dl_prod_qty.reindex(columns=sort_month_cols(dl_prod_qty.columns.tolist()))

    download_file = make_excel_file({
        "월별채널매출":  dl_ch_sales,
        "월별채널출고량": dl_ch_qty,
        "월별제품매출":  dl_prod_sales,
        "월별제품출고량": dl_prod_qty,
    })

    st.download_button(
        label="📥 분석 결과 통합 엑셀 다운로드",
        data=download_file,
        file_name="Lingtea_Dashboard_v5.3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("### 포함 시트")
    st.write("- 월별 채널 매출액")
    st.write("- 월별 채널 출고량")
    st.write("- 월별 제품 매출액")
    st.write("- 월별 제품 출고량")

st.success("🚀 Lingtea Dashboard v5.3.2 Ready")
