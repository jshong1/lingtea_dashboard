import io
import os
import requests
from datetime import datetime

import gspread
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials
import firebase_admin
from firebase_admin import credentials, firestore, auth as firebase_auth

# -----------------------------------
# 기본 설정
# -----------------------------------
st.set_page_config(page_title="Lingtea Dashboard", layout="wide")

SHEET_ID = "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"

ALL_TABS = ["월별추이", "주차별추이", "채널분석", "제품분석", "공헌이익분석(통합)", "공헌이익분석(국내)", "공헌이익분석(해외)", "제품별원가", "다운로드"]

DEFAULT_USER_TABS = {t: False for t in ALL_TABS}
DEFAULT_ADMIN_TABS = {t: True for t in ALL_TABS}

# -----------------------------------
# Firebase Admin 초기화 (1회만)
# -----------------------------------
def init_firebase():
    if not firebase_admin._apps:
        key_dict = dict(st.secrets["gcp_service_account"])
        cred = credentials.Certificate(key_dict)
        firebase_admin.initialize_app(cred)

init_firebase()
db = firestore.client()

# -----------------------------------
# 로그인/회원가입 CSS
# -----------------------------------
LOGIN_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600&family=Noto+Sans+KR:wght@300;400;500;600&display=swap');

[data-testid="stAppViewContainer"],
[data-testid="stApp"],
.main {
    background: #F7F6F3 !important;
}
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebar"] { display: none !important; }

.auth-outer {
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 48px 16px;
}
.auth-card {
    background: #FFFFFF;
    border: 1px solid #E8E4DC;
    border-radius: 20px;
    padding: 52px 48px 44px;
    width: 100%;
    max-width: 440px;
    box-shadow: 0 8px 40px rgba(60,50,30,0.08), 0 1px 4px rgba(60,50,30,0.06);
}
.auth-brand {
    display: flex;
    align-items: baseline;
    gap: 10px;
    margin-bottom: 6px;
}
.auth-logo {
    font-family: 'Playfair Display', serif;
    font-size: 30px;
    font-weight: 600;
    color: #1A1814;
    letter-spacing: -0.3px;
    line-height: 1;
}
.auth-badge {
    font-family: 'Noto Sans KR', sans-serif;
    font-size: 10px;
    font-weight: 500;
    color: #9B8E78;
    letter-spacing: 2px;
    text-transform: uppercase;
    background: #F3EFE8;
    padding: 3px 8px;
    border-radius: 20px;
}
.auth-subtitle {
    font-family: 'Noto Sans KR', sans-serif;
    font-size: 13px;
    color: #9B8E78;
    margin-bottom: 36px;
    font-weight: 300;
}
.auth-divider {
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 20px 0;
}
.auth-divider-line {
    flex: 1;
    height: 1px;
    background: #E8E4DC;
}
.auth-divider-text {
    font-family: 'Noto Sans KR', sans-serif;
    font-size: 11px;
    color: #B8AF9E;
    white-space: nowrap;
}
.auth-switch {
    text-align: center;
    margin-top: 20px;
    font-family: 'Noto Sans KR', sans-serif;
    font-size: 13px;
    color: #9B8E78;
}
.auth-switch-link {
    color: #5C7A5C;
    font-weight: 500;
    cursor: pointer;
    text-decoration: underline;
    text-underline-offset: 2px;
}
.auth-notice {
    background: #F3EFE8;
    border-left: 3px solid #C8B88A;
    border-radius: 0 8px 8px 0;
    padding: 10px 14px;
    font-family: 'Noto Sans KR', sans-serif;
    font-size: 12px;
    color: #7A6E5A;
    margin-bottom: 20px;
    line-height: 1.6;
}
.msg-error {
    background: #FEF2F2;
    border: 1px solid #FECACA;
    border-radius: 8px;
    padding: 10px 14px;
    font-family: 'Noto Sans KR', sans-serif;
    font-size: 13px;
    color: #DC2626;
    margin-top: 10px;
}
.msg-success {
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    border-radius: 8px;
    padding: 10px 14px;
    font-family: 'Noto Sans KR', sans-serif;
    font-size: 13px;
    color: #16A34A;
    margin-top: 10px;
}

/* 입력 필드 */
.stTextInput > label {
    font-family: 'Noto Sans KR', sans-serif !important;
    font-size: 11px !important;
    font-weight: 500 !important;
    letter-spacing: 0.8px !important;
    text-transform: uppercase !important;
    color: #7A6E5A !important;
}
.stTextInput > div > div > input {
    background: #FAFAF8 !important;
    border: 1.5px solid #E8E4DC !important;
    border-radius: 10px !important;
    color: #1A1814 !important;
    font-family: 'Noto Sans KR', sans-serif !important;
    font-size: 14px !important;
    padding: 10px 14px !important;
    transition: border-color 0.2s !important;
}
.stTextInput > div > div > input:focus {
    border-color: #5C7A5C !important;
    box-shadow: 0 0 0 3px rgba(92,122,92,0.10) !important;
    background: #FFFFFF !important;
}

/* 버튼 - 기본(로그인/회원가입) */
.stButton > button[kind="primary"],
.stButton > button {
    background: #2D3A2D !important;
    color: #FFFFFF !important;
    font-family: 'Noto Sans KR', sans-serif !important;
    font-weight: 500 !important;
    font-size: 14px !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 11px 0 !important;
    width: 100% !important;
    letter-spacing: 0.5px !important;
    transition: all 0.18s ease !important;
}
.stButton > button:hover {
    background: #3D4E3D !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(45,58,45,0.18) !important;
}
</style>
"""

# -----------------------------------
# Firebase Auth REST API (로그인)
# -----------------------------------
def firebase_sign_in(email: str, password: str):
    api_key = st.secrets["firebase_web"]["api_key"]
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}"
    resp = requests.post(url, json={
        "email": email,
        "password": password,
        "returnSecureToken": True
    }, timeout=10)
    return resp.json()

# -----------------------------------
# Firebase Auth REST API (회원가입)
# -----------------------------------
def firebase_sign_up(email: str, password: str):
    api_key = st.secrets["firebase_web"]["api_key"]
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signUp?key={api_key}"
    resp = requests.post(url, json={
        "email": email,
        "password": password,
        "returnSecureToken": True
    }, timeout=10)
    return resp.json()

# -----------------------------------
# Firestore 유저 관리
# -----------------------------------
def get_user_doc(uid: str):
    doc = db.collection("users").document(uid).get()
    return doc.to_dict() if doc.exists else None

def create_or_update_user(uid: str, email: str, role: str, tabs: dict):
    db.collection("users").document(uid).set({
        "email": email,
        "role": role,
        "tabs": tabs,
        "updated_at": datetime.now().isoformat()
    }, merge=True)

def get_all_users():
    docs = db.collection("users").stream()
    return [{"uid": d.id, **d.to_dict()} for d in docs]

def update_user_tabs(uid: str, tabs: dict):
    db.collection("users").document(uid).update({
        "tabs": tabs,
        "updated_at": datetime.now().isoformat()
    })

def update_user_role(uid: str, role: str):
    db.collection("users").document(uid).update({
        "role": role,
        "updated_at": datetime.now().isoformat()
    })

def disable_user(uid: str):
    """Firebase Auth에서 사용자 비활성화"""
    firebase_auth.update_user(uid, disabled=True)
    db.collection("users").document(uid).update({
        "disabled": True,
        "updated_at": datetime.now().isoformat()
    })

def delete_user(uid: str):
    """Firebase Auth + Firestore에서 완전 삭제"""
    firebase_auth.delete_user(uid)
    db.collection("users").document(uid).delete()

# -----------------------------------
# 로그인 처리
# -----------------------------------
def handle_login(email: str, password: str):
    result = firebase_sign_in(email, password)
    if "error" in result:
        msg = result["error"].get("message", "LOGIN_FAILED")
        # Firebase v2 통합 오류코드 대응
        if "INVALID_LOGIN_CREDENTIALS" in msg or "INVALID_PASSWORD" in msg or "EMAIL_NOT_FOUND" in msg:
            return False, "이메일 또는 비밀번호가 올바르지 않습니다."
        elif "USER_DISABLED" in msg:
            return False, "비활성화된 계정입니다. 관리자에게 문의하세요."
        elif "TOO_MANY_ATTEMPTS_TRY_LATER" in msg:
            return False, "로그인 시도가 너무 많습니다. 잠시 후 다시 시도해주세요."
        return False, f"로그인 오류: {msg}"

    uid      = result["localId"]
    email    = result["email"]
    id_token = result["idToken"]

    user_doc     = get_user_doc(uid)
    admin_emails = list(st.secrets["auth"]["admin_emails"])

    if user_doc is None:
        role = "admin" if email in admin_emails else "user"
        tabs = DEFAULT_ADMIN_TABS.copy() if role == "admin" else DEFAULT_USER_TABS.copy()
        create_or_update_user(uid, email, role, tabs)
        user_doc = {"email": email, "role": role, "tabs": tabs}
    else:
        if email in admin_emails and user_doc.get("role") != "admin":
            update_user_role(uid, "admin")
            user_doc["role"] = "admin"

    st.session_state["logged_in"] = True
    st.session_state["uid"]       = uid
    st.session_state["email"]     = email
    st.session_state["role"]      = user_doc.get("role", "user")
    st.session_state["id_token"]  = id_token
    st.session_state["tabs_perm"] = user_doc.get("tabs", DEFAULT_USER_TABS.copy())
    # 쿠키에 저장 (새로고침 후 복원용)
    st.query_params["sid"] = uid
    return True, "ok"

# -----------------------------------
# 회원가입 처리
# -----------------------------------
def handle_signup(email: str, password: str, password_confirm: str):
    allowed_domain = st.secrets["auth"].get("allowed_domain", "lingtea.co.kr")
    if not email.endswith(f"@{allowed_domain}"):
        return False, f"@{allowed_domain} 이메일만 가입할 수 있습니다."
    if len(password) < 8:
        return False, "비밀번호는 8자 이상이어야 합니다."
    if password != password_confirm:
        return False, "비밀번호가 일치하지 않습니다."

    result = firebase_sign_up(email, password)
    if "error" in result:
        msg = result["error"].get("message", "SIGNUP_FAILED")
        if "EMAIL_EXISTS" in msg:
            return False, "이미 가입된 이메일입니다."
        elif "WEAK_PASSWORD" in msg:
            return False, "비밀번호가 너무 단순합니다. 더 강력한 비밀번호를 사용해주세요."
        return False, f"가입 오류: {msg}"

    uid          = result["localId"]
    admin_emails = list(st.secrets["auth"]["admin_emails"])
    role         = "admin" if email in admin_emails else "user"
    tabs         = DEFAULT_ADMIN_TABS.copy() if role == "admin" else DEFAULT_USER_TABS.copy()
    create_or_update_user(uid, email, role, tabs)
    return True, "ok"

# -----------------------------------
# 로그인 / 회원가입 화면
# -----------------------------------
def show_login():
    st.markdown(LOGIN_CSS, unsafe_allow_html=True)

    if "auth_mode" not in st.session_state:
        st.session_state["auth_mode"] = "login"  # "login" | "signup"

    col1, col2, col3 = st.columns([1, 1.1, 1])
    with col2:
        # 브랜드 헤더
        st.markdown("""
        <div class="auth-brand">
            <span class="auth-logo">Lingtea</span>
            <span class="auth-badge">Dashboard</span>
        </div>
        <div class="auth-subtitle">데이터 기반 비즈니스 인사이트 플랫폼</div>
        """, unsafe_allow_html=True)

        mode = st.session_state["auth_mode"]

        if mode == "login":
            # ── 로그인 폼 ──
            email    = st.text_input("이메일", placeholder="your@lingtea.co.kr", key="li_email")
            password = st.text_input("비밀번호", type="password", placeholder="8자 이상", key="li_pw")

            if st.button("로그인", use_container_width=True, key="btn_login"):
                if not email or not password:
                    st.markdown('<div class="msg-error">이메일과 비밀번호를 입력해주세요.</div>', unsafe_allow_html=True)
                else:
                    with st.spinner("인증 중..."):
                        ok, msg = handle_login(email, password)
                    if ok:
                        st.rerun()
                    else:
                        st.markdown(f'<div class="msg-error">{msg}</div>', unsafe_allow_html=True)

            # 전환 링크
            st.markdown("""
            <div class="auth-divider">
                <div class="auth-divider-line"></div>
                <span class="auth-divider-text">계정이 없으신가요?</span>
                <div class="auth-divider-line"></div>
            </div>
            """, unsafe_allow_html=True)

            if st.button("회원가입", use_container_width=True, key="btn_go_signup"):
                st.session_state["auth_mode"] = "signup"
                st.rerun()

        else:
            # ── 회원가입 폼 ──
            st.markdown("""
            <div class="auth-notice">
                🔐 <b>@lingtea.co.kr</b> 이메일 주소로만 가입할 수 있습니다.
            </div>
            """, unsafe_allow_html=True)

            email    = st.text_input("이메일", placeholder="your@lingtea.co.kr", key="su_email")
            password = st.text_input("비밀번호", type="password", placeholder="8자 이상", key="su_pw")
            password2= st.text_input("비밀번호 확인", type="password", placeholder="동일하게 입력", key="su_pw2")

            if st.button("가입하기", use_container_width=True, key="btn_signup"):
                if not email or not password or not password2:
                    st.markdown('<div class="msg-error">모든 항목을 입력해주세요.</div>', unsafe_allow_html=True)
                else:
                    with st.spinner("계정 생성 중..."):
                        ok, msg = handle_signup(email, password, password2)
                    if ok:
                        st.markdown('<div class="msg-success">✅ 가입이 완료되었습니다! 로그인해주세요.</div>', unsafe_allow_html=True)
                        import time; time.sleep(1.2)
                        st.session_state["auth_mode"] = "login"
                        st.rerun()
                    else:
                        st.markdown(f'<div class="msg-error">{msg}</div>', unsafe_allow_html=True)

            st.markdown("""
            <div class="auth-divider">
                <div class="auth-divider-line"></div>
                <span class="auth-divider-text">이미 계정이 있으신가요?</span>
                <div class="auth-divider-line"></div>
            </div>
            """, unsafe_allow_html=True)

            if st.button("로그인으로 돌아가기", use_container_width=True, key="btn_go_login"):
                st.session_state["auth_mode"] = "login"
                st.rerun()

# -----------------------------------
# 세션 복원 (쿠키 → session_state)
# -----------------------------------
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# -----------------------------------
# query_params 기반 세션 복원 (새로고침 유지)
# 로그인 시 uid를 ?sid=에 저장 → 새로고침 후 Firestore에서 복원
# -----------------------------------
if not st.session_state["logged_in"]:
    _sid = st.query_params.get("sid", "")
    if _sid:
        try:
            user_doc = get_user_doc(_sid)
            if user_doc and not user_doc.get("disabled", False):
                _email = user_doc.get("email", "")
                admin_emails = list(st.secrets["auth"]["admin_emails"])
                role = "admin" if _email in admin_emails else user_doc.get("role", "user")
                st.session_state["logged_in"] = True
                st.session_state["uid"]       = _sid
                st.session_state["email"]     = _email
                st.session_state["role"]      = role
                st.session_state["tabs_perm"] = user_doc.get("tabs", DEFAULT_USER_TABS.copy())
        except Exception:
            st.query_params.clear()

if not st.session_state["logged_in"]:
    show_login()
    st.stop()

# -----------------------------------
# 로그아웃 버튼 (사이드바)
# -----------------------------------
with st.sidebar:
    st.markdown(f"**{st.session_state['email']}**")
    role_badge = "🔑 관리자" if st.session_state["role"] == "admin" else "👤 사용자"
    st.caption(role_badge)
    if st.button("로그아웃", use_container_width=True):
        # 쿠키 삭제
        st.query_params.clear()
        for k in ["logged_in", "uid", "email", "role", "id_token", "tabs_perm",
                  "logistics_table", "ad_cost_monthly", "channel_cost", "channel_dept_map"]:
            st.session_state.pop(k, None)
        st.rerun()

# -----------------------------------
# 공통 유틸
# -----------------------------------
def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scope
    )
    return gspread.authorize(creds)

# [신규] 쓰기 권한 포함 gspread 클라이언트 (AUTH_MASTER 업데이트용)
def get_gspread_client_rw():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scope
    )
    return gspread.authorize(creds)

def load_cost_input(sh):
    """COST_INPUT 시트 로드.
    구조:
      1행: '부서별 물류비' | 2026-01 | 2026-02 | ...
      2행~(빈행 전): 부서명 | 금액 | 금액 | ...
      (빈행)
      광고비 헤더행: '광고비' | 2026-01 | ...
      이후행: 제품명 | 금액 | ...

    반환:
      logistics_dict: { (부서명, 년월): 금액 }   ← 부서별 물류비
      ad_dict:        { (품목군, 년월): 금액 }    ← 제품명→품목군 변환 후 합산
    """
    ws   = sh.worksheet("COST_INPUT")
    data = ws.get_all_values()
    if len(data) < 2:
        return {}, {}

    # 1행에서 월 헤더 추출 (B열부터)
    months = [str(m).strip() for m in data[0][1:]]

    logistics_dict  = {}
    ad_section      = False
    ad_months       = months
    ad_dict_by_prod = {}   # 임시: { (제품명, 년월): 금액 }

    for row in data[1:]:
        label = str(row[0]).strip()

        # 빈 행 — 섹션 구분용, 스킵
        if not label:
            continue

        # 광고비 헤더 행 감지 ("광고비" 포함)
        if "광고비" in label:
            ad_section = True
            ad_months  = [str(m).strip() for m in row[1:]]
            continue

        if ad_section:
            # 광고비 섹션: 제품명 기준으로 임시 저장
            for i, m in enumerate(ad_months):
                if i + 1 >= len(row):
                    break
                val = str(row[i + 1]).replace(",", "").strip()
                try:
                    ad_dict_by_prod[(label, m)] = float(val) if val and val.lower() not in ("", "nan") else 0
                except:
                    ad_dict_by_prod[(label, m)] = 0
        else:
            # 물류비 섹션: { (부서명, 년월): 금액 }
            for i, m in enumerate(months):
                if i + 1 >= len(row):
                    break
                val = str(row[i + 1]).replace(",", "").strip()
                try:
                    logistics_dict[(label, m)] = float(val) if val and val.lower() not in ("", "nan") else 0
                except:
                    logistics_dict[(label, m)] = 0

    # A열 값이 이미 품목군명 — ITEM_MASTER 변환 불필요
    # { (품목군, 년월): 금액 } 으로 직접 저장
    ad_dict = {}
    for (prod, m), amt in ad_dict_by_prod.items():
        if not prod or prod in ("nan", ""):
            continue
        key = (prod, m)
        ad_dict[key] = ad_dict.get(key, 0) + amt

    return logistics_dict, ad_dict


def load_channel_cost(sh):
    """CHANNEL_COST 시트 로드.
    컬럼: 년월(A) | 거래처명(B) | 품목군(C) | 비용항목(D) | 금액(E)
    반환: { (년월, 거래처명, 품목군): 금액 }  ← 기존과 동일 키 구조 유지
    (담당부서 매핑은 공헌이익 계산 시점에 cust_df로 처리)
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
    total     = pivot.sum(numeric_only=True)
    total_row = pd.DataFrame([total], index=["[합계]"])
    return pd.concat([total_row, pivot])

def sort_pivot_by_last_month(pivot: pd.DataFrame, month_cols: list) -> pd.DataFrame:
    if not month_cols or month_cols[-1] not in pivot.columns:
        return pivot
    last_m     = month_cols[-1]
    total_part = pivot[pivot.index == "[합계]"]
    data_part  = pivot[pivot.index != "[합계]"].sort_values(last_m, ascending=False)
    return pd.concat([total_part, data_part])

def safe_multiindex_from_tuples(df: pd.DataFrame, tuples: list):
    if tuples:
        df.columns = pd.MultiIndex.from_tuples(tuples)
    return df

def safe_max(series: pd.Series, default=0):
    if series is None or len(series) == 0:
        return default
    value = series.max()
    return default if pd.isna(value) else value

# -----------------------------------
# 데이터 로드
# -----------------------------------
@st.cache_data(ttl=600)
def load_view_table():
    client = get_gspread_client()
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
    client = get_gspread_client()
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
    client = get_gspread_client()
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
    # [추가] 담당부서 컬럼 정규화
    if "담당부서" not in cust_df.columns:
        cust_df["담당부서"] = ""
    cust_df["담당부서"] = cust_df["담당부서"].astype(str).str.strip()
    return item_df, cust_df

# [수정] AUTH_MASTER 컬럼 구조 확장: 권한유형(C) / 품목군(D) 신규 탐색
# 반환값: (auth_df, email_col, dept_col, role_type_col, item_group_col) 5-tuple
# [수정] ttl=60 짧은 캐시 적용 — 429 Quota 초과 방지 (저장 후 .clear() 호출)
@st.cache_data(ttl=60)
def load_auth_master():
    """AUTH_MASTER 시트 로드. 캐시 없음 (캐시 오염 방지).
    컬럼 구조: e-mail(A) | 담당부서(B) | 권한유형(C) | 품목군(D) | 비고(E)
    반환값: (auth_df, email_col, dept_col, role_type_col, item_group_col) 튜플."""
    client = get_gspread_client()
    ws   = client.open_by_key(SHEET_ID).worksheet("AUTH_MASTER")
    data = ws.get_all_values()
    if not data or len(data) < 1:
        return pd.DataFrame(), None, None, None, None
    # 헤더: 리스트 컴프리헨션으로 strip — Index.str / DataFrame.str 일절 사용 안 함
    raw_headers = [str(h).strip() for h in data[0]]
    rows        = data[1:]
    auth_df = pd.DataFrame(rows, columns=raw_headers).copy()
    # 컬럼명 탐색: 'e-mail' 또는 'email' 허용, '담당부서' 포함
    email_col      = next((c for c in raw_headers if c.lower().replace("-", "") == "email"), None)
    dept_col       = next((c for c in raw_headers if "담당부서" in c), None)
    # [신규] 권한유형 / 품목군 컬럼 탐색
    role_type_col  = next((c for c in raw_headers if "권한유형" in c), None)
    item_group_col = next((c for c in raw_headers if "품목군"   in c), None)

    if email_col is None or dept_col is None:
        return auth_df, None, None, None, None

    # 중복 컬럼 제거: 같은 이름 컬럼이 2개 이상이면 첫 번째만 유지
    auth_df = auth_df.loc[:, ~auth_df.columns.duplicated()].copy()
    # 값 정규화 (컬럼명 변경 없음, Series 단일 접근 보장)
    auth_df[email_col] = auth_df[email_col].astype(str).str.strip().str.lower()
    auth_df[dept_col]  = auth_df[dept_col].astype(str).str.strip()
    # [신규] 권한유형 / 품목군 정규화
    if role_type_col:
        auth_df[role_type_col]  = auth_df[role_type_col].astype(str).str.strip()
    if item_group_col:
        auth_df[item_group_col] = auth_df[item_group_col].astype(str).str.strip()

    return auth_df, email_col, dept_col, role_type_col, item_group_col

@st.cache_data(ttl=600)
def build_dataset():
    df        = load_view_table()
    item_df, cust_df = load_master()
    cost_dict = load_cost_master()
    merged = df.merge(
        item_df[["상품명", "품목군"]].drop_duplicates("상품명"),  # 상품명 중복 시 행 증식 방지
        left_on="내품상품명", right_on="상품명", how="left"
    )
    merged["거래처분류"] = merged["거래처코드"]
    # [추가] 거래처분류(CUSTOMER_MASTER B열) 기준으로 merge → 담당부서 포함
    # VIEW_TABLE.거래처코드 == CUSTOMER_MASTER.거래처분류(B열)
    cust_cols = ["거래처분류", "수수료율", "국내여부", "담당부서"]
    cust_cols = [c for c in cust_cols if c in cust_df.columns]
    merged = merged.merge(
        cust_df[cust_cols].drop_duplicates("거래처분류"),
        on="거래처분류", how="left"
    )
    merged["수수료율"] = merged["수수료율"].fillna(0)
    merged["국내여부"] = merged["국내여부"].fillna("국내")
    merged["담당부서"] = merged["담당부서"].fillna("").astype(str).str.strip()
    merged["거래처분류"] = merged["거래처분류"].replace("", np.nan).fillna(merged["거래처코드"])
    # [추가] 시간 컬럼 확장 (출고일자 → 연도/월/주차, 출고년월 유지)
    merged["출고일자"] = pd.to_datetime(merged["출고일자"], errors="coerce")
    merged["연도"] = merged["출고일자"].dt.year
    merged["월"]   = merged["출고일자"].dt.month
    merged["주차"] = merged["출고일자"].dt.isocalendar().week.astype("Int64")
    # 출고년월 유지 (기존 컬럼 덮어쓰지 않음 — 이미 load_view_table에서 생성됨)
    def get_cost(row):
        try:
            prev_ym = (pd.to_datetime(row["출고년월"] + "-01") - relativedelta(months=1)).strftime("%Y-%m")
        except:
            return 0
        return cost_dict.get((prev_ym, row["내품상품명"]), 0)
    merged["제품원가"]  = merged.apply(get_cost, axis=1)
    merged["원가총액"]  = merged["총내품출고수량"] * merged["제품원가"]
    merged["채널수수료"] = merged["품목별매출(VAT제외)"] * merged["수수료율"]
    merged["매출총이익"]  = merged["품목별매출(VAT제외)"] - merged["원가총액"]
    merged["매출총이익률"] = safe_divide(merged["매출총이익"], merged["품목별매출(VAT제외)"])
    # 품목군 NaN·공란 행 → 제외하지 않고 "__미분류__" 레이블로 보존
    # 매출 합계 정확성 유지 + 광고비·물류비 배분 대상에서는 자동 제외됨
    merged["품목군"] = merged["품목군"].fillna("__미분류__")
    merged.loc[
        merged["품목군"].astype(str).str.strip().isin(["", "nan"]),
        "품목군"
    ] = "__미분류__"
    # 매출조정 행도 동일하게 명시적 레이블 부여
    merged.loc[merged["내품상품명"].astype(str).str.strip() == "매출조정", "품목군"] = "__매출조정__"
    return merged

# -----------------------------------
# 초기화
# -----------------------------------
df     = build_dataset()
client = get_gspread_client()
sh     = client.open_by_key(SHEET_ID)

# -----------------------------------
# [v7.3] AUTH_MASTER 기반 권한 필터링
# 권한유형: 관리자 / 부서기반 / PM
# 관리자  → 전체 데이터 (필터 없음)
# 부서기반 → 담당부서 일치 데이터만
# PM      → 품목군 기준 필터 (모든 채널 조회 가능)
# -----------------------------------
_current_email = st.session_state.get("email", "").strip().lower()

try:
    # [수정] load_auth_master 반환값 5개로 변경
    _auth_df, _email_col, _dept_col, _role_type_col, _item_group_col = load_auth_master()

    if _email_col is None or _dept_col is None:
        st.error("🚫 AUTH_MASTER 컬럼 구조를 인식할 수 없습니다. (e-mail / 담당부서 컬럼 확인 필요)")
        st.stop()

    _user_row = _auth_df[_auth_df[_email_col] == _current_email]
    if _user_row.empty:
        st.error("🚫 접근 권한이 없습니다. 관리자에게 담당부서 권한을 요청하세요.")
        st.stop()

    _row0      = _user_row.iloc[0]
    _user_dept = str(_row0[_dept_col]).strip()

    if not _user_dept or _user_dept in ("nan", ""):
        st.error("🚫 접근 권한이 없습니다. 담당부서가 지정되지 않았습니다.")
        st.stop()

    # [신규] 권한유형 읽기 — 컬럼 없으면 담당부서 값으로 하위 호환
    # (기존 "관리자" 담당부서 로직 유지)
    _role_type = str(_row0[_role_type_col]).strip() if _role_type_col else (
        "관리자" if _user_dept == "관리자" else "부서기반"
    )

    # [신규] 품목군 읽기 — 쉼표 구분 split, 공백 제거
    _item_group_raw = str(_row0[_item_group_col]).strip() if _item_group_col else "ALL"
    _user_items = [x.strip() for x in _item_group_raw.split(",") if x.strip()]

    # [신규] 권한유형 기반 3-way 필터링
    if _role_type == "관리자":
        pass  # 전체 데이터 (필터 없음)

    elif _role_type == "부서기반":
        # 자신의 담당부서 데이터만 조회
        df = df[df["담당부서"] == _user_dept].copy()

    elif _role_type == "PM":
        # 모든 채널 조회 가능, 품목군 기준 필터
        if "ALL" not in _user_items:
            df = df[df["품목군"].isin(_user_items)].copy()

    else:
        # 알 수 없는 권한유형 → 안전하게 차단
        st.error(f"🚫 알 수 없는 권한유형입니다: {_role_type}")
        st.stop()

    # 세션에 저장 (관리자 UI 등에서 활용)
    st.session_state["user_role_type"] = _role_type
    st.session_state["user_dept"]      = _user_dept
    st.session_state["user_items"]     = _user_items

except Exception as _e:
    import traceback
    st.error(f"🚫 권한 확인 중 오류가 발생했습니다: {_e}")
    st.code(traceback.format_exc())   # 정확한 줄번호/원인 화면 출력
    st.stop()

if "logistics_table" not in st.session_state or not st.session_state.get("ad_cost_monthly"):
    logistics_dict, ad_dict = load_cost_input(sh)
    st.session_state["logistics_table"] = logistics_dict
    st.session_state["ad_cost_monthly"] = ad_dict

if "channel_cost" not in st.session_state:
    st.session_state["channel_cost"] = load_channel_cost(sh)

# [신규] 거래처명 → 담당부서 매핑 (CHANNEL_COST 부서별 배분용)
# CUSTOMER_MASTER의 거래처명(거래처분류) → 담당부서 딕셔너리
if "channel_dept_map" not in st.session_state:
    _, _cust_df = load_master()
    # 거래처분류(=거래처명 키)와 담당부서 컬럼으로 매핑 생성
    # 거래처분류가 CHANNEL_COST.거래처명과 동일한 값이어야 함
    _ch_dept = (
        _cust_df[["거래처분류", "담당부서"]]
        .dropna(subset=["거래처분류"])
        .drop_duplicates("거래처분류")
    )
    st.session_state["channel_dept_map"] = dict(
        zip(_ch_dept["거래처분류"].str.strip(), _ch_dept["담당부서"].str.strip())
    )

# -----------------------------------
# 권한 체크 헬퍼
# -----------------------------------
def tab_allowed(tab_name: str) -> bool:
    if st.session_state.get("role") == "admin":
        return True
    return st.session_state.get("tabs_perm", {}).get(tab_name, False)

# -----------------------------------
# 사이드바 필터
# -----------------------------------
st.title("📊 Lingtea Dashboard")
st.caption("월별 채널/제품 분석 + 매출/출고량/공헌이익 통합 대시보드")

st.sidebar.header("📌 필터")

# 출고일자 기반 기간 필터
_valid_dates = df["출고일자"].dropna()
_min_date    = _valid_dates.min().date()
_max_date    = _valid_dates.max().date()
_today       = datetime.today().date()
_today_clamped = min(_today, _max_date)  # 오늘이 데이터 최대일보다 크면 최대일로

st.sidebar.markdown("**📅 기간 설정**")

_date_start = st.sidebar.date_input("시작일", value=_min_date, min_value=_min_date, max_value=_max_date, key="date_start")
_date_end   = st.sidebar.date_input("종료일", value=_max_date, min_value=_min_date, max_value=_max_date, key="date_end")

if _date_start > _date_end:
    st.sidebar.error("시작일이 종료일보다 늦을 수 없습니다.")
    st.stop()

import datetime as _dt
_date_start_dt = pd.Timestamp(_date_start)
_date_end_dt   = pd.Timestamp(_date_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

# all_months: 전체 데이터 기준 월 목록 (공헌이익 탭 비용표 등에서 사용)
all_months = sort_month_cols(df["출고년월"].dropna().unique().tolist())

# 기간 내 출고년월 목록 역산 (비용 배분 루프용 selected_months 유지)
_date_filtered_df  = df[(df["출고일자"] >= _date_start_dt) & (df["출고일자"] <= _date_end_dt)]
selected_months    = sort_month_cols(_date_filtered_df["출고년월"].dropna().unique().tolist())

# -----------------------------------
# 채널 필터 — 담당부서 → 채널 연동
# -----------------------------------
st.sidebar.markdown("**🏪 채널**")

_all_depts = sorted(df["담당부서"].dropna().replace("", None).dropna().unique().tolist())

# 담당부서: 빈칸으로 시작
selected_depts = st.sidebar.multiselect(
    "담당부서",
    options=_all_depts,
    default=[],
    key="filter_depts"
)

# 선택된 부서 있으면 해당 채널만, 없으면 전체
if selected_depts:
    _dept_filtered_channels = sorted(
        df[df["담당부서"].isin(selected_depts)]["거래처분류"].dropna().unique().tolist()
    )
else:
    _dept_filtered_channels = sorted(df["거래처분류"].dropna().unique().tolist())

all_channel_groups = sorted(df["거래처분류"].dropna().unique().tolist())

# 채널 Select All 체크박스
_ch_select_all = st.sidebar.checkbox("채널 전체 선택", value=True, key="ch_select_all")
if _ch_select_all:
    selected_channel_groups = _dept_filtered_channels
else:
    selected_channel_groups = st.sidebar.multiselect(
        "채널",
        options=_dept_filtered_channels,
        default=_dept_filtered_channels,
        key="filter_channels"
    )

# -----------------------------------
# 품목 필터 — 품목군 → 품목 연동
# -----------------------------------
st.sidebar.markdown("**📦 품목**")

_all_item_groups = sorted([
    g for g in df["품목군"].dropna().unique().tolist()
    if g not in ("__매출조정__", "__미분류__")
])

# 품목군: 빈칸으로 시작
selected_item_groups = st.sidebar.multiselect(
    "품목군",
    options=_all_item_groups,
    default=[],
    key="filter_item_groups"
)

# 선택된 품목군 있으면 해당 SKU만, 없으면 전체
if selected_item_groups:
    _ig_filtered_items = sorted(
        df[df["품목군"].isin(selected_item_groups)]["내품상품명"].dropna().unique().tolist()
    )
else:
    _ig_filtered_items = sorted(df["내품상품명"].dropna().unique().tolist())

all_items = sorted([
    i for i in df["내품상품명"].dropna().unique().tolist()
    if i != "매출조정"
])

# 품목 Select All 체크박스
_item_select_all = st.sidebar.checkbox("품목 전체 선택", value=True, key="item_select_all")
if _item_select_all:
    selected_items = _ig_filtered_items
else:
    selected_items = st.sidebar.multiselect(
        "품목",
        options=_ig_filtered_items,
        default=_ig_filtered_items,
        key="filter_items"
    )

# -----------------------------------
# filtered_df
# -----------------------------------
_base_mask = (
    (df["출고일자"] >= _date_start_dt) &
    (df["출고일자"] <= _date_end_dt) &
    (df["거래처분류"].isin(selected_channel_groups))
)
# 매출조정 행은 부서 구분 없이 품목 필터와 무관하게 항상 포함 (음수 매출 반영 필수)
_is_adj_mask = df["내품상품명"].astype(str).str.strip() == "매출조정"
filtered_df = df[
    _base_mask & (df["내품상품명"].isin(selected_items) | _is_adj_mask)
].copy()

# -----------------------------------
# 비용 배분
# -----------------------------------
filtered_df["물류비"] = 0.0
filtered_df["광고비"] = 0.0

for m in selected_months:
    month_mask = filtered_df["출고년월"] == m

    # [수정] 물류비: 부서별 배분
    # logistics_table 키: (부서명, 년월) → 금액
    _depts_this_month = set(
        dept for (dept, ym) in st.session_state["logistics_table"].keys() if ym == m
    )
    for dept in _depts_this_month:
        dept_logistics_amt = st.session_state["logistics_table"].get((dept, m), 0)
        if dept_logistics_amt <= 0:
            continue
        # 해당 부서 전체 매출 (df 기준, 필터 전)
        dept_total_mask = (df["출고년월"] == m) & (df["담당부서"] == dept)
        dept_total_sales = df.loc[dept_total_mask, "품목별매출(VAT제외)"].sum()
        if dept_total_sales <= 0:
            continue
        # filtered_df 내 해당 부서 행에만 비율 배분
        f_dept_mask = month_mask & (filtered_df["담당부서"] == dept)
        ratio = filtered_df.loc[f_dept_mask, "품목별매출(VAT제외)"] / dept_total_sales
        filtered_df.loc[f_dept_mask, "물류비"] = ratio * dept_logistics_amt

    # [수정] 광고비: 품목군 기준 + 국내 채널만 배분 (해외 채널 제외)
    # ad_cost_monthly 키: (품목군, 년월) → 금액
    _igs_this_month = set(
        ig for (ig, ym) in st.session_state["ad_cost_monthly"].keys() if ym == m
    )
    for ig in _igs_this_month:
        ad_amt = st.session_state["ad_cost_monthly"].get((ig, m), 0)
        if ad_amt <= 0:
            continue
        # 분모: 해당 품목군 국내 채널 매출만 (df 기준, 필터 전)
        ig_total_mask  = (df["출고년월"] == m) & (df["품목군"] == ig) & (df["국내여부"] == "국내")
        ig_total_sales = df.loc[ig_total_mask, "품목별매출(VAT제외)"].sum()
        if ig_total_sales <= 0:
            continue
        # 배분 대상: filtered_df 내 해당 품목군 + 국내 행만
        f_ig_mask = month_mask & (filtered_df["품목군"] == ig) & (filtered_df["국내여부"] == "국내")
        ratio = filtered_df.loc[f_ig_mask, "품목별매출(VAT제외)"] / ig_total_sales
        filtered_df.loc[f_ig_mask, "광고비"] = ratio * ad_amt

if filtered_df.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

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
# 탭 구성 (권한에 따라 동적 생성) — KPI보다 먼저 체크
# -----------------------------------
tab_defs = {
    "월별추이":         "📈 월별 추이",
    "주차별추이":       "📅 주차별 추이",
    "채널분석":         "🏪 채널 분석",
    "제품분석":         "📦 제품 분석",
    "공헌이익분석(국내)": "📊 공헌이익(국내)",
    "공헌이익분석(해외)": "🌏 공헌이익(해외)",
    "공헌이익분석(통합)": "📋 공헌이익(통합)",
    "제품별원가":       "💰 제품별 원가",
    "다운로드":         "📥 다운로드",
}

visible_tabs = [k for k in ALL_TABS if tab_allowed(k)]

if st.session_state["role"] == "admin":
    visible_tabs_labels = [tab_defs[k] for k in visible_tabs] + ["⚙️ 관리자"]
else:
    visible_tabs_labels = [tab_defs[k] for k in visible_tabs]

if not visible_tabs:
    st.warning("접근 가능한 탭이 없습니다. 관리자에게 권한을 요청하세요.")
    st.stop()

# -----------------------------------
# KPI (탭 권한 있는 사용자에게만 표시)
# -----------------------------------
st.markdown("""
<style>[data-testid="stMetricValue"] { font-size: 28px; }</style>
""", unsafe_allow_html=True)

total_sales        = filtered_df["품목별매출(VAT제외)"].sum()
total_qty          = filtered_df["총내품출고수량"].sum()
total_gross_profit = filtered_df["매출총이익"].sum()
gross_profit_rate  = (total_gross_profit / total_sales * 100) if total_sales else 0

monthly_kpi = filtered_df.groupby("출고년월", as_index=False)[["품목별매출(VAT제외)"]].sum()
monthly_kpi["dt"] = pd.to_datetime(monthly_kpi["출고년월"] + "-01", errors="coerce")
monthly_kpi = monthly_kpi.sort_values("dt")
if len(monthly_kpi) >= 2:
    cur, prv = monthly_kpi.iloc[-1]["품목별매출(VAT제외)"], monthly_kpi.iloc[-2]["품목별매출(VAT제외)"]
    sales_mom = ((cur - prv) / prv * 100) if prv else 0
else:
    sales_mom = 0

_top_channel_series = (
    filtered_df.groupby("거래처분류")["품목별매출(VAT제외)"]
    .sum().sort_values(ascending=False)
)
top_channel = _top_channel_series.index[0] if not _top_channel_series.empty else "-"

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("누적 매출",                    f"{total_sales:,.0f} 원")
c2.metric("누적 출고량",                   f"{total_qty:,.0f}")
c3.metric("매출 총 이익 (매출액 - 원가)",   f"{total_gross_profit:,.0f} 원")
c4.metric("매출 총 이익률",                f"{gross_profit_rate:.2f}%")
c5.metric("Top 채널", top_channel,         delta=f"{sales_mom:.2f}% MoM")

st.divider()

# -----------------------------
# 차트용 데이터: 매출조정 제외
# -----------------------------
filtered_df = filtered_df[
    filtered_df["내품상품명"].astype(str).str.strip() != "매출조정"
].copy()

if filtered_df.empty:
    st.info("차트 및 상세 분석에 표시할 데이터가 없습니다.")
    st.stop()

created_tabs = st.tabs(visible_tabs_labels)

tab_map = {k: created_tabs[i] for i, k in enumerate(visible_tabs)}
if st.session_state["role"] == "admin":
    admin_tab = created_tabs[len(visible_tabs)]

# ===================================
# TAB 1: 월별 추이
# ===================================
if "월별추이" in tab_map:
    with tab_map["월별추이"]:
        st.subheader("📈 월별 추이")

        PREV_YEAR_SALES = {
            "2025-01": 1991931168, "2025-02": 2176591336, "2025-03": 3651018838,
            "2025-04": 3872623112, "2025-05": 4175587255, "2025-06": 5201516827,
            "2025-07": 7693114287, "2025-08": 7275245463, "2025-09": 4754097322,
            "2025-10": 3676493223, "2025-11": 2421817295, "2025-12": 3094450532,
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

        current_months_26 = [m for m in monthly["출고년월"].tolist() if m.startswith("2026")]
        has_26_data = len(current_months_26) > 0

        def get_prev_year_sales(ym):
            try:
                dt = pd.to_datetime(ym + "-01")
                prev_ym = (dt - relativedelta(years=1)).strftime("%Y-%m")
                return PREV_YEAR_SALES.get(prev_ym, None)
            except:
                return None

        monthly["전년동월매출"] = monthly["출고년월"].apply(get_prev_year_sales)

        ctrl1, ctrl2 = st.columns([1, 1])
        with ctrl1:
            show_label = st.checkbox("📊 라벨 표시", value=True)
        with ctrl2:
            if has_26_data:
                show_prev_year = st.checkbox("📅 전년 동월 비교 표시", value=False)
            else:
                show_prev_year = False

        fig = go.Figure()
        if has_26_data and show_prev_year:
            prev_data = monthly.dropna(subset=["전년동월매출"])
            if not prev_data.empty:
                fig.add_trace(go.Bar(
                    x=prev_data["출고년월"], y=prev_data["전년동월매출"],
                    name="매출액 (전년 동월)", marker_color="#D1D5DB",
                    text=prev_data["전년동월매출"] if show_label else None,
                    texttemplate='%{text:,.0f}', textposition='outside', cliponaxis=False
                ))
        fig.add_trace(go.Bar(
            x=monthly["출고년월"], y=monthly["매출액"],
            name="매출액 (26년)", marker_color="#1f77b4",
            text=monthly["매출액"] if show_label else None,
            texttemplate='%{text:,.0f}', textposition='outside', cliponaxis=False
        ))
        fig.add_trace(go.Scatter(
            x=monthly["출고년월"], y=monthly["매출총이익"],
            name="매출총이익", mode="lines+markers+text" if show_label else "lines+markers",
            line=dict(width=4, color="#EF4444"),
            text=monthly["매출총이익"] if show_label else None,
            texttemplate='%{text:,.0f}', textposition="top center"
        ))

        y_vals = [safe_max(monthly["매출액"], 0), safe_max(monthly["매출총이익"], 0)]
        if has_26_data and show_prev_year and not monthly["전년동월매출"].dropna().empty:
            y_vals.append(safe_max(monthly["전년동월매출"].dropna(), 0))
        y_max = max(y_vals) * 1.25 if max(y_vals) > 0 else 1

        fig.update_layout(height=400, barmode="group",
                          yaxis=dict(range=[0, y_max]),
                          legend=dict(orientation="h"), margin=dict(t=40))
        fig.update_traces(textfont=dict(size=12, color="black"))
        st.plotly_chart(fig, use_container_width=True)

        if has_26_data and show_prev_year:
            st.markdown("##### 📊 전년 동월 대비 매출 비교")
            compare_df = monthly[monthly["출고년월"].isin(current_months_26)][
                ["출고년월", "매출액", "전년동월매출"]
            ].copy()
            compare_df = compare_df.rename(columns={
                "출고년월": "월", "매출액": "26년 매출액", "전년동월매출": "25년 동월 매출액"
            })
            compare_df["증감액"] = compare_df["26년 매출액"] - compare_df["25년 동월 매출액"].fillna(0)
            compare_df["증감률"] = np.where(
                compare_df["25년 동월 매출액"] > 0,
                (compare_df["증감액"] / compare_df["25년 동월 매출액"]) * 100, np.nan
            )
            st.dataframe(
                compare_df.style.format({
                    "26년 매출액": "{:,.0f}", "25년 동월 매출액": "{:,.0f}",
                    "증감액": "{:+,.0f}", "증감률": "{:+.1f}%",
                }),
                use_container_width=True
            )

        st.markdown("### 📦 월별 출고량")
        fig_qty = go.Figure()
        fig_qty.add_trace(go.Bar(
            x=monthly["출고년월"], y=monthly["출고량"], name="출고량",
            text=monthly["출고량"] if show_label else None,
            texttemplate='%{text:,.0f}', textposition='outside', cliponaxis=False
        ))
        fig_qty.update_layout(height=400, yaxis=dict(range=[0, safe_max(monthly["출고량"], 0) * 1.25 if safe_max(monthly["출고량"], 0) > 0 else 1]),
                               margin=dict(t=40))
        fig_qty.update_traces(textfont=dict(size=12, color="black"))
        st.plotly_chart(fig_qty, use_container_width=True)

        # ── 부서별 매출 구성비 도넛 차트 (expander) ──
        with st.expander("🏢 부서별 월별 매출 구성비", expanded=False):
            # 매출조정 행 제외하고 집계 (부서 귀속 불분명한 조정액 제외)
            _dept_df = filtered_df[filtered_df["내품상품명"].astype(str).str.strip() != "매출조정"].copy()
            _dept_monthly = (
                _dept_df.groupby(["출고년월", "담당부서"], as_index=False)["품목별매출(VAT제외)"].sum()
            )
            _dept_monthly = _dept_monthly[_dept_monthly["품목별매출(VAT제외)"] > 0]

            _months_sorted = sort_month_cols(_dept_monthly["출고년월"].dropna().unique().tolist())

            if not _months_sorted:
                st.info("표시할 부서별 데이터가 없습니다.")
            else:
                # 월이 많으면 한 줄에 3개씩 배치
                _cols_per_row = 3
                _rows = [_months_sorted[i:i+_cols_per_row] for i in range(0, len(_months_sorted), _cols_per_row)]

                # 부서 색상 고정 (최대 12개 부서)
                _dept_list = sorted(_dept_monthly["담당부서"].dropna().unique().tolist())
                _palette = [
                    "#1f77b4","#ff7f0e","#2ca02c","#d62728","#9467bd",
                    "#8c564b","#e377c2","#7f7f7f","#bcbd22","#17becf",
                    "#aec7e8","#ffbb78",
                ]
                _color_map = {d: _palette[i % len(_palette)] for i, d in enumerate(_dept_list)}

                for _row_months in _rows:
                    _cols = st.columns(len(_row_months))
                    for _ci, _ym in enumerate(_row_months):
                        _slice = _dept_monthly[_dept_monthly["출고년월"] == _ym].copy()
                        _total = _slice["품목별매출(VAT제외)"].sum()
                        with _cols[_ci]:
                            st.caption(f"**{_ym}** | 합계 {_total:,.0f}원")
                            _fig_pie = go.Figure(go.Pie(
                                labels=_slice["담당부서"],
                                values=_slice["품목별매출(VAT제외)"],
                                hole=0.45,
                                marker_colors=[_color_map.get(d, "#cccccc") for d in _slice["담당부서"]],
                                textinfo="label+percent",
                                textfont=dict(size=11),
                                hovertemplate="%{label}<br>%{value:,.0f}원<br>%{percent}<extra></extra>",
                            ))
                            _fig_pie.update_layout(
                                height=280,
                                margin=dict(t=10, b=10, l=10, r=10),
                                showlegend=False,
                            )
                            st.plotly_chart(_fig_pie, use_container_width=True, key=f"pie_{_ym}")

# ===================================
# TAB 2: 주차별 추이
# ===================================
if "주차별추이" in tab_map:
    with tab_map["주차별추이"]:
        st.subheader("📅 주차별 추이")

        # 주차 컬럼 생성: YYYY-WNN 형태 (예: 2026-W03)
        wdf = filtered_df.copy()
        wdf = wdf.dropna(subset=["출고일자", "주차"])
        wdf["연도"] = wdf["출고일자"].dt.isocalendar().year.astype(int)
        wdf["주차_int"] = wdf["주차"].astype(int)
        wdf["연도주차"] = wdf["연도"].astype(str) + "-W" + wdf["주차_int"].astype(str).str.zfill(2)

        # [수정] x축 레이블: ISO 주차 → 해당 주의 월요일~일요일 날짜 범위로 변환
        # isocalendar() 기반이므로 %G-W%V-%u (ISO 8601) 포맷 사용
        # 예) 2026-W16 → 26.04.13~26.04.19
        def week_to_date_range(yw: str) -> str:
            try:
                year, w = yw.split("-W")
                mon = datetime.strptime(f"{year}-W{int(w):02d}-1", "%G-W%V-%u")
                sun = mon + pd.Timedelta(days=6)
                return f"{mon.strftime('%y.%m.%d')}~{sun.strftime('%y.%m.%d')}"
            except Exception:
                return yw

        # 정렬 기준용 숫자키
        wdf["연도주차_key"] = wdf["연도"] * 100 + wdf["주차_int"]
        wdf["주차_label"] = wdf["연도주차"].apply(week_to_date_range)

        weekly = (
            wdf.groupby(["연도주차", "연도주차_key", "주차_label"], as_index=False)[
                ["품목별매출(VAT제외)", "매출총이익", "총내품출고수량"]
            ].sum()
        )
        weekly = weekly.sort_values("연도주차_key").rename(columns={
            "품목별매출(VAT제외)": "매출액",
            "총내품출고수량": "출고량",
        })

        if weekly.empty:
            st.info("선택한 기간에 주차 데이터가 없습니다.")
        else:
            wk_ctrl1, wk_ctrl2 = st.columns([1, 1])
            with wk_ctrl1:
                wk_show_label = st.checkbox("📊 라벨 표시", value=True, key="wk_label")
            with wk_ctrl2:
                wk_top_n = st.selectbox(
                    "최근 N주 표시 (종료일 기준)",
                    options=[4, 8, 12, 16, 24, 52],
                    index=1,          # 기본값: 8주
                    key="wk_top_n"
                )

            # 종료일 기준 최근 N주: weekly는 이미 연도주차_key 오름차순 정렬됨
            weekly_view = weekly.tail(wk_top_n).copy()

            # ── 주차별 매출액 + 매출총이익 차트 ──
            st.markdown("### 💰 주차별 매출액 / 매출총이익")
            fig_wk = go.Figure()
            fig_wk.add_trace(go.Bar(
                x=weekly_view["주차_label"],
                y=weekly_view["매출액"],
                name="매출액",
                marker_color="#1f77b4",
                text=weekly_view["매출액"] if wk_show_label else None,
                texttemplate="%{text:,.0f}",
                textposition="outside",
                cliponaxis=False,
            ))
            fig_wk.add_trace(go.Scatter(
                x=weekly_view["주차_label"],
                y=weekly_view["매출총이익"],
                name="매출총이익",
                mode="lines+markers+text" if wk_show_label else "lines+markers",
                line=dict(width=3, color="#EF4444"),
                text=weekly_view["매출총이익"] if wk_show_label else None,
                texttemplate="%{text:,.0f}",
                textposition="top center",
            ))
            wk_y_max = max(weekly_view["매출액"].max(), weekly_view["매출총이익"].max()) * 1.3
            fig_wk.update_layout(
                height=420,
                barmode="group",
                yaxis=dict(range=[0, wk_y_max]),
                legend=dict(orientation="h"),
                margin=dict(t=40, b=60),
                xaxis=dict(tickangle=0, tickfont=dict(size=11)),
            )
            fig_wk.update_traces(textfont=dict(size=11, color="black"))
            st.plotly_chart(fig_wk, use_container_width=True)

            # ── 주차별 출고량 차트 ──
            st.markdown("### 📦 주차별 출고량")
            fig_wk_qty = go.Figure()
            fig_wk_qty.add_trace(go.Bar(
                x=weekly_view["주차_label"],
                y=weekly_view["출고량"],
                name="출고량",
                marker_color="#1f77b4",
                text=weekly_view["출고량"] if wk_show_label else None,
                texttemplate="%{text:,.0f}",
                textposition="outside",
                cliponaxis=False,
            ))
            fig_wk_qty.update_layout(
                height=380,
                yaxis=dict(range=[0, weekly_view["출고량"].max() * 1.3]),
                margin=dict(t=40, b=60),
                xaxis=dict(tickangle=0, tickfont=dict(size=11)),
            )
            fig_wk_qty.update_traces(textfont=dict(size=11, color="black"))
            st.plotly_chart(fig_wk_qty, use_container_width=True)

            # ── 채널별 주차 매출 히트맵 (expander로 열고 닫기) ──
            with st.expander("🗂️ 채널별 주차 매출 히트맵", expanded=False):
                wk_ch_pivot = wdf.copy()
                wk_ch_pivot = wk_ch_pivot[wk_ch_pivot["연도주차"].isin(weekly_view["연도주차"])]
                # 주차_label 매핑 딕셔너리
                _label_map = dict(zip(weekly_view["연도주차"], weekly_view["주차_label"]))
                wk_ch_pivot["주차_label"] = wk_ch_pivot["연도주차"].map(_label_map)
                wk_ch_pivot = wk_ch_pivot.groupby(["거래처분류", "주차_label"])["품목별매출(VAT제외)"].sum().reset_index()
                wk_ch_pivot = wk_ch_pivot.pivot_table(
                    index="거래처분류", columns="주차_label", values="품목별매출(VAT제외)", aggfunc="sum", fill_value=0
                )
                sorted_labels = weekly_view["주차_label"].tolist()
                wk_ch_pivot = wk_ch_pivot.reindex(columns=sorted_labels, fill_value=0)
                wk_ch_pivot["_total"] = wk_ch_pivot.sum(axis=1)
                wk_ch_pivot = wk_ch_pivot.sort_values("_total", ascending=False).drop(columns="_total").head(15)

                fig_heat = px.imshow(
                    wk_ch_pivot,
                    color_continuous_scale="Blues",
                    aspect="auto",
                    text_auto=True,
                    labels=dict(x="주차", y="채널", color="매출액"),
                )
                fig_heat.update_traces(texttemplate="%{z:,.0f}", textfont=dict(size=10))
                fig_heat.update_layout(height=420, margin=dict(t=30), xaxis=dict(tickangle=0))
                st.plotly_chart(fig_heat, use_container_width=True)

            # ── 주차별 상세 테이블 ──
            with st.expander("📋 주차별 상세 데이터", expanded=False):
                wk_table = weekly_view[["주차_label", "매출액", "매출총이익", "출고량"]].copy()
                wk_table["매출총이익률"] = safe_divide(wk_table["매출총이익"], wk_table["매출액"]) * 100
                wk_table["매출액_WoW"] = wk_table["매출액"].pct_change() * 100
                wk_table = wk_table.rename(columns={"주차_label": "주차"})
                st.dataframe(
                    wk_table.style.format({
                        "매출액": "{:,.0f}",
                        "매출총이익": "{:,.0f}",
                        "출고량": "{:,.0f}",
                        "매출총이익률": "{:.1f}%",
                        "매출액_WoW": "{:+.1f}%",
                    }),
                    use_container_width=True,
                )

# ===================================
# TAB 3: 채널 분석
# ===================================
if "채널분석" in tab_map:
    with tab_map["채널분석"]:
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

        st.subheader("💰 월별 채널별 매출액 및 구성비")
        ch_sales_pivot = pd.pivot_table(
            filtered_df, values="품목별매출(VAT제외)", index="거래처분류",
            columns="출고년월", aggfunc="sum", fill_value=0
        )
        ch_sales_mcols = sort_month_cols(ch_sales_pivot.columns.tolist())
        ch_sales_pivot = ch_sales_pivot.reindex(columns=ch_sales_mcols)
        ch_month_totals = ch_sales_pivot.sum()
        ch_sales_pivot = add_total_row(ch_sales_pivot)
        ch_sales_pivot = sort_pivot_by_last_month(ch_sales_pivot, ch_sales_mcols)

        ch_display_cols = []
        for m in ch_sales_mcols:
            ratio_col = f"{m}_구성비"
            ch_sales_pivot[ratio_col] = ch_sales_pivot[m] / ch_month_totals[m] if ch_month_totals[m] > 0 else 0
            ch_display_cols += [m, ratio_col]

        ch_display = ch_sales_pivot[ch_display_cols].copy()
        tuples = []
        for c in ch_display_cols:
            if c in ch_sales_mcols:
                tuples.append((c, "매출액"))
            else:
                tuples.append((c.replace("_구성비", ""), "구성비"))
        ch_display = safe_multiindex_from_tuples(ch_display, tuples)
        ch_display_fmt = {}
        for m in ch_sales_mcols:
            ch_display_fmt[(m, "매출액")] = "{:,.0f}"
            ch_display_fmt[(m, "구성비")] = "{:.0%}"
        if ch_display.empty:
            st.info("표시할 채널별 매출 데이터가 없습니다.")
        else:
            st.dataframe(ch_display.style.format(ch_display_fmt), use_container_width=True)

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
if "제품분석" in tab_map:
    with tab_map["제품분석"]:
        st.subheader("📦 제품 분석")

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
        prod_month_totals = prod_s_pivot.sum()
        prod_s_pivot     = add_total_row(prod_s_pivot)
        prod_q_by_month  = add_total_row(prod_q_by_month)
        prod_s_pivot     = sort_pivot_by_last_month(prod_s_pivot, prod_s_mcols)
        prod_q_by_month  = prod_q_by_month.reindex(prod_s_pivot.index, fill_value=0)

        prod_display_cols = []
        for m in prod_s_mcols:
            ratio_col = f"{m}_구성비"
            price_col = f"{m}_개당단가"
            month_total = prod_month_totals[m] if prod_month_totals[m] > 0 else 1
            prod_s_pivot[ratio_col] = prod_s_pivot[m] / month_total
            qty_series = prod_q_by_month[m].reindex(prod_s_pivot.index).fillna(0)
            prod_s_pivot[price_col] = np.where(qty_series != 0, prod_s_pivot[m] / qty_series, 0)
            prod_display_cols += [m, ratio_col, price_col]

        prod_display = prod_s_pivot[prod_display_cols].copy()
        tuples_prod = []
        for c in prod_display_cols:
            if c in prod_s_mcols:
                tuples_prod.append((c, "매출액"))
            elif "_구성비" in c:
                tuples_prod.append((c.replace("_구성비", ""), "구성비"))
            else:
                tuples_prod.append((c.replace("_개당단가", ""), "개당단가"))
        prod_display = safe_multiindex_from_tuples(prod_display, tuples_prod)
        prod_display_fmt = {}
        for m in prod_s_mcols:
            prod_display_fmt[(m, "매출액")]   = "{:,.0f}"
            prod_display_fmt[(m, "구성비")]   = "{:.0%}"
            prod_display_fmt[(m, "개당단가")] = "{:,.0f}"
        if prod_display.empty:
            st.info("표시할 제품별 매출 데이터가 없습니다.")
        else:
            st.dataframe(prod_display.style.format(prod_display_fmt), use_container_width=True)

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
# ===================================
# TAB 4: 공헌이익 분석 (국내)
# ===================================
def _render_contrib_tab(base_df, market_filter, tab_label, ad_apply):
    """공헌이익 분석 공통 렌더링 함수.
    base_df     : filtered_df 기준 (이미 권한 필터 적용된 전체 df)
    market_filter: "국내" / "해외" / None(통합)
    tab_label   : subheader 표시용 문자열
    ad_apply    : 광고비 안분 여부 (해외는 False)
    """
    # ── 시장구분 필터 적용 ──
    if market_filter == "국내":
        mdf = base_df[base_df["국내여부"] == "국내"].copy()
    elif market_filter == "해외":
        mdf = base_df[base_df["국내여부"] != "국내"].copy()
    else:
        mdf = base_df.copy()

    st.subheader(tab_label)
    st.caption("※ 물류비 / 광고비는 COST_INPUT 시트에서 수정됩니다")

    if mdf.empty:
        st.info("해당 조건에 데이터가 없습니다.")
        return

    # ── 비용 재계산 (market_filter된 mdf 기준) ──
    mdf["물류비"] = 0.0
    mdf["광고비"] = 0.0

    for m in selected_months:
        m_mask = mdf["출고년월"] == m

        # 물류비: 부서별 배분
        _depts_m2 = set(dept for (dept, ym) in st.session_state["logistics_table"].keys() if ym == m)
        for dept in _depts_m2:
            dept_amt = st.session_state["logistics_table"].get((dept, m), 0)
            if dept_amt <= 0:
                continue
            # 분모: 해당 부서 전체 매출 (df 기준)
            dept_tot = df.loc[(df["출고년월"] == m) & (df["담당부서"] == dept), "품목별매출(VAT제외)"].sum()
            if dept_tot <= 0:
                continue
            f_mask = m_mask & (mdf["담당부서"] == dept)
            mdf.loc[f_mask, "물류비"] = mdf.loc[f_mask, "품목별매출(VAT제외)"] / dept_tot * dept_amt

        # 광고비: 국내만 안분, 해외는 0
        if ad_apply:
            _igs_m2 = set(ig for (ig, ym) in st.session_state["ad_cost_monthly"].keys() if ym == m)
            for ig in _igs_m2:
                ad_amt = st.session_state["ad_cost_monthly"].get((ig, m), 0)
                if ad_amt <= 0:
                    continue
                # 분모: 품목군 국내 전체 매출 (df 기준)
                ig_tot = df.loc[
                    (df["출고년월"] == m) & (df["품목군"] == ig) & (df["국내여부"] == "국내"),
                    "품목별매출(VAT제외)"
                ].sum()
                if ig_tot <= 0:
                    continue
                f_mask = m_mask & (mdf["품목군"] == ig) & (mdf["국내여부"] == "국내")
                mdf.loc[f_mask, "광고비"] = mdf.loc[f_mask, "품목별매출(VAT제외)"] / ig_tot * ad_amt

    mdf["공헌이익"] = (
        mdf["품목별매출(VAT제외)"]
        - mdf["원가총액"]
        - mdf["채널수수료"]
        - mdf["물류비"]
        - mdf["광고비"]
    )
    mdf["공헌이익률"] = safe_divide(mdf["공헌이익"], mdf["품목별매출(VAT제외)"])

    # ── 품목군별 공헌이익 ──
    st.markdown("### 📦 품목군별 공헌이익")
    pc = mdf.groupby("품목군", as_index=False)[[
        "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
        "매출총이익", "채널수수료", "물류비", "광고비"
    ]].sum()
    pc["비용"] = 0.0
    _ch_dept_map3 = st.session_state.get("channel_dept_map", {})
    for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
        if ym not in selected_months or ch not in selected_channel_groups:
            continue
        # 해외탭: 해당 거래처가 해외인 경우만 / 국내탭: 국내인 경우만 / 통합: 전체
        if market_filter is not None:
            ch_is_dom = (st.session_state.get("channel_dept_map", {}).get(ch, "") != "") and \
                        mdf[mdf["거래처분류"] == ch]["국내여부"].eq("국내").any() if not mdf[mdf["거래처분류"] == ch].empty else True
            # 단순하게 mdf에 해당 채널이 있는지로 판단
            if mdf[mdf["거래처분류"] == ch].empty:
                continue
        mask = pc["품목군"] == ig
        if mask.any():
            pc.loc[mask, "비용"] += amt

    pc["공헌이익"] = (
        pc["품목별매출(VAT제외)"] - pc["원가총액"] - pc["채널수수료"]
        - pc["물류비"] - pc["광고비"] - pc["비용"]
    )
    pc["공헌이익률"] = safe_divide(pc["공헌이익"], pc["품목별매출(VAT제외)"])
    # 내부 레이블 행 제외 (표시용)
    pc = pc[~pc["품목군"].isin(["__매출조정__", "__미분류__"])]
    pc = pc[["품목군", "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
             "매출총이익", "채널수수료", "물류비", "광고비", "비용", "공헌이익", "공헌이익률"]]
    st.dataframe(pc.style.format({
        "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
        "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}", "채널수수료": "{:,.0f}",
        "물류비": "{:,.0f}", "광고비": "{:,.0f}", "비용": "{:,.0f}",
        "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
    }), use_container_width=True)

    # ── 채널별 공헌이익 ──
    with st.expander("🏪 채널별 공헌이익", expanded=False):
        cc = mdf.groupby("거래처분류", as_index=False)[[
            "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
            "매출총이익", "채널수수료", "물류비", "광고비"
        ]].sum()
        cc["비용"] = 0.0
        for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
            if ym not in selected_months:
                continue
            if market_filter is not None and mdf[mdf["거래처분류"] == ch].empty:
                continue
            mask = cc["거래처분류"] == ch
            if mask.any():
                cc.loc[mask, "비용"] += amt
        cc["공헌이익"] = (
            cc["품목별매출(VAT제외)"] - cc["원가총액"] - cc["채널수수료"]
            - cc["물류비"] - cc["광고비"] - cc["비용"]
        )
        cc["공헌이익률"] = safe_divide(cc["공헌이익"], cc["품목별매출(VAT제외)"])
        cc = cc[["거래처분류", "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                 "매출총이익", "채널수수료", "물류비", "광고비", "비용", "공헌이익", "공헌이익률"]]
        st.dataframe(cc.style.format({
            "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
            "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}", "채널수수료": "{:,.0f}",
            "물류비": "{:,.0f}", "광고비": "{:,.0f}", "비용": "{:,.0f}",
            "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
        }), use_container_width=True)


if "공헌이익분석(국내)" in tab_map:
    with tab_map["공헌이익분석(국내)"]:
        _render_contrib_tab(filtered_df, "국내", "📊 공헌이익 분석 (국내)", ad_apply=True)

# ===================================
# TAB 5: 공헌이익 분석 (해외)
# ===================================
if "공헌이익분석(해외)" in tab_map:
    with tab_map["공헌이익분석(해외)"]:
        _render_contrib_tab(filtered_df, "해외", "🌏 공헌이익 분석 (해외)", ad_apply=False)

# ===================================
# TAB 6: 공헌이익 분석 (통합)
# ===================================
if "공헌이익분석(통합)" in tab_map:
    with tab_map["공헌이익분석(통합)"]:
        st.subheader("📋 공헌이익 분석 (통합)")
        st.caption("※ 물류비 / 광고비는 COST_INPUT 시트에서 수정됩니다")

        temp_df = df.copy()
        temp_df["물류비"] = 0.0

        # [수정] 물류비: 부서별 배분 (temp_df = 전체 df 기준)
        for m in all_months:
            _depts_m = set(
                dept for (dept, ym) in st.session_state["logistics_table"].keys() if ym == m
            )
            for dept in _depts_m:
                dept_logistics_amt = st.session_state["logistics_table"].get((dept, m), 0)
                if dept_logistics_amt <= 0:
                    continue
                dept_total_mask  = (temp_df["출고년월"] == m) & (temp_df["담당부서"] == dept)
                dept_total_sales = temp_df.loc[dept_total_mask, "품목별매출(VAT제외)"].sum()
                if dept_total_sales <= 0:
                    continue
                ratio = temp_df.loc[dept_total_mask, "품목별매출(VAT제외)"] / dept_total_sales
                temp_df.loc[dept_total_mask, "물류비"] = ratio * dept_logistics_amt

        temp_df["광고비"] = 0.0
        for m in all_months:
            # [수정] 광고비: 품목군 기준 + 국내 채널만 배분 (해외 채널 제외)
            _igs_m = set(
                ig for (ig, ym) in st.session_state["ad_cost_monthly"].keys() if ym == m
            )
            for ig in _igs_m:
                ad_amt = st.session_state["ad_cost_monthly"].get((ig, m), 0)
                if ad_amt <= 0:
                    continue
                ig_total_mask  = (temp_df["출고년월"] == m) & (temp_df["품목군"] == ig) & (temp_df["국내여부"] == "국내")
                ig_total_sales = temp_df.loc[ig_total_mask, "품목별매출(VAT제외)"].sum()
                if ig_total_sales <= 0:
                    continue
                ratio = temp_df.loc[ig_total_mask, "품목별매출(VAT제외)"] / ig_total_sales
                temp_df.loc[ig_total_mask, "광고비"] = ratio * ad_amt

        temp_df = temp_df[
            (temp_df["출고년월"].isin(selected_months)) &
            (temp_df["거래처분류"].isin(selected_channel_groups)) &
            (temp_df["내품상품명"].isin(selected_items))
        ].copy()
        temp_df["매출총이익"] = temp_df["품목별매출(VAT제외)"] - temp_df["원가총액"]

        product_contrib = (
            temp_df.groupby("품목군", as_index=False)[[
                "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                "매출총이익", "채널수수료", "물류비", "광고비"
            ]].sum()
        )
        product_contrib["비용"] = 0.0

        _ch_dept_map = st.session_state.get("channel_dept_map", {})
        for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
            if ym not in selected_months or ch not in selected_channel_groups:
                continue
            _ch_dept = _ch_dept_map.get(ch, "")
            mask = product_contrib["품목군"] == ig
            if mask.any():
                product_contrib.loc[mask, "비용"] += amt

        product_contrib["공헌이익"] = (
            product_contrib["품목별매출(VAT제외)"]
            - product_contrib["원가총액"]
            - product_contrib["채널수수료"]
            - product_contrib["물류비"]
            - product_contrib["광고비"]
            - product_contrib["비용"]
        )
        product_contrib["공헌이익률"] = safe_divide(product_contrib["공헌이익"], product_contrib["품목별매출(VAT제외)"])
        product_contrib = product_contrib[~product_contrib["품목군"].isin(["__매출조정__", "__미분류__"])]
        product_contrib = product_contrib[[
            "품목군", "총내품출고수량", "품목별매출(VAT제외)",
            "원가총액", "매출총이익", "채널수수료",
            "물류비", "광고비", "비용", "공헌이익", "공헌이익률"
        ]]

        # ── 1. 품목군별 공헌이익 (항상 열림) ──
        st.markdown("### 📦 품목군별 공헌이익")
        st.dataframe(
            product_contrib.style.format({
                "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
                "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}",
                "채널수수료": "{:,.0f}", "물류비": "{:,.0f}",
                "광고비": "{:,.0f}", "비용": "{:,.0f}",
                "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
            }),
            use_container_width=True
        )

        # ── 2. 채널별 공헌이익 (닫힘) ──
        with st.expander("🏪 채널별 공헌이익", expanded=False):
            channel_contrib = (
                temp_df.groupby("거래처분류", as_index=False)[[
                    "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                    "매출총이익", "채널수수료", "물류비", "광고비"
                ]].sum()
            )
            channel_contrib["비용"] = 0.0
            _ch_dept_map2 = st.session_state.get("channel_dept_map", {})
            for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
                if ym not in selected_months:
                    continue
                mask = channel_contrib["거래처분류"] == ch
                if mask.any():
                    channel_contrib.loc[mask, "비용"] += amt
            channel_contrib["공헌이익"] = (
                channel_contrib["품목별매출(VAT제외)"]
                - channel_contrib["원가총액"]
                - channel_contrib["채널수수료"]
                - channel_contrib["물류비"]
                - channel_contrib["광고비"]
                - channel_contrib["비용"]
            )
            channel_contrib["공헌이익률"] = safe_divide(channel_contrib["공헌이익"], channel_contrib["품목별매출(VAT제외)"])
            channel_contrib = channel_contrib[[
                "거래처분류", "총내품출고수량", "품목별매출(VAT제외)",
                "원가총액", "매출총이익", "채널수수료",
                "물류비", "광고비", "비용", "공헌이익", "공헌이익률"
            ]]
            st.dataframe(
                channel_contrib.style.format({
                    "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
                    "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}",
                    "채널수수료": "{:,.0f}", "물류비": "{:,.0f}",
                    "광고비": "{:,.0f}", "비용": "{:,.0f}",
                    "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
                }),
                use_container_width=True
            )

        # ── 3. 부서별 월별 물류비 (닫힘) ──
        with st.expander("🚚 부서별 월별 물류비", expanded=False):
            _lt = st.session_state["logistics_table"]
            if _lt:
                _depts_all = sorted(set(dept for (dept, ym) in _lt.keys()))
                _logistics_rows = []
                for dept in _depts_all:
                    row = {"부서": dept}
                    for m in all_months:
                        row[m] = _lt.get((dept, m), 0)
                    _logistics_rows.append(row)
                _logistics_display = pd.DataFrame(_logistics_rows).set_index("부서")
                st.dataframe(_logistics_display.style.format("{:,.0f}"), use_container_width=True)
            else:
                st.info("COST_INPUT 시트에 물류비 데이터가 없습니다.")

        # ── 4. 품목군별 광고비 (닫힘) ──
        with st.expander("📢 품목군별 광고비", expanded=False):
            ad_data = []
            item_groups_ad = sorted(set(k[0] for k in st.session_state["ad_cost_monthly"].keys()))
            for ig in item_groups_ad:
                row = {"품목군": ig}
                for m in all_months:
                    row[m] = st.session_state["ad_cost_monthly"].get((ig, m), 0)
                ad_data.append(row)
            ad_df = pd.DataFrame(ad_data)
            ad_mcols = [c for c in ad_df.columns if c != "품목군"]
            st.dataframe(ad_df.style.format({m: "{:,.0f}" for m in ad_mcols}), use_container_width=True)

        # ── 5. 채널별 후정산 비용 (닫힘) ──
        with st.expander("💸 채널별 후정산 비용", expanded=False):
            if st.session_state["channel_cost"]:
                _ch_dept_map_disp = st.session_state.get("channel_dept_map", {})
                cc_rows = [
                    {
                        "년월": ym,
                        "거래처명": ch,
                        "담당부서": _ch_dept_map_disp.get(ch, "-"),
                        "품목군": ig,
                        "비용(VAT-)": amt,
                    }
                    for (ym, ch, ig), amt in st.session_state["channel_cost"].items()
                ]
                cc_df = pd.DataFrame(cc_rows).sort_values(["년월", "담당부서", "거래처명", "품목군"])
                st.dataframe(cc_df.style.format({"비용(VAT-)": "{:,.0f}"}), use_container_width=True)
            else:
                st.info("CHANNEL_COST 시트에 데이터가 없습니다.")

# ===================================
# TAB 7: 다운로드
# ===================================
if "다운로드" in tab_map:
    with tab_map["다운로드"]:
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
            "월별채널매출": dl_ch_sales, "월별채널출고량": dl_ch_qty,
            "월별제품매출": dl_prod_sales, "월별제품출고량": dl_prod_qty,
        })
        st.download_button(
            label="📥 분석 결과 통합 엑셀 다운로드",
            data=download_file,
            file_name="Lingtea_Dashboard_v7.3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("### 포함 시트")
        st.write("- 월별 채널 매출액")
        st.write("- 월별 채널 출고량")
        st.write("- 월별 제품 매출액")
        st.write("- 월별 제품 출고량")

# ===================================
# TAB 6: 관리자 (admin only)
# ===================================
if st.session_state["role"] == "admin":
    with admin_tab:
        st.subheader("⚙️ 관리자 페이지")

        admin_sub = st.tabs(["👥 계정 권한 관리", "➕ 사용자 역할 변경", "🗑️ 계정 관리"])

        # ── 계정 권한 관리 ──
        with admin_sub[0]:
            st.markdown("### 계정별 탭 접근 권한")
            st.caption("관리자 계정은 항상 전체 접근 가능합니다. 일반 사용자의 탭별 권한을 설정하세요.")

            all_users = get_all_users()
            user_list = [u for u in all_users if u.get("role") != "admin"]

            if not user_list:
                st.info("권한 설정이 필요한 일반 사용자 계정이 없습니다.")
            else:
                for user in user_list:
                    uid   = user["uid"]
                    email = user.get("email", uid)
                    tabs  = user.get("tabs", DEFAULT_USER_TABS.copy())
                    disabled = user.get("disabled", False)

                    status_badge = "🔴 비활성" if disabled else "🟢 활성"
                    with st.expander(f"{status_badge}  {email}", expanded=False):
                        st.markdown(f"**UID:** `{uid}`")
                        new_tabs = {}
                        cols = st.columns(len(ALL_TABS))
                        for i, tab_name in enumerate(ALL_TABS):
                            with cols[i]:
                                new_tabs[tab_name] = st.toggle(
                                    tab_name,
                                    value=tabs.get(tab_name, False),
                                    key=f"toggle_{uid}_{tab_name}"
                                )
                        if st.button("💾 저장", key=f"save_{uid}", use_container_width=True):
                            update_user_tabs(uid, new_tabs)
                            st.success(f"{email} 권한이 저장되었습니다.")
                            st.rerun()

        # ── 사용자 역할 변경 ──
        with admin_sub[1]:
            st.markdown("### 사용자 역할 변경")
            all_users = get_all_users()
            admin_emails_list = list(st.secrets["auth"]["admin_emails"])

            # [신규] ITEM_MASTER에서 품목군 목록 로드 (관리자 UI용)
            @st.cache_data(ttl=600)
            def load_item_group_list():
                _client  = get_gspread_client()
                _ws      = _client.open_by_key(SHEET_ID).worksheet("ITEM_MASTER")
                _data    = _ws.get_all_values()
                if not _data or len(_data) < 2:
                    return []
                _headers = [str(h).strip() for h in _data[0]]
                _df      = pd.DataFrame(_data[1:], columns=_headers)
                _col     = next((c for c in _headers if "품목군" in c), None)
                if _col is None:
                    return []
                return sorted(_df[_col].dropna().astype(str).str.strip().unique().tolist())

            _item_group_options = load_item_group_list()

            # [신규] AUTH_MASTER 행 업데이트 함수 (권한유형 + 품목군 저장)
            def update_auth_master_row(target_email: str, role_type: str, item_groups: list):
                """AUTH_MASTER에서 해당 이메일 행의 권한유형/품목군 컬럼을 업데이트."""
                _client  = get_gspread_client_rw()
                _ws      = _client.open_by_key(SHEET_ID).worksheet("AUTH_MASTER")
                _data    = _ws.get_all_values()
                if not _data:
                    return False
                _headers       = [str(h).strip() for h in _data[0]]
                # 컬럼 인덱스 탐색 (1-based for gspread update_cell)
                _email_idx     = next((i+1 for i, h in enumerate(_headers) if h.lower().replace("-","") == "email"), None)
                _role_type_idx = next((i+1 for i, h in enumerate(_headers) if "권한유형" in h), None)
                _item_idx      = next((i+1 for i, h in enumerate(_headers) if "품목군"   in h), None)
                if not all([_email_idx, _role_type_idx, _item_idx]):
                    return False
                # 행 탐색 후 업데이트
                for row_idx, row in enumerate(_data[1:], start=2):
                    if str(row[_email_idx-1]).strip().lower() == target_email.strip().lower():
                        _ws.update_cell(row_idx, _role_type_idx, role_type)
                        _ws.update_cell(row_idx, _item_idx, ",".join(item_groups))
                        return True
                return False

            # [핵심 수정] AUTH_MASTER를 루프 밖에서 1회만 로드 — 429 Quota 방지
            _auth_df_adm, _ec, _dc, _rtc, _igc = load_auth_master()

            # ── 유저별 렌더링 ──
            for user in all_users:
                uid   = user["uid"]
                email = user.get("email", uid)
                role  = user.get("role", "user")

                # secrets에 등록된 admin은 변경 불가
                if email in admin_emails_list:
                    st.markdown(f"🔒 **{email}** — 고정 관리자 (변경 불가)")
                    continue

                col_a, col_b = st.columns([3, 1])
                with col_a:
                    new_role = st.selectbox(
                        email,
                        options=["user", "admin"],
                        index=0 if role == "user" else 1,
                        key=f"role_{uid}"
                    )
                with col_b:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("변경", key=f"change_role_{uid}"):
                        update_user_role(uid, new_role)
                        if new_role == "admin":
                            update_user_tabs(uid, DEFAULT_ADMIN_TABS.copy())
                        st.success(f"{email} 역할이 {new_role}(으)로 변경되었습니다.")
                        st.rerun()

                # [수정] expander 내부에서 load_auth_master() 재호출 제거
                # 루프 밖에서 받아온 _auth_df_adm 재사용 → API 호출 1회로 고정
                with st.expander(f"🔐 {email} — AUTH_MASTER 권한 설정", expanded=False):
                    _row_adm = (
                        _auth_df_adm[_auth_df_adm[_ec] == email.lower()]
                        if _ec else pd.DataFrame()
                    )

                    _cur_role_type   = "부서기반"
                    _cur_item_groups = []
                    if not _row_adm.empty and _rtc:
                        _cur_role_type = str(_row_adm.iloc[0][_rtc]).strip()
                    if not _row_adm.empty and _igc:
                        _raw_ig = str(_row_adm.iloc[0][_igc]).strip()
                        # ALL은 multiselect에서 제외 (저장 시 자동 처리)
                        _cur_item_groups = [
                            x.strip() for x in _raw_ig.split(",")
                            if x.strip() and x.strip() != "ALL"
                        ]

                    # (1) 권한유형 dropdown
                    _new_role_type = st.selectbox(
                        "권한유형",
                        options=["관리자", "부서기반", "PM"],
                        index=(
                            ["관리자", "부서기반", "PM"].index(_cur_role_type)
                            if _cur_role_type in ["관리자", "부서기반", "PM"] else 1
                        ),
                        key=f"auth_role_type_{uid}"
                    )

                    # (2) 품목군 multiselect — PM일 때만 의미 있음, 항상 표시
                    _new_items = st.multiselect(
                        "품목군 (PM 권한 시 적용 / 관리자·부서기반은 ALL로 자동 저장)",
                        options=_item_group_options,
                        default=[ig for ig in _cur_item_groups if ig in _item_group_options],
                        key=f"auth_items_{uid}"
                    )

                    # (3) 저장 버튼 — 관리자/부서기반은 품목군을 ALL로 저장
                    if st.button("💾 AUTH_MASTER 저장", key=f"save_auth_{uid}", use_container_width=True):
                        _save_items = _new_items if _new_role_type == "PM" else ["ALL"]
                        _ok = update_auth_master_row(email, _new_role_type, _save_items)
                        if _ok:
                            # 저장 후 캐시 무효화 → 다음 렌더링 시 최신값 반영
                            load_auth_master.clear()
                            st.success(
                                f"✅ {email} 권한 저장 완료 "
                                f"(권한유형: {_new_role_type} / 품목군: {','.join(_save_items)})"
                            )
                            st.rerun()
                        else:
                            st.error(
                                "❌ 저장 실패. AUTH_MASTER에 해당 이메일 행이 없거나 "
                                "권한유형/품목군 컬럼 구조를 확인하세요."
                            )

        # ── 계정 관리 (비활성화 / 삭제) ──
        with admin_sub[2]:
            st.markdown("### 계정 비활성화 / 삭제")
            st.warning("⚠️ 삭제된 계정은 복구할 수 없습니다.")
            all_users = get_all_users()
            admin_emails_list = list(st.secrets["auth"]["admin_emails"])

            for user in all_users:
                uid      = user["uid"]
                email    = user.get("email", uid)
                disabled = user.get("disabled", False)

                if email in admin_emails_list:
                    st.markdown(f"🔒 **{email}** — 고정 관리자 (관리 불가)")
                    continue

                col_x, col_y, col_z = st.columns([3, 1, 1])
                with col_x:
                    status = "🔴 비활성" if disabled else "🟢 활성"
                    st.markdown(f"{status} **{email}**")
                with col_y:
                    btn_label = "활성화" if disabled else "비활성화"
                    if st.button(btn_label, key=f"disable_{uid}"):
                        if disabled:
                            firebase_auth.update_user(uid, disabled=False)
                            db.collection("users").document(uid).update({
                                "disabled": False,
                                "updated_at": datetime.now().isoformat()
                            })
                            st.success(f"{email} 계정이 활성화되었습니다.")
                        else:
                            disable_user(uid)
                            st.success(f"{email} 계정이 비활성화되었습니다.")
                        st.rerun()
                with col_z:
                    if st.button("🗑️ 삭제", key=f"delete_{uid}", type="primary"):
                        delete_user(uid)
                        st.success(f"{email} 계정이 삭제되었습니다.")
                        st.rerun()

# ===================================
# TAB 7: 제품별 원가
# ===================================
if "제품별원가" in tab_map:
    with tab_map["제품별원가"]:
        st.subheader("💰 제품별 원가")

        @st.cache_data(ttl=600)
        def load_cost_master_table():
            client = get_gspread_client()
            ws   = client.open_by_key(SHEET_ID).worksheet("COST_MASTER")
            data = ws.get_all_values()
            if not data or len(data) < 2:
                return pd.DataFrame()
            raw_headers = [str(h).strip() for h in data[0]]
            cm_df = pd.DataFrame(data[1:], columns=raw_headers).copy()
            return cm_df

        cm_df = load_cost_master_table()

        if cm_df.empty:
            st.warning("COST_MASTER 시트에 데이터가 없습니다.")
        else:
            # 컬럼명 정규화
            cm_df.columns = [str(c).strip() for c in cm_df.columns]

            # 상품코드 컬럼 탐색
            code_col  = next((c for c in cm_df.columns if "상품코드" in c), None)
            name_col  = next((c for c in cm_df.columns if "상품명"  in c), None)
            cost_col  = next((c for c in cm_df.columns if "원가"    in c), None)
            ym_col    = next((c for c in cm_df.columns if "년월"    in c), None)

            if not all([code_col, name_col, cost_col, ym_col]):
                st.error(f"COST_MASTER 컬럼 인식 실패. 현재 컬럼: {cm_df.columns.tolist()}")
            else:
                # (1) 상품코드 필터: 숫자 8자리 + 3/4/7로 시작하는 코드만
                cm_df[code_col] = cm_df[code_col].astype(str).str.strip()
                cm_df = cm_df[
                    cm_df[code_col].str.isdigit() &
                    (cm_df[code_col].str.len() == 8) &
                    (cm_df[code_col].str[0].isin(["3", "4", "7"]))
                ].copy()

                # (2) 제품원가 숫자 변환 (콤마 제거)
                cm_df[cost_col] = pd.to_numeric(
                    cm_df[cost_col].astype(str).str.replace(",", "", regex=False).str.strip(),
                    errors="coerce"
                )

                # (3) Pivot 생성: index=[상품코드, 상품명], columns=년월, values=제품원가
                pivot_df = cm_df.pivot_table(
                    index=[code_col, name_col],
                    columns=ym_col,
                    values=cost_col,
                    aggfunc="mean"
                )

                # (4) 월 컬럼 YYYY-MM 기준 오름차순 정렬
                month_cols = sort_month_cols(pivot_df.columns.tolist())
                pivot_df   = pivot_df[month_cols] if month_cols else pivot_df

                # (5) 평균 컬럼 추가 (NaN 자동 제외)
                pivot_df["평균"] = pivot_df.mean(axis=1) if len(pivot_df.columns) > 0 else np.nan

                # (6) 컬럼명 rename 및 index reset
                pivot_df = pivot_df.reset_index()
                pivot_df = pivot_df.rename(columns={
                    code_col: "제품코드",
                    name_col: "제품명",
                })

                # (7) 숫자 컬럼 포맷 (천단위 콤마, NaN 유지)
                num_cols    = [c for c in pivot_df.columns if c not in ["제품코드", "제품명"]]
                fmt_dict    = {c: "{:,.2f}" for c in num_cols}

                st.caption(f"총 {len(pivot_df)}개 제품 | 기간: {month_cols[0] if month_cols else '-'} ~ {month_cols[-1] if month_cols else '-'}")
                if pivot_df.empty:
                    st.info("표시할 제품별 원가 데이터가 없습니다.")
                else:
                    st.dataframe(
                        pivot_df.style.format(fmt_dict, na_rep="-"),
                        use_container_width=True,
                        height=500
                    )

st.success("🚀 Lingtea Dashboard v8.5 Ready")
