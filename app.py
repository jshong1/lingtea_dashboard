import io
import os
import uuid
import requests
from datetime import datetime, timedelta

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
from streamlit_cookies_manager import EncryptedCookieManager

# ── AI 공용 함수 (Claude API 호출) ──
def call_claude_api(system_prompt, messages, max_tokens=2048):
    api_key = st.secrets["anthropic"]["api_key"]
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": "claude-sonnet-4-5",
            "max_tokens": max_tokens,
            "system": system_prompt,
            "messages": messages,
        },
        timeout=60,
    )
    if resp.status_code != 200:
        raise Exception(f"API 오류 {resp.status_code}: {resp.text}")
    return resp.json()["content"][0]["text"]

# ── AI 공용 함수 (데이터 컨텍스트 생성) ──
def build_common_ai_context(fdf):
    lines = []
    _ems = sorted(fdf["출고년월"].dropna().unique().tolist())
    if not _ems: return "데이터가 없습니다."
    lines.append(f"## 분석 기간: {_ems[0]} ~ {_ems[-1]} ({len(_ems)}개월)")
    _es = fdf["품목별매출(VAT제외)"].sum()
    _eq = fdf["총내품출고수량"].sum()
    _egp = fdf["매출총이익"].sum()
    _eci = fdf["공헌이익"].sum()
    lines.append("\n## 전체 KPI")
    lines.append(f"- 총매출(VAT제외): {_es:,.0f}원")
    lines.append(f"- 총출고수량: {_eq:,.0f}")
    lines.append(f"- 매출총이익: {_egp:,.0f}원 ({_egp/_es*100:.1f}%)" if _es else "- 매출총이익: 0")
    lines.append(f"- 공헌이익: {_eci:,.0f}원 ({_eci/_es*100:.1f}%)" if _es else "- 공헌이익: 0")
    _emly = fdf.groupby("출고년월")[["품목별매출(VAT제외)", "공헌이익"]].sum().reset_index().sort_values("출고년월")
    lines.append("\n## 월별 추이")
    for _, _er in _emly.iterrows():
        _ms2 = _er["품목별매출(VAT제외)"]; _ci2 = _er["공헌이익"]
        lines.append(f"- {_er['출고년월']}: 매출 {_ms2:,.0f}원 / 공헌이익 {_ci2:,.0f}원 ({_ci2/_ms2*100:.1f}%)" if _ms2 else f"- {_er['출고년월']}: 매출 0")
    _ech = fdf.groupby("거래처분류")[["품목별매출(VAT제외)", "공헌이익"]].sum().sort_values("품목별매출(VAT제외)", ascending=False).head(15)
    lines.append("\n## 주요 채널별 실적 (상위 15개)")
    for _cn, _cr in _ech.iterrows():
        _cs = _cr["품목별매출(VAT제외)"]; _cci = _cr["공헌이익"]
        lines.append(f"- {_cn}: 매출 {_cs:,.0f}원 / 공헌이익 {_cci:,.0f}원 ({_cci/_cs*100:.1f}%)" if _cs else f"- {_cn}: 매출 0")
    _eig = fdf[~fdf["품목군"].isin(["__매출조정__","__미분류__"])].groupby("품목군")[["품목별매출(VAT제외)","공헌이익","총내품출고수량"]].sum().sort_values("품목별매출(VAT제외)", ascending=False)
    lines.append("\n## 품목군별 실적")
    for _in2, _ir in _eig.iterrows():
        _is = _ir["품목별매출(VAT제외)"]; _ici = _ir["공헌이익"]
        lines.append(f"- {_in2}: 매출 {_is:,.0f}원 / 공헌이익 {_ici:,.0f}원 ({_ici/_is*100:.1f}%)" if _is else f"- {_in2}: 매출 0")
    return "\n".join(lines)

# -----------------------------------
# 기본 설정
# -----------------------------------
st.set_page_config(page_title="Lingtea Dashboard", layout="wide")

SHEET_ID = "1d_TZiPZZbETyoB61PrsXVZsP5p9qsaXFgKcEgHUC_sk"

ALL_TABS = ["경영진요약", "월별추이", "주차별추이", "채널분석", "제품분석", "YoY분석", "공헌이익분석(통합)", "공헌이익분석(국내)", "공헌이익분석(해외)", "제품별원가", "확정비교", "AI분석", "다운로드"]

DEFAULT_USER_TABS = {t: False for t in ALL_TABS}
DEFAULT_ADMIN_TABS = {t: True for t in ALL_TABS}

tab_defs = {
    "경영진요약":        "👔 경영진 요약",
    "월별추이":         "📈 월별 추이",
    "주차별추이":       "📅 주차별 추이",
    "채널분석":         "🏪 채널 분석",
    "제품분석":         "📦 제품 분석",
    "YoY분석":          "📊 YoY 분석",
    "공헌이익분석(국내)": "📊 공헌이익(국내)",
    "공헌이익분석(해외)": "🌏 공헌이익(해외)",
    "공헌이익분석(통합)": "📋 공헌이익(통합)",
    "제품별원가":       "💰 제품별 원가",
    "확정비교":         "⚖️ 확정 비교 분석",
    "AI분석":           "🤖 AI 분석",
    "다운로드":         "📥 다운로드",
}

def tab_allowed(tab_name: str) -> bool:
    if st.session_state.get("role") == "admin":
        return True
    return st.session_state.get("tabs_perm", {}).get(tab_name, False)


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
# 쿠키 매니저 초기화 (암호화된 세션 쿠키)
# secrets.toml에 [cookie] / password = "..." 필요
# -----------------------------------
cookies = EncryptedCookieManager(
    prefix="lingtea_",
    password=st.secrets["cookie"]["password"],
)
if not cookies.ready():
    st.stop()

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
# 대시보드 메인 CSS
# -----------------------------------
DASHBOARD_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap');

/* 전체 배경 및 폰트 */
[data-testid="stAppViewContainer"] {
    background-color: #F8F9FA !important;
}
[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid #E9ECEF !important;
}
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', sans-serif !important;
}

/* 사이드바 메뉴 스타일 */
.stRadio > div {
    gap: 0px !important;
}
.stRadio label {
    background-color: transparent !important;
    border-radius: 0px !important;
    padding: 12px 16px !important;
    transition: all 0.2s !important;
    border-bottom: 1px solid #F1F3F5 !important;
    margin: 0 !important;
    cursor: pointer !important;
}
.stRadio label:hover {
    background-color: #F8F9FA !important;
}
.stRadio [data-testid="stWidgetLabel"] {
    display: none !important;
}


/* 카드 스타일 컨테이너 (구분선 스타일로 변경) */
.dashboard-card {
    background-color: transparent;
    padding: 24px 0px;
    border-radius: 0px;
    box-shadow: none;
    border: none;
    border-bottom: 1px solid #E9ECEF;
    margin-bottom: 32px;
}

/* KPI 메트릭 스타일 */
[data-testid="stMetricValue"] {
    font-weight: 700 !important;
    color: #1A1C1E !important;
}
[data-testid="stMetricDelta"] {
    font-weight: 500 !important;
}

/* 헤더 스타일 */
.main-header {
    font-size: 2rem;
    font-weight: 700;
    color: #1A1C1E;
    margin-bottom: 0.5rem;
}
.sub-header {
    font-size: 1rem;
    color: #6C757D;
    margin-bottom: 2rem;
}

/* 필터 섹션 (구분선 스타일로 변경) */
.filter-container {
    background-color: transparent;
    padding: 20px 0px;
    border-radius: 0px;
    box-shadow: none;
    border: none;
    border-bottom: 1px solid #E9ECEF;
    margin-bottom: 24px;
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
# Firestore 세션 관리 (쿠키 기반)
# -----------------------------------
SESSION_TTL_DAYS = 30

def create_session(uid: str) -> str:
    """랜덤 세션 토큰 생성 → Firestore sessions 컬렉션에 저장"""
    token = str(uuid.uuid4())
    expires_at = datetime.now() + timedelta(days=SESSION_TTL_DAYS)
    db.collection("sessions").document(token).set({
        "uid": uid,
        "created_at": datetime.now().isoformat(),
        "expires_at": expires_at.isoformat(),
    })
    return token

def get_session(token: str):
    """토큰으로 세션 조회. 만료/없으면 None 반환"""
    if not token:
        return None
    try:
        doc = db.collection("sessions").document(token).get()
        if not doc.exists:
            return None
        data = doc.to_dict()
        expires_at = datetime.fromisoformat(data["expires_at"])
        if datetime.now() > expires_at:
            db.collection("sessions").document(token).delete()
            return None
        return data
    except Exception:
        return None

def delete_session(token: str):
    """세션 토큰 삭제 (로그아웃)"""
    if token:
        try:
            db.collection("sessions").document(token).delete()
        except Exception:
            pass

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
    # 쿠키 기반 세션 저장 (URL에 인증정보 노출 없음)
    _token = create_session(uid)
    cookies["session_token"] = _token
    cookies.save()
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
# 세션 복원 (쿠키 → Firestore 검증 → session_state)
# URL에 인증정보 없음 — 쿠키의 session_token으로만 복원
# -----------------------------------
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    _cookie_token = cookies.get("session_token", "")
    if _cookie_token:
        try:
            _session_data = get_session(_cookie_token)
            if _session_data:
                _uid = _session_data["uid"]
                user_doc = get_user_doc(_uid)
                if user_doc and not user_doc.get("disabled", False):
                    _email = user_doc.get("email", "")
                    admin_emails = list(st.secrets["auth"]["admin_emails"])
                    role = "admin" if _email in admin_emails else user_doc.get("role", "user")
                    st.session_state["logged_in"] = True
                    st.session_state["uid"]       = _uid
                    st.session_state["email"]     = _email
                    st.session_state["role"]      = role
                    st.session_state["tabs_perm"] = user_doc.get("tabs", DEFAULT_USER_TABS.copy())
                else:
                    # 비활성화된 계정 → 세션 삭제
                    delete_session(_cookie_token)
                    cookies["session_token"] = ""
                    cookies.save()
        except Exception:
            cookies["session_token"] = ""
            cookies.save()

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
        # Firestore 세션 삭제 + 쿠키 삭제
        _logout_token = cookies.get("session_token", "")
        delete_session(_logout_token)
        cookies["session_token"] = ""
        cookies.save()
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
def load_view_table(months=36):
    # PostgreSQL에서 view_table 조회 (3년치 기본)
    DB_URL = st.secrets["DB_URL"]
    conn = st.connection("postgresql", type="sql", url=DB_URL)
    
    # 선택된 개월 수만큼 데이터 로드
    if months >= 9999:
        query = "SELECT * FROM view_table"
    else:
        cutoff = datetime.today() - relativedelta(months=months)
        cutoff_str = cutoff.strftime("%Y-%m-%d")
        query = f"SELECT * FROM view_table WHERE 출고일자 >= '{cutoff_str}'"
    
    df = conn.query(query)
    
    df["출고일자"] = pd.to_datetime(df["출고일자"], errors="coerce")
    df["출고년월"] = df["출고년월"].astype(str).str.strip()
    df["총내품출고수량"] = pd.to_numeric(df["총내품출고수량"], errors="coerce").fillna(0)
    df["품목별매출(VAT제외)"] = pd.to_numeric(df["품목별매출(VAT제외)"], errors="coerce").fillna(0)
    
    return df

@st.cache_data(ttl=600)
def load_fin_view_table(months=36):
    """PostgreSQL에서 fin_view_table(확정데이터) 조회"""
    try:
        DB_URL = st.secrets["DB_URL"]
        conn = st.connection("postgresql", type="sql", url=DB_URL)
        
        if months >= 9999:
            query = "SELECT * FROM fin_view_table"
        else:
            cutoff = datetime.today() - relativedelta(months=months)
            cutoff_str = cutoff.strftime("%Y-%m-%d")
            query = f"SELECT * FROM fin_view_table WHERE 출고일자 >= '{cutoff_str}'"
        
        df = conn.query(query)
        
        if not df.empty:
            df["출고일자"] = pd.to_datetime(df["출고일자"], errors="coerce")
            df["출고년월"] = df["출고년월"].astype(str).str.strip()
            df["총내품출고수량"] = pd.to_numeric(df["총내품출고수량"], errors="coerce").fillna(0)
            df["품목별매출(VAT제외)"] = pd.to_numeric(df["품목별매출(VAT제외)"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"확정 데이터(fin_view_table) 로드 중 오류: {e}")
        return pd.DataFrame()

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
@st.cache_data(ttl=300)
def load_kpi_target():
    """KPI_TARGET 시트에서 연도별 목표매출액 로드. 시트 없으면 빈 dict 반환."""
    try:
        client = get_gspread_client()
        ws = client.open_by_key(SHEET_ID).worksheet("KPI_TARGET")
        data = ws.get_all_values()
        if len(data) < 2:
            return {}
        result = {}
        for row in data[1:]:
            if len(row) < 2:
                continue
            try:
                year = int(str(row[0]).strip())
                target = float(str(row[1]).replace(",", "").strip())
                result[year] = target
            except:
                pass
        return result
    except Exception:
        return {}

def save_kpi_target(year: int, target: float):
    """KPI_TARGET 시트에 연도별 목표매출액 저장 (없으면 시트 생성)."""
    client = get_gspread_client_rw()
    sh = client.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("KPI_TARGET")
    except Exception:
        ws = sh.add_worksheet(title="KPI_TARGET", rows=50, cols=5)
        ws.update("A1:B1", [["연도", "목표매출액"]])
    data = ws.get_all_values()
    # 헤더 포함 전체 읽기
    for i, row in enumerate(data[1:], start=2):
        if row and str(row[0]).strip() == str(year):
            ws.update(f"A{i}:B{i}", [[year, target]])
            load_kpi_target.clear()
            return
    # 없으면 마지막 행에 추가
    next_row = len(data) + 1
    ws.update(f"A{next_row}:B{next_row}", [[year, target]])
    load_kpi_target.clear()

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
def build_dataset(months=36):
    df        = load_view_table(months=months)
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
# 초기화 및 세션 상태 관리
if "reset_count" not in st.session_state:
    st.session_state["reset_count"] = 0

# 데이터 로드 기간 설정 포함
with st.sidebar:
    st.markdown("### ⚙️ 설정")
    _lookback_label = st.selectbox(
        "데이터 로드 기간",
        options=["최근 3년 (기본)", "최근 5년", "전체 데이터"],
        index=0,
        help="DB에서 불러올 데이터의 기간을 설정합니다. 기간이 길어지면 로딩이 느려질 수 있습니다."
    )
    _months_map = {"최근 3년 (기본)": 36, "최근 5년": 60, "전체 데이터": 9999}
    _load_months = _months_map[_lookback_label]

    df     = build_dataset(months=_load_months)
    client = get_gspread_client()
    sh     = client.open_by_key(SHEET_ID)

    st.divider()
    with st.expander("🔍 상세 필터 (채널/품목)", expanded=False):
        st.markdown("**🏪 채널 필터**")
        _all_depts = sorted(df["담당부서"].dropna().replace("", None).dropna().unique().tolist())
        selected_depts = st.multiselect(
            "담당부서",
            options=_all_depts,
            default=[],
            key=f"filter_depts_{st.session_state['reset_count']}"
        )

        if selected_depts:
            _dept_filtered_channels = sorted(
                df[df["담당부서"].isin(selected_depts)]["거래처분류"].dropna().unique().tolist()
            )
        else:
            _dept_filtered_channels = sorted(df["거래처분류"].dropna().unique().tolist())

        _ch_select_all = st.checkbox("채널 전체 선택", value=True, key=f"ch_select_all_{st.session_state['reset_count']}")
        if _ch_select_all:
            selected_channel_groups = _dept_filtered_channels
        else:
            selected_channel_groups = st.multiselect(
                "채널",
                options=_dept_filtered_channels,
                default=_dept_filtered_channels,
                key=f"filter_channels_{st.session_state['reset_count']}"
            )

        st.divider()
        st.markdown("**📦 품목 필터**")
        _all_item_groups = sorted([
            g for g in df["품목군"].dropna().unique().tolist()
            if g not in ("__매출조정__", "__미분류__")
        ])

        selected_item_groups = st.multiselect(
            "품목군",
            options=_all_item_groups,
            default=[],
            key=f"filter_item_groups_{st.session_state['reset_count']}"
        )

        if selected_item_groups:
            _ig_filtered_items = sorted(
                df[df["품목군"].isin(selected_item_groups)]["내품상품명"].dropna().unique().tolist()
            )
        else:
            _ig_filtered_items = sorted(df["내품상품명"].dropna().unique().tolist())

        _item_select_all = st.checkbox("품목 전체 선택", value=True, key=f"item_select_all_{st.session_state['reset_count']}")
        if _item_select_all:
            selected_items = _ig_filtered_items
        else:
            selected_items = st.multiselect(
                "품목",
                options=_ig_filtered_items,
                default=_ig_filtered_items,
                key=f"filter_items_{st.session_state['reset_count']}"
            )

    st.divider()
    st.markdown("### 🧭 내비게이션")
    visible_tabs = [k for k in ALL_TABS if tab_allowed(k)]
    if st.session_state["role"] == "admin":
        menu_labels = [tab_defs[k] for k in visible_tabs] + ["⚙️ 관리자"]
    else:
        menu_labels = [tab_defs[k] for k in visible_tabs]
    
    selected_menu_label = st.radio(
        "Menu",
        options=menu_labels,
        label_visibility="collapsed",
        key="main_menu"
    )
    st.divider()


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
# 메인 화면 레이아웃 시작
# -----------------------------------
st.markdown(DASHBOARD_CSS, unsafe_allow_html=True)

# 상단 헤더 영역
st.markdown(f'<div class="main-header">{selected_menu_label}</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">데이터 기반 비즈니스 인사이트 플랫폼</div>', unsafe_allow_html=True)

# -----------------------------------
# 전역 필터 (상단 배치)
# -----------------------------------
with st.container():
    st.markdown('<div class="filter-container">', unsafe_allow_html=True)
    f_col1, f_col2, f_col3, f_col4 = st.columns([1, 1, 1, 0.5])
    
    with f_col1:
        _available_years = sorted(df["연도"].unique().tolist(), reverse=True)
        _selected_year = st.selectbox("📅 분석 연도", options=_available_years, index=0, key=f"selected_year_{st.session_state['reset_count']}")

    # 선택된 연도의 데이터 범위 계산
    df_year = df[df["연도"] == _selected_year]
    _valid_dates_year = df_year["출고일자"].dropna()
    if not _valid_dates_year.empty:
        _min_date_year = _valid_dates_year.min().date()
        _max_date_year = _valid_dates_year.max().date()
    else:
        _min_date_year = datetime(_selected_year, 1, 1).date()
        _max_date_year = datetime(_selected_year, 12, 31).date()

    with f_col2:
        _date_start = st.date_input(
            "시작일", 
            value=_min_date_year, 
            min_value=datetime(_selected_year, 1, 1).date(), 
            max_value=datetime(_selected_year, 12, 31).date(), 
            key=f"date_start_{st.session_state['reset_count']}"
        )
    with f_col3:
        _date_end = st.date_input(
            "종료일", 
            value=_max_date_year, 
            min_value=datetime(_selected_year, 1, 1).date(), 
            max_value=datetime(_selected_year, 12, 31).date(), 
            key=f"date_end_{st.session_state['reset_count']}"
        )
    with f_col4:
        st.write("") # 간격 조절용
        _show_detailed = st.toggle("📊 상세 금액 (원)", value=False, key=f"show_detailed_{st.session_state['reset_count']}")
        if st.button("🔄 필터 초기화", use_container_width=True):
            st.session_state["reset_count"] += 1
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

if _date_start > _date_end:
    st.error("시작일이 종료일보다 늦을 수 없습니다.")
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
# [v7.3] AUTH_MASTER 기반 권한 필터링
# -----------------------------------

# -----------------------------------
# filtered_df
# -----------------------------------
# 매출조정 행은 부서 구분 없이 품목 필터와 무관하게 항상 포함 (음수 매출 반영 필수)
_is_adj_mask = df["내품상품명"].astype(str).str.strip() == "매출조정"

# [추가] 비교 분석용 베이스 데이터 (날짜 필터만 제외하고 채널/품목 필터 적용)
_comp_base_mask = (
    df["거래처분류"].isin(selected_channel_groups) &
    (df["내품상품명"].isin(selected_items) | _is_adj_mask)
)
comparison_base_df = df[_comp_base_mask].copy()

filtered_df = comparison_base_df[
    (comparison_base_df["출고일자"] >= _date_start_dt) &
    (comparison_base_df["출고일자"] <= _date_end_dt)
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

if not visible_tabs:
    st.warning("접근 가능한 탭이 없습니다. 관리자에게 권한을 요청하세요.")
    st.stop()

# -----------------------------------
# KPI (탭 권한 있는 사용자에게만 표시)
# -----------------------------------
st.markdown("""
<style>[data-testid="stMetricValue"] { font-size: 28px; }</style>
""", unsafe_allow_html=True)

# ── 관리자 탭 전역 CSS: 고정 헤더 + 스크롤 영역 ──
st.markdown("""
<style>
/* ──────────────────────────────────────────
   공통: 헤더 셀 줄바꿈 절대 금지
────────────────────────────────────────── */
.perm-header-cell,
.role-header-cell {
    white-space: nowrap !important;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* ──────────────────────────────────────────
   계정 권한 관리: 유저 행 스크롤 컨테이너
   - 헤더 HTML 바로 아래 st.columns 묶음에 적용
   - .perm-scroll-area 클래스를 감싸는 방식
────────────────────────────────────────── */

/* 헤더 HTML 블록 */
.perm-header-wrap {
    border: 1px solid #E0D8CC;
    border-radius: 10px 10px 0 0;
    overflow: hidden;
    background: #F3EFE8;
}

/* 스크롤 래퍼 — st.container 바로 아래 div에 높이 제한 */
.perm-scroll-area {
    max-height: 400px;
    overflow-y: auto;
    border-left: 1px solid #E0D8CC;
    border-right: 1px solid #E0D8CC;
    border-bottom: 1px solid #E0D8CC;
    border-radius: 0 0 10px 10px;
    background: #fff;
}

/* ──────────────────────────────────────────
   역할 변경: 동일 패턴
────────────────────────────────────────── */
.role-header-wrap {
    border: 1px solid #E0D8CC;
    border-radius: 10px 10px 0 0;
    overflow: hidden;
    background: #F3EFE8;
}

.role-scroll-area {
    max-height: 400px;
    overflow-y: auto;
    border-left: 1px solid #E0D8CC;
    border-right: 1px solid #E0D8CC;
    border-bottom: 1px solid #E0D8CC;
    border-radius: 0 0 10px 10px;
    background: #fff;
}

/* ──────────────────────────────────────────
   Streamlit st.columns 행을 스크롤 컨테이너로
   감싸는 핵심 CSS
   각 행의 체크박스/셀렉트박스가 포함된
   수평 블록들의 부모 div에 max-height 적용
────────────────────────────────────────── */

/* 계정 권한 관리 스크롤 영역 */
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlock"].perm-scroll {
    max-height: 400px;
    overflow-y: auto;
    border: 1px solid #E0D8CC;
    border-top: none;
    border-radius: 0 0 10px 10px;
    padding: 0 4px;
}

/* 모든 관리자 행 구분선 스타일 */
.admin-row-divider {
    border: none;
    border-top: 1px solid #F0EBE2;
    margin: 0;
    padding: 0;
}
</style>
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

def _fmt_money(v):
    """숫자를 억/천만 단위로 축약 표시. 예) 11,159,909,268 → 111.6억"""
    if v >= 1_0000_0000:
        return f"{v / 1_0000_0000:.1f}억 원"
    elif v >= 1000_0000:
        return f"{v / 1000_0000:.1f}천만 원"
    else:
        return f"{v:,.0f} 원"

st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
c1, c2, c3, c4, c5 = st.columns([1.4, 1, 1.4, 1, 1.2])
c1.metric("누적 매출",                   _fmt_money(total_sales))
c2.metric("누적 출고량",                  f"{total_qty:,.0f}")
c3.metric("매출 총 이익",                 _fmt_money(total_gross_profit))
c4.metric("매출 총 이익률",               f"{gross_profit_rate:.2f}%")
c5.metric("Top 채널", top_channel,        delta=f"{sales_mom:.2f}% MoM")
st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# 차트용 데이터: 매출조정 제외
# -----------------------------
filtered_df = filtered_df[
    filtered_df["내품상품명"].astype(str).str.strip() != "매출조정"
].copy()

if filtered_df.empty:
    st.info("차트 및 상세 분석에 표시할 데이터가 없습니다.")
    st.stop()

# 메뉴 키 매핑
reverse_tab_defs = {v: k for k, v in tab_defs.items()}
current_tab_key = reverse_tab_defs.get(selected_menu_label)
if not current_tab_key and "관리자" in selected_menu_label:
    current_tab_key = "admin_setting"

# ===================================
# TAB 0: 경영진 요약 (admin only)
# ===================================
if current_tab_key == "경영진요약":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("👔 경영진 요약 대시보드")
    st.caption("채널/품목 수익성, AI 종합 분석을 한 화면에서 확인합니다. 각 섹션은 접을 수 있습니다.")

    # ── 공통 히트맵 colorscale ──
    _HM_COLORSCALE = [
        [0.0,  "#ef4444"],
        [0.35, "#fbbf24"],
        [0.6,  "#86efac"],
        [1.0,  "#15803d"],
    ]

    # ══════════════════════════════════════
    # 섹션 1: 월별 매출 및 공헌이익 추이
    # ══════════════════════════════════════
    with st.expander("📈 월별 매출 및 공헌이익 추이", expanded=True):
        _exec_monthly = (
            filtered_df.groupby("출고년월", as_index=False)[
                ["품목별매출(VAT제외)", "매출총이익", "공헌이익"]
            ].sum()
        )
        _exec_monthly["dt"] = pd.to_datetime(_exec_monthly["출고년월"] + "-01", errors="coerce")
        _exec_monthly = _exec_monthly.sort_values("dt")

        _exec_fig = go.Figure()
        # 토글 상태에 따른 단위 변환
        _div = 1 if _show_detailed else 1e8
        _unit = "원" if _show_detailed else "억"
        _fmt = ":,.0f" if _show_detailed else ":.1f"

        _exec_fig.add_trace(go.Bar(
            x=_exec_monthly["출고년월"], y=_exec_monthly["품목별매출(VAT제외)"] / _div,
            name="매출액", marker_color="#1f77b4",
            text=_exec_monthly["품목별매출(VAT제외)"] / _div,
            texttemplate=f"%{{text{_fmt}}}{_unit}",
            textposition="outside", cliponaxis=False,
            textfont=dict(size=12, color="black"),
        ))
        _exec_fig.add_trace(go.Scatter(
            x=_exec_monthly["출고년월"], y=_exec_monthly["공헌이익"] / _div,
            name="공헌이익", mode="lines+markers+text",
            line=dict(width=4, color="#EF4444"),
            text=_exec_monthly["공헌이익"] / _div,
            texttemplate=f"%{{text{_fmt}}}{_unit}",
            textposition="top center",
            textfont=dict(size=12, color="black"),
        ))
        _exec_monthly_ymax = max(
            (_exec_monthly["품목별매출(VAT제외)"] / _div).max() if not _exec_monthly.empty else 1,
            (_exec_monthly["공헌이익"] / _div).max() if not _exec_monthly.empty else 1,
        ) * 1.25
        _exec_fig.update_layout(
            height=380, barmode="group",
            yaxis=dict(range=[0, _exec_monthly_ymax], title=f"금액 ({_unit})"),
            legend=dict(orientation="h"),
            margin=dict(t=40, b=20),
        )
        st.plotly_chart(_exec_fig, use_container_width=True)

    # ══════════════════════════════════════
    # 섹션 2: 채널 매출 집중도 리스크
    # ══════════════════════════════════════
    with st.expander("⚠️ 채널 매출 집중도 리스크", expanded=True):
        _exec_ch_sales = (
            filtered_df.groupby("거래처분류")["품목별매출(VAT제외)"]
            .sum().sort_values(ascending=False)
        )
        if not _exec_ch_sales.empty and total_sales > 0:
            _top3_sales = _exec_ch_sales.head(3).sum()
            _top3_rate  = _top3_sales / total_sales * 100
            _top1_rate  = _exec_ch_sales.iloc[0] / total_sales * 100
            _risk_label = "🔴 고위험 (분산 필요)" if _top3_rate >= 70 else "🟡 주의" if _top3_rate >= 50 else "🟢 양호"
            _conc_c1, _conc_c2, _conc_c3 = st.columns(3)
            _conc_c1.metric("Top 1 채널 집중도", f"{_top1_rate:.1f}%",
                            delta=f"최대 채널: {_exec_ch_sales.index[0]}")
            _conc_c2.metric("Top 3 채널 집중도", f"{_top3_rate:.1f}%",
                            delta=_risk_label)
            _conc_c3.metric("전체 채널 수", f"{len(_exec_ch_sales)}개")
            st.caption("📌 Top 3 집중도 70% 이상이면 특정 채널 이탈 시 매출 타격이 큽니다.")

            # 채널별 매출 비중 바차트 (각 바에 비중% 라벨 표시)
            _conc_df = _exec_ch_sales.reset_index()
            _conc_df.columns = ["채널", "매출"]
            _conc_df["비중(%)"] = _conc_df["매출"] / total_sales * 100
            _conc_colors = (
                ["#ef4444"] * 1 + ["#f97316"] * 2 + ["#60a5fa"] * (len(_conc_df) - 3)
                if len(_conc_df) >= 3 else ["#60a5fa"] * len(_conc_df)
            )
            _conc_fig = go.Figure(go.Bar(
                x=_conc_df["채널"], y=_conc_df["매출"],
                marker_color=_conc_colors,
                text=_conc_df["비중(%)"].apply(lambda v: f"{v:.1f}%"),
                textposition="outside",
                cliponaxis=False,
                hovertemplate="채널: %{x}<br>매출: %{y:,.0f}원<br>비중: %{text}<extra></extra>",
            ))
            _conc_fig.update_layout(
                height=320,
                margin=dict(t=30, b=60),
                xaxis=dict(tickangle=-35),
                yaxis=dict(title="매출액 (원)"),
                annotations=[dict(
                    x=0.01, y=1.06, xref="paper", yref="paper",
                    text="🔴 1위  🟠 2~3위  🔵 나머지",
                    showarrow=False, font=dict(size=11, color="#6b7280"),
                )],
            )
            st.plotly_chart(_conc_fig, use_container_width=True)

    # ══════════════════════════════════════
    # 섹션 3: 채널별 공헌이익률 히트맵
    # ══════════════════════════════════════
    with st.expander("🌡️ 채널별 공헌이익률 히트맵 (채널 × 월)", expanded=True):
        st.caption("색이 진할수록 공헌이익률이 높습니다. 빨간색은 적자 채널입니다. (매출 상위 15개 채널)")
        _hm_df = (
            filtered_df.groupby(["거래처분류", "출고년월"])[["품목별매출(VAT제외)", "공헌이익"]]
            .sum().reset_index()
        )
        _hm_df["공헌이익률(%)"] = np.where(
            _hm_df["품목별매출(VAT제외)"] != 0,
            _hm_df["공헌이익"] / _hm_df["품목별매출(VAT제외)"] * 100, 0,
        )
        if not _hm_df.empty:
            _hm_pivot = _hm_df.pivot_table(
                index="거래처분류", columns="출고년월", values="공헌이익률(%)", aggfunc="mean"
            )
            _hm_month_cols = sort_month_cols(_hm_pivot.columns.tolist())
            _hm_pivot = _hm_pivot[_hm_month_cols]
            _top_ch_for_hm = (
                filtered_df.groupby("거래처분류")["품목별매출(VAT제외)"].sum()
                .sort_values(ascending=False).head(15).index.tolist()
            )
            _hm_pivot = _hm_pivot.loc[_hm_pivot.index.isin(_top_ch_for_hm)]
            _hm_fig = go.Figure(go.Heatmap(
                z=_hm_pivot.values,
                x=_hm_pivot.columns.tolist(),
                y=_hm_pivot.index.tolist(),
                colorscale=_HM_COLORSCALE, zmid=0,
                text=np.round(_hm_pivot.values, 1),
                texttemplate="%{text}%", textfont={"size": 11},
                hovertemplate="채널: %{y}<br>월: %{x}<br>공헌이익률: %{z:.1f}%<extra></extra>",
                colorbar=dict(title="공헌이익률(%)", ticksuffix="%"),
            ))
            _hm_fig.update_layout(
                height=max(300, len(_hm_pivot) * 32 + 80),
                margin=dict(t=20, b=60, l=10, r=10),
                xaxis=dict(tickangle=-30),
            )
            st.plotly_chart(_hm_fig, use_container_width=True)

    # ══════════════════════════════════════
    # 섹션 4: 품목군별 공헌이익률 히트맵 (신규)
    # ══════════════════════════════════════
    with st.expander("🌡️ 품목군별 공헌이익률 히트맵 (품목군 × 월)", expanded=True):
        st.caption("색이 진할수록 공헌이익률이 높습니다. 빨간색은 적자 품목군입니다.")
        _ig_hm_df = (
            filtered_df[~filtered_df["품목군"].isin(["__매출조정__", "__미분류__"])]
            .groupby(["품목군", "출고년월"])[["품목별매출(VAT제외)", "공헌이익"]]
            .sum().reset_index()
        )
        _ig_hm_df["공헌이익률(%)"] = np.where(
            _ig_hm_df["품목별매출(VAT제외)"] != 0,
            _ig_hm_df["공헌이익"] / _ig_hm_df["품목별매출(VAT제외)"] * 100, 0,
        )
        if not _ig_hm_df.empty:
            _ig_hm_pivot = _ig_hm_df.pivot_table(
                index="품목군", columns="출고년월", values="공헌이익률(%)", aggfunc="mean"
            )
            _ig_hm_month_cols = sort_month_cols(_ig_hm_pivot.columns.tolist())
            _ig_hm_pivot = _ig_hm_pivot[_ig_hm_month_cols]
            _ig_hm_fig = go.Figure(go.Heatmap(
                z=_ig_hm_pivot.values,
                x=_ig_hm_pivot.columns.tolist(),
                y=_ig_hm_pivot.index.tolist(),
                colorscale=_HM_COLORSCALE, zmid=0,
                text=np.round(_ig_hm_pivot.values, 1),
                texttemplate="%{text}%", textfont={"size": 11},
                hovertemplate="품목군: %{y}<br>월: %{x}<br>공헌이익률: %{z:.1f}%<extra></extra>",
                colorbar=dict(title="공헌이익률(%)", ticksuffix="%"),
            ))
            _ig_hm_fig.update_layout(
                height=max(260, len(_ig_hm_pivot) * 36 + 80),
                margin=dict(t=20, b=60, l=10, r=10),
                xaxis=dict(tickangle=-30),
            )
            st.plotly_chart(_ig_hm_fig, use_container_width=True)

    # ══════════════════════════════════════
    # 섹션 5: 품목군별 전월 대비 매출 성장률
    # ══════════════════════════════════════
    with st.expander("🚀 품목군별 전월 대비 매출 성장률", expanded=True):
        _MIN_SALES_THRESHOLD = 1_000_000  # 100만 원 하한선 (옵션 A)
        st.caption(f"전월·당월 매출 모두 {_MIN_SALES_THRESHOLD:,.0f}원 이상인 품목군만 표시합니다. (소규모 노이즈 제거)")
        _exec_ig = (
            filtered_df[~filtered_df["품목군"].isin(["__매출조정__", "__미분류__"])]
            .groupby(["품목군", "출고년월"])["품목별매출(VAT제외)"].sum()
            .reset_index().sort_values("출고년월")
        )
        _exec_ig_pivot = _exec_ig.pivot_table(
            index="품목군", columns="출고년월", values="품목별매출(VAT제외)", fill_value=0
        )
        _ig_month_cols = sort_month_cols(_exec_ig_pivot.columns.tolist())
        if len(_ig_month_cols) >= 2:
            _last_m = _ig_month_cols[-1]
            _prev_m = _ig_month_cols[-2]
            _ig_growth = pd.DataFrame({
                "품목군":   _exec_ig_pivot.index,
                "전월매출": _exec_ig_pivot[_prev_m].values,
                "당월매출": _exec_ig_pivot[_last_m].values,
            })
            # 옵션 A: 전월·당월 모두 100만원 이상인 경우만 계산
            _ig_growth = _ig_growth[
                (_ig_growth["전월매출"] >= _MIN_SALES_THRESHOLD) &
                (_ig_growth["당월매출"] >= _MIN_SALES_THRESHOLD)
            ].copy()
            _ig_growth["성장률(%)"] = (
                (_ig_growth["당월매출"] - _ig_growth["전월매출"]) / _ig_growth["전월매출"] * 100
            )
            _ig_growth = _ig_growth.sort_values("성장률(%)", ascending=False)
            if not _ig_growth.empty:
                _bar_colors = ["#22c55e" if v >= 0 else "#ef4444" for v in _ig_growth["성장률(%)"]]
                _growth_fig = go.Figure(go.Bar(
                    x=_ig_growth["품목군"],
                    y=_ig_growth["성장률(%)"],
                    marker_color=_bar_colors,
                    text=_ig_growth["성장률(%)"].apply(lambda v: f"{v:+.1f}%"),
                    textposition="outside", cliponaxis=False,
                    customdata=np.stack([
                        _ig_growth["전월매출"].values,
                        _ig_growth["당월매출"].values,
                    ], axis=-1),
                    hovertemplate=(
                        "품목군: %{x}<br>"
                        f"전월({_prev_m}): %{{customdata[0]:,.0f}}원<br>"
                        f"당월({_last_m}): %{{customdata[1]:,.0f}}원<br>"
                        "성장률: %{y:.1f}%<extra></extra>"
                    ),
                ))
                _growth_fig.update_layout(
                    height=340,
                    yaxis=dict(ticksuffix="%", zeroline=True,
                               zerolinecolor="#94a3b8", zerolinewidth=2),
                    margin=dict(t=40, b=60),
                    xaxis=dict(tickangle=-30),
                    title=dict(text=f"{_prev_m} → {_last_m} 전월 대비 품목군 매출 성장률",
                               font=dict(size=13)),
                )
                st.plotly_chart(_growth_fig, use_container_width=True)
            else:
                st.info(f"기준({_MIN_SALES_THRESHOLD:,.0f}원)을 충족하는 품목군이 없습니다.")
        else:
            st.info("전월 대비 성장률 계산을 위해 2개월 이상의 데이터가 필요합니다.")

    # ══════════════════════════════════════
    # 섹션 6: 국내/해외 매출 구분
    # ══════════════════════════════════════
    with st.expander("🌏 국내 / 해외 매출 구분", expanded=True):
        _mkt_df = (
            filtered_df.groupby("국내여부")[["품목별매출(VAT제외)", "공헌이익"]].sum().reset_index()
        )
        _mkt_cols = st.columns(len(_mkt_df))
        for _mi, _mrow in _mkt_df.iterrows():
            _ms = _mrow["품목별매출(VAT제외)"]
            _mci = _mrow["공헌이익"]
            _mci_rate = (_mci / _ms * 100) if _ms else 0
            _mkt_cols[_mi].metric(
                f"{_mrow['국내여부']} 매출",
                _fmt_money(_ms),
                delta=f"공헌이익률 {_mci_rate:.1f}%",
            )
        _donut_fig = go.Figure(go.Pie(
            labels=_mkt_df["국내여부"],
            values=_mkt_df["품목별매출(VAT제외)"],
            hole=0.55,
            marker_colors=["#3b82f6", "#f59e0b"],
            textinfo="label+percent",
        ))
        _donut_fig.update_layout(height=260, margin=dict(t=10, b=10, l=10, r=10),
                                 showlegend=False)
        st.plotly_chart(_donut_fig, use_container_width=True)

    # ══════════════════════════════════════
    # 섹션 7: 🤖 AI 경영 데이터 챗봇
    # ══════════════════════════════════════
    with st.expander("🤖 AI 경영 데이터 챗봇 (데이터 기반 대화)", expanded=True):
        st.caption("현재 필터링된 실적 데이터를 바탕으로 AI와 자유롭게 대화하세요. (예: '가장 효율 좋은 채널은?', '매출이 급감한 이유는?')")

        _exec_ctx_key = f"exec_ai_chat_{hash(str(sorted(selected_months)))}_{hash(str(sorted(selected_channel_groups)))}"
        if st.session_state.get("exec_chat_key") != _exec_ctx_key:
            st.session_state["exec_chat_key"] = _exec_ctx_key
            st.session_state["exec_chat_history"] = []
            st.session_state["exec_ai_context"] = build_common_ai_context(filtered_df)

        _exec_ctx = st.session_state["exec_ai_context"]
        _exec_system = f"""당신은 링티(Lingtea) 비즈니스 데이터 전문가입니다. 다음 데이터를 바탕으로 대답하세요.
{_exec_ctx}
- 숫자를 기반으로 구체적이고 객관적인 답변을 제공하세요.
- 경영진 관점에서 인사이트를 포함하세요. 한국어로 답변하세요."""

        # 채팅 히스토리 표시
        for msg in st.session_state["exec_chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        _exec_input = st.chat_input("질문을 입력하세요 (예: 이번 달 매출 성장의 핵심 채널은?)", key="exec_chat_input")
        if _exec_input:
            with st.chat_message("user"):
                st.markdown(_exec_input)
            st.session_state["exec_chat_history"].append({"role": "user", "content": _exec_input})
            
            with st.chat_message("assistant"):
                with st.spinner("생성 중..."):
                    try:
                        _ans = call_claude_api(_exec_system, st.session_state["exec_chat_history"][-10:])
                        st.markdown(_ans)
                        st.session_state["exec_chat_history"].append({"role": "assistant", "content": _ans})
                    except Exception as e:
                        st.error(f"오류: {e}")

        if st.session_state["exec_chat_history"]:
            if st.button("🗑️ 대화 초기화", key="exec_clear_chat"):
                st.session_state["exec_chat_history"] = []
                st.rerun()

    # ══════════════════════════════════════
    # 섹션 8: AI 경영 종합 분석 리포트
    # ══════════════════════════════════════
    with st.expander("🚀 AI 경영 종합 분석 리포트", expanded=False):
        _exec_ctx = st.session_state["exec_ai_context"]
        _exec_report_system = f"""당신은 링티(Lingtea)의 비즈니스 데이터 분석 전문가로, C레벨 경영진에게 보고하는 역할입니다.
아래는 현재 대시보드 필터 기준으로 집계된 실제 판매 데이터입니다.

{_exec_ctx}

위 데이터를 바탕으로 경영진이 즉시 의사결정에 활용할 수 있는 간결하고 핵심적인 분석을 제공하세요.
- 숫자는 구체적으로 언급하되, 핵심 메시지를 먼저 제시하세요 (BLUF 원칙).
- 문제점·기회 요인을 균형있게 제시하고, 우선순위가 명확한 액션을 권고하세요.
- 전문 경영 용어를 사용하되 한국어로 간결하게 작성하세요.
- 마크다운 형식으로 구조화하세요."""

        if "exec_ai_report" not in st.session_state:
            st.session_state["exec_ai_report"] = None

        _exec_ctx = st.session_state["exec_ai_context"]
        _exec_system = f"""당신은 링티(Lingtea)의 비즈니스 데이터 분석 전문가로, C레벨 경영진에게 보고하는 역할입니다.
아래는 현재 대시보드 필터 기준으로 집계된 실제 판매 데이터입니다.

{_exec_ctx}

위 데이터를 바탕으로 경영진이 즉시 의사결정에 활용할 수 있는 간결하고 핵심적인 분석을 제공하세요.
- 숫자는 구체적으로 언급하되, 핵심 메시지를 먼저 제시하세요 (BLUF 원칙).
- 문제점·기회 요인을 균형있게 제시하고, 우선순위가 명확한 액션을 권고하세요.
- 전문 경영 용어를 사용하되 한국어로 간결하게 작성하세요.
- 마크다운 형식으로 구조화하세요."""

        _exec_col_btn, _exec_col_reset, _ = st.columns([2, 1, 4])
        with _exec_col_btn:
            _exec_run = st.button("🚀 AI 경영 분석 실행", use_container_width=True, type="primary", key="exec_ai_run")
        with _exec_col_reset:
            if st.session_state.get("exec_ai_report") and st.button("🔄 재생성", use_container_width=True, key="exec_ai_reset"):
                st.session_state["exec_ai_report"] = None
                st.rerun()

        if _exec_run:
            st.session_state["exec_ai_report"] = None

        if _exec_run or st.session_state.get("exec_ai_report") == "__running__":
            st.session_state["exec_ai_report"] = "__running__"
            with st.spinner("AI가 경영진 보고용 분석을 생성 중입니다... (15~30초 소요)"):
                try:
                    _exec_report = call_claude_api(
                        _exec_report_system,
                        [{"role": "user", "content": (
                            "현재 데이터를 경영진 보고용으로 종합 분석해주세요. 다음 항목을 포함하세요:\n"
                            "1. 📊 핵심 요약 (전체 성과 한 줄 평가, 목표 달성 여부)\n"
                            "2. 📈 주목할 매출 트렌드 (월별 변화, 이상치)\n"
                            "3. 🏪 채널 분석 (상위/하위 채널, 공헌이익 관점 효율성, 집중도 리스크)\n"
                            "4. 📦 품목군 분석 (기여도 높은/낮은 품목군, 성장률 주목 품목)\n"
                            "5. ⚠️ 리스크 & 기회 요인\n"
                            "6. 💡 경영진 추천 액션 (구체적으로 2~3가지, 우선순위 포함)"
                        )}],
                        max_tokens=2048,
                    )
                    st.session_state["exec_ai_report"] = _exec_report
                except Exception as _exec_e:
                    st.error(f"AI 분석 중 오류가 발생했습니다: {_exec_e}")
                    st.session_state["exec_ai_report"] = None

        if st.session_state.get("exec_ai_report") and st.session_state["exec_ai_report"] != "__running__":
            st.markdown("---")
            st.markdown(st.session_state["exec_ai_report"])

# ===================================
# TAB 1: 월별 추이
# ===================================
if current_tab_key == "월별추이":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("📈 월별 추이")
    month_subtabs = st.tabs(["📊 전체 월별 추이", "📺 채널별 월별 추이"])
    with month_subtabs[0]:

        # [수정] 전년 동월 실적을 비교용 베이스 데이터(전체 기간)에서 산출
        _comp_monthly = comparison_base_df.groupby("출고년월")["품목별매출(VAT제외)"].sum().to_dict()

        monthly = (
            filtered_df.groupby("출고년월", as_index=False)[
                ["품목별매출(VAT제외)", "매출총이익", "총내품출고수량"]
            ].sum()
        )
        monthly["dt"] = pd.to_datetime(monthly["출고년월"] + "-01", errors="coerce")
        monthly = monthly.sort_values("dt").rename(columns={
            "품목별매출(VAT제외)": "매출액", "총내품출고수량": "출고량"
        })

        def get_prev_year_sales(ym):
            try:
                dt = pd.to_datetime(ym + "-01")
                prev_ym = (dt - relativedelta(years=1)).strftime("%Y-%m")
                return _comp_monthly.get(prev_ym, None)
            except:
                return None

        monthly["전년동월매출"] = monthly["출고년월"].apply(get_prev_year_sales)
        
        # 레이블용 연도 추출
        _year_label = f"{str(_selected_year)[2:]}년"

        ctrl1, ctrl2 = st.columns([1, 1])
        with ctrl1:
            show_label = st.checkbox("📊 라벨 표시", value=True)
        with ctrl2:
            # 전년 데이터 존재 여부 확인
            has_prev_data = monthly["전년동월매출"].notna().any()
            if has_prev_data:
                show_prev_year = st.checkbox("📅 전년 동월 비교 표시", value=False)
            else:
                show_prev_year = False

        fig = go.Figure()
        _div = 1 if _show_detailed else 1e8
        _unit = "원" if _show_detailed else "억"
        _fmt = ":,.0f" if _show_detailed else ":.1f"

        if show_prev_year:
            prev_data = monthly.dropna(subset=["전년동월매출"])
            if not prev_data.empty:
                fig.add_trace(go.Bar(
                    x=prev_data["출고년월"], y=prev_data["전년동월매출"] / _div,
                    name="매출액 (전년 동월)", marker_color="#D1D5DB",
                    text=prev_data["전년동월매출"] / _div if show_label else None,
                    texttemplate=f"%{{text{_fmt}}}{_unit}" if show_label else None,
                    textposition='outside', cliponaxis=False
                ))
        fig.add_trace(go.Bar(
            x=monthly["출고년월"], y=monthly["매출액"] / _div,
            name=f"매출액", marker_color="#1f77b4",
            text=monthly["매출액"] / _div if show_label else None,
            texttemplate=f"%{{text{_fmt}}}{_unit}" if show_label else None,
            textposition='outside', cliponaxis=False
        ))
        fig.add_trace(go.Scatter(
            x=monthly["출고년월"], y=monthly["매출총이익"] / _div,
            name="매출총이익", mode="lines+markers+text" if show_label else "lines+markers",
            line=dict(width=4, color="#EF4444"),
            text=monthly["매출총이익"] / _div if show_label else None,
            texttemplate=f"%{{text{_fmt}}}{_unit}" if show_label else None,
            textposition="top center"
        ))

        y_vals = [safe_max(monthly["매출액"] / _div, 0), safe_max(monthly["매출총이익"] / _div, 0)]
        if show_prev_year and not monthly["전년동월매출"].dropna().empty:
            y_vals.append(safe_max(monthly["전년동월매출"].dropna() / _div, 0))
        y_max = max(y_vals) * 1.25 if max(y_vals) > 0 else 1

        fig.update_layout(height=400, barmode="group",
                          yaxis=dict(range=[0, y_max], title=f"금액 ({_unit})"),
                          legend=dict(orientation="h"), margin=dict(t=40))
        fig.update_traces(textfont=dict(size=12, color="black"))
        st.plotly_chart(fig, use_container_width=True)

        if show_prev_year:
            st.markdown("##### 📊 전년 동월 대비 매출 비교")
            # 전년 데이터가 있는 행만 필터링
            compare_df = monthly[monthly["전년동월매출"].notna()][
                ["출고년월", "매출액", "전년동월매출"]
            ].copy()
            compare_df = compare_df.rename(columns={
                "출고년월": "월", "매출액": "당해 매출액", "전년동월매출": "전년 동월 매출액"
            })
            compare_df["증감액"] = compare_df["당해 매출액"] - compare_df["전년 동월 매출액"]
            compare_df["증감률"] = (compare_df["증감액"] / compare_df["전년 동월 매출액"]) * 100
            
            st.dataframe(
                compare_df.style.format({
                    "당해 매출액": "{:,.0f}", "전년 동월 매출액": "{:,.0f}",
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


    with month_subtabs[1]:
        st.markdown("### 📺 채널별 월별 추이")

        _cm_df = filtered_df.groupby(
            ["출고년월", "거래처분류"], as_index=False
        )[["품목별매출(VAT제외)", "매출총이익", "총내품출고수량"]].sum()
        _cm_df = _cm_df.sort_values("출고년월")

        _cm_c1, _cm_c2, _cm_c3 = st.columns([2, 2, 3])
        with _cm_c1:
            _cm_metric = st.selectbox(
                "지표 선택",
                options=["매출액", "매출총이익", "출고량"],
                key="cm_metric"
            )
        with _cm_c2:
            _cm_top_mode = st.selectbox(
                "표시 방식",
                options=["Top 5 채널", "Top 10 채널", "직접 선택"],
                key="cm_top_mode"
            )

        _cm_col_map = {
            "매출액": "품목별매출(VAT제외)",
            "매출총이익": "매출총이익",
            "출고량": "총내품출고수량",
        }
        _cm_val_col = _cm_col_map[_cm_metric]

        _cm_ch_totals = (
            _cm_df.groupby("거래처분류")[_cm_val_col].sum()
            .sort_values(ascending=False)
        )

        if _cm_top_mode == "Top 5 채널":
            _cm_channels = _cm_ch_totals.head(5).index.tolist()
        elif _cm_top_mode == "Top 10 채널":
            _cm_channels = _cm_ch_totals.head(10).index.tolist()
        else:
            with _cm_c3:
                _cm_channels = st.multiselect(
                    "채널 직접 선택",
                    options=_cm_ch_totals.index.tolist(),
                    default=_cm_ch_totals.head(5).index.tolist(),
                    key="cm_channels"
                )

        _cm_plot_df = _cm_df[_cm_df["거래처분류"].isin(_cm_channels)].copy()

        if _cm_plot_df.empty:
            st.info("선택한 채널의 데이터가 없습니다.")
        else:
            fig_cm = px.line(
                _cm_plot_df,
                x="출고년월",
                y=_cm_val_col,
                color="거래처분류",
                markers=True,
                labels={
                    "출고년월": "출고년월",
                    _cm_val_col: _cm_metric,
                    "거래처분류": "채널",
                },
                title=f"채널별 월별 {_cm_metric}",
            )
            fig_cm.update_layout(
                height=480,
                legend=dict(orientation="h", yanchor="bottom", y=-0.35),
                margin=dict(t=50, b=80),
                yaxis=dict(tickformat=","),
            )
            fig_cm.update_traces(line=dict(width=2.5))
            st.plotly_chart(fig_cm, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ===================================
# TAB 2: 주차별 추이
# ===================================
if current_tab_key == "주차별추이":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
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
        _div = 1 if _show_detailed else 1e8
        _unit = "원" if _show_detailed else "억"
        _fmt = ":,.0f" if _show_detailed else ":.1f"

        fig_wk.add_trace(go.Bar(
            x=weekly_view["주차_label"],
            y=weekly_view["매출액"] / _div,
            name="매출액",
            marker_color="#1f77b4",
            text=weekly_view["매출액"] / _div if wk_show_label else None,
            texttemplate=f"%{{text{_fmt}}}{_unit}" if wk_show_label else None,
            textposition="outside",
            cliponaxis=False,
        ))
        

    




    


        fig_wk.add_trace(go.Scatter(
            x=weekly_view["주차_label"],
            y=weekly_view["매출총이익"] / _div,
            name="매출총이익",
            mode="lines+markers+text" if wk_show_label else "lines+markers",
            line=dict(width=3, color="#EF4444"),
            text=weekly_view["매출총이익"] / _div if wk_show_label else None,
            texttemplate=f"%{{text{_fmt}}}{_unit}" if wk_show_label else None,
            textposition="top center",
        ))
        wk_y_max = max((weekly_view["매출액"] / _div).max(), (weekly_view["매출총이익"] / _div).max()) * 1.3
        fig_wk.update_layout(
            height=420,
            barmode="group",
            yaxis=dict(range=[0, wk_y_max], title=f"금액 ({_unit})"),
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
    st.markdown('</div>', unsafe_allow_html=True)


    # ── 일별 상세 추이 (expander) ──
    with st.expander("📅 일별 상세 추이", expanded=False):
        _daily_df = filtered_df.copy()
        _daily_df = _daily_df.dropna(subset=["출고일자"])
        _daily_df["출고일자_date"] = _daily_df["출고일자"].dt.date

        _d_c1, _d_c2, _d_c3 = st.columns([2, 2, 3])
        with _d_c1:
            _d_days = st.selectbox(
                "최근 N일",
                options=[7, 14, 30, 60],
                index=1,
                key="daily_days"
            )
        with _d_c2:
            _d_metric = st.selectbox(
                "지표 선택",
                options=["매출액", "매출총이익", "출고량"],
                key="daily_metric"
            )
        with _d_c3:
            _d_all_ch = sorted(_daily_df["거래처분류"].dropna().unique().tolist())
            _d_channels = st.multiselect(
                "채널 선택 (비워두면 전체 합산)",
                options=_d_all_ch,
                default=[],
                key="daily_channels"
            )

        _d_col_map = {
            "매출액": "품목별매출(VAT제외)",
            "매출총이익": "매출총이익",
            "출고량": "총내품출고수량",
        }
        _d_val_col = _d_col_map[_d_metric]

        # 채널 필터
        if _d_channels:
            _daily_df = _daily_df[_daily_df["거래처분류"].isin(_d_channels)]

        # 일별 집계
        _daily_agg = (
            _daily_df.groupby("출고일자_date", as_index=False)[_d_val_col].sum()
            .sort_values("출고일자_date")
        )
        _daily_agg["출고일자_date"] = pd.to_datetime(_daily_agg["출고일자_date"])

        # 최근 N일 슬라이싱
        _daily_agg = _daily_agg.tail(_d_days).copy()

        if _daily_agg.empty:
            st.info("선택한 조건에 일별 데이터가 없습니다.")
        else:
            # 7일 이동평균
            _daily_agg["이동평균_7일"] = (
                _daily_agg[_d_val_col].rolling(window=7, min_periods=1).mean()
            )

            fig_daily = go.Figure()
            fig_daily.add_trace(go.Scatter(
                x=_daily_agg["출고일자_date"],
                y=_daily_agg[_d_val_col],
                name=_d_metric,
                mode="lines+markers",
                line=dict(width=2, color="#1f77b4"),
                marker=dict(size=5),
            ))
            fig_daily.add_trace(go.Scatter(
                x=_daily_agg["출고일자_date"],
                y=_daily_agg["이동평균_7일"],
                name="7일 이동평균",
                mode="lines",
                line=dict(width=2.5, color="#EF4444", dash="dot"),
            ))
            fig_daily.update_layout(
                height=400,
                title=f"일별 {_d_metric} (최근 {_d_days}일)",
                xaxis=dict(title="날짜", tickformat="%m/%d"),
                yaxis=dict(title=_d_metric, tickformat=","),
                legend=dict(orientation="h"),
                margin=dict(t=50),
            )
            st.plotly_chart(fig_daily, use_container_width=True)

# ===================================
# TAB 3: 채널 분석
# ===================================
if current_tab_key == "채널분석":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
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
    
    tuples_ch = []
    seen_metrics = {}
    for c in ch_display_cols:
        if c in ch_sales_mcols:
            metric = "매출액"
            seen_metrics[metric] = seen_metrics.get(metric, 0) + 1
            unique_metric = metric + ("\u200b" * (seen_metrics[metric] - 1))
            tuples_ch.append((c, unique_metric))
        else:
            metric = "구성비"
            seen_metrics[metric] = seen_metrics.get(metric, 0) + 1
            unique_metric = metric + ("\u200b" * (seen_metrics[metric] - 1))
            tuples_ch.append((c.replace("_구성비", ""), unique_metric))
    
    ch_display = safe_multiindex_from_tuples(ch_display, tuples_ch)

    # 데이터타입 강제 변환 (float)
    for col in ch_display.columns:
        ch_display[col] = pd.to_numeric(ch_display[col], errors='coerce').fillna(0).astype(float)

    # 포맷팅 딕셔너리 (바뀐 유니크 헤더 이름에 맞춤)
    ch_display_fmt = {}
    seen_fmt = {"매출액": 0, "구성비": 0}
    for m in ch_sales_mcols:
        seen_fmt["매출액"] += 1
        u_ms = "매출액" + ("\u200b" * (seen_fmt["매출액"] - 1))
        ch_display_fmt[(m, u_ms)] = "{:,.0f}"
        
        seen_fmt["구성비"] += 1
        u_rt = "구성비" + ("\u200b" * (seen_fmt["구성비"] - 1))
        ch_display_fmt[(m, u_rt)] = "{:.1%}"

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
    
    # [수정] 데이터타입 강제 변환 (안전한 포맷팅 위해)
    ch_qty_pivot = ch_qty_pivot.apply(pd.to_numeric, errors='coerce').fillna(0)
    
    st.dataframe(ch_qty_pivot.style.format("{:,.0f}"), use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ===================================
# TAB 3: 제품 분석
# ===================================
if current_tab_key == "제품분석":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
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
    seen_metrics_p = {}
    for c in prod_display_cols:
        if c in prod_s_mcols:
            metric = "매출액"
            seen_metrics_p[metric] = seen_metrics_p.get(metric, 0) + 1
            unique_metric = metric + ("\u200b" * (seen_metrics_p[metric] - 1))
            tuples_prod.append((c, unique_metric))
        elif "_구성비" in c:
            metric = "구성비"
            seen_metrics_p[metric] = seen_metrics_p.get(metric, 0) + 1
            unique_metric = metric + ("\u200b" * (seen_metrics_p[metric] - 1))
            tuples_prod.append((c.replace("_구성비", ""), unique_metric))
        else:
            metric = "개당단가"
            seen_metrics_p[metric] = seen_metrics_p.get(metric, 0) + 1
            unique_metric = metric + ("\u200b" * (seen_metrics_p[metric] - 1))
            tuples_prod.append((c.replace("_개당단가", ""), unique_metric))
            
    prod_display = safe_multiindex_from_tuples(prod_display, tuples_prod)
    
    # 데이터타입 강제 변환
    for col in prod_display.columns:
        prod_display[col] = pd.to_numeric(prod_display[col], errors='coerce').fillna(0).astype(float)

    # 포맷팅 딕셔너리 (바뀐 유니크 헤더 이름에 맞춤)
    prod_display_fmt = {}
    seen_fmt_p = {"매출액": 0, "구성비": 0, "개당단가": 0}
    for m in prod_s_mcols:
        seen_fmt_p["매출액"] += 1
        u_ms = "매출액" + ("\u200b" * (seen_fmt_p["매출액"] - 1))
        prod_display_fmt[(m, u_ms)] = "{:,.0f}"
        
        seen_fmt_p["구성비"] += 1
        u_rt = "구성비" + ("\u200b" * (seen_fmt_p["구성비"] - 1))
        prod_display_fmt[(m, u_rt)] = "{:.1%}"
        
        seen_fmt_p["개당단가"] += 1
        u_pr = "개당단가" + ("\u200b" * (seen_fmt_p["개당단가"] - 1))
        prod_display_fmt[(m, u_pr)] = "{:,.0f}"

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
    st.markdown('</div>', unsafe_allow_html=True)

# ===================================
# TAB: YoY 분석
# ===================================
if current_tab_key == "YoY분석":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("📊 YoY 성과 분석")
    
    # 1. 비교 모드 선택
    yoy_mode = st.radio(
        "비교 방식 선택",
        ["전년 동월 비교", "누적(YTD) 비교", "기간 자유 선택"],
        horizontal=True,
        key="yoy_mode_selector"
    )
    
    all_months = sorted(df["출고년월"].dropna().unique().tolist())
    
    target_df = pd.DataFrame()
    base_df = pd.DataFrame()
    target_label = ""
    base_label = ""
    
    if yoy_mode == "전년 동월 비교":
        col1, col2 = st.columns(2)
        with col1:
            sel_month = st.selectbox("비교 대상 월 선택 (Target)", all_months[::-1], index=0)
        
        # 전년 동월 계산
        try:
            target_dt = datetime.strptime(sel_month, "%Y-%m")
            base_dt = target_dt - relativedelta(years=1)
            base_month = base_dt.strftime("%Y-%m")
        except:
            base_month = None
            
        with col2:
            st.info(f"기준 월 (Base): {base_month if base_month in all_months else '데이터 없음'}")
            
        target_df = df[df["출고년월"] == sel_month]
        base_df = df[df["출고년월"] == base_month]
        target_label = sel_month
        base_label = base_month

    elif yoy_mode == "누적(YTD) 비교":
        sel_year = st.selectbox("연도 선택", sorted(list(set(df["출고년월"].str[:4])), reverse=True))
        max_month_in_year = df[df["출고년월"].str.startswith(sel_year)]["출고년월"].max()
        
        year_months = sorted([m for m in all_months if m.startswith(sel_year)])
        sel_end_month = st.select_slider("누적 종료월 선택", options=year_months, value=max_month_in_year)
        
        target_year = sel_year
        base_year = str(int(sel_year)-1)
        end_mm = sel_end_month.split("-")[1]
        
        target_mask = (df["출고년월"].str[:4] == target_year) & (df["출고년월"].str[5:7] <= end_mm)
        base_mask = (df["출고년월"].str[:4] == base_year) & (df["출고년월"].str[5:7] <= end_mm)
        
        target_df = df[target_mask]
        base_df = df[base_mask]
        target_label = f"{target_year}-01 ~ {end_mm}"
        base_label = f"{base_year}-01 ~ {end_mm}"

    else: # 기간 자유 선택
        st.info("비교하고 싶은 현재 기간을 선택하면, 자동으로 1년 전 동일 기간과 비교합니다.")
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            start_d = st.date_input("시작일", datetime.now() - timedelta(days=30))
        with d_col2:
            end_d = st.date_input("종료일", datetime.now())
            
        base_start = start_d - relativedelta(years=1)
        base_end = end_d - relativedelta(years=1)
        
        st.caption(f"비교 기준 기간: {base_start} ~ {base_end}")
        
        df_tmp = df.copy()
        df_tmp["출고일자_dt"] = pd.to_datetime(df_tmp["출고일자"], errors='coerce')
        
        target_df = df_tmp[(df_tmp["출고일자_dt"].dt.date >= start_d) & (df_tmp["출고일자_dt"].dt.date <= end_d)]
        base_df = df_tmp[(df_tmp["출고일자_dt"].dt.date >= base_start) & (df_tmp["출고일자_dt"].dt.date <= base_end)]
        target_label = f"{start_d}~{end_d}"
        base_label = f"{base_start}~{base_end}"

    if target_df.empty and base_df.empty:
        st.warning("선택한 기간에 데이터가 없습니다.")
    else:
        # 공통 분석 데이터 생성
        def get_agg(df):
            return df.groupby("거래처분류")["품목별매출(VAT제외)"].sum().reset_index()

        t_agg = get_agg(target_df).rename(columns={"품목별매출(VAT제외)": "현재 매출"})
        b_agg = get_agg(base_df).rename(columns={"품목별매출(VAT제외)": "전년 매출"})
        
        merged = pd.merge(t_agg, b_agg, on="거래처분류", how="outer").fillna(0)
        merged["증감액"] = merged["현재 매출"] - merged["전년 매출"]
        merged["성장률"] = merged.apply(lambda x: (x["현재 매출"] - x["전년 매출"]) / x["전년 매출"] if x["전년 매출"] > 0 else (1.0 if x["현재 매출"] > 0 else 0), axis=1)
        
        # 상태 분류
        merged["상태"] = "유지"
        merged.loc[(merged["전년 매출"] == 0) & (merged["현재 매출"] > 0), "상태"] = "신규"
        merged.loc[(merged["전년 매출"] > 0) & (merged["현재 매출"] == 0), "상태"] = "이탈"
        
        # KPI 섹션
        t_total = merged["현재 매출"].sum()
        b_total = merged["전년 매출"].sum()
        total_diff = t_total - b_total
        total_growth = (total_diff / b_total) if b_total > 0 else (1.0 if t_total > 0 else 0)
        
        k1, k2, k3 = st.columns(3)
        with k1:
            st.metric(f"현재 매출 ({target_label})", f"{t_total/1e8:.2f}억", f"{total_diff/1e8:+.2f}억")
        with k2:
            st.metric("전년 대비 성장률", f"{total_growth:.1%}")
        with k3:
            if not merged.empty:
                top_ch_row = merged.sort_values("증감액", ascending=False).iloc[0]
                st.metric("최대 기여 채널", top_ch_row["거래처분류"], f"{top_ch_row['증감액']/1e8:+.2f}억")

        # 서브 탭 구성
        sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs(["📊 종합 분석", "✨ 신규 채널", "👋 이탈 채널", "🔗 유지 채널"])
        
        with sub_tab1:
            # 폭포수 차트
            st.markdown("#### 🌊 채널별 성과 기여도 (Waterfall)")
            waterfall_data = merged[merged["증감액"] != 0].sort_values("증감액", ascending=False)
            if not waterfall_data.empty:
                # 상위 7개, 하위 3개, 나머지 '기타'
                top_w = waterfall_data.head(7)
                bot_w = waterfall_data.tail(3)
                
                # 중복 방지를 위한 처리
                combined_ids = set(top_w.index) | set(bot_w.index)
                others_w = waterfall_data.drop(index=list(combined_ids))
                
                w_labels = list(top_w["거래처분류"]) + ["기타"] + list(bot_w["거래처분류"])
                w_values = list(top_w["증감액"]) + [others_w["증감액"].sum()] + list(bot_w["증감액"])
                
                fig_w = go.Figure(go.Waterfall(
                    name="YoY", orientation="v",
                    measure=["relative"] * len(w_labels),
                    x=w_labels,
                    textposition="outside",
                    text=[f"{v/1e8:+.1f}억" for v in w_values],
                    y=w_values,
                    connector={"line": {"color": "rgb(63, 63, 63)"}},
                ))
                fig_w.update_layout(title=f"전년 대비 매출 변동 요인 ({base_label} -> {target_label})", showlegend=False)
                st.plotly_chart(fig_w, use_container_width=True)
            else:
                st.info("변동 내역이 없습니다.")
            
            # 스캐터 차트
            st.markdown("#### 🎯 성장성 vs 매출 규모 (유지 채널)")
            retained_df = merged[merged["상태"] == "유지"].copy()
            if not retained_df.empty:
                # [수정] Plotly scatter에서 size는 반드시 양수여야 하므로 음수(반품 등)는 0으로 처리
                retained_df["display_size"] = retained_df["현재 매출"].clip(lower=0)
                # [수정] 성장률이 무한대(inf)인 경우를 대비해 처리 (색상 스케일 오류 방지)
                retained_df["성장률_plot"] = retained_df["성장률"].replace([np.inf, -np.inf], np.nan).fillna(0)
                
                fig_s = px.scatter(
                    retained_df, x="전년 매출", y="성장률_plot", text="거래처분류",
                    size="display_size", color="성장률_plot",
                    color_continuous_scale="RdYlGn",
                    labels={"전년 매출": "전년 매출액", "성장률_plot": "성장률", "display_size": "현재 매출액"}
                )
                fig_s.update_traces(textposition='top center')
                fig_s.add_hline(y=0, line_dash="dash", line_color="gray")
                st.plotly_chart(fig_s, use_container_width=True)

        def show_sub_table(m_df, title):
            st.markdown(f"#### {title}")
            display_df = m_df.copy()
            if "현재 매출" in display_df.columns:
                display_df = display_df.sort_values("현재 매출", ascending=False)
            st.dataframe(
                display_df.style.format({
                    "현재 매출": "{:,.0f}",
                    "전년 매출": "{:,.0f}",
                    "증감액": "{:+,.0f}",
                    "성장률": "{:.1%}"
                }),
                use_container_width=True
            )

        with sub_tab2:
            new_df = merged[merged["상태"] == "신규"]
            if new_df.empty: st.info("신규 채널이 없습니다.")
            else: show_sub_table(new_df[["거래처분월분류" if "거래처분월분류" in new_df.columns else "거래처분류", "현재 매출"]], "신규 진입 채널 실적")
            
        with sub_tab3:
            exit_df = merged[merged["상태"] == "이탈"]
            if exit_df.empty: st.info("이탈 채널이 없습니다.")
            else: show_sub_table(exit_df[["거래처분류", "전년 매출"]], "이탈 채널 과거 실적")
            
        with sub_tab4:
            ret_df = merged[merged["상태"] == "유지"]
            if ret_df.empty: st.info("유지 중인 채널이 없습니다.")
            else: show_sub_table(ret_df[["거래처분류", "전년 매출", "현재 매출", "증감액", "성장률"]], "유지 채널 성과 비교")

    st.markdown('</div>', unsafe_allow_html=True)

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

    # ── 월별 품목군별 공헌이익 (닫힘) ──
    with st.expander("📅 월별 품목군별 공헌이익", expanded=False):
        _month_ig_rows = []
        for m in sorted(selected_months):
            _m_mask = mdf["출고년월"] == m
            _m_ig = mdf[_m_mask].groupby("품목군", as_index=False)[[
                "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                "매출총이익", "채널수수료", "물류비", "광고비"
            ]].sum()
            _m_ig["비용"] = 0.0
            for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
                if ym != m:
                    continue
                if market_filter is not None and mdf[mdf["거래처분류"] == ch].empty:
                    continue
                _mask_ig = _m_ig["품목군"] == ig
                if _mask_ig.any():
                    _m_ig.loc[_mask_ig, "비용"] += amt
            _m_ig["공헌이익"] = (
                _m_ig["품목별매출(VAT제외)"] - _m_ig["원가총액"]
                - _m_ig["채널수수료"] - _m_ig["물류비"] - _m_ig["광고비"] - _m_ig["비용"]
            )
            _m_ig["공헌이익률"] = safe_divide(_m_ig["공헌이익"], _m_ig["품목별매출(VAT제외)"])
            _m_ig = _m_ig[~_m_ig["품목군"].isin(["__매출조정__", "__미분류__"])]
            _m_label = f"{int(m.split('-')[1]):02d}월" if "-" in str(m) else str(m)
            _m_ig.insert(0, "월", _m_label)
            _month_ig_rows.append(_m_ig)
        if _month_ig_rows:
            _monthly_ig_df = pd.concat(_month_ig_rows, ignore_index=True)
            _monthly_ig_df = _monthly_ig_df[[
                "월", "품목군", "총내품출고수량", "품목별매출(VAT제외)",
                "원가총액", "매출총이익", "채널수수료", "물류비", "광고비",
                "비용", "공헌이익", "공헌이익률"
            ]]
            st.dataframe(_monthly_ig_df.style.format({
                "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
                "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}", "채널수수료": "{:,.0f}",
                "물류비": "{:,.0f}", "광고비": "{:,.0f}", "비용": "{:,.0f}",
                "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
            }), use_container_width=True)
        else:
            st.info("월별 품목군별 공헌이익 데이터가 없습니다.")

    # ── 월별 채널별 공헌이익 (닫힘) ──
    with st.expander("📅 월별 채널별 공헌이익", expanded=False):
        _month_ch_rows = []
        for m in sorted(selected_months):
            _m_mask = mdf["출고년월"] == m
            _m_cc = mdf[_m_mask].groupby("거래처분류", as_index=False)[[
                "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                "매출총이익", "채널수수료", "물류비", "광고비"
            ]].sum()
            _m_cc["비용"] = 0.0
            for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
                if ym != m:
                    continue
                if market_filter is not None and mdf[mdf["거래처분류"] == ch].empty:
                    continue
                _mask_ch = _m_cc["거래처분류"] == ch
                if _mask_ch.any():
                    _m_cc.loc[_mask_ch, "비용"] += amt
            _m_cc["공헌이익"] = (
                _m_cc["품목별매출(VAT제외)"] - _m_cc["원가총액"]
                - _m_cc["채널수수료"] - _m_cc["물류비"] - _m_cc["광고비"] - _m_cc["비용"]
            )
            _m_cc["공헌이익률"] = safe_divide(_m_cc["공헌이익"], _m_cc["품목별매출(VAT제외)"])
            _m_label = f"{int(m.split('-')[1]):02d}월" if "-" in str(m) else str(m)
            _m_cc.insert(0, "월", _m_label)
            _month_ch_rows.append(_m_cc)
        if _month_ch_rows:
            _monthly_ch_df = pd.concat(_month_ch_rows, ignore_index=True)
            _monthly_ch_df = _monthly_ch_df[[
                "월", "거래처분류", "총내품출고수량", "품목별매출(VAT제외)",
                "원가총액", "매출총이익", "채널수수료", "물류비", "광고비",
                "비용", "공헌이익", "공헌이익률"
            ]]
            st.dataframe(_monthly_ch_df.style.format({
                "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
                "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}", "채널수수료": "{:,.0f}",
                "물류비": "{:,.0f}", "광고비": "{:,.0f}", "비용": "{:,.0f}",
                "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
            }), use_container_width=True)
        else:
            st.info("월별 채널별 공헌이익 데이터가 없습니다.")


if current_tab_key == "공헌이익분석(국내)":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    _render_contrib_tab(filtered_df, "국내", "📊 공헌이익 분석 (국내)", ad_apply=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ===================================
# TAB 5: 공헌이익 분석 (해외)
# ===================================
if current_tab_key == "공헌이익분석(해외)":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    _render_contrib_tab(filtered_df, "해외", "🌏 공헌이익 분석 (해외)", ad_apply=False)
    st.markdown('</div>', unsafe_allow_html=True)

# ===================================
# TAB 6: 공헌이익 분석 (통합)
# ===================================
if current_tab_key == "공헌이익분석(통합)":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
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

    # ── 2-1. 월별 품목군별 공헌이익 (닫힘) ──
    with st.expander("📅 월별 품목군별 공헌이익", expanded=False):
        _t_month_ig_rows = []
        for m in sorted(selected_months):
            _t_m_mask = temp_df["출고년월"] == m
            _t_m_ig = temp_df[_t_m_mask].groupby("품목군", as_index=False)[[
                "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                "매출총이익", "채널수수료", "물류비", "광고비"
            ]].sum()
            _t_m_ig["비용"] = 0.0
            for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
                if ym != m:
                    continue
                _t_mask_ig = _t_m_ig["품목군"] == ig
                if _t_mask_ig.any():
                    _t_m_ig.loc[_t_mask_ig, "비용"] += amt
            _t_m_ig["공헌이익"] = (
                _t_m_ig["품목별매출(VAT제외)"] - _t_m_ig["원가총액"]
                - _t_m_ig["채널수수료"] - _t_m_ig["물류비"] - _t_m_ig["광고비"] - _t_m_ig["비용"]
            )
            _t_m_ig["공헌이익률"] = safe_divide(_t_m_ig["공헌이익"], _t_m_ig["품목별매출(VAT제외)"])
            _t_m_ig = _t_m_ig[~_t_m_ig["품목군"].isin(["__매출조정__", "__미분류__"])]
            _t_m_label = f"{int(m.split('-')[1]):02d}월" if "-" in str(m) else str(m)
            _t_m_ig.insert(0, "월", _t_m_label)
            _t_month_ig_rows.append(_t_m_ig)
        if _t_month_ig_rows:
            _t_monthly_ig_df = pd.concat(_t_month_ig_rows, ignore_index=True)
            _t_monthly_ig_df = _t_monthly_ig_df[[
                "월", "품목군", "총내품출고수량", "품목별매출(VAT제외)",
                "원가총액", "매출총이익", "채널수수료", "물류비", "광고비",
                "비용", "공헌이익", "공헌이익률"
            ]]
            st.dataframe(_t_monthly_ig_df.style.format({
                "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
                "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}", "채널수수료": "{:,.0f}",
                "물류비": "{:,.0f}", "광고비": "{:,.0f}", "비용": "{:,.0f}",
                "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
            }), use_container_width=True)
        else:
            st.info("월별 품목군별 공헌이익 데이터가 없습니다.")

    # ── 2-2. 월별 채널별 공헌이익 (닫힘) ──
    with st.expander("📅 월별 채널별 공헌이익", expanded=False):
        _t_month_ch_rows = []
        for m in sorted(selected_months):
            _t_m_mask = temp_df["출고년월"] == m
            _t_m_cc = temp_df[_t_m_mask].groupby("거래처분류", as_index=False)[[
                "총내품출고수량", "품목별매출(VAT제외)", "원가총액",
                "매출총이익", "채널수수료", "물류비", "광고비"
            ]].sum()
            _t_m_cc["비용"] = 0.0
            for (ym, ch, ig), amt in st.session_state["channel_cost"].items():
                if ym != m:
                    continue
                _t_mask_ch = _t_m_cc["거래처분류"] == ch
                if _t_mask_ch.any():
                    _t_m_cc.loc[_t_mask_ch, "비용"] += amt
            _t_m_cc["공헌이익"] = (
                _t_m_cc["품목별매출(VAT제외)"] - _t_m_cc["원가총액"]
                - _t_m_cc["채널수수료"] - _t_m_cc["물류비"] - _t_m_cc["광고비"] - _t_m_cc["비용"]
            )
            _t_m_cc["공헌이익률"] = safe_divide(_t_m_cc["공헌이익"], _t_m_cc["품목별매출(VAT제외)"])
            _t_m_label = f"{int(m.split('-')[1]):02d}월" if "-" in str(m) else str(m)
            _t_m_cc.insert(0, "월", _t_m_label)
            _t_month_ch_rows.append(_t_m_cc)
        if _t_month_ch_rows:
            _t_monthly_ch_df = pd.concat(_t_month_ch_rows, ignore_index=True)
            _t_monthly_ch_df = _t_monthly_ch_df[[
                "월", "거래처분류", "총내품출고수량", "품목별매출(VAT제외)",
                "원가총액", "매출총이익", "채널수수료", "물류비", "광고비",
                "비용", "공헌이익", "공헌이익률"
            ]]
            st.dataframe(_t_monthly_ch_df.style.format({
                "총내품출고수량": "{:,.0f}", "품목별매출(VAT제외)": "{:,.0f}",
                "원가총액": "{:,.0f}", "매출총이익": "{:,.0f}", "채널수수료": "{:,.0f}",
                "물류비": "{:,.0f}", "광고비": "{:,.0f}", "비용": "{:,.0f}",
                "공헌이익": "{:,.0f}", "공헌이익률": "{:.2%}",
            }), use_container_width=True)
        else:
            st.info("월별 채널별 공헌이익 데이터가 없습니다.")
    st.markdown('</div>', unsafe_allow_html=True)

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
# TAB AI: AI 분석
# ===================================
# ===================================
# TAB: 확정 비교 분석 (Variance Analysis)
# ===================================
if current_tab_key == "확정비교":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("⚖️ 확정 비교 분석 (가집계 vs ERP 확정)")
    st.caption("가집계(view_table)와 ERP 확정마감(fin_view_table) 데이터를 비교하여 오차를 분석합니다.")

    # 확정 데이터 로드
    fin_df_raw = load_fin_view_table(months=_load_months)
    
    if fin_df_raw.empty:
        st.warning("확정마감 데이터(fin_view_table)가 없습니다. Supabase를 확인해주세요.")
    else:
        # 필터 적용 (가집계와 동일한 날짜 기준)
        fin_df = fin_df_raw[
            (fin_df_raw["출고일자"] >= _date_start_dt) &
            (fin_df_raw["출고일자"] <= _date_end_dt)
        ].copy()
        
        # 권한 필터 적용 (현재 사용자 권한에 맞게)
        if st.session_state.get("user_role_type") == "부서기반":
            _user_dept_name = st.session_state.get("user_dept")
            # 가집계 데이터(filtered_df)에서 이미 필터링된 채널 목록 추출
            valid_channels = filtered_df["거래처분류"].unique()
            fin_df = fin_df[fin_df["거래처코드"].isin(valid_channels)]
        
        # 비교용 집계 (월별/채널별)
        # 가집계 요약: '매출조정'이 포함된 데이터를 사용하기 위해 comparison_base_df 활용
        comp_pre_df = comparison_base_df[
            (comparison_base_df["출고일자"] >= _date_start_dt) &
            (comparison_base_df["출고일자"] <= _date_end_dt)
        ].copy()
        
        pre_summary = comp_pre_df.groupby(["출고년월", "거래처분류"])[["품목별매출(VAT제외)", "총내품출고수량"]].sum().reset_index()
        pre_summary.columns = ["출고년월", "채널", "가집계_매출", "가집계_수량"]
        
        # 확정 요약
        fin_summary = fin_df.groupby(["출고년월", "거래처코드"])[["품목별매출(VAT제외)", "총내품출고수량"]].sum().reset_index()
        fin_summary.columns = ["출고년월", "채널", "확정_매출", "확정_수량"]
        
        # 머지
        comp_df = pd.merge(pre_summary, fin_summary, on=["출고년월", "채널"], how="outer").fillna(0)
        
        # [추가] 확정 데이터가 존재하는 마지막 월까지만 데이터 필터링
        # (예: 확정이 3월까지면, 가집계 4~5월 데이터는 비교에서 제외)
        if not fin_summary.empty:
            max_fin_month = fin_summary["출고년월"].max()
            comp_df = comp_df[comp_df["출고년월"] <= max_fin_month].copy()
            st.info(f"💡 현재 ERP 확정마감 데이터가 존재하는 **{max_fin_month}**까지만 비교 분석합니다.")
        
        # 오차 계산
        comp_df["매출오차(Δ)"] = comp_df["확정_매출"] - comp_df["가집계_매출"]
        comp_df["오차율(%)"] = safe_divide(comp_df["매출오차(Δ)"], comp_df["가집계_매출"]) * 100
        
        # KPI 카드
        total_pre_sales = comp_df["가집계_매출"].sum()
        total_fin_sales = comp_df["확정_매출"].sum()
        total_diff = total_fin_sales - total_pre_sales
        total_var_rate = safe_divide(total_diff, total_pre_sales) * 100
        
        v_c1, v_c2, v_c3 = st.columns(3)
        v_c1.metric("전체 가집계 매출", _fmt_money(total_pre_sales))
        v_c2.metric("전체 확정마감 매출", _fmt_money(total_fin_sales), delta=f"{total_diff:+,.0f}")
        v_c3.metric("전체 오차율", f"{total_var_rate:+.2f}%", delta_color="inverse")
        
        st.divider()
        
        # 상세 비교 테이블 (월별 멀티헤더 구조)
        st.markdown("#### 📊 채널별 상세 비교 (월별)")
        
        # 피벗 테이블 생성 (Index: 채널, Columns: [출고년월, 지표])
        pivot_comp = comp_df.pivot(index="채널", columns="출고년월", values=["가집계_매출", "확정_매출", "매출오차(Δ)", "오차율(%)"])
        
        # 컬럼 순서 조정: (지표, 월) -> (월, 지표)
        pivot_comp.columns = pivot_comp.columns.swaplevel(0, 1)
        
        # 월별 정렬 및 선택된 지표 순서 정의
        sorted_months_comp = sort_month_cols(list(pivot_comp.columns.levels[0]))
        metric_order = ["가집계_매출", "확정_매출", "매출오차(Δ)", "오차율(%)"]
        
        # 멀티인덱스 재배열
        new_cols = []
        for m in sorted_months_comp:
            for met in metric_order:
                if (m, met) in pivot_comp.columns:
                    new_cols.append((m, met))
        
        pivot_comp = pivot_comp.reindex(columns=pd.MultiIndex.from_tuples(new_cols))
        
        # 합계 행 추가
        pivot_comp.loc["[합계]"] = pivot_comp.sum()
        # [합계] 행의 오차율은 다시 계산
        for m in sorted_months_comp:
            if (m, "가집계_매출") in pivot_comp.columns:
                pre_tot = pivot_comp.loc["[합계]", (m, "가집계_매출")]
                fin_tot = pivot_comp.loc["[합계]", (m, "확정_매출")]
                diff_tot = fin_tot - pre_tot
                pivot_comp.loc["[합계]", (m, "매출오차(Δ)")] = diff_tot
                pivot_comp.loc["[합계]", (m, "오차율(%)")] = safe_divide(diff_tot, pre_tot) * 100

        # 마지막 월 오차액 기준 내림차순 정렬 (합계 제외)
        if sorted_months_comp:
            last_m = sorted_months_comp[-1]
            if (last_m, "매출오차(Δ)") in pivot_comp.columns:
                data_part = pivot_comp.drop("[합계]").sort_values((last_m, "매출오차(Δ)"), key=abs, ascending=False)
                pivot_comp = pd.concat([pivot_comp.loc[["[합계]"]], data_part])

        # 포맷팅 설정
        comp_fmt = {}
        for col in pivot_comp.columns:
            if col[1] == "오차율(%)":
                comp_fmt[col] = "{:+.2f}%"
            elif col[1] == "매출오차(Δ)":
                comp_fmt[col] = "{:+,.0f}"
            else:
                comp_fmt[col] = "{:,.0f}"

        if pivot_comp.empty:
            st.info("비교할 데이터가 없습니다.")
        else:
            st.dataframe(pivot_comp.style.format(comp_fmt), use_container_width=True)
            
            # 차트: 주요 채널별 가집계 vs 확정 비교
            st.divider()
            c_col1, c_col2, c_col3 = st.columns([2, 2, 3])
            with c_col1:
                top_n_comp = st.selectbox("표시 채널 수", [10, 20, 30, 50], index=0, key="top_n_comp")
            with c_col2:
                show_full_amt = st.checkbox("전체 금액으로 표시", value=False, key="show_full_amt")
            
            st.markdown(f"#### 📈 주요 채널별 차이 시각화")
            
            # 차트용 데이터 (최근 월 기준 정렬된 comp_df 데이터 활용)
            chart_df = comp_df.copy()
            chart_df["abs_diff"] = chart_df["매출오차(Δ)"].abs()
            top_channels_data = chart_df.sort_values("abs_diff", ascending=False).head(top_n_comp)
            
            # 라벨 포맷 함수
            def format_label(val, full=False):
                if full:
                    return f"{val:,.0f}"
                else:
                    if abs(val) >= 1e8:
                        return f"{val/1e8:.1f}억"
                    elif abs(val) >= 1e4:
                        return f"{val/1e4:,.0f}만"
                    else:
                        return f"{val:,.0f}"

            fig_comp = go.Figure()
            fig_comp.add_trace(go.Bar(
                x=top_channels_data["채널"], 
                y=top_channels_data["가집계_매출"],
                name="가집계", 
                marker_color="#94a3b8",
                text=[format_label(v, show_full_amt) for v in top_channels_data["가집계_매출"]],
                textposition='auto',
            ))
            fig_comp.add_trace(go.Bar(
                x=top_channels_data["채널"], 
                y=top_channels_data["확정_매출"],
                name="확정마감", 
                marker_color="#1e293b",
                text=[format_label(v, show_full_amt) for v in top_channels_data["확정_매출"]],
                textposition='auto',
            ))
            fig_comp.update_layout(
                barmode='group',
                height=500,
                xaxis=dict(tickangle=-45),
                yaxis=dict(title="매출액 (원)", tickformat=","), # Y축도 쉼표 표기
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(b=100)
            )
            st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

if current_tab_key == "AI분석":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("🤖 AI 분석")

    # ── 컨텍스트 캐시 (필터 바뀔 때마다 초기화) ──
    _ctx_key = f"ai_ctx_{hash(str(sorted(selected_months)))}_{hash(str(sorted(selected_channel_groups)))}"
    if st.session_state.get("ai_context_key") != _ctx_key:
        st.session_state["ai_context_key"]  = _ctx_key
        st.session_state["ai_context"]       = build_common_ai_context(filtered_df)
        st.session_state["ai_chat_history"]  = []
        st.session_state["ai_auto_report"]   = None

    _ctx = st.session_state["ai_context"]

    # 데이터 컨텍스트 포함 system prompt (종합분석 + 채팅 공용)
    _system_with_data = f"""당신은 링티(Lingtea) 비즈니스 데이터 분석 전문가입니다.
아래는 현재 대시보드 필터 기준으로 집계된 실제 판매 데이터입니다.

{_ctx}

위 데이터를 바탕으로 사용자의 질문에 답변하거나 분석을 제공하세요.
- 숫자는 구체적으로 언급하고 변화율·비율을 함께 제시하세요.
- 문제점과 기회 요인을 균형 있게 제시하세요.
- 한국어로 답변하세요.
- 마크다운 형식으로 구조화해서 작성하세요."""

    # 데이터 컨텍스트 없는 system prompt (순수 질문용)
    _system_no_data = """당신은 비즈니스·마케팅·데이터 분석 전문가입니다.
사용자의 질문에 친절하고 명확하게 답변하세요.
- 한국어로 답변하세요.
- 마크다운 형식으로 구조화해서 작성하세요."""

    

    # ══════════════════════════════════════
    # 섹션 1: 🤖 AI 데이터 챗봇
    # ══════════════════════════════════════
    with st.expander("🤖 AI 데이터 챗봇 (데이터 기반 대화)", expanded=True):
        st.markdown("##### 💡 이런 질문은 어떠세요?")
        
        # 입력창 처리 (추천 질문 클릭 시 해당 값 사용을 위해 최상단 배치)
        _chat_placeholder = "궁금한 점을 입력하세요... (예: 전월 대비 매출이 급감한 품목군이 있어?)"
        user_input = st.chat_input(_chat_placeholder, key="ai_analysis_chat_input")

        # 추천 질문 처리 로직
        if "pending_question" not in st.session_state:
            st.session_state["pending_question"] = None

        if st.session_state["pending_question"]:
            user_input = st.session_state["pending_question"]
            st.session_state["pending_question"] = None # 소모

        # 추천 질문 버튼 (입력창 아래 배치)
        ex_cols = st.columns(2)
        examples = [
            "가장 매출 기여도가 높은 채널은 어디야?",
            "전월 대비 매출이 급감한 품목군이 있어?",
            "공헌이익률이 낮은 채널들의 특징은?",
            "최근 매출 트렌드를 3줄로 요약해줘."
        ]
        for i, ex_text in enumerate(examples):
            with ex_cols[i % 2]:
                if st.button(f"❓ {ex_text}", key=f"ex_btn_{i}", use_container_width=True):
                    st.session_state["pending_question"] = ex_text
                    st.rerun()

        st.markdown("---")

        # 채팅 히스토리 표시
        for msg in st.session_state.get("ai_chat_history", []):
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        if user_input:
            with st.chat_message("user"):
                st.markdown(user_input)
            st.session_state["ai_chat_history"].append({"role": "user", "content": user_input})
            with st.chat_message("assistant"):
                with st.spinner("답변 생성 중..."):
                    try:
                        recent_history = st.session_state["ai_chat_history"][-10:]
                        # 데이터 컨텍스트가 포함된 시스템 프롬프트 사용
                        answer = call_claude_api(_system_with_data, recent_history, max_tokens=1024)
                        st.markdown(answer)
                        st.session_state["ai_chat_history"].append({"role": "assistant", "content": answer})
                    except Exception as e:
                        err_msg = f"오류가 발생했습니다: {e}"
                        st.error(err_msg)
                        st.session_state["ai_chat_history"].append({"role": "assistant", "content": err_msg})

        if st.session_state.get("ai_chat_history"):
            if st.button("🗑️ 대화 초기화", key="clear_chat_analysis"):
                st.session_state["ai_chat_history"] = []
                st.rerun()


    st.divider()

    # ══════════════════════════════════════
    # 섹션 2: 📋 종합 분석
    # ══════════════════════════════════════
    st.markdown("### 📋 종합 분석")

    # 종합 분석 미리보기 (버튼 누르기 전 항상 표시)
    st.markdown("""
<div style="background:#F8F6F2; border:1px solid #E8E4DC; border-radius:12px; padding:20px 24px; margin-bottom:16px;">
<p style="font-size:13px; color:#7A6E5A; margin:0 0 12px 0; font-weight:600;">📌 종합 분석 버튼을 누르면 아래 항목들이 생성됩니다</p>
<table style="width:100%; border-collapse:collapse; font-size:13px; color:#3A3530;">
<tr style="border-bottom:1px solid #E8E4DC;">
  <td style="padding:8px 12px; font-weight:600; width:30%;">📊 핵심 요약</td>
  <td style="padding:8px 12px; color:#7A6E5A;">전체 기간 성과를 한 문장으로 평가 · 전월 대비 매출 증감률 · 공헌이익률 수준 진단</td>
</tr>
<tr style="border-bottom:1px solid #E8E4DC;">
  <td style="padding:8px 12px; font-weight:600;">📈 매출 트렌드</td>
  <td style="padding:8px 12px; color:#7A6E5A;">월별 매출 흐름 해석 · 이상치(급등/급락) 구간 짚기 · 계절성 패턴 언급</td>
</tr>
<tr style="border-bottom:1px solid #E8E4DC;">
  <td style="padding:8px 12px; font-weight:600;">🏪 채널 분석</td>
  <td style="padding:8px 12px; color:#7A6E5A;">매출 상위·하위 채널 · 공헌이익률 기준 효율 채널 vs 적자 채널 · 국내/해외 비중</td>
</tr>
<tr style="border-bottom:1px solid #E8E4DC;">
  <td style="padding:8px 12px; font-weight:600;">📦 품목군 분석</td>
  <td style="padding:8px 12px; color:#7A6E5A;">매출·공헌이익 기여도 높은 품목군 · 출고량 대비 수익성 낮은 품목군 경고</td>
</tr>
<tr style="border-bottom:1px solid #E8E4DC;">
  <td style="padding:8px 12px; font-weight:600;">⚠️ 리스크 & 기회</td>
  <td style="padding:8px 12px; color:#7A6E5A;">공헌이익 마이너스 채널/품목군 · 매출 집중도 리스크 · 성장 가능성 있는 채널</td>
</tr>
<tr>
  <td style="padding:8px 12px; font-weight:600;">💡 추천 액션</td>
  <td style="padding:8px 12px; color:#7A6E5A;">데이터 기반 구체적 액션 2~3가지 · 우선순위 채널/품목군 제안</td>
</tr>
</table>
</div>
""", unsafe_allow_html=True)

    # 종합 분석 버튼
    col_btn, col_reset, _ = st.columns([2, 1, 4])
    with col_btn:
        run_report = st.button("🚀 종합 분석 실행", use_container_width=True, type="primary")
    with col_reset:
        if st.session_state.get("ai_auto_report") and st.button("🔄 재생성", use_container_width=True):
            st.session_state["ai_auto_report"] = None
            st.rerun()

    if run_report:
        st.session_state["ai_auto_report"] = None

    # 종합 분석 결과 표시
    if run_report or st.session_state.get("ai_auto_report") == "__running__":
        st.session_state["ai_auto_report"] = "__running__"
        with st.spinner("AI가 전체 데이터를 종합 분석 중입니다... (15~30초 소요)"):
            try:
                report = call_claude_api(
                    _system_with_data,
                    [{"role": "user", "content": (
                        "현재 데이터를 종합 분석해서 다음 항목을 포함한 리포트를 작성해주세요:\n"
                        "1. 📊 핵심 요약 (전체 성과 한 줄 평가, 주요 지표 수치 포함)\n"
                        "2. 📈 주목할 매출 트렌드 (월별 변화, 이상치, 계절성)\n"
                        "3. 🏪 채널 분석 (상위/하위 채널, 공헌이익 관점 효율성)\n"
                        "4. 📦 품목군 분석 (기여도 높은/낮은 품목군, 수익성 진단)\n"
                        "5. ⚠️ 리스크 & 기회 요인\n"
                        "6. 💡 추천 액션 (구체적으로 2~3가지, 우선순위 포함)"
                    )}],
                    max_tokens=2048,
                )
                st.session_state["ai_auto_report"] = report
            except Exception as e:
                st.error(f"AI 분석 중 오류가 발생했습니다: {e}")
                st.session_state["ai_auto_report"] = None

    if st.session_state.get("ai_auto_report") and st.session_state["ai_auto_report"] != "__running__":
        st.markdown("---")
        st.markdown(st.session_state["ai_auto_report"])
    st.markdown('</div>', unsafe_allow_html=True)


# ===================================
# TAB 7: 다운로드
# ===================================
if current_tab_key == "다운로드":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
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
    st.markdown('</div>', unsafe_allow_html=True)

# ===================================
# TAB 6: 관리자 (admin only)
# ===================================
if current_tab_key == "admin_setting":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
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
            # ── CSS: 헤더 셀 줄바꿈 금지 + 폰트 ──
            st.markdown("""
            <style>
            .adm-hdr {
                display: flex;
                align-items: center;
                background: #F3EFE8;
                border: 1px solid #E0D8CC;
                border-radius: 8px 8px 0 0;
                border-bottom: 2px solid #C8B88A;
                padding: 0 2px;
            }
            .adm-hdr-cell {
                font-family: 'Noto Sans KR', sans-serif;
                font-size: 10.5px;
                font-weight: 700;
                color: #5C4A2A;
                white-space: nowrap;
                padding: 9px 4px;
                text-align: center;
                letter-spacing: 0.2px;
                flex-shrink: 0;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .adm-hdr-cell.h-status { width: 52px; }
            .adm-hdr-cell.h-email  { width: 200px; min-width: 140px; flex: 1; text-align: left; padding-left: 10px; }
            .adm-hdr-cell.h-tab    { width: 66px; }
            .adm-hdr-cell.h-all    { width: 58px; background: #EAE4D6; border-radius: 0 6px 0 0; }
            /* 스크롤 컨테이너 하단 테두리 */
            [data-testid="stVerticalBlock"] iframe { display: none; }
            </style>
            """, unsafe_allow_html=True)

            # ── 유저 행 렌더링: st.container(height=) 으로 스크롤 ──
            COL_RATIOS = [0.6, 2.5] + [1.0] * len(ALL_TABS) + [0.8]
            hdr_cols = st.columns(COL_RATIOS)
            hdr_cols[0].markdown("<div style='font-size:12px; font-weight:bold; padding-top:4px;'>상태</div>", unsafe_allow_html=True)
            hdr_cols[1].markdown("<div style='font-size:12px; font-weight:bold; padding-top:4px;'>계정</div>", unsafe_allow_html=True)
            for i, t in enumerate(ALL_TABS):
                hdr_cols[2 + i].markdown(f"<div style='font-size:11px; font-weight:bold; word-break:keep-all; line-height:1.3;' title='{t}'>{t}</div>", unsafe_allow_html=True)
            hdr_cols[-1].markdown("<div style='font-size:12px; font-weight:bold; padding-top:4px;'>전체</div>", unsafe_allow_html=True)
            st.markdown("<hr style='margin:0;border:none;border-top:1px solid #F0EBE2'>", unsafe_allow_html=True)

            all_new_tabs = {}
            with st.container(height=400, border=True):
                for idx, user in enumerate(user_list):
                    uid      = user["uid"]
                    email    = user.get("email", uid)
                    tabs     = user.get("tabs", DEFAULT_USER_TABS.copy())
                    disabled = user.get("disabled", False)

                    row_cols = st.columns(COL_RATIOS)

                    with row_cols[0]:
                        st.markdown(
                            f"<div style='font-size:10px;white-space:nowrap;padding-top:8px'>"
                            f"{'🔴 비활' if disabled else '🟢 활성'}</div>",
                            unsafe_allow_html=True
                        )
                    with row_cols[1]:
                        st.markdown(
                            f"<div style='font-size:10.5px;word-break:break-all;padding-top:8px'>{email}</div>",
                            unsafe_allow_html=True
                        )

                    new_tabs = {}
                    for i, tab_name in enumerate(ALL_TABS):
                        with row_cols[2 + i]:
                            new_tabs[tab_name] = st.checkbox(
                                label="",
                                value=tabs.get(tab_name, False),
                                key=f"ck_{uid}_{tab_name}",
                                label_visibility="collapsed"
                            )

                    with row_cols[-1]:
                        all_checked = all(new_tabs.get(t, False) for t in ALL_TABS)
                        select_all = st.checkbox(
                            label="",
                            value=all_checked,
                            key=f"selall_{uid}",
                            label_visibility="collapsed"
                        )
                        if select_all != all_checked:
                            for t in ALL_TABS:
                                new_tabs[t] = select_all

                    all_new_tabs[uid] = (email, new_tabs)

                    if idx < len(user_list) - 1:
                        st.markdown(
                            "<hr style='margin:0;border:none;border-top:1px solid #F0EBE2'>",
                            unsafe_allow_html=True
                        )

            # ── 일괄 저장 버튼 (하단) ──
            col_save_bot, _ = st.columns([1, 5])
            with col_save_bot:
                save_all_bot = st.button("💾 전체 일괄 저장", key="save_all_tabs_bot", use_container_width=True, type="primary")

            if save_all_bot:
                errors = []
                for uid, (email, new_tabs) in all_new_tabs.items():
                    try:
                        update_user_tabs(uid, new_tabs)
                    except Exception as e:
                        errors.append(f"{email}: {e}")
                if errors:
                    st.error("일부 저장 실패:\n" + "\n".join(errors))
                else:
                    st.success(f"✅ {len(all_new_tabs)}개 계정의 탭 권한이 일괄 저장되었습니다.")
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
            _email_idx     = next((i+1 for i, h in enumerate(_headers) if h.lower().replace("-","") == "email"), None)
            _role_type_idx = next((i+1 for i, h in enumerate(_headers) if "권한유형" in h), None)
            _item_idx      = next((i+1 for i, h in enumerate(_headers) if "품목군"   in h), None)
            if not all([_email_idx, _role_type_idx, _item_idx]):
                return False
            for row_idx, row in enumerate(_data[1:], start=2):
                if str(row[_email_idx-1]).strip().lower() == target_email.strip().lower():
                    _ws.update_cell(row_idx, _role_type_idx, role_type)
                    _ws.update_cell(row_idx, _item_idx, ",".join(item_groups))
                    return True
            return False

        # AUTH_MASTER를 루프 밖에서 1회만 로드 — 429 Quota 방지
        _auth_df_adm, _ec, _dc, _rtc, _igc = load_auth_master()

        # 변경 대상 유저 목록 (고정 관리자 제외)
        role_user_list = [u for u in all_users if u.get("email", u["uid"]) not in admin_emails_list]

        # 고정 관리자 표시
        for user in all_users:
            email = user.get("email", user["uid"])
            if email in admin_emails_list:
                st.markdown(f"🔒 **{email}** — 고정 관리자 (변경 불가)")

        if not role_user_list:
            st.info("역할 변경이 가능한 계정이 없습니다.")
        else:
            # ── CSS: 역할 변경 헤더 (줄바꿈 금지) ──
            st.markdown("""
            <style>
            .role-hdr {
                display: flex;
                align-items: center;
                background: #F3EFE8;
                border: 1px solid #E0D8CC;
                border-radius: 8px 8px 0 0;
                border-bottom: 2px solid #C8B88A;
                padding: 0 4px;
            }
            .role-hdr-cell {
                font-family: 'Noto Sans KR', sans-serif;
                font-size: 10.5px;
                font-weight: 700;
                color: #5C4A2A;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                padding: 9px 6px;
                letter-spacing: 0.2px;
                flex-shrink: 0;
            }
            .role-hdr-cell.rh-email  { flex: 2.5; min-width: 160px; }
            .role-hdr-cell.rh-role   { flex: 1.2; min-width: 100px; text-align: center; }
            .role-hdr-cell.rh-auth   { flex: 1.5; min-width: 100px; text-align: center; }
            .role-hdr-cell.rh-items  { flex: 2.5; min-width: 160px; }
            </style>
            """, unsafe_allow_html=True)

            ROLE_COL = [2.5, 1.2, 1.5, 2.5]
            hdr_cols2 = st.columns(ROLE_COL)
            hdr_cols2[0].markdown("**계정**")
            hdr_cols2[1].markdown("**역할 (Firestore)**")
            hdr_cols2[2].markdown("**AUTH 권한유형**")
            hdr_cols2[3].markdown("**품목군 (PM 전용)**")
            st.markdown("<hr style='margin:0;border:none;border-top:1px solid #F0EBE2'>", unsafe_allow_html=True)
            role_changes = {}
            auth_changes = {}

            # ── 유저 행 스크롤 컨테이너 ──
            with st.container(height=400, border=True):
                for idx, user in enumerate(role_user_list):
                    uid   = user["uid"]
                    email = user.get("email", uid)
                    role  = user.get("role", "user")

                    # AUTH_MASTER 현재값
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
                        _cur_item_groups = [
                            x.strip() for x in _raw_ig.split(",")
                            if x.strip() and x.strip() != "ALL"
                        ]

                    rcols = st.columns(ROLE_COL)
                    with rcols[0]:
                        st.markdown(
                            f"<div style='font-size:10.5px;word-break:break-all;padding-top:8px'>{email}</div>",
                            unsafe_allow_html=True
                        )
                    with rcols[1]:
                        new_role = st.selectbox(
                            label="",
                            options=["user", "admin"],
                            index=0 if role == "user" else 1,
                            key=f"role_{uid}",
                            label_visibility="collapsed"
                        )
                    with rcols[2]:
                        new_role_type = st.selectbox(
                            label="",
                            options=["관리자", "부서기반", "PM"],
                            index=(
                                ["관리자", "부서기반", "PM"].index(_cur_role_type)
                                if _cur_role_type in ["관리자", "부서기반", "PM"] else 1
                            ),
                            key=f"auth_role_type_{uid}",
                            label_visibility="collapsed"
                        )
                    with rcols[3]:
                        new_items = st.multiselect(
                            label="",
                            options=_item_group_options,
                            default=[ig for ig in _cur_item_groups if ig in _item_group_options],
                            key=f"auth_items_{uid}",
                            placeholder="PM 권한 시 품목군 선택",
                            label_visibility="collapsed"
                        )

                    role_changes[uid] = (email, new_role)
                    auth_changes[uid] = (email, new_role_type, new_items)

                    if idx < len(role_user_list) - 1:
                        st.markdown(
                            "<hr style='margin:0;border:none;border-top:1px solid #F0EBE2'>",
                            unsafe_allow_html=True
                        )

            # ── 일괄 저장 버튼 (하단) ──
            col_rsave_bot, _ = st.columns([1, 5])
            with col_rsave_bot:
                role_save_bot = st.button("💾 전체 일괄 저장", key="save_all_roles_bot", use_container_width=True, type="primary")

            if role_save_bot:
                r_errors = []
                r_ok_count = 0
                for uid, (email, new_role) in role_changes.items():
                    try:
                        update_user_role(uid, new_role)
                        if new_role == "admin":
                            update_user_tabs(uid, DEFAULT_ADMIN_TABS.copy())
                        r_ok_count += 1
                    except Exception as e:
                        r_errors.append(f"{email} (역할): {e}")

                a_errors = []
                a_ok_count = 0
                for uid, (email, role_type, items) in auth_changes.items():
                    try:
                        _save_items = items if role_type == "PM" else ["ALL"]
                        ok = update_auth_master_row(email, role_type, _save_items)
                        if ok:
                            a_ok_count += 1
                        else:
                            a_errors.append(f"{email}: AUTH_MASTER 행 없음 또는 컬럼 구조 오류")
                    except Exception as e:
                        a_errors.append(f"{email} (AUTH): {e}")

                load_auth_master.clear()

                all_errors = r_errors + a_errors
                if all_errors:
                    st.warning(f"⚠️ 일부 저장 실패:\n" + "\n".join(all_errors))
                else:
                    st.success(
                        f"✅ 역할 저장 {r_ok_count}건 / AUTH_MASTER 저장 {a_ok_count}건 완료"
                    )
                    st.rerun()

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
    st.markdown('</div>', unsafe_allow_html=True)




# ===================================
# TAB 7: 제품별 원가
# ===================================
if current_tab_key == "제품별원가":
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
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
    st.markdown('</div>', unsafe_allow_html=True)

st.success("🚀 Lingtea Dashboard v10 Ready")