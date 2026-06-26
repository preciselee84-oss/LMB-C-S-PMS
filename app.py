# VERSION: 20260518_staff_filter_deadline
import streamlit as st
import pandas as pd
import numpy as np
import time
import json
import os
import html
import re
import base64
import hashlib
import requests as _requests
from io import BytesIO
from datetime import datetime, timedelta
from streamlit_cookies_controller import CookieController

st.set_page_config(page_title="내부 관리", layout="wide", initial_sidebar_state="expanded")


def resolve_template_file(file_name):
    base_dir = os.path.dirname(__file__)
    candidates = [
        os.path.join(base_dir, "templates", file_name),
        os.path.join(base_dir, "github-work", "templates", file_name),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0]


DB_FILE = "users.json"
PERF_FILE = "manual_perf.json"
SAVED_STATE_FILE = "saved_state.json"
SERVER_CONNECTION_FILE = "server_connection.json"
DELEGATED_WORKPLACE_FILE = "delegated_workplaces.json"
APPROVAL_FILE = "approvals.json"
BANK_ACCOUNT_FILE = "bank_accounts.json"
USAGE_REPORT_FILE = "usage_reports.json"
COMPANY_PROFILE_FILE = "company_profile.json"
EXCEL_SAMPLE_FILE = resolve_template_file("LMB월간 활동실적_000000(샘플).xlsx")

DEFAULT_URL_ANALYSIS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9XPHqrqcaFf9bCOVya7yHORr-c1R4KCF0eEpdE3ESn8qJELP0BkqTOslur9bsGcVabRUIcyOa877R/pub?output=csv"
DEFAULT_URL_SYNC = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9F7R7oLA2B02H-I25kVv2JeYHFgWQq0CT7TeW61hrNpJLdHWJFhFR_iDQGCFAW044o8rRwBDeovKG/pub?gid=1533424484&single=true&output=csv"
DEFAULT_URL_HANA = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQgRHnTZD4eDW2UeODQuGxmxFrflKpbQda3sBsVjj1s3qAFWMKcpke2U58UuT6VEDlkbXveZlaroTCr/pub?gid=0&single=true&output=csv"
DEFAULT_URL_HANA_BILLING = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQgRHnTZD4eDW2UeODQuGxmxFrflKpbQda3sBsVjj1s3qAFWMKcpke2U58UuT6VEDlkbXveZlaroTCr/pub?gid=1172734914&single=true&output=csv"
DEFAULT_URL_HANA_PERFORMANCE = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS2CUE3No1cptBOTehN8r1xoTQyUni07sjbut-f1Teo9mpB-rcJgpE5xfI6dTy0M4IUxSg8Mv5_uT4l/pub?gid=1749034066&single=true&output=csv"


@st.cache_data(ttl=300, show_spinner=False)
def read_csv_cached(url, header=0, cache_buster=0):
    url = normalize_google_sheet_csv_url(url)
    return pd.read_csv(url, header=header)


def read_google_csv(url, header=0, force_refresh=False):
    cache_buster = int(datetime.utcnow().timestamp()) if force_refresh else 0
    return read_csv_cached(url, header=header, cache_buster=cache_buster).copy()


def normalize_google_sheet_csv_url(url):
    """Google Sheets pubhtml/pub links are normalized to direct CSV export."""
    if not isinstance(url, str):
        return url
    normalized = url.strip()
    if "docs.google.com/spreadsheets" not in normalized:
        return normalized
    normalized = normalized.replace("/pubhtml?", "/pub?")
    normalized = normalized.replace("/pubhtml", "/pub")
    if "/pub?" in normalized and "output=csv" not in normalized:
        separator = "&" if "?" in normalized else "?"
        normalized = f"{normalized}{separator}output=csv"
    return normalized


GITHUB_REPO = "preciselee84-oss/LMB-C-S-PMS"
GITHUB_BRANCH = "main"
GITHUB_DATA_DIR = "data"
SESSION_UID_COOKIE = "auto_login_uid"
LAST_MENU_COOKIE = "last_menu"
BILLING_MENU = "청구자료 생성"


def _get_github_token():
    try:
        return st.secrets.get("GITHUB_TOKEN", "")
    except Exception:
        return os.environ.get("GITHUB_TOKEN", "")


# ── DART 전자공시 API ─────────────────────────────────────────────

def _get_dart_api_key():
    # 1) 설정 메뉴에서 입력한 값 우선
    session_key = st.session_state.get("dart_api_key", "")
    if session_key:
        return session_key
    # 2) Streamlit secrets → 환경변수
    try:
        return st.secrets.get("DART_API_KEY", "")
    except Exception:
        return os.environ.get("DART_API_KEY", "")


@st.cache_data(ttl=86400)
def _load_dart_corp_map(api_key):
    """DART 기업코드 목록 로드 ({정규화기업명: corp_code}). 24시간 캐시."""
    if not api_key:
        return {}
    try:
        import zipfile, io
        import xml.etree.ElementTree as ET
        resp = _requests.get(
            f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={api_key}",
            timeout=30,
        )
        if resp.status_code != 200:
            return {}
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            xml_bytes = zf.read("CORPCODE.xml")
        root = ET.fromstring(xml_bytes.decode("utf-8"))
        return {
            item.findtext("corp_name", "").strip().replace(" ", "").lower(): item.findtext("corp_code", "").strip()
            for item in root.findall("list")
            if item.findtext("corp_name", "").strip() and item.findtext("corp_code", "").strip()
        }
    except Exception:
        return {}


@st.cache_data(ttl=3600)
def _dart_get_filings(api_key, corp_code, months=6):
    """DART 최근 N개월 공시 목록. 1시간 캐시."""
    if not api_key or not corp_code:
        return []
    try:
        today = datetime.utcnow() + timedelta(hours=9)
        bgn = (today - timedelta(days=months * 31)).strftime("%Y%m%d")
        end = today.strftime("%Y%m%d")
        resp = _requests.get(
            "https://opendart.fss.or.kr/api/list.json",
            params={
                "crtfc_key": api_key, "corp_code": corp_code,
                "bgn_de": bgn, "end_de": end, "page_count": 10,
            },
            timeout=10,
        )
        d = resp.json()
        return d.get("list", []) if d.get("status") == "000" else []
    except Exception:
        return []


# 유통활동 관련 공시 키워드 → (가산점수, 사유 설명)
_DART_KW_SCORE = [
    ("유상증자",  30, "유상증자 공시 → 자금 유입, 이체 활성화 기대"),
    ("무상증자",  15, "무상증자 공시 → 재무구조 개선"),
    ("합병",      25, "합병 공시 → 계좌 통합 니즈 발생"),
    ("분할",      20, "기업분할 공시 → 신규 계좌 가능성"),
    ("수주",      20, "수주 공시 → 매출 증가·자금 흐름 활성화"),
    ("신규사업",  15, "신규사업 공시 → 향후 거래 확대 가능"),
    ("투자",      10, "투자 공시 → 자금 유출입 예상"),
]


def _dart_enrich(api_key, corp_map, company_name):
    """기업명으로 DART 공시 조회 후 (가산점수, 사유 목록) 반환."""
    key = str(company_name).strip().replace(" ", "").lower()
    corp_code = corp_map.get(key)
    if not corp_code:
        return 0, []
    filings = _dart_get_filings(api_key, corp_code)
    if not filings:
        return 0, []
    bonus, reasons, seen = 0, [], set()
    for f in filings:
        nm = f.get("report_nm", "")
        dt = str(f.get("rcept_dt", ""))[:8]
        for kw, pts, desc in _DART_KW_SCORE:
            if kw in nm and kw not in seen:
                bonus = max(bonus, pts)
                reasons.append(f"[DART공시] {desc} ({dt[:4]}.{dt[4:6]}.{dt[6:]})")
                seen.add(kw)
    return bonus, reasons


# ── 국세청 사업자등록 상태조회 API ─────────────────────────────────

def _get_nts_api_key():
    session_key = st.session_state.get("nts_api_key", "")
    if session_key:
        return session_key
    try:
        return st.secrets.get("NTS_API_KEY", "")
    except Exception:
        return os.environ.get("NTS_API_KEY", "")


def _check_business_status(api_key, biz_numbers):
    """사업자등록번호 목록의 휴폐업 상태를 조회. {사업자번호(숫자만): 상태} 반환."""
    clean_numbers = sorted({re.sub(r"\D", "", str(b)) for b in biz_numbers if str(b).strip()})
    if not api_key or not clean_numbers:
        return {}
    try:
        resp = _requests.post(
            f"https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey={api_key}",
            json={"b_no": clean_numbers},
            headers={"Content-Type": "application/json"},
            timeout=10,
        )
        resp.raise_for_status()
        data = resp.json().get("data", [])
        return {item.get("b_no", ""): item.get("b_stt", "") or "조회결과 없음" for item in data}
    except Exception:
        return {}


# ── 예금주조회 (계좌 실명조회 API 연동 전 임시 구현) ────────────────

def _lookup_account_holder_name(bank_name, account_number, expected_holder=""):
    """계좌 실명조회 API 연동 전까지, 등록된 입금은행/계좌번호가 있으면 예상예금주를 조회 결과로 반환."""
    if not bank_name or not account_number:
        return ""
    return expected_holder


def _github_save(file_path, data):
    token = _get_github_token()
    if not token:
        return False
    try:
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        gh_path = f"{GITHUB_DATA_DIR}/{file_path}"
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{gh_path}"
        resp = _requests.get(url, headers=headers, params={"ref": GITHUB_BRANCH}, timeout=10)
        sha = resp.json().get("sha") if resp.status_code == 200 else None
        content = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=4).encode()).decode()
        payload = {"message": f"auto-save {file_path}", "content": content, "branch": GITHUB_BRANCH}
        if sha:
            payload["sha"] = sha
        resp = _requests.put(url, json=payload, headers=headers, timeout=15)
        return resp.status_code in (200, 201)
    except Exception:
        return False


def _github_load(file_path, default_data):
    token = _get_github_token()
    if not token:
        return default_data
    try:
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        gh_path = f"{GITHUB_DATA_DIR}/{file_path}"
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{gh_path}"
        resp = _requests.get(url, headers=headers, params={"ref": GITHUB_BRANCH}, timeout=10)
        if resp.status_code == 200:
            content = base64.b64decode(resp.json()["content"]).decode()
            return json.loads(content)
    except Exception:
        pass
    return default_data


def load_db(file_path, default_data):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    data = _github_load(file_path, default_data)
    if data != default_data:
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
    return data


def save_db(file_path, data, allow_shrink=False):
    if file_path == DB_FILE and not allow_shrink:
        existing = load_user_db()
        if len(_real_users(existing)) > len(_real_users(data)):
            data = merge_user_db(existing, data)
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    # 직원DB는 .bak 백업 파일도 항상 동기화 (로컬 데이터 유실 방지)
    if file_path == DB_FILE:
        try:
            with open(file_path + ".bak", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
    _github_save(file_path, data)


def safe_cookie_controller():
    try:
        return CookieController()
    except Exception:
        return None


def cookie_get(cookie_manager, key, default=""):
    if cookie_manager is None:
        return default
    try:
        return cookie_manager.get(key) or default
    except Exception:
        return default


def cookie_set(cookie_manager, key, value, **kwargs):
    if cookie_manager is None:
        return False
    try:
        cookie_manager.set(key, value, **kwargs)
        return True
    except TypeError:
        try:
            cookie_manager.set(key, value)
            return True
        except Exception:
            return False
    except Exception:
        return False


def cookie_remove(cookie_manager, key):
    if cookie_manager is None:
        return False
    try:
        cookie_manager.remove(key)
        return True
    except Exception:
        return False


def restore_saved_work_state(user_name):
    saved_db = load_db(SAVED_STATE_FILE, {})
    user_saved = saved_db.get(user_name)
    if user_saved:
        if user_saved.get("user_excel_data"):
            st.session_state.user_excel_data = pd.DataFrame.from_dict(user_saved["user_excel_data"])
        if user_saved.get("user_prev_month_sel"):
            st.session_state.user_prev_month_sel = user_saved["user_prev_month_sel"]

    admin_analysis = saved_db.get("admin_analysis")
    if admin_analysis and admin_analysis.get("sent_df"):
        st.session_state.analysis_result = pd.DataFrame.from_dict(admin_analysis["sent_df"])
    if admin_analysis and admin_analysis.get("adm_prev_month"):
        st.session_state.adm_prev_month = admin_analysis["adm_prev_month"]
    deadline_info = saved_db.get("deadline")
    if deadline_info:
        st.session_state.deadline_time = deadline_info.get("time", "")
    report_closed_info = saved_db.get("report_closed")
    if report_closed_info:
        st.session_state.report_closed = report_closed_info.get("time", "")


def restore_login_from_cookie():
    if st.session_state.get("logged_in"):
        return
    cookie_manager = safe_cookie_controller()
    uid = str(cookie_get(cookie_manager, SESSION_UID_COOKIE, "")).strip()
    if not uid:
        return

    db = st.session_state.get("user_db", {})
    is_super = uid == "1"
    is_user = uid in db and db[uid].get("access") == "허용"
    if not (is_super or is_user):
        cookie_remove(cookie_manager, SESSION_UID_COOKIE)
        return

    user = db.get(uid, {"role": "관리자", "name": "최고관리자"})
    st.session_state.logged_in = True
    st.session_state.user_role = user.get("role", "관리자")
    st.session_state.user_name = user.get("name", "최고관리자")
    if not st.session_state.get("login_time"):
        st.session_state.login_time = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")

    last_menu = str(cookie_get(cookie_manager, LAST_MENU_COOKIE, "")).strip()
    if last_menu:
        st.session_state.current_menu = last_menu

    restore_saved_work_state(st.session_state.user_name)


def persist_login_session(user_id):
    cookie_manager = safe_cookie_controller()
    cookie_set(cookie_manager, SESSION_UID_COOKIE, str(user_id), max_age=60 * 60 * 24 * 30)


def persist_current_menu():
    cookie_manager = safe_cookie_controller()
    menu = st.session_state.get("current_menu", "")
    if menu:
        cookie_set(cookie_manager, LAST_MENU_COOKIE, str(menu), max_age=60 * 60 * 24 * 30)


def send_kakao_notify(message):
    try:
        token = ""
        try:
            token = st.secrets.get("KAKAO_TOKEN", "")
        except Exception:
            token = os.environ.get("KAKAO_TOKEN", "")
        if not token:
            return False
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/x-www-form-urlencoded"}
        data = {"template_object": json.dumps({"object_type": "text", "text": message, "link": {"web_url": "", "mobile_web_url": ""}, "button_title": "확인"})}
        resp = _requests.post("https://kapi.kakao.com/v2/api/talk/memo/default/send", headers=headers, data=data, timeout=10)
        return resp.status_code == 200
    except Exception:
        return False


def get_staff_names():
    return {
        info.get("name")
        for uid, info in st.session_state.get("user_db", {}).items()
        if uid != "1"
        and info.get("name")
        and info.get("dept_type", "사업부") == "C&S"
    }


def get_current_user_rank():
    name = st.session_state.get("user_name", "")
    for uid, info in st.session_state.get("user_db", {}).items():
        if info.get("name") == name:
            return info.get("rank", "직원")
    return "직원"


def filter_by_staff(df, name_col="담당자"):
    if df.empty or name_col not in df.columns:
        return df
    staff_names = get_staff_names()
    if not staff_names:
        return df
    return df[df[name_col].isin(staff_names)].reset_index(drop=True)


def hide_department_heads(df):
    if df is None or df.empty or "직급" not in df.columns:
        return df
    return df[df["직급"].astype(str).str.strip() != "부서장"].reset_index(drop=True)


def dataframe_to_upload_payload(df):
    safe_df = strip_activity_time_columns(df).copy()
    safe_df = safe_df.replace({np.nan: ""})
    return {
        "columns": [str(c) for c in safe_df.columns],
        "rows": safe_df.astype(str).values.tolist(),
        "saved_at": (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S"),
    }


def upload_payload_to_dataframe(payload):
    if not payload:
        return pd.DataFrame()
    return pd.DataFrame(payload.get("rows", []), columns=payload.get("columns", []))


_USER_DB_DEFAULT = {
    "1": {
        "pw": "1",
        "name": "최고관리자",
        "email": "",
        "access": "허용",
        "role": "관리자",
        "dept_type": "사업부",
        "staff_type": "정규직",
        "outsource": "아니오",
        "outsource_period": "해당없음",
    },
    "T": {
        "pw": "1111",
        "name": "이성환",
        "email": "",
        "access": "허용",
        "role": "관리자",
        "dept_type": "사업부",
        "staff_type": "정규직",
        "outsource": "아니오",
        "outsource_period": "해당없음",
    },
    "iminjee": {
        "pw": "1111",
        "name": "임인지",
        "email": "",
        "access": "허용",
        "role": "사용자",
        "dept_type": "사업부",
        "staff_type": "정규직",
        "outsource": "아니오",
        "outsource_period": "해당없음",
    },
    "kangkt": {
        "pw": "1111",
        "name": "강경태",
        "email": "",
        "access": "허용",
        "role": "사용자",
        "dept_type": "사업부",
        "staff_type": "정규직",
        "outsource": "아니오",
        "outsource_period": "해당없음",
    },
    "leesh": {
        "pw": "1111",
        "name": "이수현",
        "email": "",
        "access": "허용",
        "role": "사용자",
        "dept_type": "사업부",
        "staff_type": "정규직",
        "outsource": "아니오",
        "outsource_period": "해당없음",
    },
    "gilmj": {
        "pw": "1111",
        "name": "길민종",
        "email": "",
        "access": "허용",
        "role": "사용자",
        "dept_type": "사업부",
        "staff_type": "정규직",
        "outsource": "아니오",
        "outsource_period": "해당없음",
    },
    "maengks": {
        "pw": "1111",
        "name": "맹국성",
        "email": "",
        "access": "허용",
        "role": "사용자",
        "dept_type": "사업부",
        "staff_type": "정규직",
        "outsource": "아니오",
        "outsource_period": "해당없음",
    },
}


def _real_users(d):
    """_deleted 같은 메타키를 제외한 실제 사용자 수 반환."""
    return {k: v for k, v in d.items() if not k.startswith("_")}


def load_user_db():
    """users.json → users.json.bak → GitHub → {} 순서로 로드 후 기본 계정과 병합하여 반환.
    _deleted 목록에 있는 ID는 기본 계정이라도 복원하지 않는다."""
    candidates = []

    # 1) 메인 파일
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict):
                candidates.append(d)
        except Exception:
            pass

    # 2) 백업 파일
    bak = DB_FILE + ".bak"
    if os.path.exists(bak):
        try:
            with open(bak, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict):
                candidates.append(d)
        except Exception:
            pass

    # 3) GitHub
    gh = _github_load(DB_FILE, None)
    if isinstance(gh, dict):
        candidates.append(gh)

    loaded = {}
    for candidate in candidates:
        if len(_real_users(candidate)) > len(_real_users(loaded)):
            loaded = candidate

    # 삭제된 ID 목록 추출 (파일에만 존재하는 메타 키)
    deleted_ids = set(loaded.get("_deleted", []))

    if len(_real_users(loaded)) > 1:
        try:
            with open(DB_FILE, "w", encoding="utf-8") as f:
                json.dump(loaded, f, ensure_ascii=False, indent=4)
            with open(DB_FILE + ".bak", "w", encoding="utf-8") as f:
                json.dump(loaded, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    # 기본 계정 베이스 + 파일 데이터 병합, _deleted ID는 제외
    merged = {**_USER_DB_DEFAULT, **_real_users(loaded)}
    for uid in deleted_ids:
        merged.pop(uid, None)
    return merged


def merge_user_db(existing, incoming):
    deleted_ids = set(incoming.get("_deleted", [])) | set(existing.get("_deleted", []))
    merged = {**_USER_DB_DEFAULT}
    if isinstance(existing, dict):
        merged.update(_real_users(existing))
    if isinstance(incoming, dict):
        for uid, info in _real_users(incoming).items():
            if uid not in merged or len(_real_users(incoming)) >= len(_real_users(existing or {})):
                merged[uid] = info
    for uid in deleted_ids:
        merged.pop(uid, None)
    return merged


def init_state():
    if "user_db" not in st.session_state:
        st.session_state.user_db = load_user_db()

    defaults = {
        "logged_in": False,
        "user_role": "사용자",
        "user_name": "",
        "auth_mode": "login",
        "current_menu": BILLING_MENU,
        "url_analysis": DEFAULT_URL_ANALYSIS,
        "url_sync": DEFAULT_URL_SYNC,
        "url_hana": DEFAULT_URL_HANA,
        "url_hana_billing": DEFAULT_URL_HANA_BILLING,
        "url_hana_performance": DEFAULT_URL_HANA_PERFORMANCE,
        "dart_api_key": "",
        "hana_sheet_df": None,
        "hana_billing_df": None,
        "analysis_lookup_df": None,
        "cloud_sheet_df": None,
        "analysis_result": None,
        "user_excel_data": None,
        "temp_cloud_df": None,
        "auto_prev_df": None,
        "deadline_time": "",
        "report_closed": "",
        "dark_mode": False,
        "login_time": "",
        "_prev_menu": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if st.session_state.current_menu == "구글 스트레드시트 연동":
        st.session_state.current_menu = "서버 접속 정보"

    transfer_result_upload_menus = {
        "이체 결과 엑셀 업로드",
        "이체결과 엑셀 업로드",
        "이체 결과 업로드",
        "지급 결과 엑셀 업로드",
    }
    if st.session_state.current_menu in transfer_result_upload_menus:
        st.session_state.current_menu = "지급 결과 확인"

    removed_menus = {
        "이력확인 및 작성",
        "은행 이력 업로드",
        "업로드 및 실적 확인",
        "이번달 활동 대상고객 추천",
        "최종 실적 확인",
        "관리자용 실적 확인",
        "실적 분석/계산",
        "실적 보고서",
        "청구자료 작성",
        "주간보고 이력 작성",
        "방문이력 작성",
        "주간보고 취합",
        "운영계획",
    }
    if st.session_state.current_menu in removed_menus:
        st.session_state.current_menu = "전도금 요청"

    RENAMED_MENUS = {
        "사업장 정보 등록": "위탁 사업장 관리",
        "사업장 예측/보고": BILLING_MENU,
        "보고서": BILLING_MENU,
        "직원 및 권한설정": "위탁 사업장 관리",
        "계좌 관리": "위탁 사업장 관리",
        "담당자 관리": "위탁 사업장 관리",
        "사업장 관리": "위탁 사업장 관리",
        "이체 자료 생성": "이체 자료 확정",
        "사용품의서 보고": "전도금 사용 결의 보고",
    }
    if st.session_state.current_menu in RENAMED_MENUS:
        st.session_state.current_menu = RENAMED_MENUS[st.session_state.current_menu]
    if st.session_state.current_menu != BILLING_MENU:
        st.session_state.current_menu = BILLING_MENU


init_state()


def clean_header_logic(df):
    try:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        if len(df.columns) > 0 and str(df.columns[0]).startswith("Unnamed"):
            for i in range(min(len(df), 10)):
                row_text = " ".join(df.iloc[i].astype(str).tolist())
                if any(k in row_text for k in ["등록자", "담당자", "성명", "업체명", "사업자번호", "고객번호"]):
                    df.columns = [str(c).strip() for c in df.iloc[i]]
                    df = df.iloc[i + 1:].reset_index(drop=True)
                    break

        keep = ~pd.Series(df.columns).astype(str).str.contains("^Unnamed|^nan", case=False, na=False).values
        df = df.loc[:, keep]

        # 중요한 컬럼은 값이 비어있어도 유지
        important_keywords = ["업무번호", "플로우", "식권", "비즈플레이"]
        important_cols = [col for col in df.columns if any(keyword in str(col) for keyword in important_keywords)]

        # 빈 컬럼 제거 (단, 중요 컬럼은 제외)
        empty_cols = df.columns[df.isna().all()]
        cols_to_drop = [col for col in empty_cols if col not in important_cols]
        df = df.drop(columns=cols_to_drop, errors="ignore")

        # 빈 행 제거
        df = df.dropna(how="all", axis=0)
        return strip_activity_time_columns(df)
    except Exception:
        return df


def strip_activity_time_columns(df):
    try:
        df = df.copy()
        for col in df.columns:
            col_name = str(col).replace(" ", "").replace("　", "")
            if "활동일" not in col_name:
                continue

            converted = pd.to_datetime(df[col], errors="coerce")
            valid = converted.notna()
            if valid.any():
                df.loc[valid, col] = converted.loc[valid].dt.strftime("%Y-%m-%d")
        return df
    except Exception:
        return df


def find_col(df, keys, fallback=None):
    for c in df.columns:
        c_normalized = str(c).replace(" ", "").replace("　", "").lower()
        if any(k.replace(" ", "").replace("　", "").lower() in c_normalized for k in keys):
            return c
    return fallback


def is_blank_value(value):
    text = "" if pd.isna(value) else str(value).strip()
    return text == "" or text.lower() in ["nan", "nat", "none", "-", "null"]


def parse_sheet_date(value):
    if is_blank_value(value):
        return pd.NaT
    text = str(value).strip()
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 8:
        parsed = pd.to_datetime(digits[:8], format="%Y%m%d", errors="coerce")
        if pd.notna(parsed):
            return parsed
    return pd.to_datetime(text, errors="coerce")


def filter_visit_rows(df):
    if df is None:
        return pd.DataFrame()
    if df.empty:
        return df.copy()
    visit_col = find_col(df, ["활동구분", "접수유형"])
    if not visit_col or visit_col not in df.columns:
        return df.copy()
    mask = df[visit_col].astype(str).str.strip().str.contains("방문", na=False)
    return df[mask].copy()


def render_history_search_filters(source_df, key_prefix):
    st.markdown("##### 검색 조건")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        company = st.text_input("업체명", key=f"{key_prefix}_company")

    date_options = ["전체"]
    category_options = ["전체"]
    detail_options = ["전체"]
    if source_df is not None and not source_df.empty:
        date_col = find_col(source_df, ["활동일자", "활동일", "초과일자", "일자"])
        category_col = find_col(source_df, ["활동구분", "접수유형"])
        detail_col = find_col(source_df, ["활동상세", "활동내용"])
        if date_col and date_col in source_df.columns:
            _dates = pd.to_datetime(source_df[date_col], errors="coerce").dropna().dt.strftime("%Y-%m-%d").unique()
            date_options += sorted(_dates, reverse=True)
        if category_col and category_col in source_df.columns:
            category_options += sorted(v for v in source_df[category_col].astype(str).str.strip().unique() if v)
        if detail_col and detail_col in source_df.columns:
            detail_options += sorted(v for v in source_df[detail_col].astype(str).str.strip().unique() if v)

    with c2:
        activity_date = st.selectbox("활동일자", date_options, key=f"{key_prefix}_date")
    with c3:
        activity_category = st.selectbox("활동구분", category_options, key=f"{key_prefix}_category")
    with c4:
        activity_detail = st.selectbox("활동상세", detail_options, key=f"{key_prefix}_detail")

    return {
        "company": company.strip(),
        "date": activity_date,
        "category": activity_category,
        "detail": activity_detail,
    }


def apply_history_search_filters(df, filters):
    if df is None or df.empty:
        return df
    if not filters:
        return df
    result = df.copy()
    company_col = find_col(result, ["업체명", "상호", "고객명"])
    date_col = find_col(result, ["활동일자", "활동일", "초과일자", "일자"])
    category_col = find_col(result, ["활동구분", "접수유형"])
    detail_col = find_col(result, ["활동상세", "활동내용"])

    if filters.get("company") and company_col in result.columns:
        result = result[result[company_col].astype(str).str.contains(filters["company"], case=False, na=False)]
    if filters.get("date") and filters["date"] != "전체" and date_col in result.columns:
        _date_values = pd.to_datetime(result[date_col], errors="coerce").dt.strftime("%Y-%m-%d")
        result = result[_date_values == filters["date"]]
    if filters.get("category") and filters["category"] != "전체" and category_col in result.columns:
        result = result[result[category_col].astype(str).str.strip() == filters["category"]]
    if filters.get("detail") and filters["detail"] != "전체" and detail_col in result.columns:
        result = result[result[detail_col].astype(str).str.strip() == filters["detail"]]
    return result


def has_active_history_filters(filters):
    if not filters:
        return False
    return bool(filters.get("company")) or filters.get("date") != "전체" or filters.get("category") != "전체" or filters.get("detail") != "전체"


def history_filter_signature(filters):
    if not filters:
        return "all"
    raw = "|".join(str(filters.get(key, "")) for key in ["company", "date", "category", "detail"])
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:10]


def normalize_biz(series):
    def normalize_one(value):
        if pd.isna(value):
            return ""

        # 숫자 타입이면 정수로 변환 (과학적 표기법 방지)
        if isinstance(value, (int, float, np.integer, np.floating)):
            try:
                value = int(float(value))
            except (ValueError, OverflowError):
                pass

        text = str(value).strip().replace(",", "")
        if re.fullmatch(r"\d+\.0+", text):
            text = text.split(".", 1)[0]

        digits = re.sub(r"[^0-9]", "", text)
        if "." in str(value) and digits.endswith("0") and len(digits) == 11:
            digits = digits[:-1]
        return digits

    if isinstance(series, pd.Series):
        return series.apply(normalize_one)
    return normalize_one(series)


def get_uploaded_month(df):
    try:
        d_col = find_col(df, ["활동일", "일자"])
        if d_col and d_col in df.columns:
            dates = pd.to_datetime(df[d_col], errors="coerce").dropna()
            if not dates.empty:
                return dates.dt.strftime("%Y-%m").value_counts().idxmax()
    except Exception:
        pass
    return ""


def attach_cloud_dates(user_df):
    df = user_df.copy()
    cloud = st.session_state.get("cloud_sheet_df")

    if cloud is None or df.empty:
        if cloud is None and not df.empty:
            st.warning("⚠️ 본사 구글시트 데이터를 불러올 수 없습니다. 관리자 > 본사 URL 설정에서 URL을 확인해주세요.")
        return df

    cloud = clean_header_logic(cloud.copy())

    biz_col_user = find_col(df, ["사업자번호"])
    biz_col_cloud = find_col(cloud, ["사업자번호"])
    open_col = find_col(cloud, ["개설완료일자", "개설일"])
    erp_col = find_col(cloud, ["ERP연계일자", "연계일자"])

    if not biz_col_user or not biz_col_cloud:
        return df

    div_col = find_col(cloud, ["신규/이행구분", "이행구분", "신규이행"])
    add_col = find_col(cloud, ["이행추가연계"])

    cloud_cols = [biz_col_cloud]
    rename_map = {}

    if open_col:
        cloud_cols.append(open_col)
        rename_map[open_col] = "본사 개설완료일자"
    if erp_col:
        cloud_cols.append(erp_col)
        rename_map[erp_col] = "본사 ERP연계일자"
    if div_col:
        cloud_cols.append(div_col)
        rename_map[div_col] = "본사 신규이행구분"
    if add_col:
        cloud_cols.append(add_col)
        rename_map[add_col] = "본사 이행추가연계"

    if len(cloud_cols) == 1:
        return df

    df["_biz_key"] = normalize_biz(df[biz_col_user])
    cloud["_biz_key"] = normalize_biz(cloud[biz_col_cloud])

    cloud_map = cloud[["_biz_key"] + [c for c in cloud_cols if c != biz_col_cloud]].rename(columns=rename_map)
    cloud_map = cloud_map.drop_duplicates("_biz_key")

    df = pd.merge(df, cloud_map, on="_biz_key", how="left", suffixes=("", "_cloud"))
    for col in ["본사 개설완료일자", "본사 ERP연계일자", "본사 신규이행구분", "본사 이행추가연계"]:
        cloud_col = f"{col}_cloud"
        if cloud_col in df.columns:
            if col in df.columns:
                df[col] = df[col].where(df[col].notna() & (df[col].astype(str).str.strip() != ""), df[cloud_col])
            else:
                df[col] = df[cloud_col]
            df = df.drop(columns=[cloud_col], errors="ignore")
    return df.drop(columns=["_biz_key"], errors="ignore")


def build_other_validation_errors(df):
    other_errors = []
    if df is None or df.empty:
        return pd.DataFrame()

    biz_col = find_col(df, ["사업자번호"], "사업자번호")
    comp_col = find_col(df, ["업체명", "상호"], "업체명")
    user_col = find_col(df, ["등록자", "담당자", "성명"], "등록자")
    open_col = "본사 개설완료일자"
    erp_col = "본사 ERP연계일자"

    if open_col not in df.columns or erp_col not in df.columns:
        return pd.DataFrame()

    def parse_compare_date(value):
        if pd.isna(value):
            return pd.NaT
        text = str(value).strip()
        digits = re.sub(r"[^0-9]", "", text)
        if len(digits) >= 8:
            parsed = pd.to_datetime(digits[:8], format="%Y%m%d", errors="coerce")
            if pd.notna(parsed):
                return parsed
        return pd.to_datetime(text, errors="coerce")

    df_check = df.copy()
    df_check["_open_date"] = df_check[open_col].apply(parse_compare_date)
    df_check["_erp_date"] = df_check[erp_col].apply(parse_compare_date)

    for _, row in df_check.iterrows():
        open_date = row["_open_date"]
        erp_date = row["_erp_date"]
        if pd.notna(open_date) and pd.notna(erp_date) and erp_date.normalize() < open_date.normalize():
            other_errors.append({
                "업체명": row.get(comp_col, "") if comp_col else "",
                "사업자번호": row.get(biz_col, "") if biz_col else "",
                "등록자": row.get(user_col, "") if user_col else "",
                open_col: open_date.strftime("%Y-%m-%d"),
                erp_col: erp_date.strftime("%Y-%m-%d"),
                "오류 사유": "본사 ERP연계일자가 개설완료일자보다 이전",
            })

    return pd.DataFrame(other_errors)


def normalize_customer_no(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().replace(",", "")
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    digits = re.sub(r"[^0-9]", "", text)
    if digits and len(digits) < 8:
        digits = digits.zfill(8)
    return digits


def read_excel_history_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df = clean_header_logic(df)
    return df.replace({np.nan: ""})


def parse_history_date(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    if isinstance(value, (int, float, np.integer, np.floating)):
        num = float(value)
        if 20000 <= num <= 60000:
            return (datetime(1899, 12, 30) + timedelta(days=int(num))).strftime("%Y-%m-%d")
    text = str(value).strip()
    digits = re.sub(r"[^0-9]", "", text)
    if len(digits) >= 8:
        parsed = pd.to_datetime(digits[:8], format="%Y%m%d", errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%Y-%m-%d")
    parsed = pd.to_datetime(text, errors="coerce")
    return parsed.strftime("%Y-%m-%d") if pd.notna(parsed) else text


def infer_activity_detail(row, source_cols):
    text = " ".join(str(row.get(col, "")) for col in source_cols if col)
    if "개설" in text:
        return "개설"
    if "연계" in text or "ERP" in text.upper():
        return "연계"
    return "운영"


def title_from_activity_detail(value):
    detail = str(value).strip()
    if "연계" in detail:
        return "연계 방문"
    if "개설" in detail:
        return "개설 방문"
    return "운영방문"


def normalize_converted_history_df(df):
    if df is None or df.empty:
        return pd.DataFrame()
    result = df.copy()
    for col in list(result.columns):
        if "업무번호" in str(col):
            result = result.rename(columns={col: "업무번호"})
            break
    result["지사"] = "HANA지사"
    if "_is_manual" not in result.columns:
        result["_is_manual"] = False
    if "활동구분" in result.columns:
        result["활동구분"] = result["활동구분"].astype(str).map(
            lambda v: "방문" if "방문" in v else ("원격" if "원격" in v else ("상담" if v == "상담" else "상담"))
        )
    if "활동상세" in result.columns:
        result["활동상세"] = result["활동상세"].astype(str).map(
            lambda v: "연계" if "연계" in v else ("개설" if "개설" in v else "운영")
        )
        default_titles = result["활동상세"].map(title_from_activity_detail)
        if "제목" in result.columns:
            result["제목"] = result["제목"].where(
                result["제목"].notna() & result["제목"].astype(str).str.strip().ne(""),
                default_titles,
            )
        else:
            result["제목"] = default_titles
    if "활동일" in result.columns and "활동일자" not in result.columns:
        result = result.rename(columns={"활동일": "활동일자"})
    if "활동내용" in result.columns and "활동내역" not in result.columns:
        result = result.rename(columns={"활동내용": "활동내역"})
    return result


def hana_customer_biz_map():
    hana = st.session_state.get("hana_sheet_df")
    if hana is None or hana.empty:
        hana = read_google_csv(st.session_state.get("url_hana", DEFAULT_URL_HANA), header=2)
        st.session_state.hana_sheet_df = hana

    hana = clean_header_logic(hana.copy())
    customer_col = find_col(hana, ["고객번호", "고개번호", "고객NO", "고객 No", "고객"])
    biz_col = find_col(hana, ["사업자번호"])
    if not customer_col or customer_col not in hana.columns or not biz_col or biz_col not in hana.columns:
        return {}, "하나은행 구글 시트에서 고객번호 또는 사업자번호 컬럼을 찾을 수 없습니다."

    mapping_df = hana[[customer_col, biz_col]].copy()
    mapping_df["_customer_key"] = mapping_df[customer_col].apply(normalize_customer_no)
    mapping_df["_biz_value"] = normalize_biz(mapping_df[biz_col])
    mapping_df = mapping_df[(mapping_df["_customer_key"] != "") & (mapping_df["_biz_value"] != "")]
    mapping_df = mapping_df.drop_duplicates("_customer_key")
    return dict(zip(mapping_df["_customer_key"], mapping_df["_biz_value"])), ""


def convert_history_to_sample_df(history_df, user_name):
    history_df = clean_header_logic(history_df.copy()).replace({np.nan: ""})
    customer_col = find_col(history_df, ["고객번호", "고개번호", "고객NO", "고객 No", "고객"])
    staff_col = find_col(history_df, ["담당자"]) or find_col(history_df, ["처리자", "접수자", "등록자", "성명"])
    date_col = find_col(history_df, ["활동일", "처리일자", "접수일자", "일자"])
    company_col = find_col(history_df, ["업체명", "고객명", "상호", "회사명"])
    product_col = find_col(history_df, ["상품", "서비스", "제품"])
    category_col = find_col(history_df, ["활동구분", "상담구분", "접수구분", "접수유형"])
    detail_col = find_col(history_df, ["활동상세", "처리유형", "진행상태", "접수유형"])
    location_col = find_col(history_df, ["방문장소", "주소", "지역"])
    work_no_col = find_col(history_df, ["업무번호", "플로우", "작성번호"])
    title_col = find_col(history_df, ["제목", "접수내용", "문의내용", "진행상태"])
    content_col = find_col(history_df, ["활동내용", "처리내용", "상담내용", "내용"])

    missing = []
    if not customer_col or customer_col not in history_df.columns:
        missing.append("고객번호")
    if not staff_col or staff_col not in history_df.columns:
        missing.append("담당자")
    if missing:
        return pd.DataFrame(), {"error": f"이력 파일에서 {', '.join(missing)} 컬럼을 찾을 수 없습니다."}

    filtered = history_df[history_df[staff_col].astype(str).str.strip() == str(user_name).strip()].copy()
    if filtered.empty:
        return pd.DataFrame(), {"error": f"로그인 사용자({user_name})와 담당자가 일치하는 이력이 없습니다."}

    biz_map, map_error = hana_customer_biz_map()
    if map_error:
        return pd.DataFrame(), {"error": map_error}

    rows = []
    unmatched = 0
    for _, row in filtered.iterrows():
        customer_key = normalize_customer_no(row.get(customer_col, ""))
        biz_no = biz_map.get(customer_key, "")
        if not biz_no:
            unmatched += 1

        activity_detail = infer_activity_detail(row, [detail_col, title_col, content_col, category_col])
        activity_category = str(row.get(category_col, "")).strip() if category_col else ""
        activity_text = " ".join(str(row.get(col, "")) for col in [category_col, title_col, content_col, detail_col] if col)
        if "방문" in activity_text:
            activity_category = "방문"
        elif "원격" in activity_category:
            activity_category = "원격"
        elif activity_category not in ["방문", "상담", "원격"]:
            activity_category = "상담"

        rows.append({
            "지사": "HANA지사",
            "상품": row.get(product_col, "통합CMS") if product_col else "통합CMS",
            "업체명": row.get(company_col, "") if company_col else "",
            "사업자번호": biz_no,
            "등록자": user_name,
            "활동일자": parse_history_date(row.get(date_col, "")) if date_col else "",
            "방문장소 (시, 군, 구까지)": row.get(location_col, "") if location_col else "",
            "활동구분": activity_category,
            "활동상세": activity_detail,
            "업무번호": row.get(work_no_col, "") if work_no_col else "",
            "제목": title_from_activity_detail(activity_detail),
            "활동내역": row.get(content_col, "") if content_col else "",
        })

    return normalize_converted_history_df(pd.DataFrame(rows)), {"total": len(rows), "unmatched": unmatched}


def sample_format_excel_bytes(df):
    from openpyxl import load_workbook

    output = BytesIO()
    if os.path.exists(EXCEL_SAMPLE_FILE):
        wb = load_workbook(EXCEL_SAMPLE_FILE)
        ws = wb.worksheets[0]
        if ws.max_row > 2:
            ws.delete_rows(3, ws.max_row - 2)
        def header_key(value):
            return re.sub(r"\s+", "", str(value or ""))

        header_map = {header_key(ws.cell(row=2, column=col).value): col for col in range(1, ws.max_column + 1)}
        aliases = {
            "활동내역": "활동내용",
            "활동일자": "활동일",
            "업무번호": "업무번호\n(플로우에 식권, 비즈플레이, 플로우 작성번호)",
        }
        for row_idx, (_, row) in enumerate(df.iterrows(), start=3):
            for header, value in row.items():
                col_idx = header_map.get(header_key(header)) or header_map.get(header_key(aliases.get(header, "")))
                if not col_idx:
                    continue
                ws.cell(row=row_idx, column=col_idx, value=value)
        wb.save(output)
    else:
        output.write(dataframe_to_excel_bytes({"Sheet0": df}))
    output.seek(0)
    return output.getvalue()


def criteria_df():
    data = [
        ["교차판매", "타겟고객선별", 5, 100],
        ["교차판매", "메일발송", 2, 100],
        ["교차판매", "제안서전달", 5, 100],
        ["교차판매", "방문설명회&견적발송", 30, "-"],
        ["교차판매", "계약진행 시", 50, "-"],
        ["교차판매", "유선 (해피콜)", 5, 100],
        ["교차판매", "활성화 (조회업무)", 10, 100],
        ["교차판매", "이체, 집금 활성화", 30, 150],
        ["교차판매", "계열사 추가도입", 30, "-"],
        ["교차판매", "신규연계도입", 60, "-"],
        ["추가활동", "문서 작성 (본사)", 100, 100],
        ["추가활동", "문서 작성 (가이드)", 50, "-"],
        ["추가활동", "문서 작성 (기타)", 20, "-"],
        ["추가활동", "VOC (아이디어)", 10, 50],
        ["추가활동", "운영활동(원격)", 10, 200],
    ]
    return pd.DataFrame(data, columns=["활동구분", "구분", "단위 점수", "월 최대점수"])


def render_manual_perf_input_table(base):
    """표는 보고서 표와 동일하게 렌더링하고, 실적 입력은 표 아래에서 처리한다."""
    result_df = base.copy()
    result_df["입력(건)"] = pd.to_numeric(result_df["입력(건)"], errors="coerce").fillna(0).astype(int)

    style_report_logic(result_df.drop(columns=["입력(건)"], errors="ignore"))

    st.markdown("#### 실적 입력")
    st.info("월 최대점수가 있는 항목은 `월 최대점수 ÷ 단위 점수`까지만 입력할 수 있습니다. 예: 타겟고객선별은 100 ÷ 5 = 최대 20건입니다.")
    st.markdown(
        """<style>
        .manual-input-row {
            color:#4A5568;
            font-size:13px;
            font-weight:700;
            padding-top:9px;
            white-space:nowrap;
        }
        .manual-input-help {
            color:#718096;
            font-size:12px;
            margin-top:-4px;
            margin-bottom:4px;
        }
        body:has(#pms-d:checked) .manual-input-row,
        body:has(#pms-d:checked) .manual-input-help {
            color:#ffffff !important;
        }
        </style>""",
        unsafe_allow_html=True,
    )

    results = {}
    for i in range(0, len(result_df), 3):
        cols = st.columns(3)
        for col, (_, row) in zip(cols, result_df.iloc[i:i + 3].iterrows()):
            item = str(row["구분"])
            unit_score = int(row["단위 점수"])
            monthly_limit = row["월 최대점수"]
            max_count = None
            if str(monthly_limit).strip() != "-":
                try:
                    max_count = int(float(monthly_limit) // unit_score)
                except Exception:
                    max_count = None

            value = int(row["입력(건)"])
            if max_count is not None:
                value = min(value, max_count)

            with col:
                st.markdown(f"<div class='manual-input-row'>{html.escape(item)}</div>", unsafe_allow_html=True)
                if max_count is not None:
                    st.markdown(f"<div class='manual-input-help'>최대 {max_count}건 입력 가능</div>", unsafe_allow_html=True)
                else:
                    st.markdown("<div class='manual-input-help'>월 최대점수 제한 없음</div>", unsafe_allow_html=True)
                widget_key = f"perf_{item}"
                kwargs = {
                    "label": "입력(건)",
                    "min_value": 0,
                    "value": value,
                    "key": widget_key,
                    "label_visibility": "collapsed",
                    "step": 1,
                }
                if max_count is not None:
                    kwargs["max_value"] = max_count
                st.number_input(**kwargs)
                results[item] = int(st.session_state.get(widget_key, value) or 0)

    result_df["입력(건)"] = result_df["구분"].map(results).fillna(0).astype(int)
    return result_df


def manual_points_for_user(name):
    override_db = st.session_state.get("manual_perf_preview_override", {})
    saved = override_db.get(name)
    if saved is None:
        perf_db = load_db(PERF_FILE, {})
        saved = perf_db.get(name, {})

    total = 0
    criteria = criteria_df()
    criteria["_key"] = criteria["구분"].astype(str).str.strip()
    point_map = criteria.set_index("_key")["단위 점수"].to_dict()
    limit_map = criteria.set_index("_key")["월 최대점수"].to_dict()
    for item, count in saved.items():
        try:
            key = str(item).strip()
            score = int(point_map.get(key, 0)) * int(count)
            limit_value = limit_map.get(key, "-")
            if str(limit_value).strip() != "-":
                score = min(score, int(float(limit_value)))
            total += score
        except Exception:
            pass
    return total


def calculate_manual_perf_total(edited_df):
    if edited_df is None or edited_df.empty:
        return 0
    total = 0
    for _, row in edited_df.iterrows():
        try:
            score = int(float(row.get("단위 점수", 0))) * int(float(row.get("입력(건)", 0)))
            limit_value = row.get("월 최대점수", "-")
            if str(limit_value).strip() != "-":
                score = min(score, int(float(limit_value)))
            total += score
        except Exception:
            pass
    return int(total)


def apply_rs_allowance_formula(perf_df, user_db, return_debug=False):
    if perf_df is None or perf_df.empty or "합계포인트" not in perf_df.columns:
        return (perf_df, {}) if return_debug else perf_df

    def num(value):
        try:
            if pd.isna(value):
                return 0
            return int(float(str(value).replace(",", "").strip() or 0))
        except Exception:
            return 0

    def payable_points(row):
        open_link_points = num(row.get("개설포인트", 0)) + num(row.get("연계포인트", 0))
        operation_points = num(row.get("운영포인트 (실제 활동)", 0))
        operation_points += num(row.get("운영포인트 (추가 활동)", 0))
        operation_points += num(row.get("운영포인트(추가 활동)", 0))
        return min(open_link_points, 1000) + min(operation_points, 1800)

    result = perf_df.copy()
    result["지급포인트"] = 0
    result["지급예상금액"] = 0

    work_df = result[result["담당자"].astype(str) != "합계"].copy() if "담당자" in result.columns else result.copy()
    if work_df.empty:
        return (result, {}) if return_debug else result

    # 직원 정보 매핑 (이름 -> 정보)
    name_to_info = {}
    for uid, info in user_db.items():
        if isinstance(info, dict) and info.get("name"):
            name_to_info[info["name"]] = {
                "rank": info.get("rank", "직원"),
                "outsource": info.get("outsource", "아니오"),
                "staff_type": info.get("staff_type", "정규직"),
                "period": info.get("outsource_period", "해당없음"),
                "dept_type": info.get("dept_type", "사업부")
            }

    # C&S 부서만 대상
    cs_staff = []
    outsource_staff = []
    regular_staff = []

    for idx, row in work_df.iterrows():
        staff_name = str(row.get("담당자", ""))
        if staff_name in name_to_info:
            staff_info = name_to_info[staff_name]
            if staff_info["dept_type"] == "C&S":
                cs_staff.append((idx, row, staff_info))
                if staff_info["outsource"] == "예" or staff_info["staff_type"] == "외주":
                    outsource_staff.append((idx, row, staff_info))
                else:
                    regular_staff.append((idx, row, staff_info))

    if not cs_staff:
        return (result, {}) if return_debug else result

    # BU 평균 계산 (C&S 전체 지급 산정 포인트 / C&S 인원)
    total_bu_points = sum(payable_points(row) for _, row, _ in cs_staff)
    bu_count = len(cs_staff)
    bu_average = total_bu_points / bu_count if bu_count > 0 else 0

    # 외주직원 가감포인트 계산
    outsource_total_points = 0
    for idx, row, staff_info in outsource_staff:
        staff_points = payable_points(row)

        # 근무기간 계수
        period = staff_info["period"]
        if period == "1년 미만":
            period_factor = 0.8
        elif period == "1년 이상":
            period_factor = 0.9
        else:  # 2년 이상
            period_factor = 1.0

        # 외주직원 가감포인트 = 외주직원포인트 - (BU평균 × 근무기간)
        adjustment = staff_points - (bu_average * period_factor)
        outsource_total_points += adjustment

    # 일반직원 수
    regular_count = len(regular_staff)
    outsource_point_per_regular = outsource_total_points / regular_count if regular_count > 0 else 0

    # 팀장수당
    team_leader_bonus = 267

    # 디버깅 정보
    debug_info = {
        "BU합산": total_bu_points,
        "BU인원": bu_count,
        "BU평균": bu_average,
        "외주직원수": len(outsource_staff),
        "일반직원수": regular_count,
        "외주가감총합": outsource_total_points,
        "일반직원1인당분배": outsource_point_per_regular,
        "팀장수당": team_leader_bonus
    }

    # 최종 지급액 계산
    for idx, row, staff_info in cs_staff:
        total_points = payable_points(row)

        # 팀장 여부 확인
        is_team_leader = staff_info["rank"] == "팀장"

        # 외주직원 여부 확인
        is_outsource = staff_info["outsource"] == "예" or staff_info["staff_type"] == "외주"

        if is_outsource:
            # 외주직원: 외주실적 가감포인트 분배 대상에서는 제외
            final_pay_point = total_points - 1000
        else:
            # 팀장/직원 공식: (개설+연계+운영(+팀장수당)-1000)+(외주실적가감포인트/일반직원수)
            final_pay_point = total_points - 1000 + outsource_point_per_regular
            if is_team_leader:
                final_pay_point += team_leader_bonus

        final_pay_point = int(max(0, final_pay_point))
        result.at[idx, "지급포인트"] = final_pay_point
        result.at[idx, "지급예상금액"] = int(final_pay_point * 500)

    if return_debug:
        return result, debug_info
    return result


def may_2026_business_dates():
    holidays = {"2026-05-01", "2026-05-05", "2026-05-24", "2026-05-25"}
    days = pd.date_range("2026-05-01", "2026-05-31", freq="D")
    return [d.strftime("%Y-%m-%d") for d in days if d.weekday() < 5 and d.strftime("%Y-%m-%d") not in holidays]


def prepare_random_history_source_df(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    result = df.copy()
    result.columns = [str(col).strip() for col in result.columns]
    result = result.loc[:, ~pd.Index(result.columns).duplicated()]
    for col in result.columns:
        if not isinstance(result[col], pd.Series):
            result[col] = pd.Series(result[col], index=result.index)
    return clean_header_logic(result)


def prepare_display_dataframe(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    result = df.copy()
    result.columns = [str(col).strip() for col in result.columns]
    result = result.loc[:, ~pd.Index(result.columns).duplicated()]
    result = result[
        [
            col for col in result.columns
            if not str(col).strip().lower().startswith("unnamed")
            and not str(col).strip().startswith("_")
        ]
    ]
    if result.empty:
        return result
    return result.loc[
        :,
        result.apply(lambda col: col.astype(str).str.strip().replace("nan", "").ne("").any(), axis=0)
    ]


def limit_history_to_total_point_cap(history_df, perf_df):
    if history_df is None or not isinstance(history_df, pd.DataFrame) or history_df.empty:
        return history_df
    if perf_df is None or not isinstance(perf_df, pd.DataFrame) or perf_df.empty:
        return history_df

    df = clean_header_logic(history_df.copy())
    user_col = find_col(df, ["등록자", "담당자", "성명"])
    detail_col = find_col(df, ["활동상세", "활동내용"])
    category_col = find_col(df, ["활동구분", "접수유형"])
    if not user_col or user_col not in df.columns:
        return history_df

    base_points_by_user = {}
    for _, row in perf_df.iterrows():
        name = str(row.get("담당자", "")).strip()
        if not name or name == "합계":
            continue
        open_points = int(float(row.get("개설포인트", 0) or 0))
        link_points = int(float(row.get("연계포인트", 0) or 0))
        base_points_by_user[name] = max(0, 2800 - open_points - link_points)

    used_points_by_user = {name: 0 for name in base_points_by_user}
    keep_indices = []
    for idx, row in df.iterrows():
        name = str(row.get(user_col, "")).strip()
        if name not in base_points_by_user:
            keep_indices.append(idx)
            continue

        detail = str(row.get(detail_col, "")).strip() if detail_col else ""
        category = str(row.get(category_col, "")).strip() if category_col else ""
        is_operation = "운영" in detail or "방문" in category or "원격" in category
        if not is_operation:
            keep_indices.append(idx)
            continue

        point = 10 if "원격" in category else 30
        if used_points_by_user[name] + point > base_points_by_user[name]:
            continue
        used_points_by_user[name] += point
        keep_indices.append(idx)

    return df.loc[keep_indices].reset_index(drop=True)


def build_random_admin_extra_history(source_df, existing_df, target_count):
    if source_df is None or not isinstance(source_df, pd.DataFrame) or source_df.empty:
        return pd.DataFrame(), "하나지사 활동이력 데이터가 없습니다."

    source = prepare_random_history_source_df(source_df)
    existing = prepare_random_history_source_df(existing_df) if isinstance(existing_df, pd.DataFrame) else pd.DataFrame()

    detail_col = find_col(source, ["활동상세", "활동내용"])
    if not detail_col or detail_col not in source.columns:
        return pd.DataFrame(), "하나지사 활동이력에서 활동상세 컬럼을 찾을 수 없습니다."

    source = source[source[detail_col].astype(str).str.contains("운영", na=False)].copy()
    if source.empty:
        return pd.DataFrame(), "활동상세가 운영인 데이터가 없습니다."

    company_col = find_col(source, ["업체명", "상호", "고객명"])
    biz_col = find_col(source, ["사업자번호"])
    user_col = find_col(source, ["등록자", "담당자", "성명"])
    place_col = find_col(source, ["방문장소", "주소", "지역"])
    category_col = find_col(source, ["활동구분", "접수유형"])
    work_col = find_col(source, ["업무번호"])
    note_col = find_col(source, ["활동내역", "활동내용", "비고", "제목"])

    if not user_col or user_col not in source.columns:
        return pd.DataFrame(), "하나지사 활동이력에서 담당자 컬럼을 찾을 수 없습니다."

    business_dates = may_2026_business_dates()
    if not business_dates:
        return pd.DataFrame(), "2026년 5월 영업일을 만들 수 없습니다."

    existing_visit_counts = {}
    existing_keys = set()
    if not existing.empty:
        e_user_col = find_col(existing, ["등록자", "담당자", "성명"])
        e_biz_col = find_col(existing, ["사업자번호"])
        e_date_col = find_col(existing, ["활동일자", "활동일", "일자"])
        e_detail_col = find_col(existing, ["활동상세", "활동내용"])
        e_category_col = find_col(existing, ["활동구분", "접수유형"])
        if e_user_col and e_date_col and e_user_col in existing.columns and e_date_col in existing.columns:
            for _, row in existing.iterrows():
                user = str(row.get(e_user_col, "")).strip()
                date = str(row.get(e_date_col, "")).strip()
                category = str(row.get(e_category_col, "")) if e_category_col else ""
                if user and date and "방문" in category:
                    existing_visit_counts[(user, date)] = existing_visit_counts.get((user, date), 0) + 1
                biz = normalize_biz(row.get(e_biz_col, "")) if e_biz_col else ""
                detail = str(row.get(e_detail_col, "")).strip() if e_detail_col else ""
                if biz and user and date and detail:
                    existing_keys.add((biz, user, date, detail))

    rng_state = int((datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m%d%H%M%S")) % (2**32 - 1)
    source = source.sample(frac=1, random_state=rng_state).reset_index(drop=True)
    target_count = max(1, int(target_count or 1))
    rows = []

    for _, row in source.iterrows():
        if len(rows) >= target_count:
            break

        user = str(row.get(user_col, "")).strip()
        if not user:
            continue

        raw_category = str(row.get(category_col, "")).strip() if category_col else ""
        activity_category = "원격" if "원격" in raw_category else "방문"
        biz = normalize_biz(row.get(biz_col, "")) if biz_col else ""

        assigned_date = ""
        for date in np.random.default_rng().permutation(business_dates):
            date = str(date)
            if activity_category == "방문" and existing_visit_counts.get((user, date), 0) >= 5:
                continue
            if biz and (biz, user, date, "운영") in existing_keys:
                continue
            assigned_date = date
            break

        if not assigned_date:
            continue

        if activity_category == "방문":
            existing_visit_counts[(user, assigned_date)] = existing_visit_counts.get((user, assigned_date), 0) + 1
        if biz:
            existing_keys.add((biz, user, assigned_date, "운영"))

        rows.append({
            "업체명": row.get(company_col, "") if company_col else "",
            "사업자번호": row.get(biz_col, "") if biz_col else "",
            "등록자": user,
            "활동일": assigned_date,
            "방문장소 (시, 군, 구까지)": row.get(place_col, "") if place_col else "",
            "활동구분": activity_category,
            "활동상세": "운영",
            "업무번호": row.get(work_col, "") if work_col else "",
            "제목": "운영방문" if activity_category == "방문" else "운영원격",
            "활동내역": row.get(note_col, "") if note_col else "",
            "본사 개설완료일자": assigned_date,
            "본사 ERP연계일자": assigned_date,
            "_is_manual": False,
        })

    if not rows:
        return pd.DataFrame(), "중복/초과방문 조건을 피해서 추가할 수 있는 데이터가 없습니다."

    return pd.DataFrame(rows), ""


def save_manual_perf_override_for_current_user():
    name = st.session_state.get("user_name")
    override = st.session_state.get("manual_perf_preview_override", {}).get(name)
    if name and override is not None:
        db = load_db(PERF_FILE, {})
        db[name] = override
        save_db(PERF_FILE, db)


def has_performance_required_columns(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return False
    return all(
        find_col(df, keys) in df.columns
        for keys in [["등록자", "담당자", "성명"], ["활동상세", "활동내용"], ["활동일자", "활동일", "일자"]]
    )


def prepare_history_analysis_df(raw_df):
    if raw_df is None or not isinstance(raw_df, pd.DataFrame) or raw_df.empty:
        return pd.DataFrame()

    df = normalize_converted_history_df(clean_header_logic(raw_df.copy()))
    if has_performance_required_columns(df):
        return df

    try:
        converted_df, _ = convert_history_to_sample_df(df, st.session_state.get("user_name", ""))
        converted_df = normalize_converted_history_df(converted_df)
        if has_performance_required_columns(converted_df):
            return converted_df
    except Exception:
        pass

    return df


def current_history_analysis_df():
    preview_df = st.session_state.get("history_convert_preview_data")
    if isinstance(preview_df, pd.DataFrame):
        preview_df = prepare_history_analysis_df(preview_df)
        if has_performance_required_columns(preview_df):
            return preview_df
    excel_df = st.session_state.get("user_excel_data")
    if isinstance(excel_df, pd.DataFrame):
        excel_df = prepare_history_analysis_df(excel_df)
        if has_performance_required_columns(excel_df):
            return excel_df
    return pd.DataFrame()


def apply_rank_from_user_db(df):
    if "담당자" not in df.columns or "직급" not in df.columns:
        return df
    rank_map = {
        info.get("name", ""): info.get("rank", "직원")
        for uid, info in st.session_state.get("user_db", {}).items()
        if uid != "1" and info.get("name")
    }
    if not rank_map:
        return df
    df = df.copy()
    df["직급"] = df["담당자"].astype(str).str.strip().map(lambda n: rank_map.get(n, "직원"))
    return df


_RANK_ORDER = {"부서장": 0, "팀장": 1, "과장": 2, "대리": 3, "주임": 4, "직원": 5}


def sort_by_rank_name(df):
    if df.empty:
        return df
    name_col = next((c for c in ["담당자", "성명"] if c in df.columns), None)
    if "직급" not in df.columns or name_col is None:
        return df
    df = df.copy()
    df["_rank_order"] = df["직급"].map(_RANK_ORDER).fillna(9)
    df = df.sort_values(["_rank_order", name_col]).drop(columns=["_rank_order"]).reset_index(drop=True)
    return df


def process_performance_analysis(curr_df_raw, prev_df_raw=None):
    try:
        if curr_df_raw is None:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        df = normalize_converted_history_df(clean_header_logic(curr_df_raw))
        if df.empty and len(df.columns) == 0:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        u_col = find_col(df, ["등록자", "담당자", "성명"])
        d_col = find_col(df, ["활동상세", "활동내용"])
        date_col = find_col(df, ["활동일자", "활동일", "일자"])
        biz_col = find_col(df, ["사업자번호"], "사업자번호")
        comp_col = find_col(df, ["업체명", "상호"], "업체명")

        if not all([u_col, d_col, date_col]):
            prepared_df = prepare_history_analysis_df(curr_df_raw)
            if not prepared_df.empty:
                df = prepared_df
                u_col = find_col(df, ["등록자", "담당자", "성명"])
                d_col = find_col(df, ["활동상세", "활동내용"])
                date_col = find_col(df, ["활동일자", "활동일", "일자"])
                biz_col = find_col(df, ["사업자번호"], "사업자번호")
                comp_col = find_col(df, ["업체명", "상호"], "업체명")

        required_cols = [("등록자", u_col), ("활동상세", d_col), ("활동일", date_col)]
        missing = [label for label, col in required_cols if not col or col not in df.columns]
        if missing:
            return f"필수 컬럼이 없습니다: {', '.join(missing)}", None, None

        _udb = st.session_state.get("user_db", {})
        member_db = {
            info.get("name", ""): info.get("rank", "직원")
            for uid, info in _udb.items()
            if uid != "1" and info.get("name")
        }

        df_clean = df.dropna(subset=[u_col, date_col]).copy()
        _staff = get_staff_names()
        if _staff and u_col in df_clean.columns:
            df_clean = df_clean[df_clean[u_col].isin(_staff)]
        df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors="coerce")
        df_clean = df_clean.dropna(subset=[date_col])
        df_clean[date_col] = df_clean[date_col].dt.strftime("%Y-%m-%d")

        daily_counts = df_clean.groupby([u_col, date_col]).size().reset_index(name="일방문횟수")
        monthly_counts = df_clean.groupby(u_col).size().reset_index(name="월총방문")

        if comp_col in df_clean.columns and biz_col in df_clean.columns:
            error_raw = pd.merge(
                daily_counts,
                df_clean[[u_col, date_col, comp_col, biz_col]].drop_duplicates(),
                on=[u_col, date_col],
                how="left",
            )
            error_df = pd.merge(error_raw, monthly_counts, on=u_col, how="left")
            error_df = error_df[[comp_col, biz_col, u_col, date_col, "일방문횟수", "월총방문"]].drop_duplicates()
            error_df.columns = ["업체명", "사업자번호", "담당자", "초과일자", "일방문", "월총방문"]
        else:
            error_df = pd.DataFrame(columns=["업체명", "사업자번호", "담당자", "초과일자", "일방문", "월총방문"])

        # 은행 이력(실제 활동)과 이력 추가(추가 활동) 분리
        _flag = "_is_manual"
        if _flag in df.columns:
            df_real = df[~df[_flag].fillna(False).astype(bool)].copy()
            df_added = df[df[_flag].fillna(False).astype(bool)].copy()
        else:
            df_real = df
            df_added = pd.DataFrame()

        # 중복 이력: 변환파일 미리보기 기준 사업자번호+등록자+활동일자+활동상세가 모두 같은 행
        if biz_col in df_clean.columns and u_col in df_clean.columns and date_col in df_clean.columns and d_col in df_clean.columns:
            _dup_df = df_clean.copy()
            _dup_df["_dup_biz"] = normalize_biz(_dup_df[biz_col])
            _dup_df["_dup_user"] = _dup_df[u_col].astype(str).str.strip()
            _dup_df["_dup_date"] = pd.to_datetime(_dup_df[date_col], errors="coerce").dt.strftime("%Y-%m-%d")
            _dup_df["_dup_detail"] = _dup_df[d_col].astype(str).str.strip()
            _dup_keys = ["_dup_biz", "_dup_user", "_dup_date", "_dup_detail"]
            _dup_df = _dup_df[(_dup_df[_dup_keys] != "").all(axis=1)]
            dup_biz_df = _dup_df[_dup_df.duplicated(subset=_dup_keys, keep=False)].drop(columns=_dup_keys, errors="ignore")
            dup_biz_df = dup_biz_df.sort_values(by=[date_col, biz_col, u_col, d_col])
        else:
            dup_biz_df = pd.DataFrame()

        # 이력 추가 행의 포인트를 담당자별로 집계
        added_pts_by_name = {}
        if not df_added.empty and u_col in df_added.columns and d_col in df_added.columns:
            for _n, _g in df_added.dropna(subset=[u_col, d_col]).groupby(u_col):
                _det = _g[d_col].astype(str)
                _pts = _det.str.contains("운영|방문|점검").sum() * 30
                _pts += _det.str.contains("개설").sum() * 90
                _pts += _det.str.contains("연계").sum() * 120
                added_pts_by_name[_n] = int(_pts)

        summary = (
            df_real.dropna(subset=[u_col, d_col])
            .groupby(u_col)
            .agg(
                o=(d_col, lambda x: x.astype(str).str.contains("개설").sum()),
                l=(d_col, lambda x: x.astype(str).str.contains("연계").sum()),
                v=(d_col, lambda x: x.astype(str).str.contains("운영|방문|점검").sum()),
            )
            .reset_index()
        )

        if summary.empty:
            return pd.DataFrame(), error_df, dup_biz_df

        name_to_info = {
            info.get("name"): {
                "staff_type": info.get("staff_type", "정규직"),
                "period": info.get("outsource_period", "해당없음"),
            }
            for uid, info in st.session_state.user_db.items()
            if uid != "1"
        }

        member_stats = {}
        total_points = 0

        for _, row in summary.iterrows():
            name = row[u_col]
            o_p = int(row["o"]) * 90
            l_p = int(row["l"]) * 120
            v_actual_p = int(row["v"]) * 30
            manual_p = manual_points_for_user(name) + added_pts_by_name.get(name, 0)
            p_sum = min(2800, o_p + l_p + v_actual_p + manual_p)

            member_stats[name] = {
                "o_p": o_p,
                "l_p": l_p,
                "v_actual_p": v_actual_p,
                "manual_p": manual_p,
                "p_sum": p_sum,
            }
            total_points += p_sum

        bu_avg = total_points / len(summary) if len(summary) > 0 else 0
        outsource_adj_pool = 0
        regular_count = 0

        for name, stats in member_stats.items():
            info = name_to_info.get(name, {})
            if info.get("staff_type", "정규직") == "외주":
                tenure = {"1년 미만": 0.8, "1년 이상": 0.9, "2년 이상": 1.0}.get(info.get("period"), 1.0)
                outsource_adj_pool += stats["p_sum"] - (bu_avg * tenure)
            else:
                regular_count += 1

        adj_per_regular = outsource_adj_pool / regular_count if regular_count else 0

        rows = []
        for _, row in summary.iterrows():
            name = row[u_col]
            stats = member_stats[name]
            rank = member_db.get(name, "직원")
            is_outsource = name_to_info.get(name, {}).get("staff_type", "정규직") == "외주"

            rows.append(
                {
                    "담당자": name,
                    "직급": rank,
                    "개설건수": int(row["o"]),
                    "개설포인트": stats["o_p"],
                    "연계건수": int(row["l"]),
                    "연계포인트": stats["l_p"],
                    "운영건수 (실제 활동)": int(row["v"]),
                    "운영포인트 (실제 활동)": stats["v_actual_p"],
                    "운영건수 (추가 활동)": round(manual_points_for_user(name) / 30),
                    "운영포인트(추가 활동)": manual_points_for_user(name),
                    "합계포인트": stats["p_sum"],
                    "지급포인트": 0,
                    "지급예상금액": 0,
                    "전월대비": 0,
                }
            )

        res_df = pd.DataFrame(rows)
        res_df = apply_rank_from_user_db(res_df)
        res_df = apply_rs_allowance_formula(res_df, st.session_state.user_db)
        res_df = hide_department_heads(res_df)
        res_df = sort_by_rank_name(res_df)

        if prev_df_raw is not None:
            try:
                prev_res_df, _, _ = process_performance_analysis(prev_df_raw)
                if isinstance(prev_res_df, pd.DataFrame) and not prev_res_df.empty:
                    p_map = prev_res_df.set_index("담당자")["지급예상금액"].to_dict()
                    res_df["전월대비"] = res_df.apply(
                        lambda r: int(r["지급예상금액"] - p_map.get(r["담당자"], 0)),
                        axis=1,
                    )
            except Exception:
                pass

        return res_df, error_df, dup_biz_df

    except Exception as e:
        return f"ERR: {str(e)}", None, None


def render_plain_html_table(
    df, max_rows=500, center_align=True, merge_cols=None, stretch=True, max_width=None, border=True
):
    """AG Grid 없이 순수 HTML 테이블로 렌더링 — 다크모드 완전 호환."""
    if df is None or df.empty:
        st.info("표시할 데이터가 없습니다.")
        return
    df = df.head(max_rows).reset_index(drop=True)
    merge_cols = [c for c in (merge_cols or []) if c in df.columns]

    rowspan_map = {}
    for col in merge_cols:
        rowspan_map[col] = {}
        last_value = None
        start_idx = None
        span = 0
        for i, value in enumerate(df[col].tolist()):
            text = "" if pd.isna(value) else str(value).strip()
            if text and start_idx is not None and text == last_value:
                span += 1
                rowspan_map[col][i] = 0
            elif text:
                if start_idx is not None:
                    rowspan_map[col][start_idx] = span
                last_value = text
                start_idx = i
                span = 1
            elif start_idx is not None and last_value:
                span += 1
                rowspan_map[col][i] = 0
            else:
                rowspan_map[col][i] = 1
        if start_idx is not None:
            rowspan_map[col][start_idx] = span

    th = "background:#EDF2F7;color:#4A5568;font-weight:700;font-size:12px;padding:6px 10px;white-space:nowrap;border-bottom:2px solid #E2E8F0;text-align:center;"
    headers = "".join(f"<th style='{th}'>{html.escape(str(c))}</th>" for c in df.columns)
    body = ""
    for i, row in df.iterrows():
        bg = "#FFFFFF" if i % 2 == 0 else "#F7FAFC"
        tds = ""
        for col in df.columns:
            if col in rowspan_map and rowspan_map[col].get(i, 1) == 0:
                continue
            val = "" if pd.isna(row[col]) else html.escape(str(row[col]))
            align = "center" if center_align else "left"
            td_align = f"text-align:{align};"
            rowspan = ""
            if col in rowspan_map and rowspan_map[col].get(i, 1) > 1:
                rowspan = f" rowspan='{rowspan_map[col][i]}'"
            tds += f"<td{rowspan} style='background:{bg};padding:5px 10px;border-bottom:1px solid #EDF2F7;font-size:12px;color:#2D3748;white-space:nowrap;vertical-align:middle;{td_align}'>{val}</td>"
        body += f"<tr>{tds}</tr>"
    wrapper_display = "display:block;" if stretch else "display:inline-block;max-width:100%;"
    if max_width:
        wrapper_display += f"max-width:{max_width};"
    table_width = "100%" if stretch else "auto"
    wrapper_border = "border:1px solid #E2E8F0;box-shadow:0 2px 6px rgba(0,0,0,0.05);" if border else "border:none;box-shadow:none;"
    st.markdown(
        f"""<div class="pms-report-table" style="overflow-x:auto;{wrapper_border}border-radius:8px;margin-bottom:1rem;{wrapper_display}">
        <table style="width:{table_width};border-collapse:collapse;">
            <thead><tr>{headers}</tr></thead>
            <tbody>{body}</tbody>
        </table></div>""",
        unsafe_allow_html=True,
    )


def style_report_logic(df, compact=False, align_overrides=None, default_align=None):
    if df is None or df.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    df = strip_activity_time_columns(df)
    df = hide_department_heads(df)
    if df.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    # 지사, 상품 컬럼 제거
    df = df.drop(columns=["지사", "상품"], errors="ignore")

    diff_cols = [c for c in df.columns if "전월대비" in str(c) or "증감" in str(c) or "대비" in str(c)]
    exclude_from_num = ["담당자", "직급", "전송시각", "등록월", "항목", "활동구분", "구분", "일치여부"] + [c for c in df.columns if "사업자번호" in str(c)]
    num_cols = [c for c in df.columns if c not in exclude_from_num + diff_cols]

    def fmt_diff(x):
        try:
            v = int(float(x))
            if v > 0:
                return f"<span style='color:#E53E3E;font-weight:800;'>▲ {v:,}</span>"
            if v < 0:
                return f"<span style='color:#3182CE;font-weight:800;'>▼ {abs(v):,}</span>"
            return f"{v:,}"
        except Exception:
            return ""

    def fmt_num(x):
        try:
            if pd.isna(x):
                return ""
            if isinstance(x, (int, float, np.integer, np.floating)):
                return f"{int(x):,}"
            return html.escape(str(x))
        except Exception:
            return html.escape(str(x))

    def fmt_match(x):
        text = "" if pd.isna(x) else str(x)
        if text == "불일치":
            return "<span style='color:#E53E3E;font-weight:900;'>불일치</span>"
        if text == "일치":
            return "<span style='color:#2F855A;font-weight:900;'>일치</span>"
        return html.escape(text)

    # 전월대비 컬럼명 치환
    user_sel = st.session_state.get("user_prev_month_sel") or st.session_state.get("adm_prev_month")
    if user_sel and user_sel != "선택안함":
        prev_m = str(int(user_sel.split("-")[1])) + "월"
        비교월_label = f"{prev_m} 대비"
    else:
        비교월_label = "전월 대비"

    def format_col_name(col):
        if "전월대비" in str(col):
            return str(col).replace("전월대비", 비교월_label)
        return str(col)

    th_pad = "5px 6px" if compact else "9px 12px"
    th_font = "12px" if compact else "13px"
    td_pad = "5px 6px" if compact else "8px 12px"
    td_font = "12px" if compact else "13px"
    th = f"background:#EDF2F7;color:#4A5568;font-weight:800;font-size:{th_font};padding:{th_pad};text-align:center;border-bottom:2px solid #E2E8F0;white-space:nowrap;"
    headers = "".join(f"<th style='{th}'>{html.escape(format_col_name(c))}</th>" for c in df.columns)
    align_overrides = align_overrides or {}

    body = ""
    for i, row in df.reset_index(drop=True).iterrows():
        tds = ""
        for col in df.columns:
            align = default_align or ("right" if col in num_cols or col in diff_cols else "center")
            if str(col).strip() == "활동내역":
                align = "left"
            align = align_overrides.get(col, align)
            bg = "#FFFFFF" if i % 2 == 0 else "#F7FAFC"

            if col == "일치여부":
                value = fmt_match(row[col])
            elif col in diff_cols:
                value = fmt_diff(row[col])
            elif col in num_cols:
                value = fmt_num(row[col])
            else:
                value = "" if pd.isna(row[col]) else html.escape(str(row[col]))

            _ws = "white-space:nowrap;"
            tds += (
                f"<td style='background:{bg};padding:{td_pad};border-bottom:1px solid #EDF2F7;"
                f"font-size:{td_font};color:#2D3748;text-align:{align};{_ws}'>{value}</td>"
            )
        body += f"<tr>{tds}</tr>"

    _ov = "visible" if compact else "auto"
    _tbl_style = "width:100%;border-collapse:collapse;table-layout:fixed;" if compact else "width:100%;border-collapse:collapse;"
    st.markdown(
        f"""
        <div class="pms-report-table" style="overflow-x:{_ov};border:1px solid #E2E8F0;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);margin-bottom:1rem;">
            <table style="{_tbl_style}">
                <thead><tr>{headers}</tr></thead>
                <tbody>{body}</tbody>
            </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_auth_page():
    inject_theme_toggle()
    st.markdown(
        """
        <style>
        .stAppHeader, header, [data-testid="stHeader"], .stDecoration {
            display: none !important;
            height: 0 !important;
            visibility: hidden !important;
        }
        #MainMenu, footer,
        [data-testid="stToolbar"],
        [data-testid="stStatusWidget"],
        [data-testid="stDecoration"],
        .stDeployButton,
        ._profileContainer_gzau3_53,
        ._profilePreview_gzau3_63,
        [class*="viewerBadge"],
        [class*="StatusWidget"] {
            display: none !important;
            visibility: hidden !important;
        }
        [data-testid="stSidebar"] { display: none !important; }
        body:not(:has(#pms-d:checked)) .stApp,
        body:not(:has(#pms-d:checked)) [data-testid="stAppViewContainer"],
        body:not(:has(#pms-d:checked)) .main {
            background: #FFFFFF !important;
        }
        .main .block-container {
            max-width: 100% !important;
            padding: 64px 0 48px !important;
        }
        .auth-logo-card {
            background: transparent;
            text-align: center;
            padding: 22px 32px 18px;
            margin-bottom: 16px;
        }
        [data-testid="stColumn"] {
            background: transparent !important;
            box-shadow: none !important;
        }
        .auth-logo-title {
            color: #1E1A3A;
            font-size: 56px;
            font-weight: 900;
            letter-spacing: 0;
            line-height: 1.05;
            margin-bottom: 5px;
        }
        .auth-logo-sub {
            color: #7B79AA;
            font-size: 16px;
            font-weight: 600;
            letter-spacing: 0.3px;
        }
        .auth-small {
            text-align: center;
            color: #A09AC5;
            font-size: 14px;
            margin: 16px 0 8px;
        }
        div[data-testid="stTextInput"] { margin-bottom: 14px !important; }
        div[data-testid="stTextInput"] label {
            color: #3A3660 !important;
            font-size: 14px !important;
            font-weight: 700 !important;
            margin-bottom: 4px !important;
        }
        div[data-testid="stTextInput"] > div[data-baseweb="input"] {
            min-height: 52px !important;
            border-radius: 12px !important;
            border: none !important;
            background: #F2F2F5 !important;
            box-shadow: none !important;
        }
        div[data-testid="stTextInput"] input {
            color: #1E1A3A !important;
            caret-color: #1E1A3A !important;
            font-size: 15px !important;
            background: #F2F2F5 !important;
        }
        div[data-testid="stTextInput"] > div[data-baseweb="input"]:focus-within {
            border: 1.5px solid #7B6FD4 !important;
            box-shadow: 0 0 0 2px rgba(123,111,212,0.18) !important;
        }
        div[data-testid="stTextInput"] input::placeholder {
            color: #AEACC8 !important;
        }
        [data-testid="InputInstructions"] { display: none !important; }
        div[data-testid="stCheckbox"] {
            margin-bottom: 16px !important;
        }
        div[data-testid="stCheckbox"] label {
            color: #4B466D !important;
            font-size: 14px !important;
            font-weight: 600 !important;
        }
        div.stButton > button {
            height: 54px !important;
            border-radius: 12px !important;
            font-size: 15px !important;
            font-weight: 700 !important;
            width: 100% !important;
            transition: all 0.18s !important;
        }
        [data-testid="stBaseButton-primary"] {
            background: #655CF0 !important;
            color: #FFFFFF !important;
            border: 1.5px solid #655CF0 !important;
            box-shadow: 0 8px 18px rgba(101,92,240,0.22) !important;
        }
        [data-testid="stBaseButton-primary"]:hover {
            background: #5A52DF !important;
            border-color: #5A52DF !important;
            color: #FFFFFF !important;
        }
        [data-testid="stBaseButton-secondary"] {
            background: #FAFAFC !important;
            color: #3D3580 !important;
            border: 1.5px solid #EEEFF5 !important;
            box-shadow: 0 2px 8px rgba(30,26,58,0.05) !important;
        }
        [data-testid="stBaseButton-secondary"]:hover {
            background: #FFFFFF !important;
            border-color: #D8D3F5 !important;
            color: #493E9A !important;
        }
        [data-testid="stBaseButton-primary"] p,
        [data-testid="stBaseButton-primary"] span {
            color: #FFFFFF !important;
        }
        [data-testid="stBaseButton-secondary"] p,
        [data-testid="stBaseButton-secondary"] span {
            color: #3D3580 !important;
        }
        hr {
            border: 0 !important;
            border-top: 1px solid #EDEBF8 !important;
            margin: 18px 0 0 !important;
        }
        /* 로그인 페이지 다크모드 */
        body:has(#pms-d:checked) .auth-logo-title { color: #ffffff !important; }
        body:has(#pms-d:checked) .auth-logo-sub   { color: #a6adc8 !important; }
        body:has(#pms-d:checked) .auth-small       { color: #a6adc8 !important; }
        body:has(#pms-d:checked) div[data-testid="stTextInput"] label { color: #cdd6f4 !important; }
        body:has(#pms-d:checked) div[data-testid="stTextInput"] > div[data-baseweb="input"] {
            background: #252535 !important; border: 1.5px solid #45475a !important;
        }
        body:has(#pms-d:checked) div[data-testid="stTextInput"] input {
            background: #252535 !important; color: #ffffff !important;
            caret-color: #ffffff !important;
        }
        body:has(#pms-d:checked) div[data-testid="stTextInput"] > div[data-baseweb="input"]:focus-within {
            border-color: #818cf8 !important;
            box-shadow: 0 0 0 2px rgba(129,140,248,0.22) !important;
        }
        body:has(#pms-d:checked) div[data-testid="stCheckbox"] label { color: #cdd6f4 !important; }
        body:has(#pms-d:checked) hr { border-top-color: #45475a !important; }
        /* 모바일: 로그인 화면 */
        @media (max-width: 768px) {
            .main .block-container {
                padding: 24px 16px 32px !important;
                max-width: 100% !important;
            }
            /* 컬럼 가로 배치 해제 → 전체 폭 사용 */
            [data-testid="stHorizontalBlock"] {
                flex-direction: column !important;
            }
            [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
                width: 100% !important;
                min-width: 100% !important;
                flex: 1 1 100% !important;
            }
            .auth-logo-title {
                font-size: 36px !important;
            }
            .auth-logo-sub {
                font-size: 14px !important;
                word-break: keep-all !important;
            }
            .auth-logo-card {
                padding: 12px 16px 10px !important;
                margin-bottom: 8px !important;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    _, center, _ = st.columns([1.5, 1.1, 1.5])

    with center:
        st.markdown(
            """
            <div class="auth-logo-card">
                <div class="auth-logo-title">내부 관리</div>
                <div class="auth-logo-sub">Webcash We · 360° Control</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.session_state.auth_mode = "login"

        if st.session_state.auth_mode == "login":
            cookie_manager = safe_cookie_controller()
            sid = cookie_get(cookie_manager, "saved_id", "")

            with st.form("login_form", border=False):
                u_id = st.text_input("아이디", value=sid, placeholder="아이디를 입력하세요", key="l_id")
                u_pw = st.text_input("비밀번호", type="password", placeholder="비밀번호를 입력하세요", key="l_pw")
                save_id_cb = st.checkbox("아이디 저장", value=bool(sid))
                _login_submitted = st.form_submit_button("로그인", use_container_width=True, type="primary")

            if _login_submitted:
                db = st.session_state.user_db
                u_id_str = str(u_id).strip() if u_id else ""
                u_pw_str = str(u_pw).strip() if u_pw else ""

                is_super = u_id_str == "1" and u_pw_str == "1"
                is_user = (
                    u_id_str in db
                    and db[u_id_str].get("pw", "") == u_pw_str
                    and db[u_id_str].get("access") == "허용"
                )

                if is_super or is_user:
                    if save_id_cb:
                        cookie_set(cookie_manager, "saved_id", u_id_str)
                    else:
                        try:
                            cookie_remove(cookie_manager, "saved_id")
                        except Exception:
                            pass
                    persist_login_session(u_id_str)

                    user = db.get(u_id_str, {"role": "관리자", "name": "최고관리자"})
                    st.session_state.logged_in = True
                    st.session_state.user_role = user.get("role", "관리자")
                    st.session_state.user_name = user.get("name", "최고관리자")
                    st.session_state.login_time = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")

                    # 저장된 데이터 로드
                    saved_db = load_db(SAVED_STATE_FILE, {})
                    user_saved = saved_db.get(st.session_state.user_name)
                    if user_saved:
                        if user_saved.get("user_excel_data"):
                            st.session_state.user_excel_data = pd.DataFrame.from_dict(user_saved["user_excel_data"])
                        if user_saved.get("user_prev_month_sel"):
                            st.session_state.user_prev_month_sel = user_saved["user_prev_month_sel"]

                    admin_analysis = saved_db.get("admin_analysis")
                    if admin_analysis and admin_analysis.get("sent_df"):
                        st.session_state.analysis_result = pd.DataFrame.from_dict(admin_analysis["sent_df"])
                    if admin_analysis and admin_analysis.get("adm_prev_month"):
                        st.session_state.adm_prev_month = admin_analysis["adm_prev_month"]
                    deadline_info = saved_db.get("deadline")
                    if deadline_info:
                        st.session_state.deadline_time = deadline_info.get("time", "")
                    report_closed_info = saved_db.get("report_closed")
                    if report_closed_info:
                        st.session_state.report_closed = report_closed_info.get("time", "")

                    st.session_state.current_menu = BILLING_MENU
                    persist_current_menu()
                    st.rerun()
                elif not u_id_str:
                    st.error("아이디를 입력해주세요.")
                elif not u_pw_str:
                    st.error("비밀번호를 입력해주세요.")
                else:
                    st.error("아이디 또는 비밀번호가 올바르지 않습니다.")

            st.divider()
            st.caption("계정은 관리자가 발급한 ID와 비밀번호로만 로그인할 수 있습니다.")


def show_sidebar():
    # 모바일: 메뉴 클릭 후 사이드바를 CSS로 숨김
    if st.session_state.pop("_close_sidebar_mobile", False):
        st.markdown("""
        <style>
        @media (max-width: 768px) {
            [data-testid="stSidebar"] {
                transform: translateX(-100vw) !important;
                visibility: hidden !important;
                pointer-events: none !important;
                transition: none !important;
            }
            [data-testid="stSidebarCollapsedControl"] {
                display: flex !important;
                visibility: visible !important;
            }
        }
        </style>
        """, unsafe_allow_html=True)
    with st.sidebar:
        st.markdown(
            """
            <style>
            #MainMenu, footer,
            [data-testid="stStatusWidget"],
            [data-testid="stDecoration"],
            [data-testid="stAppViewBlockContainer"] > div:last-child,
            .stDeployButton,
            ._profileContainer_gzau3_53,
            ._profilePreview_gzau3_63,
            .viewerBadge_container__r5tak,
            .viewerBadge_link__qRIco,
            [class*="viewerBadge"],
            [class*="StatusWidget"],
            [data-testid="stToolbar"] > * {
                display: none !important;
                visibility: hidden !important;
            }
            header[data-testid="stHeader"] {
                background: transparent !important;
                height: 0px !important;
                min-height: 0px !important;
                overflow: hidden !important;
            }
            [data-testid="stSidebar"] {
                display: flex !important;
                transform: none !important;
                visibility: visible !important;
                min-width: 272px !important;
                width: 272px !important;
            }
            [data-testid="stSidebarCollapseButton"],
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"] {
                display: none !important;
            }
            /* ══ Zoho CRM 스타일 네비게이션 팔레트 ══
               배경: #16284A  카드: #1F3358  테두리: #24395E
               강조: #2F6FED  텍스트: #E8EEF8  보조: #9FB3D6
               섹션: #7689AD  액션 호버: #5A2A38              */
            [data-testid="stSidebar"] {
                background-color: #16284A !important;
                border-right: 1px solid #24395E !important;
            }
            [data-testid="stSidebar"] > div:first-child {
                padding-top: 0 !important;
            }
            [data-testid="stSidebarContent"] {
                padding: 14px 12px 16px !important;
                background-color: #16284A !important;
            }
            [data-testid="stSidebar"] * {
                color: #E8EEF8 !important;
            }
            [data-testid="stSidebar"] .gpt-side-shell {
                display: flex;
                flex-direction: column;
                gap: 12px;
                padding: 4px 2px 8px;
            }
            [data-testid="stSidebar"] .gpt-brand {
                display: flex;
                align-items: center;
                gap: 10px;
                min-height: 44px;
                padding: 6px 8px;
                border-radius: 10px;
            }
            [data-testid="stSidebar"] .gpt-brand-mark {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 48px;
                height: 48px;
                border: none;
                border-radius: 12px;
                color: #ffffff !important;
                font-size: 20px;
                font-weight: 900;
                letter-spacing: 0;
                background: linear-gradient(135deg, #2F6FED 0%, #1B4FC4 100%);
                box-shadow: 0 4px 10px rgba(47,111,237,0.35);
            }
            [data-testid="stSidebar"] .gpt-brand-text {
                display: flex;
                flex-direction: column;
                line-height: 1.2;
            }
            [data-testid="stSidebar"] .gpt-brand-title {
                font-size: 30px;
                font-weight: 800;
                color: #ffffff !important;
                line-height: 1.1;
                word-break: keep-all;
            }
            [data-testid="stSidebar"] .gpt-brand-subtitle {
                margin-top: 2px;
                font-size: 12px;
                color: #8FA3C7 !important;
                font-weight: 500;
            }
            [data-testid="stSidebar"] .gpt-user-card {
                margin: 2px 2px 8px;
                padding: 12px 14px;
                border: 1px solid #29406B;
                border-radius: 14px;
                background: #1F3358;
                box-shadow: 0 4px 14px rgba(0,0,0,0.18);
            }
            [data-testid="stSidebar"] .gpt-user-name {
                font-size: 14px;
                font-weight: 800;
                color: #ffffff !important;
                line-height: 1.25;
            }
            [data-testid="stSidebar"] .gpt-user-meta {
                margin-top: 4px;
                font-size: 11px;
                color: #9FB3D6 !important;
                font-weight: 500;
                line-height: 1.35;
            }
            [data-testid="stSidebar"] .gpt-section {
                margin: 4px 8px 1px;
                font-size: 12px;
                color: #7689AD !important;
                font-weight: 700;
                letter-spacing: 0.02em;
                text-transform: uppercase;
            }
            [data-testid="stSidebar"] .gpt-section:first-of-type {
                margin-top: 2px;
            }
            [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
                margin: 0 !important;
                padding: 0 !important;
            }
            [data-testid="stSidebar"] .gpt-sidebar-spacer {
                height: 10px;
            }
            [data-testid="stSidebar"] .gpt-sidebar-divider {
                height: 1px;
                margin: 12px 8px;
                background: #24395E;
            }
            [data-testid="stSidebar"] div.stButton {
                margin: 0 !important;
                padding: 0 !important;
            }
            [data-testid="stSidebar"] div.stButton > button {
                margin-bottom: 0 !important;
                min-height: 34px !important;
                height: auto !important;
            }
            [data-testid="stSidebar"] [data-testid="stElementContainer"] {
                padding: 0 !important;
                margin: 0 !important;
            }
            [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
                gap: 2px !important;
            }
            [data-testid="stSidebar"] hr {
                border-color: #24395E !important;
                margin: 10px 0 !important;
            }
            [data-testid="stSidebar"] div.stButton > button {
                background: transparent !important;
                border: none !important;
                box-shadow: none !important;
                color: #C9D6EC !important;
                font-size: 14px !important;
                font-weight: 600 !important;
                text-align: left !important;
                padding: 8px 10px !important;
                border-radius: 8px !important;
                justify-content: flex-start !important;
                transition: background 0.12s ease, color 0.12s ease !important;
            }
            [data-testid="stSidebar"] div.stButton > button p,
            [data-testid="stSidebar"] div.stButton > button span,
            [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] p,
            [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] span {
                font-size: 14px !important;
                font-weight: 600 !important;
                color: #C9D6EC !important;
                white-space: normal !important;
                line-height: 1.25 !important;
            }
            [data-testid="stSidebar"] div.stButton > button:hover {
                background: #233A61 !important;
                color: #ffffff !important;
            }
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + div [data-testid="stButton"] button,
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + [data-testid="stElementContainer"] [data-testid="stButton"] button {
                background: #2F6FED !important;
                color: #ffffff !important;
                box-shadow: 0 4px 12px rgba(47,111,237,0.35);
            }
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + div [data-testid="stButton"] button p,
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + [data-testid="stElementContainer"] [data-testid="stButton"] button p {
                color: #ffffff !important;
                font-weight: 700 !important;
            }
            [data-testid="stSidebar"] .gpt-logout-marker + div [data-testid="stButton"] button,
            [data-testid="stSidebar"] .gpt-logout-marker + [data-testid="stElementContainer"] [data-testid="stButton"] button {
                color: #B9C8E4 !important;
            }
            [data-testid="stSidebar"] .gpt-logout-marker + div [data-testid="stButton"] button:hover,
            [data-testid="stSidebar"] .gpt-logout-marker + [data-testid="stElementContainer"] [data-testid="stButton"] button:hover {
                background: #5A2A38 !important;
                color: #ffffff !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        def render_nav_button(menu_name):
            active_class = " gpt-nav-active" if st.session_state.current_menu == menu_name else ""
            st.markdown(f"<div class='gpt-nav-marker{active_class}'></div>", unsafe_allow_html=True)
            if st.button(menu_name, use_container_width=True, key=f"nav_{menu_name}"):
                st.session_state.current_menu = menu_name
                persist_current_menu()
                st.session_state["_close_sidebar_mobile"] = True
                st.rerun()

        st.markdown(
            f"<div class='gpt-side-shell'>"
            f"<div class='gpt-brand'>"
            f"<div class='gpt-brand-mark'>AX</div>"
            f"<div class='gpt-brand-text'>"
            f"<div class='gpt-brand-title'>내부 관리</div>"
            f"<div class='gpt-brand-subtitle'>Webcash We · 360° Control</div>"
            f"</div></div>"
            f"</div>",
            unsafe_allow_html=True,
        )

        render_nav_button(BILLING_MENU)


def _load_delegated_workplaces():
    default_data = {"workplaces": [], "requests": []}
    data = load_db(DELEGATED_WORKPLACE_FILE, default_data)
    if not isinstance(data, dict):
        return default_data
    data.setdefault("workplaces", [])
    data.setdefault("requests", [])
    return data


def _save_delegated_workplaces(data):
    data.setdefault("workplaces", [])
    data.setdefault("requests", [])
    save_db(DELEGATED_WORKPLACE_FILE, data, allow_shrink=True)


def _load_approvals():
    default_data = {"documents": []}
    data = load_db(APPROVAL_FILE, default_data)
    if not isinstance(data, dict):
        return default_data
    data.setdefault("documents", [])
    return data


def _save_approvals(data):
    data.setdefault("documents", [])
    save_db(APPROVAL_FILE, data, allow_shrink=True)


def _load_usage_reports():
    default_data = {"reports": []}
    data = load_db(USAGE_REPORT_FILE, default_data)
    if not isinstance(data, dict):
        return default_data
    data.setdefault("reports", [])
    return data


def _save_usage_reports(data):
    data.setdefault("reports", [])
    save_db(USAGE_REPORT_FILE, data, allow_shrink=True)


def _load_bank_accounts():
    default_data = {"accounts": []}
    data = load_db(BANK_ACCOUNT_FILE, default_data)
    if not isinstance(data, dict):
        return default_data
    data.setdefault("accounts", [])
    return data


def _save_bank_accounts(data):
    data.setdefault("accounts", [])
    save_db(BANK_ACCOUNT_FILE, data, allow_shrink=True)


def _account_company_label(account, workplaces_by_id):
    site = workplaces_by_id.get(account.get("linked_workplace_id"))
    if site:
        return site.get("workplace_name", "")
    return account.get("account_type", "")


def _my_workplaces(workplaces):
    """로그인 사용자가 관리자가 아닐 경우, 본인 담당/소속 사업장만 반환한다."""
    if st.session_state.user_role == "관리자":
        return workplaces

    user_name = st.session_state.get("user_name", "")
    user_dept = ""
    for info in _real_users(st.session_state.user_db).values():
        if info.get("name") == user_name:
            user_dept = info.get("dept_type", "")
            break
    return [
        row for row in workplaces
        if row.get("manager_name") == user_name or row.get("workplace_name") == user_dept
    ]


def _load_company_profile():
    default_data = {
        "name": "",
        "business_number": "",
        "ceo_name": "",
        "address": "",
        "contact": "",
        "memo": "",
        "updated_at": "",
    }
    data = load_db(COMPANY_PROFILE_FILE, default_data)
    if not isinstance(data, dict):
        return default_data
    for key, value in default_data.items():
        data.setdefault(key, value)
    return data


def _save_company_profile(data):
    save_db(COMPANY_PROFILE_FILE, data, allow_shrink=True)


def _normalize_name(value):
    return re.sub(r"\s+", "", str(value or "").strip().lower())


def _format_won(value):
    try:
        return f"{int(float(value or 0)):,}원"
    except Exception:
        return "0원"


def _current_kst():
    return datetime.utcnow() + timedelta(hours=9)


def _parse_kst_datetime(value):
    if not value:
        return None
    parsed = pd.to_datetime(str(value), errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime().replace(tzinfo=None)


def _normalize_account_number(value):
    return re.sub(r"[^0-9A-Za-z]", "", str(value or ""))


def _coerce_int_amount(value):
    try:
        cleaned = re.sub(r"[^0-9.-]", "", str(value or "0"))
        return int(float(cleaned or 0))
    except Exception:
        return 0


def _history_amount_value(history_row, previous_balance=None):
    for key in ["amount", "deposit_amount", "transaction_amount", "입금액", "거래금액"]:
        if key in history_row:
            amount = _coerce_int_amount(history_row.get(key))
            if amount:
                return amount
    if previous_balance is not None and "balance" in history_row:
        balance = _coerce_int_amount(history_row.get("balance"))
        diff = balance - previous_balance
        if diff != 0:
            return diff
    if previous_balance is None and "balance" in history_row:
        return _coerce_int_amount(history_row.get("balance"))
    return 0


def _deposit_match_label(account, target_amount):
    history = account.get("balance_history", []) or []
    if not history:
        return ""

    sorted_history = sorted(history, key=lambda item: str(item.get("at", "")))
    previous_balance = None
    for item in sorted_history:
        amount = _history_amount_value(item, previous_balance)
        if "balance" in item:
            previous_balance = _coerce_int_amount(item.get("balance"))
        if amount != int(target_amount or 0):
            continue
        deposited_at = item.get("at") or item.get("date") or item.get("transaction_at") or ""
        deposited_date = str(deposited_at).split(" ")[0] if deposited_at else "입금일자 확인"
        return f"{deposited_date} 입금"
    return ""


def _reconcile_transfer_deposits(pending_rows, bank_accounts):
    accounts_by_number = {
        _normalize_account_number(row.get("account_number")): row
        for row in bank_accounts
        if row.get("account_number")
    }
    matched_count = 0
    for request_row in pending_rows:
        account_number = request_row.get("transfer_account_number", "")
        if not account_number:
            linked_account = next(
                (
                    row for row in bank_accounts
                    if row.get("linked_workplace_id") == request_row.get("workplace_id")
                ),
                {},
            )
            account_number = linked_account.get("account_number", "")
        account = accounts_by_number.get(_normalize_account_number(account_number))
        label = _deposit_match_label(account or {}, request_row.get("request_amount", 0))
        request_row["deposit_reconciliation"] = label or "미확인"
        request_row["deposit_reconciled_at"] = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
        if label:
            matched_count += 1
    return matched_count


def _is_recent_request(value, days=7):
    requested_at = _parse_kst_datetime(value)
    if not requested_at:
        return False
    cutoff = _current_kst() - timedelta(days=days)
    return requested_at >= cutoff


def _month_key(value):
    if not value:
        return ""
    return str(value)[:7]


def _workplace_request_metrics(requests):
    current_month = _current_kst().strftime("%Y-%m")
    month_requests = [row for row in requests if _month_key(row.get("requested_at")) == current_month]
    paid = [row for row in month_requests if row.get("status") == "이체 완료"]
    pending = [row for row in requests if row.get("status") == "요청"]
    approved = [row for row in requests if row.get("status") == "품의 확정"]
    return {
        "pending": len(pending),
        "approved": len(approved),
        "paid_amount": sum(int(row.get("request_amount", 0) or 0) for row in paid),
        "month_requests": len(month_requests),
    }


def _workplace_forecast_rows(workplaces, requests):
    current_month = _current_kst().strftime("%Y-%m")
    today_day = int(_current_kst().strftime("%d"))
    rows = []
    for site in workplaces:
        site_requests = [
            row for row in requests
            if row.get("workplace_id") == site.get("id") and row.get("status") == "이체 완료"
        ]
        month_totals = {}
        for row in site_requests:
            key = _month_key(row.get("paid_at") or row.get("requested_at"))
            if not key:
                continue
            month_totals[key] = month_totals.get(key, 0) + int(row.get("request_amount", 0) or 0)
        historical = [amount for key, amount in month_totals.items() if key != current_month]
        current_amount = month_totals.get(current_month, 0)
        average_amount = int(sum(historical) / len(historical)) if historical else current_amount
        recommended = max(average_amount, current_amount)
        regular_day = int(site.get("regular_payment_day", 0) or 0)
        days_to_payment = regular_day - today_day if regular_day else None
        risk_notes = []
        if average_amount and current_amount > average_amount * 1.3:
            risk_notes.append("평상시 대비 사용 증가")
        if days_to_payment is not None and 0 <= days_to_payment <= 3:
            risk_notes.append("정기 지급일 도래")
        if not site_requests:
            risk_notes.append("지급 이력 없음")
        rows.append(
            {
                "사업장명": site.get("workplace_name", ""),
                "담당자": site.get("manager_name", ""),
                "정기 지급일": f"매월 {regular_day}일" if regular_day else "미등록",
                "당월 지급액": _format_won(current_amount),
                "평균 지급액": _format_won(average_amount),
                "추천 지급액": _format_won(recommended),
                "리스크": ", ".join(risk_notes) if risk_notes else "정상",
            }
        )
    return rows


def _render_page_chrome(breadcrumb_items):
    crumbs = " &gt; ".join(html.escape(item) for item in breadcrumb_items[:-1])
    current = html.escape(breadcrumb_items[-1])
    st.markdown(
        f"""
        <style>
        .block-container {{
            max-width: 1180px !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }}
        .pms-breadcrumb {{
            text-align: right;
            color: #64748b;
            font-size: 12px;
            font-weight: 600;
            margin: -4px 0 14px;
        }}
        .pms-breadcrumb span {{
            color: #008c78;
        }}
        [data-testid="stCheckbox"] label {{
            white-space: nowrap !important;
        }}
        </style>
        <div class="pms-breadcrumb">{crumbs} &gt; <span>{current}</span></div>
        """,
        unsafe_allow_html=True,
    )


def _workplace_dashboard_css():
    st.markdown(
        """
        <style>
        .sales-home-shell {
            position: relative;
            overflow: hidden;
            border-radius: 22px;
            padding: 14px 18px 22px;
            background:
                radial-gradient(circle at 48% 42%, rgba(16, 185, 173, 0.30) 0 32%, transparent 33%),
                linear-gradient(135deg, #f7f8fb 0%, #ffffff 52%, #f4f6fa 100%);
            box-shadow: 0 18px 50px rgba(15, 23, 42, 0.12);
            border: 1px solid rgba(226, 232, 240, 0.9);
        }
        .sales-topbar {
            height: 46px;
            display: flex;
            align-items: center;
            gap: 18px;
            border-radius: 24px;
            padding: 0 18px;
            background: rgba(255, 255, 255, 0.86);
            box-shadow: inset 0 0 0 1px rgba(226, 232, 240, 0.85);
            margin-bottom: 18px;
        }
        .sales-cloud {
            width: 34px;
            height: 22px;
            border-radius: 999px;
            background: #1d9bd7;
            position: relative;
            flex: 0 0 auto;
        }
        .sales-cloud::before, .sales-cloud::after {
            content: "";
            position: absolute;
            background: #1d9bd7;
            border-radius: 999px;
        }
        .sales-cloud::before { width: 18px; height: 18px; left: 5px; top: -7px; }
        .sales-cloud::after { width: 18px; height: 18px; right: 4px; top: -5px; }
        .sales-nav-pill {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            font-weight: 800;
            color: #111827;
            white-space: nowrap;
        }
        .sales-app-icon {
            display: inline-grid;
            place-items: center;
            width: 18px;
            height: 18px;
            border-radius: 5px;
            background: #14b8a6;
            color: #ffffff;
            font-size: 12px;
            font-weight: 900;
        }
        .sales-home-title {
            font-size: 22px;
            font-weight: 850;
            color: #111827;
            margin: 8px 0 14px;
            letter-spacing: 0;
        }
        .work-card {
            min-height: 170px;
            border-radius: 16px;
            background: rgba(255, 255, 255, 0.90);
            border: 1px solid rgba(226, 232, 240, 0.8);
            box-shadow: 0 16px 35px rgba(15, 23, 42, 0.10);
            padding: 18px 18px 14px;
        }
        .work-card-title {
            color: #111827;
            font-size: 15px;
            font-weight: 850;
            margin-bottom: 12px;
        }
        .donut {
            width: 118px;
            height: 118px;
            border-radius: 50%;
            display: grid;
            place-items: center;
            margin: 0 auto 10px;
            background: conic-gradient(#8b5cf6 var(--p), #2f8eea var(--p) 72%, #7bb9f4 72% 100%);
        }
        .donut-inner {
            width: 76px;
            height: 76px;
            border-radius: 50%;
            background: #ffffff;
            display: grid;
            place-items: center;
            text-align: center;
            color: #111827;
            line-height: 1.12;
            font-size: 13px;
            box-shadow: inset 0 0 0 1px rgba(226, 232, 240, 0.9);
        }
        .donut-value {
            display: block;
            font-size: 20px;
            font-weight: 900;
            margin-bottom: 2px;
        }
        .card-dots {
            display: flex;
            gap: 6px;
            justify-content: flex-end;
            margin-top: 4px;
        }
        .card-dots span {
            width: 5px;
            height: 5px;
            border-radius: 50%;
            background: #2f8eea;
            display: block;
        }
        .card-dots span:nth-child(3), .card-dots span:nth-child(4) { background: #8b5cf6; }
        .suggestion-panel {
            min-height: 386px;
            border-radius: 18px;
            background: rgba(255, 255, 255, 0.88);
            border: 1px solid rgba(226, 232, 240, 0.85);
            box-shadow: 0 16px 35px rgba(15, 23, 42, 0.10);
            padding: 18px;
        }
        .suggestion-title {
            color: #6b7280;
            font-size: 15px;
            font-weight: 850;
            margin-bottom: 10px;
        }
        .suggestion-row {
            display: grid;
            grid-template-columns: 28px 1fr;
            gap: 12px;
            padding: 13px 0;
        }
        .suggestion-icon {
            width: 24px;
            height: 24px;
            border-radius: 5px;
            background: #8b7cf6;
            color: #fff;
            display: grid;
            place-items: center;
            font-size: 13px;
            font-weight: 900;
        }
        .skeleton {
            height: 6px;
            border-radius: 99px;
            background: #d7d7d7;
            margin: 8px 0;
        }
        .skeleton.short { width: 38%; }
        .skeleton.mid { width: 62%; }
        .skeleton.long { width: 96%; }
        .workflow-card {
            border-radius: 16px;
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(226, 232, 240, 0.85);
            box-shadow: 0 12px 30px rgba(15, 23, 42, 0.10);
            padding: 16px;
            margin-top: 14px;
        }
        .workflow-title {
            font-size: 15px;
            font-weight: 850;
            color: #111827;
            margin-bottom: 12px;
        }
        .request-card {
            border-radius: 14px;
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(226, 232, 240, 0.9);
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08);
            padding: 13px 14px;
            min-height: 122px;
            margin-bottom: 12px;
        }
        .request-title {
            color: #111827;
            font-size: 15px;
            font-weight: 850;
            margin-bottom: 4px;
        }
        .request-meta {
            color: #6b7280;
            font-size: 13px;
            line-height: 1.4;
        }
        .status-chip {
            display: inline-flex;
            align-items: center;
            border-radius: 999px;
            padding: 4px 9px;
            font-size: 12px;
            font-weight: 800;
            background: #eef2ff;
            color: #4f46e5;
            margin-top: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _donut_card(title, value, label, percent):
    pct = max(0, min(100, int(percent)))
    return f"""
    <div class="work-card">
        <div class="work-card-title">{html.escape(str(title))}</div>
        <div class="donut" style="--p:{pct}%;">
            <div class="donut-inner">
                <div>
                    <span class="donut-value">{html.escape(str(value))}</span>
                    {html.escape(str(label))}
                </div>
            </div>
        </div>
        <div class="card-dots"><span></span><span></span><span></span><span></span></div>
    </div>
    """


def show_advance_payment_request():
    _render_page_chrome(["홈", "업무", "전도금 요청", "전도금 요청"])
    st.markdown("### 전도금 요청")
    st.caption("위탁 사업장의 전도금을 요청하고 처리 현황을 확인합니다.")

    data = _load_delegated_workplaces()
    workplaces = data.get("workplaces", [])
    requests = data.get("requests", [])

    workplaces = _my_workplaces(workplaces)

    my_requests = [
        row
        for row in requests
        if row.get("requested_by") == st.session_state.get("user_name", "")
        and _is_recent_request(row.get("requested_at"), days=7)
    ]
    if my_requests:
        st.markdown("#### 최근 요청현황")
        _workplace_dashboard_css()
        request_cols = st.columns(3)
        for index, row in enumerate(sorted(my_requests, key=lambda item: item.get("requested_at", ""), reverse=True)):
            with request_cols[index % 3]:
                reject_line = ""
                if row.get("status") == "반려" and row.get("reject_reason"):
                    reject_line = f"<div class='request-meta'>반려 사유: {html.escape(str(row.get('reject_reason')))}</div>"
                st.markdown(
                    (
                        "<div class='request-card'>"
                        f"<div class='request-title'>{html.escape(str(row.get('workplace_name', '')))}</div>"
                        f"<div class='request-meta'>{html.escape(str(row.get('request_reason') or '요청 사유 없음'))}</div>"
                        f"<div class='request-meta'>{html.escape(_format_won(row.get('request_amount')))} · {html.escape(str(row.get('requested_at', '')))}</div>"
                        f"{reject_line}"
                        f"<span class='status-chip'>{html.escape(str(row.get('status', '')))}</span>"
                        "</div>"
                    ),
                    unsafe_allow_html=True,
                )
    else:
        st.info("최근 7일간 요청한 전도금 내역이 없습니다.")

    if not workplaces:
        if st.session_state.user_role != "관리자":
            st.info("담당으로 등록된 사업장이 없습니다. 관리자에게 문의해주세요.")
        else:
            st.info("관리자 메뉴의 [위탁 사업장 관리]에서 사업장을 먼저 등록해주세요.")
    else:
        workplace_options = {row.get("workplace_name", ""): row for row in workplaces}

        if st.session_state.get("_show_advance_payment_form"):
            with st.form("delegated_fund_request_form", clear_on_submit=True):
                selected_name = st.selectbox("사업장", list(workplace_options.keys()))
                requested_by = st.text_input(
                    "요청자명",
                    value=st.session_state.get("user_name", ""),
                    placeholder="요청자명",
                )
                requester_phone = st.text_input("핸드폰번호", placeholder="010-0000-0000")
                request_amount = st.number_input("요청 금액", min_value=0, step=100000, format="%d")
                request_reason = st.text_area("요청 사유", placeholder="전도금 사용 목적 및 필요 사유")
                col_submit, col_cancel = st.columns(2)
                requested = col_submit.form_submit_button("전도금 요청 등록", type="primary", use_container_width=True)
                cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

            if requested:
                site = workplace_options[selected_name]
                if request_amount <= 0:
                    st.warning("요청 금액을 입력해주세요.")
                else:
                    now_str = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
                    request_id = int(time.time() * 1000)
                    requests.append(
                        {
                            "id": request_id,
                            "workplace_id": site.get("id"),
                            "workplace_name": site.get("workplace_name"),
                            "request_amount": int(request_amount),
                            "requested_by": requested_by.strip(),
                            "requester_phone": requester_phone.strip(),
                            "request_reason": request_reason.strip(),
                            "status": "요청",
                            "requested_at": now_str,
                            "approved_at": "",
                            "paid_at": "",
                            "approval_doc_id": request_id,
                            "reject_reason": "",
                            "transfer_file_generated_at": "",
                        }
                    )
                    _save_delegated_workplaces(data)

                    approvals = _load_approvals()
                    approvals["documents"].append(
                        {
                            "id": request_id,
                            "doc_type": "전도금요청",
                            "title": f"{site.get('workplace_name', '')} 전도금 요청",
                            "amount": int(request_amount),
                            "requester": requested_by.strip(),
                            "requester_phone": requester_phone.strip(),
                            "reason": request_reason.strip(),
                            "ref_request_id": request_id,
                            "ref_workplace_id": site.get("id"),
                            "status": "결재대기",
                            "requested_at": now_str,
                            "processed_at": "",
                            "processed_by": "",
                            "reject_reason": "",
                        }
                    )
                    _save_approvals(approvals)

                    st.session_state["_show_advance_payment_form"] = False
                    st.success("전도금 요청이 등록되었습니다. [전자결재]에서 처리 현황을 확인할 수 있습니다.")
                    st.rerun()
            if cancelled:
                st.session_state["_show_advance_payment_form"] = False
                st.rerun()
        else:
            col_add, _ = st.columns([35, 165])
            if col_add.button("+ 전도금 요청 등록", key="show_advance_payment_form_btn", use_container_width=True):
                st.session_state["_show_advance_payment_form"] = True
                st.rerun()


def show_e_approval():
    st.markdown("### 전자결재")
    st.caption("전도금 요청 등 결재 문서를 승인/반려 처리합니다.")

    approvals = _load_approvals()
    documents = approvals.get("documents", [])
    wp_data = _load_delegated_workplaces()
    wp_requests = wp_data.get("requests", [])
    requests_by_id = {row.get("id"): row for row in wp_requests}

    # 결재함 정보 표시
    pending_count = len([doc for doc in documents if doc.get("status") == "결재대기"])
    done_count = len([doc for doc in documents if doc.get("status") in ["승인", "반려"]])

    st.markdown(
        f"""
        <div style='background: linear-gradient(135deg, #2F6FED 0%, #1B4FC4 100%);
                    padding: 20px; border-radius: 12px; margin-bottom: 20px; color: white;'>
            <div style='font-size: 18px; font-weight: 700; margin-bottom: 10px;'>
                {html.escape(st.session_state.user_name)}님 반갑습니다
            </div>
            <div style='display: flex; gap: 30px; font-size: 14px;'>
                <div><strong>결재대기:</strong> {pending_count}건</div>
                <div><strong>처리완료:</strong> {done_count}건</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    pending_tab, done_tab, stats_tab = st.tabs(["결재대기", "처리 완료", "처리 통계"])

    with pending_tab:
        pending_docs = [doc for doc in documents if doc.get("status") == "결재대기"]
        if not pending_docs:
            st.info("결재 대기 중인 문서가 없습니다.")
        else:
            _workplace_dashboard_css()
            cols = st.columns(3)
            for index, doc in enumerate(sorted(pending_docs, key=lambda item: item.get("requested_at", ""), reverse=True)):
                with cols[index % 3]:
                    st.markdown(
                        (
                            "<div class='request-card'>"
                            f"<div class='request-title'>{html.escape(str(doc.get('title', '')))}</div>"
                            f"<div class='request-meta'>{html.escape(str(doc.get('reason') or '사유 없음'))}</div>"
                            f"<div class='request-meta'>{html.escape(_format_won(doc.get('amount')))} · {html.escape(str(doc.get('requester', '')))}</div>"
                            f"<div class='request-meta'>{html.escape(str(doc.get('requested_at', '')))}</div>"
                            f"<span class='status-chip'>{html.escape(str(doc.get('status', '')))}</span>"
                            "</div>"
                        ),
                        unsafe_allow_html=True,
                    )
                    if st.session_state.user_role == "관리자":
                        now_str = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
                        approve_col, reject_col = st.columns(2)
                        if approve_col.button("승인", key=f"approve_doc_{doc.get('id')}", use_container_width=True, type="primary"):
                            doc["status"] = "승인"
                            doc["processed_at"] = now_str
                            doc["processed_by"] = st.session_state.user_name
                            linked = requests_by_id.get(doc.get("ref_request_id"))
                            if doc.get("doc_type") == "사용품의서":
                                usage_data = _load_usage_reports()
                                usage_by_id = {row.get("id"): row for row in usage_data.get("reports", [])}
                                usage_report = usage_by_id.get(doc.get("ref_usage_report_id"))
                                if usage_report:
                                    usage_report["status"] = "승인"
                                    usage_report["processed_at"] = now_str
                                    usage_report["processed_by"] = st.session_state.user_name
                                _save_usage_reports(usage_data)
                                if linked:
                                    linked["usage_report_status"] = "승인"
                            elif linked:
                                linked["status"] = "품의 확정"
                                linked["approved_at"] = now_str
                            _save_approvals(approvals)
                            _save_delegated_workplaces(wp_data)
                            st.rerun()
                        if reject_col.button("반려", key=f"reject_doc_{doc.get('id')}", use_container_width=True):
                            st.session_state[f"_reject_target_{doc.get('id')}"] = True
                        if st.session_state.get(f"_reject_target_{doc.get('id')}"):
                            reason = st.text_input("반려 사유", key=f"reject_reason_{doc.get('id')}")
                            if st.button("반려 확정", key=f"reject_confirm_{doc.get('id')}", use_container_width=True):
                                doc["status"] = "반려"
                                doc["processed_at"] = now_str
                                doc["processed_by"] = st.session_state.user_name
                                doc["reject_reason"] = reason.strip()
                                linked = requests_by_id.get(doc.get("ref_request_id"))
                                if doc.get("doc_type") == "사용품의서":
                                    usage_data = _load_usage_reports()
                                    usage_by_id = {row.get("id"): row for row in usage_data.get("reports", [])}
                                    usage_report = usage_by_id.get(doc.get("ref_usage_report_id"))
                                    if usage_report:
                                        usage_report["status"] = "반려"
                                        usage_report["processed_at"] = now_str
                                        usage_report["processed_by"] = st.session_state.user_name
                                        usage_report["reject_reason"] = reason.strip()
                                    _save_usage_reports(usage_data)
                                    if linked:
                                        linked["usage_report_status"] = "반려"
                                elif linked:
                                    linked["status"] = "반려"
                                    linked["reject_reason"] = reason.strip()
                                    linked["processed_at"] = now_str
                                _save_approvals(approvals)
                                _save_delegated_workplaces(wp_data)
                                st.session_state.pop(f"_reject_target_{doc.get('id')}", None)
                                st.rerun()
                    else:
                        st.caption("관리자 승인 대기 중입니다.")

    with done_tab:
        done_docs = [doc for doc in documents if doc.get("status") != "결재대기"]
        if not done_docs:
            st.info("처리된 결재 문서가 없습니다.")
        else:
            done_df = pd.DataFrame(done_docs)
            done_df["amount_won"] = done_df["amount"].apply(_format_won)
            for col in ["processed_at", "processed_by", "reject_reason"]:
                if col not in done_df.columns:
                    done_df[col] = ""
            done_view = done_df[
                ["requested_at", "title", "amount_won", "requester", "status", "processed_at", "processed_by", "reject_reason"]
            ].rename(
                columns={
                    "requested_at": "요청일시",
                    "title": "문서명",
                    "amount_won": "금액",
                    "requester": "요청자",
                    "status": "처리결과",
                    "processed_at": "처리일시",
                    "processed_by": "처리자",
                    "reject_reason": "반려사유",
                }
            )
            render_plain_html_table(done_view.sort_values("요청일시", ascending=False), center_align=True)

    with stats_tab:
        if documents:
            pending_n = len([doc for doc in documents if doc.get("status") == "결재대기"])
            approved_n = len([doc for doc in documents if doc.get("status") == "승인"])
            rejected_n = len([doc for doc in documents if doc.get("status") == "반려"])
            s1, s2, s3 = st.columns(3)
            s1.metric("결재대기", f"{pending_n:,}건")
            s2.metric("승인", f"{approved_n:,}건")
            s3.metric("반려", f"{rejected_n:,}건")

            month_counts = {}
            for doc in documents:
                key = _month_key(doc.get("requested_at"))
                if key:
                    month_counts[key] = month_counts.get(key, 0) + 1
            if month_counts:
                import plotly.graph_objects as go
                months = sorted(month_counts.keys())
                fig = go.Figure(go.Bar(x=months, y=[month_counts[m] for m in months]))
                fig.update_layout(**_chart_layout(height=280))
                st.plotly_chart(fig, use_container_width=True, theme=None)
        else:
            st.info("전자결재 이력이 없습니다.")


def show_approval_result():
    _render_page_chrome(["홈", "조회", "품의 결과", "품의 결과"])
    st.markdown("### 품의 결과")
    st.caption("전도금 요청의 품의(승인/반려) 처리 결과를 조회합니다.")

    data = _load_delegated_workplaces()
    requests = data.get("requests", [])
    workplaces = data.get("workplaces", [])

    if st.session_state.user_role != "관리자":
        my_ids = {row.get("id") for row in _my_workplaces(workplaces)}
        requests = [row for row in requests if row.get("workplace_id") in my_ids]

    processed = [row for row in requests if row.get("status") != "요청"]
    if not processed:
        st.info("처리된 품의 결과가 없습니다.")
        return

    def _processed_at(row):
        return row.get("approved_at") or row.get("paid_at") or row.get("processed_at") or ""

    requester_names = ["전체"] + sorted({row.get("requested_by", "") for row in processed if row.get("requested_by")})
    status_options = ["전체"] + sorted({row.get("status", "") for row in processed if row.get("status")})
    today = _current_kst().date()
    if "approval_result_start_date" not in st.session_state:
        st.session_state.approval_result_start_date = today - timedelta(days=6)
    if "approval_result_end_date" not in st.session_state:
        st.session_state.approval_result_end_date = today
    period_presets = {
        "오늘": (today, today),
        "어제": (today - timedelta(days=1), today - timedelta(days=1)),
        "1주일": (today - timedelta(days=6), today),
        "1개월": (today - timedelta(days=30), today),
    }

    with st.container():
        request_date_label_col, request_date_input_col, _ = st.columns([0.1, 0.54, 0.36])
        with request_date_label_col:
            st.markdown("**조회기간**")
        with request_date_input_col:
            preset_cols = st.columns([0.13, 0.13, 0.17, 0.17, 0.40])
            for index, (label, date_range) in enumerate(period_presets.items()):
                button_type = "primary" if label == st.session_state.get("approval_result_preset", "1주일") else "secondary"
                if preset_cols[index].button(label, key=f"approval_result_preset_{label}", type=button_type, use_container_width=True):
                    st.session_state.approval_result_preset = label
                    st.session_state.approval_result_start_date = date_range[0]
                    st.session_state.approval_result_end_date = date_range[1]
                    st.rerun()

            date_start_col, tilde_col, date_end_col, month_col, request_check_col, process_check_col = st.columns(
                [0.18, 0.03, 0.18, 0.17, 0.17, 0.17]
            )
            with date_start_col:
                request_start_date = st.date_input(
                    "조회 시작일",
                    key="approval_result_start_date",
                    label_visibility="collapsed",
                )
            with tilde_col:
                st.markdown("<div style='padding-top:8px;text-align:center;'>~</div>", unsafe_allow_html=True)
            with date_end_col:
                request_end_date = st.date_input(
                    "조회 종료일",
                    key="approval_result_end_date",
                    label_visibility="collapsed",
                )
            with month_col:
                month_options = ["월별 선택"] + [
                    (today.replace(day=1) - pd.DateOffset(months=idx)).strftime("%Y-%m")
                    for idx in range(12)
                ]
                selected_month = st.selectbox(
                    "월별 선택",
                    month_options,
                    key="approval_result_month",
                    label_visibility="collapsed",
                )
            with request_check_col:
                date_by_request = st.checkbox("요청일자", value=True, key="approval_result_date_by_request")
            with process_check_col:
                date_by_processed = st.checkbox("처리일자", value=False, key="approval_result_date_by_processed")

        requester_label_col, requester_input_col, _ = st.columns([0.1, 0.24, 0.66])
        with requester_label_col:
            st.markdown("**요청자**")
        with requester_input_col:
            selected_requester = st.selectbox(
                "요청자",
                requester_names,
                key="approval_result_requester",
                label_visibility="collapsed",
            )

        status_label_col, status_input_col, _ = st.columns([0.1, 0.24, 0.66])
        with status_label_col:
            st.markdown("**처리결과**")
        with status_input_col:
            selected_status = st.selectbox(
                "처리결과",
                status_options,
                key="approval_result_status",
                label_visibility="collapsed",
            )

    button_left, button_center, button_right = st.columns([0.25, 0.1, 0.65])
    with button_center:
        search_clicked = st.button("조회", key="approval_result_search", type="primary", use_container_width=True)

    if selected_month != "월별 선택":
        month_start = pd.to_datetime(f"{selected_month}-01").date()
        month_end = (pd.to_datetime(f"{selected_month}-01") + pd.offsets.MonthEnd(0)).date()
        request_start_date, request_end_date = month_start, month_end

    if search_clicked or "_approval_result_query" not in st.session_state:
        st.session_state["_approval_result_query"] = {
            "request_date_range": (request_start_date, request_end_date),
            "date_by_request": date_by_request,
            "date_by_processed": date_by_processed,
            "selected_requester": selected_requester,
            "selected_status": selected_status,
        }

    query = st.session_state.get("_approval_result_query", {})
    request_date_range = query.get("request_date_range", ())
    date_by_request = query.get("date_by_request", True)
    date_by_processed = query.get("date_by_processed", False)
    selected_requester = query.get("selected_requester", "전체")
    selected_status = query.get("selected_status", "전체")

    def _date_in_range(value, date_range):
        if not isinstance(date_range, tuple) or len(date_range) != 2:
            return True
        parsed = _parse_kst_datetime(value)
        if not parsed:
            return False
        start, end = date_range
        return start <= parsed.date() <= end

    approvals_data = _load_approvals()
    docs_by_request_id = {
        doc.get("ref_request_id"): doc
        for doc in approvals_data.get("documents", [])
        if doc.get("ref_request_id") is not None
    }

    filtered = processed
    if request_date_range and (date_by_request or date_by_processed):
        if date_by_request:
            filtered = [row for row in filtered if _date_in_range(row.get("requested_at"), request_date_range)]
        if date_by_processed:
            filtered = [row for row in filtered if _date_in_range(_processed_at(row), request_date_range)]
    if selected_requester != "전체":
        filtered = [row for row in filtered if row.get("requested_by") == selected_requester]
    if selected_status != "전체":
        filtered = [row for row in filtered if row.get("status") == selected_status]

    if not filtered:
        st.info("조건에 맞는 품의 결과가 없습니다.")
        return

    result_df = pd.DataFrame(filtered)
    result_df["request_amount_won"] = result_df["request_amount"].apply(_format_won)
    for col in ["approved_at", "paid_at", "processed_at", "reject_reason"]:
        if col not in result_df.columns:
            result_df[col] = ""
    result_df["처리일시"] = result_df.apply(_processed_at, axis=1)
    result_df["문서구분"] = result_df.apply(
        lambda row: "전도금 사용 결의 보고"
        if (docs_by_request_id.get(row.get("id"), {}).get("doc_type") == "사용품의서" or row.get("usage_report_id"))
        else "전도금 요청",
        axis=1,
    )
    type_col1, type_col2, _ = st.columns([0.13, 0.2, 0.67])
    with type_col1:
        show_advance_requests = st.checkbox("전도금 요청", value=True, key="approval_result_show_advance_requests")
    with type_col2:
        show_usage_reports = st.checkbox("전도금 사용 결의 보고", value=True, key="approval_result_show_usage_reports")

    visible_doc_types = []
    if show_advance_requests:
        visible_doc_types.append("전도금 요청")
    if show_usage_reports:
        visible_doc_types.append("전도금 사용 결의 보고")
    result_df = result_df[result_df["문서구분"].isin(visible_doc_types)]
    if result_df.empty:
        st.info("선택한 문서구분에 해당하는 품의 결과가 없습니다.")
        return

    result_view = result_df[
        ["문서구분", "requested_at", "workplace_name", "request_amount_won", "requested_by", "status", "처리일시", "reject_reason"]
    ].rename(
        columns={
            "requested_at": "요청일시",
            "workplace_name": "사업장명",
            "request_amount_won": "요청 금액",
            "requested_by": "요청자",
            "status": "처리결과",
            "reject_reason": "반려사유",
        }
    )
    render_plain_html_table(
        result_view.sort_values("요청일시", ascending=False),
        center_align=True,
        stretch=True,
        max_width="760px",
        border=False,
    )


def show_usage_report():
    st.markdown("### 전도금 사용 결의 보고")
    st.caption("계좌 거래내역을 기반으로 출금 내역의 사용 사유를 작성하여 본사에 보고합니다.")

    tab1, tab2 = st.tabs(["보고서 작성", "제출 내역"])

    with tab1:
        _render_usage_report_form()

    with tab2:
        _render_usage_report_history()


def _render_usage_report_form():
    """전도금 사용 결의 보고서 작성 폼"""
    st.markdown("#### 보고서 작성")

    # 폼 너비 제한
    st.markdown(
        """
        <style>
        [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {
            max-width: 1000px;
            margin: 0 auto;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # 계좌 선택
    bank_data = _load_bank_accounts()
    accounts = bank_data.get("accounts", [])

    if not accounts:
        st.warning("등록된 계좌가 없습니다. [계좌 관리]에서 계좌를 먼저 등록해주세요.")
        return

    wp_data = _load_delegated_workplaces()
    workplaces_by_id = {site.get("id"): site for site in wp_data.get("workplaces", [])}

    account_options = {
        f"{row.get('account_name')} ({row.get('bank_name')} {row.get('account_number')})": row
        for row in accounts
    }

    with st.container(border=False):
        # 계좌번호
        account_label_col, account_input_col = st.columns([0.12, 0.88])
        with account_label_col:
            st.markdown("**계좌번호 <span style='color:#008c78'>*</span>**", unsafe_allow_html=True)
        with account_input_col:
            account_select_col, _ = st.columns([0.48, 0.52])
            with account_select_col:
                selected_account_label = st.selectbox(
                    "계좌번호",
                    list(account_options.keys()),
                    key="usage_report_account",
                    label_visibility="collapsed"
                )
                selected_account = account_options[selected_account_label]

        # 조회기간
        today = _current_kst().date()
        if "usage_report_start_date" not in st.session_state:
            st.session_state.usage_report_start_date = today.replace(day=1)
        if "usage_report_end_date" not in st.session_state:
            st.session_state.usage_report_end_date = today

        period_presets = {
            "오늘": (today, today),
            "어제": (today - timedelta(days=1), today - timedelta(days=1)),
            "1주일": (today - timedelta(days=6), today),
            "1개월": (today - timedelta(days=30), today),
        }

        period_label_col, period_input_col = st.columns([0.12, 0.88])
        with period_label_col:
            st.markdown("**조회기간**")
        with period_input_col:
            preset_cols = st.columns([0.13, 0.13, 0.17, 0.17, 0.40])
            for index, (label, date_range) in enumerate(period_presets.items()):
                button_type = "primary" if label == st.session_state.get("usage_report_preset", "1주일") else "secondary"
                if preset_cols[index].button(label, key=f"usage_report_preset_{label}", type=button_type, use_container_width=True):
                    st.session_state.usage_report_preset = label
                    st.session_state.usage_report_start_date = date_range[0]
                    st.session_state.usage_report_end_date = date_range[1]
                    st.rerun()

            date_start_col, tilde_col, date_end_col, month_col, _ = st.columns([0.22, 0.03, 0.22, 0.18, 0.35])
            with date_start_col:
                start_date = st.date_input("시작일", key="usage_report_start_date", label_visibility="collapsed")
            with tilde_col:
                st.markdown("<div style='padding-top:8px;text-align:center;'>~</div>", unsafe_allow_html=True)
            with date_end_col:
                end_date = st.date_input("종료일", key="usage_report_end_date", label_visibility="collapsed")
            with month_col:
                month_options = ["월별 선택"] + [
                    (today.replace(day=1) - pd.DateOffset(months=idx)).strftime("%Y-%m")
                    for idx in range(12)
                ]
                selected_month = st.selectbox(
                    "월별 선택",
                    month_options,
                    key="usage_report_month",
                    label_visibility="collapsed",
                )

    # 조회 버튼
    button_left, button_center, button_right = st.columns([0.42, 0.16, 0.42])
    with button_center:
        if st.button("조회", key="usage_report_search", type="primary", use_container_width=True):
            if selected_month != "월별 선택":
                month_start = pd.to_datetime(f"{selected_month}-01").date()
                month_end = (pd.to_datetime(f"{selected_month}-01") + pd.offsets.MonthEnd(0)).date()
                st.session_state.usage_report_start_date = month_start
                st.session_state.usage_report_end_date = month_end

            st.session_state.withdrawal_query_done = True
            st.rerun()

    if not st.session_state.get("withdrawal_query_done"):
        st.info("계좌와 조회기간을 선택한 후 '조회' 버튼을 클릭해주세요.")
        return

    st.markdown("---")
    st.markdown("#### 출금 내역 및 사용 사유 입력")

    # 계좌 거래내역에서 출금 데이터 가져오기
    # 실제 거래내역이 있다면 여기서 필터링
    # 임시로 샘플 데이터 또는 balance_history 활용
    transactions = selected_account.get("transactions", [])

    # 조회기간에 해당하는 출금 내역만 필터링
    query_start = st.session_state.usage_report_start_date
    query_end = st.session_state.usage_report_end_date

    withdrawals = [
        t for t in transactions
        if t.get("type") == "출금" and query_start <= pd.to_datetime(t.get("date")).date() <= query_end
    ]

    if not withdrawals:
        st.warning(
            "선택한 기간에 출금 내역이 없습니다.\n\n"
            "💡 참고: 거래내역은 [계좌 잔고 확인] 메뉴의 거래내역조회에서 확인할 수 있습니다."
        )
        st.info(
            "**임시 입력 모드**\n\n"
            "아래 버튼을 클릭하여 출금 내역을 직접 입력할 수 있습니다."
        )

    # 출금 항목별 사용 사유 입력
    if "usage_report_items" not in st.session_state:
        st.session_state.usage_report_items = {}

    # 거래내역이 있으면 표시하고 사용 사유 입력 받기
    if withdrawals:
        st.markdown(f"**조회된 출금 내역: {len(withdrawals)}건**")

        for idx, withdrawal in enumerate(withdrawals):
            with st.expander(
                f"📤 {withdrawal.get('date')} | {_format_won(withdrawal.get('amount'))} | {withdrawal.get('recipient', '-')}"
            ):
                trans_id = f"{withdrawal.get('date')}_{withdrawal.get('amount')}"

                usage_reason = st.text_area(
                    "사용 사유",
                    key=f"reason_{idx}",
                    placeholder="출금 사용 목적을 상세히 작성해주세요.",
                    value=st.session_state.usage_report_items.get(trans_id, {}).get("reason", "")
                )

                usage_file = st.file_uploader(
                    "증빙서류 첨부 (선택)",
                    key=f"file_{idx}",
                    type=["pdf", "jpg", "jpeg", "png"]
                )

                if st.button("저장", key=f"save_{idx}", type="primary"):
                    file_name = usage_file.name if usage_file else ""
                    st.session_state.usage_report_items[trans_id] = {
                        "date": withdrawal.get("date"),
                        "amount": withdrawal.get("amount"),
                        "recipient": withdrawal.get("recipient", ""),
                        "reason": usage_reason.strip(),
                        "file_name": file_name,
                    }
                    st.success("저장되었습니다.")

    # 수동 입력 옵션
    with st.expander("➕ 출금 내역 수동 추가"):
        with st.form("add_withdrawal_manual"):
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                manual_date = st.date_input("출금일자")
            with col_b:
                manual_amount = st.number_input("출금금액", min_value=0, step=1000, format="%d")
            with col_c:
                manual_recipient = st.text_input("거래처")

            manual_reason = st.text_area("사용 사유", placeholder="출금 사용 목적을 상세히 작성해주세요.")
            manual_file = st.file_uploader("증빙서류 첨부 (선택)", type=["pdf", "jpg", "jpeg", "png"])

            if st.form_submit_button("추가", type="primary", use_container_width=True):
                if not manual_reason.strip():
                    st.warning("사용 사유를 입력해주세요.")
                else:
                    trans_id = f"{manual_date.strftime('%Y-%m-%d')}_{manual_amount}_manual"
                    file_name = manual_file.name if manual_file else ""

                    st.session_state.usage_report_items[trans_id] = {
                        "date": manual_date.strftime("%Y-%m-%d"),
                        "amount": int(manual_amount),
                        "recipient": manual_recipient.strip(),
                        "reason": manual_reason.strip(),
                        "file_name": file_name,
                    }
                    st.success("출금 항목이 추가되었습니다.")
                    st.rerun()

    st.markdown("---")

    # 입력된 항목 표시
    if st.session_state.usage_report_items:
        st.markdown("#### 작성된 사용 사유")

        for trans_id, item in st.session_state.usage_report_items.items():
            with st.container():
                col_info, col_del = st.columns([0.9, 0.1])
                with col_info:
                    file_badge = f" 📎 {item['file_name']}" if item['file_name'] else ""
                    st.markdown(
                        f"""
                        **{item['date']}** | {_format_won(item['amount'])} | {item['recipient']}
                        사유: {item['reason']}{file_badge}
                        """
                    )
                with col_del:
                    if st.button("🗑️", key=f"del_{trans_id}"):
                        del st.session_state.usage_report_items[trans_id]
                        st.rerun()
                st.markdown("---")

        # 보고서 제출
        total_amount = sum(item['amount'] for item in st.session_state.usage_report_items.values())
        st.markdown(f"**총 출금액:** {_format_won(total_amount)}")
        st.markdown(f"**작성 완료 건수:** {len(st.session_state.usage_report_items)}건")

        if st.button("보고서 제출", type="primary", use_container_width=True):
            # 사용 사유가 입력되지 않은 항목 확인
            incomplete = [item for item in st.session_state.usage_report_items.values() if not item.get('reason', '').strip()]

            if incomplete:
                st.warning(f"사용 사유가 입력되지 않은 항목이 {len(incomplete)}건 있습니다. 모든 항목의 사용 사유를 입력해주세요.")
            else:
                # 보고서 저장
                usage_data = _load_usage_reports()
                if "reports" not in usage_data:
                    usage_data["reports"] = []

                workplace_name = workplaces_by_id.get(selected_account.get("linked_workplace_id"), {}).get("workplace_name", "")

                query_start = st.session_state.usage_report_start_date
                query_end = st.session_state.usage_report_end_date

                usage_data["reports"].append({
                    "id": int(time.time() * 1000),
                    "workplace_id": selected_account.get("linked_workplace_id"),
                    "workplace_name": workplace_name,
                    "account_id": selected_account.get("id"),
                    "report_period": f"{query_start.strftime('%Y-%m-%d')} ~ {query_end.strftime('%Y-%m-%d')}",
                    "total_amount": total_amount,
                    "items": list(st.session_state.usage_report_items.values()),
                    "status": "제출완료",
                    "requested_at": _current_kst().strftime("%Y-%m-%d %H:%M:%S"),
                })

                _save_usage_reports(usage_data)
                st.session_state.usage_report_items = {}
                st.session_state.withdrawal_query_done = False
                st.success("전도금 사용 결의 보고서가 제출되었습니다.")
                st.rerun()
    else:
        st.info("출금 내역의 사용 사유를 입력해주세요.")


def _render_usage_report_history():
    """제출된 전도금 사용 결의 보고 내역"""
    wp_data = _load_delegated_workplaces()
    workplaces = _my_workplaces(wp_data.get("workplaces", []))
    my_ids = {row.get("id") for row in workplaces}

    usage_data = _load_usage_reports()
    usage_reports = usage_data.get("reports", [])
    my_reports = [row for row in usage_reports if row.get("workplace_id") in my_ids]

    if not my_reports:
        st.info("제출된 전도금 사용 결의 보고가 없습니다.")
        return

    _workplace_dashboard_css()
    cols = st.columns(3)
    for index, row in enumerate(sorted(my_reports, key=lambda item: item.get("requested_at", ""), reverse=True)):
        with cols[index % 3]:
            reject_line = ""
            if row.get("status") == "반려" and row.get("reject_reason"):
                reject_line = f"<div class='request-meta'>반려 사유: {html.escape(str(row.get('reject_reason')))}</div>"

            items_count = len(row.get("items", []))
            report_period = row.get('report_period', row.get('report_month', ''))  # 하위 호환성
            st.markdown(
                (
                    "<div class='request-card'>"
                    f"<div class='request-title'>{html.escape(str(row.get('workplace_name', '')))}</div>"
                    f"<div class='request-meta'>{html.escape(report_period)} · {items_count}건</div>"
                    f"<div class='request-meta'>{html.escape(_format_won(row.get('total_amount')))} · {html.escape(str(row.get('requested_at', '')))}</div>"
                    f"{reject_line}"
                    f"<span class='status-chip'>{html.escape(str(row.get('status', '')))}</span>"
                    "</div>"
                ),
                unsafe_allow_html=True,
            )


def show_approval_line_settings():
    st.markdown("### 결재선 설정")
    st.caption("결재 문서의 승인자와 결재 단계를 관리합니다.")

    # 결재선 데이터 로드 (임시로 세션 스테이트 사용)
    if "approval_lines" not in st.session_state:
        st.session_state.approval_lines = []

    approval_lines = st.session_state.approval_lines

    # 등록된 결재선 표시
    if approval_lines:
        st.markdown("#### 등록된 결재선")
        line_df = pd.DataFrame(approval_lines)
        line_view = line_df[["line_name", "approver1", "approver2", "approver3"]].rename(
            columns={
                "line_name": "결재선명",
                "approver1": "1차 승인자",
                "approver2": "2차 승인자",
                "approver3": "3차 승인자",
            }
        )
        st.dataframe(line_view, use_container_width=True, hide_index=True)
    else:
        st.info("등록된 결재선이 없습니다.")

    # 결재선 추가 폼
    if st.session_state.get("_show_approval_line_add_form"):
        with st.form("approval_line_add_form"):
            line_name = st.text_input("결재선명", placeholder="예: 전도금 결재선")
            col1, col2, col3 = st.columns(3)
            approver1 = col1.text_input("1차 승인자", placeholder="필수")
            approver2 = col2.text_input("2차 승인자", placeholder="선택")
            approver3 = col3.text_input("3차 승인자", placeholder="선택")

            col_submit, col_cancel = st.columns(2)
            submitted = col_submit.form_submit_button("등록", type="primary", use_container_width=True)
            cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

        if submitted:
            if not line_name.strip() or not approver1.strip():
                st.warning("결재선명과 1차 승인자는 필수입니다.")
            else:
                approval_lines.append({
                    "id": int(time.time() * 1000),
                    "line_name": line_name.strip(),
                    "approver1": approver1.strip(),
                    "approver2": approver2.strip(),
                    "approver3": approver3.strip(),
                })
                st.session_state["_show_approval_line_add_form"] = False
                st.success("결재선이 등록되었습니다.")
                st.rerun()
        if cancelled:
            st.session_state["_show_approval_line_add_form"] = False
            st.rerun()

    # 결재선 삭제
    elif st.session_state.get("_show_approval_line_delete_select"):
        if approval_lines:
            delete_options = {line.get("line_name"): line.get("id") for line in approval_lines}
            target_label = st.selectbox("삭제 대상 결재선", list(delete_options.keys()), key="approval_line_delete_target")
            target_id = delete_options[target_label]

            col_confirm, col_cancel = st.columns(2)
            if col_confirm.button("삭제 확인", type="primary", use_container_width=True):
                st.session_state.approval_lines = [line for line in approval_lines if line.get("id") != target_id]
                st.session_state["_show_approval_line_delete_select"] = False
                st.success("결재선이 삭제되었습니다.")
                st.rerun()
            if col_cancel.button("취소", use_container_width=True):
                st.session_state["_show_approval_line_delete_select"] = False
                st.rerun()

    # 버튼들
    else:
        col_add, col_delete, _ = st.columns([20, 20, 160])
        if col_add.button("+ 결재선 추가", key="show_approval_line_add_form_btn", use_container_width=True):
            st.session_state["_show_approval_line_add_form"] = True
            st.rerun()

        if col_delete.button("삭제", key="show_approval_line_delete_btn", use_container_width=True, disabled=not approval_lines):
            st.session_state["_show_approval_line_delete_select"] = True
            st.rerun()


def show_company_profile():
    st.markdown("### 회사 관리")
    st.caption("당사(우리 회사)의 기본 정보를 관리합니다.")

    profile = _load_company_profile()
    profile.setdefault("withdrawal_accounts", [])

    # 등록된 회사 정보 표시
    if profile.get("name"):
        st.markdown("#### 등록된 회사 정보")
        info_data = {
            "회사명": profile.get("name", "-"),
            "사업자번호": profile.get("business_number", "-"),
            "대표자": profile.get("ceo_name", "-"),
            "연락처": profile.get("contact", "-"),
            "주소": profile.get("address", "-"),
            "비고": profile.get("memo", "-"),
        }
        info_df = pd.DataFrame([info_data]).T.reset_index()
        info_df.columns = ["항목", "내용"]
        render_plain_html_table(info_df, center_align=True)
        if profile.get("updated_at"):
            st.caption(f"최종 수정: {profile.get('updated_at')}")
    else:
        st.info("등록된 회사 정보가 없습니다.")

    st.markdown("#### 출금계좌")
    withdrawal_accounts = profile.get("withdrawal_accounts", [])
    withdrawal_accounts.sort(key=lambda row: (not row.get("is_main", False), row.get("created_at", "")))

    if withdrawal_accounts:
        account_df = pd.DataFrame(withdrawal_accounts)
        for col in ["account_name", "bank_name", "account_number", "holder_name", "is_main", "memo"]:
            if col not in account_df.columns:
                account_df[col] = "" if col != "is_main" else False
        account_df["main_label"] = account_df["is_main"].apply(lambda value: "메인" if value else "")
        account_view = account_df[
            ["main_label", "account_name", "bank_name", "account_number", "holder_name", "memo"]
        ].rename(
            columns={
                "main_label": "메인 출금계좌",
                "account_name": "계좌명",
                "bank_name": "은행명",
                "account_number": "계좌번호",
                "holder_name": "예금주",
                "memo": "비고",
            }
        ).reset_index(drop=True)
        account_view.insert(0, "순번", range(1, len(account_view) + 1))
        render_plain_html_table(account_view, center_align=True)
    else:
        st.info("등록된 출금계좌가 없습니다.")

    if st.session_state.get("_show_withdrawal_add_form"):
        with st.form("withdrawal_account_add_form", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            account_name = col_a.text_input("계좌명", placeholder="예: 본사 운영계좌")
            bank_name = col_b.text_input("은행명", placeholder="예: 국민은행")
            col_c, col_d = st.columns(2)
            account_number = col_c.text_input("계좌번호")
            holder_name = col_d.text_input("예금주", value=profile.get("name", ""))
            is_main = st.checkbox("메인 출금계좌로 설정", value=not bool(withdrawal_accounts))
            memo = st.text_input("비고")
            col_submit, col_cancel = st.columns(2)
            submitted = col_submit.form_submit_button("등록", type="primary", use_container_width=True)
            cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

        if submitted:
            if not account_name.strip() or not bank_name.strip() or not account_number.strip():
                st.warning("계좌명, 은행명, 계좌번호를 입력해주세요.")
            else:
                if is_main:
                    for row in withdrawal_accounts:
                        row["is_main"] = False
                withdrawal_accounts.append(
                    {
                        "id": int(time.time() * 1000),
                        "account_name": account_name.strip(),
                        "bank_name": bank_name.strip(),
                        "account_number": account_number.strip(),
                        "holder_name": holder_name.strip(),
                        "is_main": bool(is_main),
                        "memo": memo.strip(),
                        "created_at": _current_kst().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                profile["withdrawal_accounts"] = withdrawal_accounts
                profile["updated_at"] = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
                _save_company_profile(profile)
                st.session_state["_show_withdrawal_add_form"] = False
                st.success("출금계좌가 등록되었습니다.")
                st.rerun()
        if cancelled:
            st.session_state["_show_withdrawal_add_form"] = False
            st.rerun()

    elif st.session_state.get("_show_withdrawal_manage_form"):
        if withdrawal_accounts:
            account_options = {
                f"{'[메인] ' if row.get('is_main') else ''}{row.get('account_name', '')} ({row.get('bank_name', '')} {row.get('account_number', '')})": row.get("id")
                for row in withdrawal_accounts
            }
            selected_label = st.selectbox("대상 출금계좌", list(account_options.keys()), key="withdrawal_manage_target")
            selected_id = account_options[selected_label]
            target = next((row for row in withdrawal_accounts if row.get("id") == selected_id), None)

            col_main, col_delete, col_cancel = st.columns(3)
            if col_main.button("메인 출금계좌 설정", type="primary", use_container_width=True):
                for row in withdrawal_accounts:
                    row["is_main"] = row.get("id") == selected_id
                profile["withdrawal_accounts"] = withdrawal_accounts
                profile["updated_at"] = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
                _save_company_profile(profile)
                st.session_state["_show_withdrawal_manage_form"] = False
                st.success("메인 출금계좌가 설정되었습니다.")
                st.rerun()
            if col_delete.button("삭제", use_container_width=True, disabled=not target):
                profile["withdrawal_accounts"] = [row for row in withdrawal_accounts if row.get("id") != selected_id]
                if profile["withdrawal_accounts"] and not any(row.get("is_main") for row in profile["withdrawal_accounts"]):
                    profile["withdrawal_accounts"][0]["is_main"] = True
                profile["updated_at"] = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
                _save_company_profile(profile)
                st.session_state["_show_withdrawal_manage_form"] = False
                st.success("출금계좌가 삭제되었습니다.")
                st.rerun()
            if col_cancel.button("취소", use_container_width=True):
                st.session_state["_show_withdrawal_manage_form"] = False
                st.rerun()

    else:
        col_add, col_manage, _ = st.columns([0.18, 0.22, 0.6])
        if col_add.button("+ 출금계좌 등록", key="show_withdrawal_add_btn", use_container_width=True):
            st.session_state["_show_withdrawal_add_form"] = True
            st.rerun()
        if col_manage.button("메인/삭제 관리", key="show_withdrawal_manage_btn", use_container_width=True, disabled=not withdrawal_accounts):
            st.session_state["_show_withdrawal_manage_form"] = True
            st.rerun()

    # 수정 폼
    if st.session_state.get("_show_company_edit_form"):
        with st.form("company_profile_form"):
            col_a, col_b = st.columns(2)
            name = col_a.text_input("회사명", value=profile.get("name", ""), placeholder="예: (주)회사명")
            business_number = col_b.text_input("사업자번호", value=profile.get("business_number", ""), placeholder="예: 123-45-67890")
            col_c, col_d = st.columns(2)
            ceo_name = col_c.text_input("대표자", value=profile.get("ceo_name", ""), placeholder="예: 홍길동")
            contact = col_d.text_input("연락처", value=profile.get("contact", ""), placeholder="예: 02-1234-5678")
            address = st.text_input("주소", value=profile.get("address", ""), placeholder="예: 서울특별시 강남구 ...")
            memo = st.text_area("비고", value=profile.get("memo", ""), placeholder="회사 관련 특이사항")
            col_submit, col_cancel = st.columns(2)
            submitted = col_submit.form_submit_button("저장", type="primary", use_container_width=True)
            cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

        if submitted:
            if not name.strip():
                st.warning("회사명을 입력해주세요.")
            else:
                profile.update(
                    {
                        "name": name.strip(),
                        "business_number": business_number.strip(),
                        "ceo_name": ceo_name.strip(),
                        "address": address.strip(),
                        "contact": contact.strip(),
                        "memo": memo.strip(),
                        "updated_at": _current_kst().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                _save_company_profile(profile)
                st.session_state["_show_company_edit_form"] = False
                st.success("회사 정보가 저장되었습니다.")
                st.rerun()
        if cancelled:
            st.session_state["_show_company_edit_form"] = False
            st.rerun()

    # 수정 버튼
    else:
        if st.button("수정", key="show_company_edit_btn", use_container_width=False):
            st.session_state["_show_company_edit_form"] = True
            st.rerun()


def show_workplace_admin():
    tab1, tab2, tab3 = st.tabs(["사업장 정보 관리", "계좌 관리", "사용자 관리"])
    with tab1:
        _render_workplace_info_admin()
    with tab2:
        _render_bank_account_management()
    with tab3:
        _render_staff_admin()


def _render_bank_account_management():
    st.caption("위탁 사업장의 계좌 정보를 관리합니다.")

    data = _load_bank_accounts()
    accounts = data.get("accounts", [])

    wp_data = _load_delegated_workplaces()
    workplaces = wp_data.get("workplaces", [])
    workplaces_by_id = {site.get("id"): site for site in workplaces}

    # 등록된 계좌 표시
    if accounts:
        st.markdown("#### 등록된 계좌")
        account_df = pd.DataFrame(accounts)
        account_df["balance_won"] = account_df["balance"].apply(_format_won)
        for col in ["holder_name", "balance_updated_at", "memo"]:
            if col not in account_df.columns:
                account_df[col] = ""
        account_df["회사"] = [_account_company_label(row, workplaces_by_id) for row in accounts]
        account_view = account_df[
            ["회사", "account_name", "bank_name", "account_number", "holder_name", "balance_won", "balance_updated_at", "memo"]
        ].rename(
            columns={
                "account_name": "계좌명",
                "bank_name": "은행명",
                "account_number": "계좌번호",
                "holder_name": "예금주",
                "balance_won": "현재잔고",
                "balance_updated_at": "최종업데이트",
                "memo": "메모",
            }
        )
        render_plain_html_table(account_view, center_align=True)
    else:
        st.info("등록된 계좌가 없습니다.")

    # 계좌 수정 폼
    if st.session_state.get("_show_account_edit_form"):
        edit_target_id = st.session_state.get("_account_edit_target_id")
        target = next((row for row in accounts if row.get("id") == edit_target_id), None)
        if target is None:
            st.session_state["_show_account_edit_form"] = False
            st.rerun()

        company_options = {site.get("workplace_name", ""): site for site in workplaces}
        current_workplace_id = target.get("linked_workplace_id")
        current_workplace = workplaces_by_id.get(current_workplace_id, {})
        current_workplace_name = current_workplace.get("workplace_name", "")

        with st.form("bank_account_edit_form", clear_on_submit=False):
            company_label = st.selectbox("회사 선택", list(company_options.keys()),
                                        index=list(company_options.keys()).index(current_workplace_name) if current_workplace_name in company_options.keys() else 0)
            account_name = st.text_input("계좌명", value=target.get("account_name", ""))
            bank_name = st.text_input("은행명", value=target.get("bank_name", ""))
            account_number = st.text_input("계좌번호", value=target.get("account_number", ""))
            holder_name = st.text_input("예금주", value=target.get("holder_name", ""))
            memo = st.text_area("메모", value=target.get("memo", ""))
            col_submit, col_cancel = st.columns(2)
            submitted = col_submit.form_submit_button("수정 완료", type="primary", use_container_width=True)
            cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

        if submitted:
            if not account_name.strip() or not account_number.strip():
                st.warning("계좌명과 계좌번호를 입력해주세요.")
            else:
                site = company_options[company_label]
                target["account_name"] = account_name.strip()
                target["bank_name"] = bank_name.strip()
                target["account_number"] = account_number.strip()
                target["holder_name"] = holder_name.strip()
                target["linked_workplace_id"] = site.get("id")
                target["memo"] = memo.strip()
                _save_bank_accounts(data)
                st.session_state["_show_account_edit_form"] = False
                st.success("계좌 정보가 수정되었습니다.")
                st.rerun()
        if cancelled:
            st.session_state["_show_account_edit_form"] = False
            st.rerun()

    # 계좌 추가 폼
    elif st.session_state.get("_show_account_add_form"):
        if not workplaces:
            st.warning("[사업장 정보 관리]에서 사업장을 먼저 등록해주세요.")
            st.session_state["_show_account_add_form"] = False
            st.rerun()
        else:
            company_options = {site.get("workplace_name", ""): site for site in workplaces}
            with st.form("bank_account_form", clear_on_submit=True):
                col_a, col_b = st.columns(2)
                company_label = col_a.selectbox("회사 선택", list(company_options.keys()))
                account_name = col_b.text_input("계좌명", placeholder="예: 본사 지급계좌")
                col_c, col_d, col_e = st.columns([1, 1, 0.4])
                bank_name = col_c.text_input("은행명", placeholder="예: 하나은행")
                account_number = col_d.text_input("계좌번호", placeholder="예: 123-456789-01234")
                col_e.markdown("<div style='height:1.8em'></div>", unsafe_allow_html=True)
                holder_lookup = col_e.form_submit_button("예금주조회", use_container_width=True)
                holder_name = st.text_input("예금주", placeholder="예: (주)회사명")
                memo = st.text_area("메모", placeholder="계좌 관련 특이사항")
                col_submit, col_cancel = st.columns(2)
                submitted = col_submit.form_submit_button("등록", type="primary", use_container_width=True)
                cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

            if holder_lookup:
                st.info("예금주조회 기능은 준비 중입니다. (API 연동 예정)")
            if submitted:
                if not account_name.strip() or not account_number.strip():
                    st.warning("계좌명과 계좌번호를 입력해주세요.")
                else:
                    site = company_options[company_label]
                    accounts.append(
                        {
                            "id": int(time.time() * 1000),
                            "account_type": "위탁사업장",
                            "account_name": account_name.strip(),
                            "bank_name": bank_name.strip(),
                            "account_number": account_number.strip(),
                            "holder_name": holder_name.strip(),
                            "linked_workplace_id": site.get("id"),
                            "balance": 0,
                            "balance_updated_at": "",
                            "balance_history": [],
                            "memo": memo.strip(),
                            "created_at": _current_kst().strftime("%Y-%m-%d %H:%M:%S"),
                        }
                    )
                    _save_bank_accounts(data)
                    st.session_state["_show_account_add_form"] = False
                    st.success("계좌가 등록되었습니다.")
                    st.rerun()
            if cancelled:
                st.session_state["_show_account_add_form"] = False
                st.rerun()

    # 삭제 선택
    elif st.session_state.get("_show_account_delete_select"):
        if accounts:
            delete_options = {f"{row.get('account_name')} ({row.get('bank_name')} {row.get('account_number')})": row.get("id") for row in accounts}
            target_label = st.selectbox("삭제 대상 계좌", list(delete_options.keys()), key="bank_account_delete_target")
            target_id = delete_options[target_label]

            col_confirm, col_cancel = st.columns(2)
            if col_confirm.button("삭제 확인", type="primary", use_container_width=True):
                data["accounts"] = [row for row in accounts if row.get("id") != target_id]
                _save_bank_accounts(data)
                st.session_state["_show_account_delete_select"] = False
                st.success("계좌가 삭제되었습니다.")
                st.rerun()
            if col_cancel.button("취소", use_container_width=True):
                st.session_state["_show_account_delete_select"] = False
                st.rerun()

    # 수정 선택
    elif st.session_state.get("_show_account_edit_select"):
        if accounts:
            edit_options = {f"{row.get('account_name')} ({row.get('bank_name')} {row.get('account_number')})": row.get("id") for row in accounts}
            target_label = st.selectbox("수정 대상 계좌", list(edit_options.keys()), key="bank_account_edit_target")
            target_id = edit_options[target_label]

            col_confirm, col_cancel = st.columns(2)
            if col_confirm.button("수정 진행", type="primary", use_container_width=True):
                st.session_state["_show_account_edit_select"] = False
                st.session_state["_show_account_edit_form"] = True
                st.session_state["_account_edit_target_id"] = target_id
                st.rerun()
            if col_cancel.button("취소", use_container_width=True):
                st.session_state["_show_account_edit_select"] = False
                st.rerun()

    # 버튼들
    else:
        col_add, col_edit, col_delete, _ = st.columns([20, 20, 20, 140])
        if col_add.button("+ 계좌 추가", key="show_account_add_form_btn", use_container_width=True):
            st.session_state["_show_account_add_form"] = True
            st.rerun()

        if col_edit.button("수정", key="show_account_edit_form_btn", use_container_width=True, disabled=not accounts):
            st.session_state["_show_account_edit_select"] = True
            st.rerun()

        if col_delete.button("삭제", key="show_account_delete_btn", use_container_width=True, disabled=not accounts):
            st.session_state["_show_account_delete_select"] = True
            st.rerun()


def _render_transfer_confirmed_section(confirmed_rows, accounts_by_workplace_id):
    st.markdown("#### 이체 자료 확정 건")
    if not confirmed_rows:
        st.info("이체 자료 확정 건이 없습니다.")
        return
    st.caption("이체 자료가 확정되어 [지급 결과 확인]에서 처리 대기 중인 건입니다.")
    confirmed_view = pd.DataFrame(
        [
            {
                "사업장명": row.get("workplace_name", ""),
                "입금은행": accounts_by_workplace_id.get(row.get("workplace_id"), {}).get("bank_name", ""),
                "입금계좌번호": accounts_by_workplace_id.get(row.get("workplace_id"), {}).get("account_number", ""),
                "입금액": _format_won(row.get("request_amount")),
                "이체자료확정일시": row.get("transfer_file_generated_at", ""),
            }
            for row in sorted(confirmed_rows, key=lambda item: item.get("transfer_file_generated_at", ""), reverse=True)
        ]
    )
    render_plain_html_table(confirmed_view, center_align=True, stretch=True, max_width="760px", border=False)


def show_transfer_file_generation():
    _render_page_chrome(["홈", "지급관리", "이체 자료 확정", "이체 자료 확정"])
    st.markdown("### 이체 자료 확정")
    st.caption("품의가 확정된 전도금 요청을 선택하여 이체 자료를 확정하고 엑셀로 다운로드합니다.")

    # 다크모드 스타일
    st.markdown(
        """
        <style>
        /* 이체 대상 선택 data_editor 다크모드 */
        body:has(#pms-d:checked) [data-testid="stDataEditor"],
        body:has(#pms-d:checked) [data-testid="stDataEditor"] > div,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="grid"] {
            background-color: #1e1e2e !important;
            color: #ffffff !important;
            border-color: #45475a !important;
        }
        body:has(#pms-d:checked) [data-testid="stDataEditor"] th,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="columnheader"] {
            background-color: #252535 !important;
            color: #ffffff !important;
            border-color: #45475a !important;
        }
        body:has(#pms-d:checked) [data-testid="stDataEditor"] td,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="gridcell"] {
            background-color: #1e1e2e !important;
            color: #ffffff !important;
            border-color: #313244 !important;
        }
        body:has(#pms-d:checked) [data-testid="stDataEditor"] input {
            background-color: #252535 !important;
            color: #ffffff !important;
            border-color: #45475a !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    wp_data = _load_delegated_workplaces()
    requests = wp_data.get("requests", [])

    bank_data = _load_bank_accounts()
    accounts_by_workplace_id = {
        row.get("linked_workplace_id"): row
        for row in bank_data.get("accounts", [])
        if row.get("linked_workplace_id")
    }

    company_profile = _load_company_profile()
    withdrawal_accounts = company_profile.get("withdrawal_accounts", [])
    withdrawal_accounts = sorted(
        withdrawal_accounts,
        key=lambda row: (not row.get("is_main", False), row.get("created_at", "")),
    )

    targets = [row for row in requests if row.get("status") == "품의 확정"]
    confirmed_rows = [row for row in requests if row.get("status") == "이체 대상"]

    if not targets:
        st.info("이체 자료를 생성할 품의 확정 건이 없습니다.")
        _render_transfer_confirmed_section(confirmed_rows, accounts_by_workplace_id)
        return

    if not withdrawal_accounts:
        st.warning("[회사 관리]에서 출금계좌를 먼저 등록해주세요.")
        _render_transfer_confirmed_section(confirmed_rows, accounts_by_workplace_id)
        return

    source_options = {
        f"{'[메인] ' if row.get('is_main') else ''}{row.get('account_name')} ({row.get('bank_name')} {row.get('account_number')})": row
        for row in withdrawal_accounts
    }
    source_label = st.selectbox("출금 계좌", list(source_options.keys()))

    st.markdown("#### 이체 대상 선택")
    lookup_map = st.session_state.get("_transfer_holder_lookup", {})
    target_rows = []
    lookup_inputs = {}
    for row in sorted(targets, key=lambda item: item.get("requested_at", "")):
        account = accounts_by_workplace_id.get(row.get("workplace_id"), {})
        workplace_name = row.get("workplace_name", "")
        passbook_label = f"{workplace_name} 전도금"
        expected_holder = account.get("holder_name", "") or workplace_name
        lookup_inputs[row.get("id")] = (account.get("bank_name", ""), account.get("account_number", ""), expected_holder)
        target_rows.append(
            {
                "id": row.get("id"),
                "사업장명": workplace_name,
                "선택": True,
                "입금은행": account.get("bank_name", ""),
                "입금계좌번호": account.get("account_number", ""),
                "입금액": _format_won(row.get("request_amount")),
                "예상예금주": expected_holder,
                "조회한예금주": lookup_map.get(row.get("id"), ""),
                "입금통장표시": passbook_label,
                "출금통장표시": passbook_label,
            }
        )

    target_df = pd.DataFrame(target_rows)
    column_order = [
        "id", "선택", "사업장명", "입금은행", "입금계좌번호", "입금액",
        "예상예금주", "조회한예금주", "입금통장표시", "출금통장표시",
    ]
    edited_df = st.data_editor(
        target_df[column_order],
        column_config={
            "id": None,
            "선택": st.column_config.CheckboxColumn("선택", width="small"),
            "사업장명": st.column_config.TextColumn("사업장명", disabled=True, width="small"),
            "입금은행": st.column_config.TextColumn("입금은행", disabled=True, width="small"),
            "입금계좌번호": st.column_config.TextColumn("입금계좌번호", disabled=True, width="small"),
            "입금액": st.column_config.TextColumn("입금액", disabled=True, width="small"),
            "예상예금주": st.column_config.TextColumn("예상예금주", disabled=True, width="small"),
            "조회한예금주": st.column_config.TextColumn("조회한예금주", disabled=True, width="small"),
            "입금통장표시": st.column_config.TextColumn("입금통장표시", width="small"),
            "출금통장표시": st.column_config.TextColumn("출금통장표시", width="small"),
        },
        hide_index=True,
        use_container_width=True,
        key="transfer_target_editor",
    )
    selected_ids = edited_df[edited_df["선택"]]["id"].tolist()

    col_spacer, col_holder_lookup = st.columns([165, 35])
    if col_holder_lookup.button("예금주조회", key="transfer_holder_lookup_btn", use_container_width=True):
        new_lookup = {}
        for row_id, (bank_name, account_number, expected_holder) in lookup_inputs.items():
            holder = _lookup_account_holder_name(bank_name, account_number, expected_holder)
            if holder:
                new_lookup[row_id] = holder
        st.session_state["_transfer_holder_lookup"] = new_lookup
        if new_lookup:
            st.success(f"{len(new_lookup)}건의 예금주 정보를 조회했습니다.")
        else:
            st.warning("조회 가능한 입금계좌 정보가 없습니다. [계좌 관리]에서 계좌를 먼저 등록해주세요.")
        st.rerun()

    lookup_map = st.session_state.get("_transfer_holder_lookup", {})
    lookup_complete = bool(selected_ids) and all(row_id in lookup_map for row_id in selected_ids)

    source = source_options[source_label]
    edited_by_id = {item["id"]: item for item in edited_df.to_dict("records")}
    transfer_rows = []
    for row in targets:
        if row.get("id") not in selected_ids:
            continue
        account = accounts_by_workplace_id.get(row.get("workplace_id"), {})
        edited_row = edited_by_id.get(row.get("id"), {})
        transfer_rows.append(
            {
                "출금은행": source.get("bank_name", ""),
                "출금계좌번호": source.get("account_number", ""),
                "출금통장표시": edited_row.get("출금통장표시", ""),
                "입금은행": account.get("bank_name", ""),
                "입금계좌번호": account.get("account_number", ""),
                "입금액": row.get("request_amount", 0),
                "예상예금주": lookup_map.get(row.get("id")) or account.get("holder_name", "") or row.get("workplace_name", ""),
                "입금통장표시": edited_row.get("입금통장표시", ""),
                "사업장명": row.get("workplace_name", ""),
            }
        )

    col_confirm, col_download = st.columns(2)
    confirm_clicked = col_confirm.button(
        "이체 자료 확정", type="primary", disabled=not lookup_complete, use_container_width=True
    )
    excel_bytes = dataframe_to_excel_bytes({"이체자료": pd.DataFrame(transfer_rows)}) if transfer_rows else b""
    col_download.download_button(
        "📥 엑셀 다운로드",
        data=excel_bytes,
        file_name=f"이체자료_{_current_kst().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=not transfer_rows,
    )

    if confirm_clicked:
        now_str = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
        for row in targets:
            if row.get("id") in selected_ids:
                account = accounts_by_workplace_id.get(row.get("workplace_id"), {})
                edited_row = edited_by_id.get(row.get("id"), {})
                row["status"] = "이체 대상"
                row["transfer_file_generated_at"] = now_str
                row["transfer_expected_holder"] = (
                    lookup_map.get(row.get("id")) or account.get("holder_name", "") or row.get("workplace_name", "")
                )
                row["transfer_verified_holder"] = lookup_map.get(row.get("id"), "")
                row["transfer_deposit_passbook"] = edited_row.get("입금통장표시", "")
                row["transfer_withdrawal_passbook"] = edited_row.get("출금통장표시", "")
                row["transfer_bank_name"] = account.get("bank_name", "")
                row["transfer_account_number"] = account.get("account_number", "")
                row["transfer_amount"] = row.get("request_amount", 0)
                row["deposit_reconciliation"] = ""
                row["deposit_reconciled_at"] = ""

        st.session_state["_transfer_holder_lookup"] = {
            row_id: name for row_id, name in lookup_map.items() if row_id not in selected_ids
        }
        _save_delegated_workplaces(wp_data)
        st.success(f"{len(transfer_rows)}건이 이체 대상으로 확정되었습니다. [지급 결과 확인]에서 처리 결과를 확정해주세요.")
        st.rerun()

    _render_transfer_confirmed_section(confirmed_rows, accounts_by_workplace_id)


def show_transfer_result_confirmation():
    _render_page_chrome(["홈", "지급관리", "지급 결과 확인", "지급 결과 확인"])
    st.markdown("### 지급 결과 확인")
    st.caption("이체 자료가 확정된 건의 이체 완료 여부를 확인하고, 완료 이력을 관리합니다.")

    wp_data = _load_delegated_workplaces()
    requests = wp_data.get("requests", [])

    bank_data = _load_bank_accounts()
    accounts_by_workplace_id = {
        row.get("linked_workplace_id"): row
        for row in bank_data.get("accounts", [])
        if row.get("linked_workplace_id")
    }

    pending = [row for row in requests if row.get("status") == "이체 대상"]

    title_col, action_col = st.columns([0.78, 0.22])
    with title_col:
        st.markdown("#### 이체 대상")
    with action_col:
        check_clicked = st.button(
            "이체 결과 확인",
            type="primary",
            disabled=not pending,
            use_container_width=True,
            key="transfer_deposit_reconcile",
        )

    if pending:
        if check_clicked:
            matched_count = _reconcile_transfer_deposits(pending, bank_data.get("accounts", []))
            _save_delegated_workplaces(wp_data)
            st.success(f"입금 대사 완료: {matched_count}건 일치, {len(pending) - matched_count}건 미확인")
            st.rerun()

        pending_sorted = sorted(pending, key=lambda item: item.get("transfer_file_generated_at", ""), reverse=True)
        pending_rows = []
        for row in pending_sorted:
            account = accounts_by_workplace_id.get(row.get("workplace_id"), {})
            bank_name = row.get("transfer_bank_name") or account.get("bank_name", "")
            account_number = row.get("transfer_account_number") or account.get("account_number", "")
            pending_rows.append(
                {
                    "id": row.get("id"),
                    "선택": False,
                    "사업장명": row.get("workplace_name", ""),
                    "입금은행": bank_name,
                    "입금계좌번호": account_number,
                    "입금액": _format_won(row.get("request_amount")),
                    "예상예금주": row.get("transfer_expected_holder", "") or account.get("holder_name", "") or row.get("workplace_name", ""),
                    "조회한예금주": row.get("transfer_verified_holder", ""),
                    "입금통장표시": row.get("transfer_deposit_passbook", ""),
                    "출금통장표시": row.get("transfer_withdrawal_passbook", ""),
                    "입금 대사": row.get("deposit_reconciliation", ""),
                }
            )

        pending_df = pd.DataFrame(pending_rows)
        column_order = [
            "id", "선택", "사업장명", "입금은행", "입금계좌번호", "입금액",
            "예상예금주", "조회한예금주", "입금통장표시", "출금통장표시", "입금 대사",
        ]
        st.markdown(
            """
            <style>
            [data-testid="stDataFrame"] [role="columnheader"],
            [data-testid="stDataFrame"] [role="gridcell"],
            [data-testid="stDataEditor"] [role="columnheader"],
            [data-testid="stDataEditor"] [role="gridcell"] {
                text-align: center !important;
                justify-content: center !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
        edited_confirm = st.data_editor(
            pending_df[column_order],
            column_config={
                "id": None,
                "선택": st.column_config.CheckboxColumn("선택", width="small"),
                "사업장명": st.column_config.TextColumn("사업장명", disabled=True, width="small"),
                "입금은행": st.column_config.TextColumn("입금은행", disabled=True, width="small"),
                "입금계좌번호": st.column_config.TextColumn("입금계좌번호", disabled=True, width="small"),
                "입금액": st.column_config.TextColumn("입금액", disabled=True, width="small"),
                "예상예금주": st.column_config.TextColumn("예상예금주", disabled=True, width="small"),
                "조회한예금주": st.column_config.TextColumn("조회한예금주", disabled=True, width="small"),
                "입금통장표시": st.column_config.TextColumn("입금통장표시", disabled=True, width="small"),
                "출금통장표시": st.column_config.TextColumn("출금통장표시", disabled=True, width="small"),
                "입금 대사": st.column_config.TextColumn(
                    "입금 대사",
                    help="계좌번호와 입금액이 일치하는 입금 내역이 확인된 날짜입니다.",
                    disabled=True,
                    width="small",
                ),
            },
            hide_index=True,
            use_container_width=True,
            key="transfer_confirm_editor",
        )
        confirm_ids = edited_confirm[edited_confirm["선택"]]["id"].tolist()
        st.caption("입금 대사는 이체 대상의 입금계좌번호와 입금액이 실제 입금 내역과 일치하는지 확인한 결과입니다. 일치하면 입금일이 표시되고, 없으면 미확인으로 표시됩니다.")
        if st.button("선택 건 이체 완료 확인", type="primary", disabled=not confirm_ids):
            now_str = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
            for row in pending:
                if row.get("id") in confirm_ids:
                    row["status"] = "이체 완료"
                    row["paid_at"] = now_str
            _save_delegated_workplaces(wp_data)
            st.success(f"{len(confirm_ids)}건을 이체 완료로 처리했습니다.")
            st.rerun()
    else:
        st.info("이체 확정 대기 중인 건이 없습니다.")

    st.markdown("#### 이체 완료 이력")
    done = [row for row in requests if row.get("status") == "이체 완료"]
    if done:
        done_df = pd.DataFrame(done)
        done_df["request_amount_won"] = done_df["request_amount"].apply(_format_won)
        if "paid_at" not in done_df.columns:
            done_df["paid_at"] = ""
        done_view = done_df[
            ["paid_at", "workplace_name", "request_amount_won", "requested_by", "requested_at"]
        ].rename(
            columns={
                "paid_at": "이체완료일시",
                "workplace_name": "사업장명",
                "request_amount_won": "이체금액",
                "requested_by": "요청자",
                "requested_at": "요청일시",
            }
        )
        render_plain_html_table(
            done_view.sort_values("이체완료일시", ascending=False),
            center_align=True,
            stretch=True,
            max_width="760px",
            border=False,
        )
    else:
        st.info("이체 완료 이력이 없습니다.")


def show_account_balance_check():
    _render_page_chrome(["홈", "조회", "계좌 잔고 확인", "계좌 잔고 확인"])
    st.markdown("### 계좌 잔고 확인")
    st.caption("위탁 사업장 계좌의 잔고를 수동으로 관리하고 변동 추이를 확인합니다.")

    if not st.session_state.get("_balance_update_form_initialized"):
        st.session_state.show_balance_update_form = False
        st.session_state["_balance_update_form_initialized"] = True

    data = _load_bank_accounts()
    accounts = data.get("accounts", [])

    wp_data = _load_delegated_workplaces()
    workplaces = wp_data.get("workplaces", [])
    workplaces_by_id = {site.get("id"): site for site in workplaces}

    if st.session_state.user_role != "관리자":
        my_ids = {row.get("id") for row in _my_workplaces(workplaces)}
        accounts = [row for row in accounts if row.get("linked_workplace_id") in my_ids]

    if not accounts:
        st.info("등록된 계좌가 없습니다. [계좌 관리]에서 계좌를 먼저 등록해주세요.")
        return

    account_df = pd.DataFrame(accounts)
    account_df["balance_won"] = account_df["balance"].apply(_format_won)
    if "balance_updated_at" not in account_df.columns:
        account_df["balance_updated_at"] = ""
    account_df["회사"] = [_account_company_label(row, workplaces_by_id) for row in accounts]
    account_view = account_df[
        ["회사", "account_name", "bank_name", "account_number", "balance_won", "balance_updated_at"]
    ].rename(
        columns={
            "account_name": "계좌명",
            "bank_name": "은행명",
            "account_number": "계좌번호",
            "balance_won": "현재잔고",
            "balance_updated_at": "최종업데이트",
        }
    )
    title_col, button_col = st.columns([0.78, 0.22])
    with title_col:
        st.markdown("#### 계좌 리스트")
    with button_col:
        if st.button("잔액 정보 업데이트", key="open_balance_update_form", use_container_width=True, type="primary"):
            st.session_state.show_balance_update_form = True
            st.rerun()

    render_plain_html_table(account_view, center_align=True, stretch=True, max_width="100%", border=False)

    st.markdown("#### 거래내역조회")

    history_options = {f"{row.get('account_name')} ({row.get('bank_name')} {row.get('account_number')})": row for row in accounts}
    today = _current_kst().date()
    if "balance_history_start_date" not in st.session_state:
        st.session_state.balance_history_start_date = today - timedelta(days=6)
    if "balance_history_end_date" not in st.session_state:
        st.session_state.balance_history_end_date = today

    period_presets = {
        "오늘": (today, today),
        "어제": (today - timedelta(days=1), today - timedelta(days=1)),
        "1주일": (today - timedelta(days=6), today),
        "1개월": (today - timedelta(days=30), today),
    }

    st.markdown(
        """
        <style>
        .balance-history-help {
            color:#7a8599;
            font-size:12px;
            margin-top:-6px;
        }
        div[data-testid="stRadio"] > label,
        div[data-testid="stSelectbox"] > label,
        div[data-testid="stDateInput"] > label,
        div[data-testid="stTextInput"] > label {
            font-weight:700 !important;
            color:#0f172a !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=False):
        account_label_col, account_input_col = st.columns([0.12, 0.88])
        with account_label_col:
            st.markdown("**계좌번호 <span style='color:#008c78'>*</span>**", unsafe_allow_html=True)
        with account_input_col:
            account_select_col, _ = st.columns([0.48, 0.52])
            with account_select_col:
                history_label = st.selectbox(
                    "계좌번호",
                    list(history_options.keys()),
                    key="balance_history_account",
                    label_visibility="collapsed",
                    placeholder="계좌 선택",
                )

        period_label_col, period_input_col = st.columns([0.12, 0.88])
        with period_label_col:
            st.markdown("**조회기간**")
        with period_input_col:
            preset_cols = st.columns([0.13, 0.13, 0.17, 0.17, 0.40])
            for index, (label, date_range) in enumerate(period_presets.items()):
                button_type = "primary" if label == st.session_state.get("balance_history_preset", "1주일") else "secondary"
                if preset_cols[index].button(label, key=f"balance_history_preset_{label}", type=button_type, use_container_width=True):
                    st.session_state.balance_history_preset = label
                    st.session_state.balance_history_start_date = date_range[0]
                    st.session_state.balance_history_end_date = date_range[1]
                    st.rerun()

            date_start_col, tilde_col, date_end_col, month_col, _ = st.columns([0.22, 0.03, 0.22, 0.18, 0.35])
            with date_start_col:
                start_date = st.date_input("시작일", key="balance_history_start_date", label_visibility="collapsed")
            with tilde_col:
                st.markdown("<div style='padding-top:8px;text-align:center;'>~</div>", unsafe_allow_html=True)
            with date_end_col:
                end_date = st.date_input("종료일", key="balance_history_end_date", label_visibility="collapsed")
            with month_col:
                month_options = ["월별 선택"] + [
                    (today.replace(day=1) - pd.DateOffset(months=idx)).strftime("%Y-%m")
                    for idx in range(12)
                ]
                selected_month = st.selectbox(
                    "월별 선택",
                    month_options,
                    key="balance_history_month",
                    label_visibility="collapsed",
                )
            st.markdown("<div class='balance-history-help'>ㆍ 직접입력 예시 : YYYYMMDD</div>", unsafe_allow_html=True)

        content_label_col, content_input_col = st.columns([0.12, 0.88])
        with content_label_col:
            st.markdown("**조회내용**")
        with content_input_col:
            history_type = st.radio(
                "조회내용",
                ["전체(입금+출금)", "입금내역", "출금내역"],
                horizontal=True,
                key="balance_history_type",
                label_visibility="collapsed",
            )

        sort_label_col, sort_input_col = st.columns([0.12, 0.88])
        with sort_label_col:
            st.markdown("**정렬방식**")
        with sort_input_col:
            sort_col, count_col = st.columns([0.22, 0.5])
            with sort_col:
                sort_order = st.selectbox(
                    "정렬방식",
                    ["최근거래먼저", "과거거래먼저"],
                    key="balance_history_sort",
                    label_visibility="collapsed",
                )
            with count_col:
                result_count = st.radio(
                    "조회건수",
                    [15, 30, 50, 100],
                    index=1,
                    horizontal=True,
                    key="balance_history_limit",
                    label_visibility="collapsed",
                    format_func=lambda value: f"{value}건",
                )

        search_label_col, search_input_col = st.columns([0.12, 0.88])
        with search_label_col:
            st.markdown("**검색조건**")
        with search_input_col:
            search_type_col, keyword_col, _ = st.columns([0.22, 0.40, 0.16])
            with search_type_col:
                search_type = st.selectbox(
                    "검색조건",
                    ["적요"],
                    key="balance_history_search_type",
                    label_visibility="collapsed",
                )
            with keyword_col:
                search_keyword = st.text_input(
                    "검색어",
                    key="balance_history_keyword",
                    label_visibility="collapsed",
                    placeholder="적요(통장 메모) 최대 25자까지 입력가능",
                    max_chars=25,
                )

    button_left, button_center, button_right = st.columns([0.42, 0.16, 0.42])
    with button_center:
        history_clicked = st.button("조회", key="balance_history_search", type="primary", use_container_width=True)

    if selected_month != "월별 선택":
        month_start = pd.to_datetime(f"{selected_month}-01").date()
        month_end = (pd.to_datetime(f"{selected_month}-01") + pd.offsets.MonthEnd(0)).date()
        start_date, end_date = month_start, month_end

    if history_clicked:
        st.session_state["_balance_history_target"] = history_label
        st.session_state["_balance_history_query"] = {
            "start_date": start_date,
            "end_date": end_date,
            "history_type": history_type,
            "sort_order": sort_order,
            "result_count": result_count,
            "search_type": search_type,
            "search_keyword": search_keyword,
        }

    history_target_label = st.session_state.get("_balance_history_target")
    history_query = st.session_state.get("_balance_history_query", {})
    if history_target_label and history_target_label in history_options:
        target_account = history_options[history_target_label]
        history_rows = target_account.get("balance_history", [])
        start_date = history_query.get("start_date", start_date)
        end_date = history_query.get("end_date", end_date)
        history_type = history_query.get("history_type", history_type)
        sort_order = history_query.get("sort_order", sort_order)
        result_count = history_query.get("result_count", result_count)
        search_keyword = str(history_query.get("search_keyword", search_keyword) or "").strip()

        enriched_history = []
        previous_balance = None
        for item in sorted(history_rows, key=lambda row: str(row.get("at", ""))):
            amount = _history_amount_value(item, previous_balance)
            if "balance" in item:
                previous_balance = _coerce_int_amount(item.get("balance"))
            enriched = dict(item)
            enriched["_amount"] = amount
            enriched_history.append(enriched)

        filtered_history = []
        for item in enriched_history:
            parsed = _parse_kst_datetime(item.get("at"))
            if parsed and not (start_date <= parsed.date() <= end_date):
                continue
            if history_type == "입금내역" and item.get("_amount", 0) <= 0:
                continue
            if history_type == "출금내역" and item.get("_amount", 0) >= 0:
                continue
            if search_keyword and search_keyword not in str(item.get("memo", "")):
                continue
            filtered_history.append(item)

        reverse_sort = sort_order == "최근거래먼저"
        filtered_history = sorted(filtered_history, key=lambda row: str(row.get("at", "")), reverse=reverse_sort)
        filtered_history = filtered_history[: int(result_count)]

        if filtered_history:
            history_df = pd.DataFrame(filtered_history)
            history_df["거래금액"] = history_df["_amount"].apply(_format_won)
            history_df["잔고"] = history_df["balance"].apply(_format_won)
            if "memo" not in history_df.columns:
                history_df["memo"] = ""
            history_view = history_df[["at", "거래금액", "잔고", "memo"]].rename(
                columns={"at": "거래일시", "memo": "적요"}
            )
            render_plain_html_table(history_view, center_align=True, stretch=True, max_width="760px", border=False)
        else:
            st.info("조건에 맞는 거래내역이 없습니다.")

    if not st.session_state.get("show_balance_update_form"):
        return

    st.markdown("#### 잔액 정보 업데이트")
    account_options = {f"{row.get('account_name')} ({row.get('bank_name')} {row.get('account_number')})": row for row in accounts}
    selected_label = st.selectbox("계좌 선택", list(account_options.keys()), key="balance_target_account")
    selected_account = account_options[selected_label]

    with st.form("balance_update_form", clear_on_submit=True):
        new_balance = st.number_input(
            "현재 잔고", min_value=0, step=10000, format="%d", value=int(selected_account.get("balance", 0) or 0)
        )
        memo = st.text_input("메모", placeholder="예: 정기 확인")
        form_col1, form_col2 = st.columns(2)
        with form_col1:
            submitted = st.form_submit_button("업데이트", type="primary", use_container_width=True)
        with form_col2:
            cancelled = st.form_submit_button("취소", use_container_width=True)

    if submitted:
        now_str = _current_kst().strftime("%Y-%m-%d %H:%M:%S")
        selected_account["balance"] = int(new_balance)
        selected_account["balance_updated_at"] = now_str
        selected_account.setdefault("balance_history", []).append(
            {"at": now_str, "balance": int(new_balance), "memo": memo.strip()}
        )
        _save_bank_accounts(data)
        st.session_state.show_balance_update_form = False
        st.success("잔고가 업데이트되었습니다.")
        st.rerun()

    if cancelled:
        st.session_state.show_balance_update_form = False
        st.rerun()

    history = selected_account.get("balance_history", [])
    if history:
        st.markdown("#### 잔고 변동 추이")
        import plotly.graph_objects as go
        fig = go.Figure(go.Scatter(
            x=[item.get("at") for item in history],
            y=[item.get("balance") for item in history],
            mode="lines+markers",
        ))
        fig.update_layout(**_chart_layout(height=300))
        st.plotly_chart(fig, use_container_width=True, theme=None)


def _render_workplace_info_admin():
    st.caption("위탁 사업장의 기본 정보를 관리합니다.")

    data = _load_delegated_workplaces()
    workplaces = data.get("workplaces", [])

    if workplaces:
        st.markdown("#### 등록된 사업장")
        site_df = pd.DataFrame(workplaces)
        if "business_number" not in site_df.columns:
            site_df["business_number"] = ""
        if "business_alias" not in site_df.columns:
            site_df["business_alias"] = ""
        if "regular_payment_day" not in site_df.columns:
            site_df["regular_payment_day"] = 0
        if "manager_name" not in site_df.columns:
            site_df["manager_name"] = ""
        if "memo" not in site_df.columns:
            site_df["memo"] = ""
        site_df["regular_payment_day_label"] = site_df["regular_payment_day"].apply(
            lambda value: f"매월 {int(float(value or 0))}일" if int(float(value or 0)) else "미등록"
        )
        site_view = site_df[
            ["workplace_name", "business_number", "business_alias", "regular_payment_day_label", "manager_name", "memo"]
        ].rename(
            columns={
                "workplace_name": "사업장명",
                "business_number": "사업자번호",
                "business_alias": "사업자별칭",
                "regular_payment_day_label": "정기지급일자",
                "manager_name": "담당자",
                "memo": "비고",
            }
        ).reset_index(drop=True)
        site_view.insert(0, "순번", range(1, len(site_view) + 1))

        status_map = st.session_state.get("_biz_status_map", {})
        if status_map:
            site_view["휴폐업상태"] = site_df["business_number"].apply(
                lambda v: status_map.get(re.sub(r"\D", "", str(v)), "")
            )

        render_plain_html_table(site_view, center_align=True)

        col_spacer, col_status_btn = st.columns([165, 35])
        if col_status_btn.button("휴폐업조회", key="biz_status_check_btn", use_container_width=True):
            api_key = _get_nts_api_key()
            biz_numbers = site_df["business_number"].tolist()
            if not api_key:
                st.warning("국세청 API 키(NTS_API_KEY)가 설정되지 않았습니다.")
            elif not any(str(b).strip() for b in biz_numbers):
                st.warning("조회할 사업자번호가 없습니다.")
            else:
                result = _check_business_status(api_key, biz_numbers)
                if result:
                    st.session_state["_biz_status_map"] = result
                    st.rerun()
                else:
                    st.warning("휴폐업 조회 결과를 가져오지 못했습니다.")
    else:
        st.info("등록된 사업장이 없습니다.")

    staff_names = sorted({
        info.get("name") for info in _real_users(load_user_db()).values()
        if isinstance(info, dict) and info.get("name")
    })
    manager_options = ["미지정"] + staff_names

    if st.session_state.get("_show_workplace_edit_form"):
        edit_target_id = st.session_state.get("_workplace_edit_target_id")
        target = next((row for row in workplaces if row.get("id") == edit_target_id), None)
        if target is None:
            st.session_state["_show_workplace_edit_form"] = False
            st.rerun()

        current_manager = target.get("manager_name", "")
        manager_index = manager_options.index(current_manager) if current_manager in manager_options else 0

        with st.form("delegated_workplace_edit_form", clear_on_submit=False):
            workplace_name = st.text_input("사업장명", value=target.get("workplace_name", ""))
            business_number = st.text_input("사업자번호", value=target.get("business_number", ""))
            business_alias = st.text_input("사업자별칭", value=target.get("business_alias", ""))
            regular_payment_day = st.number_input(
                "정기지급일자", min_value=0, max_value=31, step=1, value=int(target.get("regular_payment_day", 0) or 0)
            )
            manager_name = st.selectbox("담당자", manager_options, index=manager_index)
            memo = st.text_area("비고", value=target.get("memo", ""))
            col_submit, col_cancel = st.columns(2)
            submitted = col_submit.form_submit_button("수정 완료", type="primary", use_container_width=True)
            cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

        if submitted:
            normalized = _normalize_name(workplace_name)
            duplicate = any(
                row.get("id") != edit_target_id and _normalize_name(row.get("workplace_name")) == normalized
                for row in workplaces
            )
            if not workplace_name.strip():
                st.warning("사업장명을 입력해주세요.")
            elif duplicate:
                st.error("이미 등록된 사업장입니다.")
            else:
                target["workplace_name"] = workplace_name.strip()
                target["business_number"] = business_number.strip()
                target["business_alias"] = business_alias.strip()
                target["regular_payment_day"] = int(regular_payment_day)
                target["manager_name"] = "" if manager_name == "미지정" else manager_name
                target["memo"] = memo.strip()
                _save_delegated_workplaces(data)
                st.session_state["_show_workplace_edit_form"] = False
                st.success("사업장 정보가 수정되었습니다.")
                st.rerun()
        if cancelled:
            st.session_state["_show_workplace_edit_form"] = False
            st.rerun()
    elif st.session_state.get("_show_workplace_add_form"):
        with st.form("delegated_workplace_form", clear_on_submit=True):
            workplace_name = st.text_input("사업장명", placeholder="예: 강남 위탁사업장")
            business_number = st.text_input("사업자번호", placeholder="예: 123-45-67890")
            business_alias = st.text_input("사업자별칭", placeholder="예: 강남점")
            regular_payment_day = st.number_input("정기지급일자", min_value=0, max_value=31, step=1, value=0)
            manager_name = st.selectbox("담당자", manager_options)
            memo = st.text_area("비고", placeholder="사업장 관련 특이사항")
            col_submit, col_cancel = st.columns(2)
            submitted = col_submit.form_submit_button("등록", type="primary", use_container_width=True)
            cancelled = col_cancel.form_submit_button("취소", use_container_width=True)

        if submitted:
            normalized = _normalize_name(workplace_name)
            duplicate = any(_normalize_name(row.get("workplace_name")) == normalized for row in workplaces)
            if not workplace_name.strip():
                st.warning("사업장명을 입력해주세요.")
            elif duplicate:
                st.error("이미 등록된 사업장입니다.")
            else:
                workplaces.append(
                    {
                        "id": int(time.time() * 1000),
                        "workplace_name": workplace_name.strip(),
                        "business_number": business_number.strip(),
                        "business_alias": business_alias.strip(),
                        "bank_name": "",
                        "account_number": "",
                        "regular_payment_day": int(regular_payment_day),
                        "manager_name": "" if manager_name == "미지정" else manager_name,
                        "manager_contact": "",
                        "memo": memo.strip(),
                        "created_at": _current_kst().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                _save_delegated_workplaces(data)
                st.session_state["_show_workplace_add_form"] = False
                st.success("사업장 정보가 등록되었습니다.")
                st.rerun()
        if cancelled:
            st.session_state["_show_workplace_add_form"] = False
            st.rerun()
    elif st.session_state.get("_show_workplace_delete_select"):
        if workplaces:
            site_options = {
                f"{row.get('workplace_name', '')} ({row.get('business_number', '') or '-'})": row.get("id")
                for row in workplaces
            }
            target_label = st.selectbox("삭제 대상 사업장", list(site_options.keys()), key="workplace_delete_target")
            target_id = site_options[target_label]

            col_confirm, col_cancel = st.columns(2)
            if col_confirm.button("삭제 확인", type="primary", use_container_width=True):
                data["workplaces"] = [row for row in workplaces if row.get("id") != target_id]
                _save_delegated_workplaces(data)
                st.session_state["_show_workplace_delete_select"] = False
                st.success("사업장 정보가 삭제되었습니다.")
                st.rerun()
            if col_cancel.button("취소", use_container_width=True):
                st.session_state["_show_workplace_delete_select"] = False
                st.rerun()
    elif st.session_state.get("_show_workplace_edit_select"):
        if workplaces:
            site_options = {
                f"{row.get('workplace_name', '')} ({row.get('business_number', '') or '-'})": row.get("id")
                for row in workplaces
            }
            target_label = st.selectbox("수정 대상 사업장", list(site_options.keys()), key="workplace_modify_target")
            target_id = site_options[target_label]

            col_confirm, col_cancel = st.columns(2)
            if col_confirm.button("수정 진행", type="primary", use_container_width=True):
                st.session_state["_show_workplace_edit_select"] = False
                st.session_state["_show_workplace_edit_form"] = True
                st.session_state["_workplace_edit_target_id"] = target_id
                st.rerun()
            if col_cancel.button("취소", use_container_width=True):
                st.session_state["_show_workplace_edit_select"] = False
                st.rerun()
    else:
        col_add, col_edit, col_delete, _ = st.columns([20, 20, 20, 140])
        if col_add.button("+ 사업장 추가", key="show_workplace_add_form_btn", use_container_width=True):
            st.session_state["_show_workplace_add_form"] = True
            st.rerun()

        if col_edit.button("수정", key="show_workplace_edit_form_btn", use_container_width=True, disabled=not workplaces):
            st.session_state["_show_workplace_edit_select"] = True
            st.rerun()

        if col_delete.button("삭제", key="delete_workplace_btn", use_container_width=True, disabled=not workplaces):
            st.session_state["_show_workplace_delete_select"] = True
            st.rerun()


def _dashboard_ai_messages(metrics, forecast_rows):
    """대시보드 상단에 표시할 AI 인사이트 메시지 목록을 생성."""
    messages = []
    if metrics["pending"]:
        messages.append(("📥", f"결재 대기 중인 전도금 요청이 {metrics['pending']}건 있습니다. [전자결재]에서 확인해주세요."))
    if metrics["approved"]:
        messages.append(("💸", f"품의 확정되어 이체 대상인 요청이 {metrics['approved']}건 있습니다. [이체 자료 확정]에서 처리해주세요."))
    for row in forecast_rows:
        risk = row.get("리스크", "")
        if risk and risk != "정상":
            messages.append(("⚠️", f"{row.get('사업장명', '')}: {risk} (추천 지급액 {row.get('추천 지급액', '')})"))
    if not messages:
        messages.append(("✅", "현재 특이사항 없이 정상적으로 운영되고 있습니다."))
    return messages


def show_dashboard():
    st.markdown("### 대시보드")
    st.caption("위탁 사업장 운영 현황과 전도금 지급 이력, 계좌 잔고 현황을 확인합니다.")

    wp_data = _load_delegated_workplaces()
    workplaces = wp_data.get("workplaces", [])
    requests = wp_data.get("requests", [])
    metrics = _workplace_request_metrics(requests)
    paid_count = len([row for row in requests if row.get("status") == "이체 완료"])
    total_count = max(len(requests), 1)
    forecast_rows = _workplace_forecast_rows(workplaces, requests)

    _workplace_dashboard_css()
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(_donut_card("Workplaces", f"{len(workplaces):,}", "Sites", min(len(workplaces) * 12, 100)), unsafe_allow_html=True)
    c2.markdown(_donut_card("결재대기", f"{metrics['pending']:,}", "Pending", min(metrics["pending"] * 18, 100)), unsafe_allow_html=True)
    c3.markdown(_donut_card("품의확정", f"{metrics['approved']:,}", "Approved", min(metrics["approved"] * 18, 100)), unsafe_allow_html=True)
    c4.markdown(_donut_card("이체완료", f"{paid_count:,}", "Transfers", int((paid_count / total_count) * 100)), unsafe_allow_html=True)

    ai_messages = _dashboard_ai_messages(metrics, forecast_rows)
    rows_html = "".join(
        f"<div class='suggestion-row'><div class='suggestion-icon'>{html.escape(icon)}</div><div>{html.escape(text)}</div></div>"
        for icon, text in ai_messages
    )
    st.markdown(
        f"<div class='suggestion-panel'><div class='suggestion-title'>🤖 AI 인사이트</div>{rows_html}</div>",
        unsafe_allow_html=True,
    )

    st.markdown("#### 전도금 지급 이력")
    if requests:
        request_df = pd.DataFrame(requests)
        request_df["request_amount_won"] = request_df["request_amount"].apply(_format_won)
        for col in ["approved_at", "paid_at", "reject_reason"]:
            if col not in request_df.columns:
                request_df[col] = ""
        request_view = request_df[
            ["requested_at", "workplace_name", "request_amount_won", "requested_by", "status", "approved_at", "paid_at", "reject_reason"]
        ].rename(
            columns={
                "requested_at": "요청일시",
                "workplace_name": "사업장명",
                "request_amount_won": "요청 금액",
                "requested_by": "요청자",
                "status": "상태",
                "approved_at": "품의 확정일시",
                "paid_at": "이체 완료일시",
                "reject_reason": "반려사유",
            }
        )
        render_plain_html_table(request_view.sort_values("요청일시", ascending=False), center_align=True)
    else:
        st.info("전도금 지급 이력이 없습니다.")

    if forecast_rows:
        st.markdown("#### AI 예측 안내")
        st.caption("현재 MVP는 지급 이력 기반의 규칙형 예측입니다. 충분한 이력이 쌓이면 모델 기반 예측으로 확장할 수 있습니다.")
        render_plain_html_table(pd.DataFrame(forecast_rows), center_align=True)
    elif not workplaces:
        st.info("예측을 위해 먼저 사업장 정보를 등록해주세요.")

    st.markdown("#### 계좌별 잔고 현황")
    bank_data = _load_bank_accounts()
    accounts = bank_data.get("accounts", [])
    if accounts:
        workplaces_by_id = {site.get("id"): site for site in workplaces}
        account_df = pd.DataFrame(accounts)
        account_df["balance_won"] = account_df["balance"].apply(_format_won)
        if "balance_updated_at" not in account_df.columns:
            account_df["balance_updated_at"] = ""
        account_df["회사"] = [_account_company_label(row, workplaces_by_id) for row in accounts]
        account_view = account_df[
            ["회사", "account_name", "bank_name", "balance_won", "balance_updated_at"]
        ].rename(
            columns={
                "account_name": "계좌명",
                "bank_name": "은행명",
                "balance_won": "현재잔고",
                "balance_updated_at": "최종업데이트",
            }
        )
        render_plain_html_table(account_view, center_align=True)
    else:
        st.info("등록된 계좌가 없습니다.")


def _chart_layout(height=300, **overrides):
    """Plotly 차트 기본 레이아웃. 배경은 앱 테마가 보이도록 투명 처리."""
    bg = "rgba(0,0,0,0)"
    plot = "rgba(0,0,0,0)"
    txt = "#1e293b"
    grid = "#e2e8f0"
    axis = "#cbd5e1"
    legend_bg = "rgba(255,255,255,0)"
    layout = dict(
        paper_bgcolor=bg, plot_bgcolor=plot, height=height,
        font=dict(color=txt, size=12),
        xaxis=dict(gridcolor=grid, linecolor=axis,
                   tickfont=dict(color=txt), title_font=dict(color=txt),
                   showgrid=True, zeroline=False),
        yaxis=dict(gridcolor=grid, linecolor=axis,
                   tickfont=dict(color=txt), title_font=dict(color=txt),
                   showgrid=True, zeroline=False),
        legend=dict(font=dict(color=txt), bgcolor=legend_bg,
                    bordercolor=axis, borderwidth=1),
        margin=dict(t=20, b=20, l=10, r=10),
    )
    layout.update(overrides)
    return layout


def apply_global_table_css():
    st.markdown(
        """
        <style>
        /* 전체 버튼 색상 통일 */
        div.stButton > button,
        [data-testid="stFormSubmitButton"] button,
        [data-testid="stDownloadButton"] button {
            background-color: #34495E !important;
            color: #FFFFFF !important;
            border-color: #34495E !important;
        }
        div.stButton > button p, div.stButton > button span,
        [data-testid="stFormSubmitButton"] button p, [data-testid="stFormSubmitButton"] button span,
        [data-testid="stDownloadButton"] button p, [data-testid="stDownloadButton"] button span {
            color: #FFFFFF !important;
        }
        div.stButton > button:hover,
        [data-testid="stFormSubmitButton"] button:hover,
        [data-testid="stDownloadButton"] button:hover {
            background-color: #2C3E50 !important;
            border-color: #2C3E50 !important;
            color: #FFFFFF !important;
        }
        div.stButton > button:disabled,
        [data-testid="stFormSubmitButton"] button:disabled,
        [data-testid="stDownloadButton"] button:disabled {
            background-color: #A0AEC0 !important;
            border-color: #A0AEC0 !important;
            color: #F1F5F9 !important;
            opacity: 0.7;
        }
        div[data-testid="stDataFrame"] th,
        div[data-testid="stDataFrame"] td,
        div[data-testid="stDataFrame"] [role="columnheader"],
        div[data-testid="stDataFrame"] [role="gridcell"],
        div[data-testid="stDataEditor"] [role="columnheader"],
        div[data-testid="stDataEditor"] [role="gridcell"] {
            white-space: nowrap !important;
            word-break: keep-all !important;
        }
        div[data-testid="stDataFrame"] [role="columnheader"],
        div[data-testid="stDataEditor"] [role="columnheader"] {
            height: 20px !important;
            min-height: 20px !important;
            max-height: 20px !important;
            line-height: 20px !important;
        }
        /* 표 전체 가운데 정렬 */
        div[data-testid="stDataFrame"] [role="columnheader"],
        div[data-testid="stDataFrame"] [role="gridcell"],
        div[data-testid="stDataEditor"] [role="columnheader"],
        div[data-testid="stDataEditor"] [role="gridcell"] {
            text-align: center !important;
            justify-content: center !important;
        }
        /* 버튼 크기 공통 (높이 30px) */
        div.stButton > button,
        [data-testid="stFormSubmitButton"] button,
        [data-testid="stDownloadButton"] button {
            min-height: 30px !important;
            height: 30px !important;
            padding: 0 0.75rem !important;
            font-size: 13px !important;
            line-height: 1.2 !important;
        }
        .action-btn button p, .action-btn button span {
            font-size: 10px !important;
        }
        /* 직원 현황 카드 (라이트 모드) */
        .pms-staff-card {
            border: 1px solid #e2e8f0; border-radius: 10px;
            padding: 14px 18px; margin-bottom: 10px; background: #f8fafc;
        }
        .pms-card-name { font-size: 16px; font-weight: 700; margin-bottom: 8px; color: #1e293b; }
        .pms-card-stats { display: flex; gap: 12px; flex-wrap: wrap; }
        .pms-card-stats span { font-size: 13px; color: #475569; }
        .pms-card-points { margin-top: 6px; font-size: 15px; font-weight: 600; color: #4F46E5; }
        .pms-delta { font-size: 12px; margin-left: 8px; }
        .pms-delta-up   { color: #16a34a; }
        .pms-delta-down { color: #dc2626; }
        /* 직원 현황 카드 (다크 모드) */
        body:has(#pms-d:checked) .pms-staff-card {
            background: #252535 !important; border-color: #45475a !important;
        }
        body:has(#pms-d:checked) .pms-card-name  { color: #ffffff !important; }
        body:has(#pms-d:checked) .pms-card-stats span { color: #ffffff !important; }
        body:has(#pms-d:checked) .pms-card-points { color: #ffffff !important; }
        body:has(#pms-d:checked) .pms-delta-up   { color: #a6e3a1 !important; }
        body:has(#pms-d:checked) .pms-delta-down { color: #f38ba8 !important; }
        /* 기타 인라인 배경 패턴 (흰/연회색 배경 다크 처리) */
        body:has(#pms-d:checked) [style*="background:#f8fafc"],
        body:has(#pms-d:checked) [style*="background: #f8fafc"],
        body:has(#pms-d:checked) [style*="background:#EBF8FF"],
        body:has(#pms-d:checked) [style*="background: #EBF8FF"],
        body:has(#pms-d:checked) [style*="background:#fff"],
        body:has(#pms-d:checked) [style*="background: #fff"],
        body:has(#pms-d:checked) [style*="background:white"],
        body:has(#pms-d:checked) [style*="background: white"] {
            background: #252535 !important;
        }
        body:has(#pms-d:checked) [style*="color:#2B6CB0"] { color: #89dceb !important; }
        body:has(#pms-d:checked) [style*="color: #2B6CB0"] { color: #89dceb !important; }

        /* 실적 보고 테이블 (style_report_logic 래퍼) 다크모드 */
        body:has(#pms-d:checked) .pms-report-table {
            border-color: #45475a !important;
            box-shadow: none !important;
            background: #252535 !important;
        }
        body:has(#pms-d:checked) .pms-report-table th {
            background: #0f0f1f !important;
            color: #ffffff !important;
            border-color: #45475a !important;
        }
        body:has(#pms-d:checked) .pms-report-table td {
            background: #252535 !important;
            color: #ffffff !important;
            border-color: #313244 !important;
        }
        body:has(#pms-d:checked) .pms-report-table tr:nth-child(odd) td {
            background: #1e1e30 !important;
        }
        body:has(#pms-d:checked) .pms-report-table tr:hover td {
            background: #2d2d45 !important;
        }

        /* 다운로드 버튼 다크모드 */
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] button,
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] a {
            background-color: #2a2a3e !important;
            color: #ffffff !important;
            border-color: #45475a !important;
        }
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] button:hover,
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] a:hover {
            background-color: #313244 !important;
            border-color: #ffffff !important;
        }
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] button p,
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] button span,
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] a p,
        body:has(#pms-d:checked) [data-testid="stDownloadButton"] a span {
            color: #ffffff !important;
        }
        /* 익스팬더(메뉴 이용 안내 등) 헤더 크기 공통 (높이 40px) */
        [data-testid="stExpander"] summary {
            min-height: 40px !important;
            height: 40px !important;
            padding: 0.25rem 0.75rem !important;
        }
        [data-testid="stExpander"] summary p {
            font-size: 0.85rem !important;
            margin: 0 !important;
        }
        [data-testid="stExpander"] summary svg,
        [data-testid="stExpander"] [data-testid="stExpanderToggleIcon"] {
            width: 1rem !important;
            height: 1rem !important;
        }
        [data-testid="stExpanderDetails"] {
            padding: 0.5rem 0.75rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


MENU_GUIDES = {
    "대시보드": [
        "🤖 AI 인사이트에서 결재 대기, 이체 대상, 리스크 사업장 등 주요 알림을 확인합니다.",
        "📊 사업장 현황, 전도금 지급 이력, AI 예측 안내를 한 화면에서 확인합니다.",
        "🏦 계좌별 잔고 현황도 함께 제공합니다.",
    ],
    "전자결재": [
        "📥 결재대기 문서를 승인/반려 처리합니다. (관리자)",
        "✅ 승인 시 해당 전도금 요청이 [품의 결과]에 '품의 확정'으로 반영됩니다.",
        "🗂️ 처리 완료 탭에서 과거 결재 이력을 조회할 수 있습니다.",
        "📈 처리 통계 탭에서 결재대기/승인/반려 건수와 월별 처리 추이를 확인합니다.",
    ],
    "회사 관리": [
        "🏬 당사(우리 회사)의 회사명, 사업자번호, 대표자, 연락처, 주소 등 기본 정보를 등록/수정합니다.",
        "🏦 여러 출금계좌를 등록하고 메인 출금계좌를 지정할 수 있습니다.",
        "💾 저장 시 최종 수정 일시가 자동으로 기록됩니다.",
        "📄 등록된 회사 정보는 문서 출력 및 엑셀 양식 등에서 활용됩니다.",
    ],
    "위탁 사업장 관리": [
        "🏢 [사업장 정보 관리] 탭에서 사업장 정보와 정기 지급일을 등록합니다.",
        "🏦 [계좌 관리] 탭에서 [사업장 정보 관리]에 등록된 사업장을 선택하여 계좌 정보를 등록/관리합니다.",
        "🔄 계좌 관리에서 '위탁 사업장 계좌 가져오기'로 사업장 계좌를 일괄 등록할 수 있습니다.",
        "👤 [담당자 관리] 탭에서 직원 계정의 접근 권한, 비밀번호, 역할을 관리합니다.",
    ],
    "서버 접속 정보": [
        "🗄️ MSSQL 서버 접속 정보를 등록합니다.",
        "🔐 비밀번호는 화면에 마스킹되어 표시됩니다.",
        "⚙️ 운영 환경에 맞게 인증 방식과 암호화 옵션을 설정합니다.",
    ],
    "이체 자료 확정": [
        "✅ 품의 확정된 전도금 요청을 선택해 이체 자료(엑셀)를 다운로드합니다.",
        "🏦 출금 계좌는 [회사 관리]에 등록된 출금계좌 중에서 선택하며, 메인 출금계좌가 먼저 표시됩니다.",
        "➡️ '이체 자료 확정' 클릭 시 해당 요청은 '이체 대상' 상태로 전환되어 [지급 결과 확인]에 표시됩니다.",
    ],
    "지급 결과 확인": [
        "💸 '이체 대상' 건을 선택해 이체 완료 여부를 확정합니다.",
        "📜 이체 완료된 지급 이력을 조회할 수 있습니다.",
    ],
    "전도금 요청": [
        "🧾 전도금 요청을 등록하면 [전자결재]로 결재 요청이 전달됩니다.",
        "📋 최근 요청현황에서 최근 7일간 요청한 내역과 처리 상태를 확인합니다.",
    ],
    "품의 결과": [
        "📊 전도금 요청의 품의(승인/반려) 처리 결과를 조회합니다.",
        "🔍 사업장/처리결과로 필터링할 수 있습니다.",
    ],
    "전도금 사용 결의 보고": [
        "🧾 제출된 전도금 사용 결의 보고 내역과 처리 상태를 확인합니다.",
    ],
    "계좌 잔고 확인": [
        "🏦 등록된 계좌의 현재 잔고를 확인하고 수동으로 업데이트합니다.",
        "📈 잔고 변동 추이를 그래프로 확인할 수 있습니다.",
    ],
}


def render_page_title(menu):
    st.markdown(
        """
        <style>
            .block-container { padding-top: 1rem !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # 메뉴 이용 안내를 우측 상단에 배치
    _, col_guide = st.columns([0.82, 0.18])

    with col_guide:
        if menu in MENU_GUIDES:
            with st.expander("📌 메뉴 이용 안내", expanded=False):
                for line in MENU_GUIDES[menu]:
                    st.markdown(f"- {line}")

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)


def load_csv_to_state(url_key, state_key, force_refresh=False):
    st.session_state[state_key] = clean_header_logic(read_google_csv(st.session_state[url_key], force_refresh=force_refresh))


def refresh_google_sheets_action():
    st.session_state.cloud_sheet_df = clean_header_logic(read_google_csv(st.session_state.url_sync, force_refresh=True))
    st.session_state.analysis_lookup_df = clean_header_logic(read_google_csv(st.session_state.url_analysis, force_refresh=True))
    st.toast("구글시트 데이터를 다시 조회했습니다.")
    time.sleep(0.3)
    st.rerun()


def validation_tabs_with_refresh(key):
    st.markdown(
        """
        <style>
        /* .refresh-tab-button div과 button도 형제 — :has()+인접형제로 타겟팅 */
        [data-testid="stElementContainer"]:has(.refresh-tab-button) + [data-testid="stElementContainer"] [data-testid="stButton"] button {
            min-width: 42px !important;
            height: 38px !important;
            padding: 0 !important;
            margin-top: 2px !important;
            border-radius: 8px !important;
            font-size: 18px !important;
            font-weight: 900 !important;
        }
        body:has(#pms-d:checked) [data-testid="stElementContainer"]:has(.refresh-tab-button) + [data-testid="stElementContainer"] [data-testid="stButton"] button {
            background: #2a2a3e !important;
            background-color: #2a2a3e !important;
            color: #cdd6f4 !important;
            border-color: #45475a !important;
        }
        body:has(#pms-d:checked) [data-testid="stElementContainer"]:has(.refresh-tab-button) + [data-testid="stElementContainer"] [data-testid="stButton"] button:hover {
            background: #313244 !important;
            background-color: #313244 !important;
            color: #ffffff !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    tabs_col, link_col, refresh_col = st.columns([0.91, 0.045, 0.045])
    with tabs_col:
        tabs = st.tabs(["중복 이력", "초과 방문", "본사 개설완료일자 누락", "본사 ERP연계일자 누락", "기타 오류"])
    with link_col:
        _sheet_url = "https://docs.google.com/spreadsheets/d/1yS4gaES-iuzt1NSRTSdj9Ivg1fjbN5mIyX4pGnvEYN0/edit?gid=1533424484#gid=1533424484"
        st.link_button("🔗", _sheet_url, use_container_width=True, help="본사 구글시트 바로가기")
    with refresh_col:
        st.markdown("<div class='refresh-tab-button'>", unsafe_allow_html=True)
        if st.button("↻", use_container_width=True, key=key, help="구글시트 데이터 다시 조회"):
            try:
                refresh_google_sheets_action()
            except Exception as e:
                st.error(f"구글시트 갱신 실패: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
    return tabs


def select_prev_month(state_key, widget_key):
    # 초기 로드 시 구글시트 자동 조회
    prev_sel_key = f"{widget_key}_prev"

    if st.session_state.analysis_lookup_df is None:
        try:
            load_csv_to_state("url_analysis", "analysis_lookup_df")
        except Exception:
            pass

    if st.session_state.analysis_lookup_df is not None:
        c_df = st.session_state.analysis_lookup_df.copy()
        d_col = find_col(c_df, ["활동일", "일자"])
        if d_col and d_col in c_df.columns:
            c_df[d_col] = pd.to_datetime(c_df[d_col], errors="coerce")
            opts = sorted(c_df[d_col].dropna().dt.strftime("%Y-%m").unique(), reverse=True)
            sel = st.selectbox("비교할 전월 선택", ["선택안함"] + list(opts), key=widget_key)

            # 선택값이 변경되었을 때 구글시트 재조회
            prev_sel = st.session_state.get(prev_sel_key)
            if prev_sel != sel and prev_sel is not None:
                try:
                    load_csv_to_state("url_analysis", "analysis_lookup_df")
                    c_df = st.session_state.analysis_lookup_df.copy()
                    c_df[d_col] = pd.to_datetime(c_df[d_col], errors="coerce")
                except Exception:
                    pass

            st.session_state[prev_sel_key] = sel
            st.session_state[state_key] = c_df[c_df[d_col].dt.strftime("%Y-%m") == sel] if sel != "선택안함" else None


def convert_bank_excel_to_activity(bank_df):
    """
    은행 엑셀 파일을 활동실적 양식으로 변환

    Args:
        bank_df: 은행 엑셀 DataFrame (접수일자, CMS구분, 고객번호, 업체명, 접수유형, 업무유형, 접수자, 담당자, 진행상태, 요청사항, 처리내용)

    Returns:
        변환된 DataFrame (등록자, 활동일, 활동상세, 업체명, 사업자번호)
    """
    try:
        # 1. 필요한 컬럼 찾기
        receipt_type_col = find_col(bank_df, ["접수유형"], "접수유형")
        receipt_date_col = find_col(bank_df, ["접수일자"], "접수일자")
        customer_no_col = find_col(bank_df, ["고객번호"], "고객번호")
        company_col = find_col(bank_df, ["업체명"], "업체명")
        manager_col = find_col(bank_df, ["담당자"], "담당자")
        request_col = find_col(bank_df, ["요청사항"], "요청사항")

        # 2. 접수유형이 "방문"인 것만 필터링
        if receipt_type_col not in bank_df.columns:
            st.error("접수유형 컬럼을 찾을 수 없습니다.")
            return pd.DataFrame()

        filtered_df = bank_df[bank_df[receipt_type_col].astype(str).str.contains("방문", na=False)].copy()

        if filtered_df.empty:
            st.warning("접수유형이 '방문'인 데이터가 없습니다.")
            return pd.DataFrame()

        # 3. 하나은행 시트 로드 (고객번호 → 사업자번호 매칭)
        try:
            load_csv_to_state("url_analysis", "analysis_lookup_df")
            hana_df = st.session_state.analysis_lookup_df.copy()
        except Exception as e:
            st.error(f"하나은행 시트를 불러올 수 없습니다: {str(e)}")
            return pd.DataFrame()

        # 하나은행 시트에서 고객번호와 사업자번호 컬럼 찾기
        hana_customer_col = find_col(hana_df, ["고객번호"], "고객번호")
        hana_biz_col = find_col(hana_df, ["사업자번호"], "사업자번호")

        if hana_customer_col not in hana_df.columns or hana_biz_col not in hana_df.columns:
            st.error("하나은행 시트에서 고객번호 또는 사업자번호 컬럼을 찾을 수 없습니다.")
            return pd.DataFrame()

        # 고객번호로 사업자번호 매칭
        hana_lookup = hana_df[[hana_customer_col, hana_biz_col]].drop_duplicates(subset=[hana_customer_col])
        hana_lookup.columns = ["고객번호_매칭", "사업자번호"]

        # 4. 본사 시트 로드 (사업자번호 → 개설완료일자, ERP연계일자 확인)
        try:
            load_csv_to_state("url_sync", "cloud_sheet_df")
            sync_df = st.session_state.cloud_sheet_df.copy()
        except Exception as e:
            st.error(f"본사 시트를 불러올 수 없습니다: {str(e)}")
            return pd.DataFrame()

        # 본사 시트에서 필요한 컬럼 찾기
        sync_biz_col = find_col(sync_df, ["사업자번호"], "사업자번호")
        sync_open_col = find_col(sync_df, ["개설완료일자"], "개설완료일자")
        sync_erp_col = find_col(sync_df, ["ERP연계일자"], "ERP연계일자")

        if sync_biz_col not in sync_df.columns:
            st.error("본사 시트에서 사업자번호 컬럼을 찾을 수 없습니다.")
            return pd.DataFrame()

        # 사업자번호 정규화
        sync_df[sync_biz_col] = normalize_biz(sync_df[sync_biz_col])
        sync_lookup = sync_df[[sync_biz_col, sync_open_col, sync_erp_col]].drop_duplicates(subset=[sync_biz_col]) if sync_open_col and sync_erp_col else sync_df[[sync_biz_col]].drop_duplicates(subset=[sync_biz_col])
        sync_lookup.columns = ["사업자번호_매칭", "개설완료일자", "ERP연계일자"] if sync_open_col and sync_erp_col else ["사업자번호_매칭"]

        # 5. 변환 로직 적용
        result_rows = []

        for idx, row in filtered_df.iterrows():
            customer_no = str(row.get(customer_no_col, "")).strip()
            receipt_date = row.get(receipt_date_col, "")
            company = row.get(company_col, "")
            manager = row.get(manager_col, "")
            request = str(row.get(request_col, "")).strip()

            # 고객번호로 사업자번호 찾기
            biz_no = ""
            matched_hana = hana_lookup[hana_lookup["고객번호_매칭"].astype(str).str.strip() == customer_no]
            if not matched_hana.empty:
                biz_no = str(matched_hana.iloc[0]["사업자번호"]).strip()

            # 사업자번호 정규화
            biz_no_normalized = normalize_biz(pd.Series([biz_no])).iloc[0]

            # 본사 시트에서 개설완료일자, ERP연계일자 확인
            open_date = ""
            erp_date = ""

            if biz_no_normalized and sync_open_col and sync_erp_col:
                matched_sync = sync_lookup[sync_lookup["사업자번호_매칭"] == biz_no_normalized]
                if not matched_sync.empty:
                    open_date = str(matched_sync.iloc[0].get("개설완료일자", "")).strip()
                    erp_date = str(matched_sync.iloc[0].get("ERP연계일자", "")).strip()

            # 활동상세 결정
            activity_detail = "운영"  # 기본값

            # 개설 조건: 요청사항에 "개설" or "구축" or "신규" 포함 AND 개설완료일자 공백
            if any(keyword in request for keyword in ["개설", "구축", "신규"]):
                if not open_date or open_date in ["", "nan", "None", "NaT"]:
                    activity_detail = "개설"

            # 연계 조건: 요청사항에 "연계" or "ERP" 포함 AND ERP연계일자 공백
            if any(keyword in request for keyword in ["연계", "ERP"]):
                if not erp_date or erp_date in ["", "nan", "None", "NaT"]:
                    activity_detail = "연계"

            # 활동일 포맷팅
            try:
                if isinstance(receipt_date, pd.Timestamp):
                    activity_date = receipt_date.strftime("%Y-%m-%d")
                else:
                    activity_date = pd.to_datetime(receipt_date).strftime("%Y-%m-%d")
            except:
                activity_date = str(receipt_date)

            result_rows.append({
                "등록자": manager,
                "활동일": activity_date,
                "활동상세": activity_detail,
                "업체명": company,
                "사업자번호": biz_no
            })

        result_df = pd.DataFrame(result_rows)

        return result_df

    except Exception as e:
        st.error(f"변환 중 오류가 발생했습니다: {str(e)}")
        return pd.DataFrame()


def render_converted_preview_editor(converted_preview_df, filters=None):
    if converted_preview_df is None or converted_preview_df.empty:
        return None

    st.markdown("#### 변환파일 미리보기")
    st.markdown(
        """
        <style>
        /* ── 변환파일 미리보기 다크모드 ───────────────────────────────────────
           canvas 에는 background-color 를 주지 않아야 filter 와 충돌하지 않음 */
        body[data-pms-theme="d"] [data-testid="stDataEditor"],
        body[data-pms-theme="d"] [data-testid="stDataEditor"] > div,
        body[data-pms-theme="d"] [data-testid="stDataEditor"] [role="grid"],
        body:has(#pms-d:checked) [data-testid="stDataEditor"],
        body:has(#pms-d:checked) [data-testid="stDataEditor"] > div,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="grid"] {
            background-color: #13131f !important;
            color: #e2e8f0 !important;
            border-color: #2d2d4a !important;
        }
        body[data-pms-theme="d"] [data-testid="stDataEditor"] canvas,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] canvas {
            filter: invert(1) hue-rotate(180deg) brightness(0.92) contrast(0.88) saturate(0.9) !important;
        }
        body[data-pms-theme="d"] [data-testid="stDataEditor"] input,
        body[data-pms-theme="d"] [data-testid="stDataEditor"] textarea,
        body[data-pms-theme="d"] [data-testid="stDataEditor"] [contenteditable="true"],
        body[data-pms-theme="d"] [data-testid="stDataEditor"] [data-baseweb="input"] input,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] input,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] textarea,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] [contenteditable="true"],
        body:has(#pms-d:checked) [data-testid="stDataEditor"] [data-baseweb="input"] input {
            background-color: #1e1e34 !important;
            color: #ffffff !important;
            caret-color: #ffffff !important;
            border-color: #6366f1 !important;
            -webkit-text-fill-color: #ffffff !important;
        }
        body[data-pms-theme="d"] [data-testid="stDataEditor"] input::selection,
        body[data-pms-theme="d"] [data-testid="stDataEditor"] textarea::selection,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] input::selection,
        body:has(#pms-d:checked) [data-testid="stDataEditor"] textarea::selection {
            background-color: #4f46e5 !important;
            color: #ffffff !important;
        }
        @media (prefers-color-scheme: dark) {
            body:has(#pms-s:checked) [data-testid="stDataEditor"],
            body:has(#pms-s:checked) [data-testid="stDataEditor"] > div,
            body:has(#pms-s:checked) [data-testid="stDataEditor"] [role="grid"] {
                background-color: #13131f !important;
                border-color: #2d2d4a !important;
            }
            body:has(#pms-s:checked) [data-testid="stDataEditor"] canvas {
                filter: invert(1) hue-rotate(180deg) brightness(0.92) contrast(0.88) saturate(0.9) !important;
            }
            body:has(#pms-s:checked) [data-testid="stDataEditor"] input,
            body:has(#pms-s:checked) [data-testid="stDataEditor"] textarea,
            body:has(#pms-s:checked) [data-testid="stDataEditor"] [contenteditable="true"],
            body:has(#pms-s:checked) [data-testid="stDataEditor"] [data-baseweb="input"] input {
                background-color: #1e1e34 !important;
                color: #ffffff !important;
                caret-color: #ffffff !important;
                border-color: #6366f1 !important;
                -webkit-text-fill-color: #ffffff !important;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    converted_preview_df = normalize_converted_history_df(converted_preview_df)
    data_key = "history_convert_preview_data"
    editor_source_df = st.session_state.get(data_key, converted_preview_df)
    if not isinstance(editor_source_df, pd.DataFrame) or list(editor_source_df.columns) != list(converted_preview_df.columns):
        editor_source_df = converted_preview_df
    editor_source_df = normalize_converted_history_df(editor_source_df)
    full_editor_df = editor_source_df.copy()
    active_filters = has_active_history_filters(filters)
    if active_filters:
        editor_source_df = apply_history_search_filters(editor_source_df, filters)
        st.caption(f"검색 결과 {len(editor_source_df):,}건 / 전체 {len(full_editor_df):,}건")
    editor_key = f"history_convert_preview_editor_{history_filter_signature(filters)}"
    # _sel 체크박스 컬럼 (맨 왼쪽) — 행 선택용, 저장 시 제거
    editor_source_df.insert(0, "_sel", False)
    edited_with_sel = st.data_editor(
        editor_source_df,
        key=editor_key,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        disabled=[col for col in editor_source_df.columns if col not in ["_sel", "활동일자", "활동구분", "활동상세"]],
        column_config={
            "_sel": st.column_config.CheckboxColumn("선택", default=False, width="small"),
            "지사": st.column_config.TextColumn("지사", disabled=True),
            "활동일자": st.column_config.TextColumn("활동일자"),
            "활동구분": st.column_config.SelectboxColumn("활동구분", options=["방문", "상담", "원격"], required=True),
            "활동상세": st.column_config.SelectboxColumn("활동상세", options=["운영", "개설", "연계"], required=True),
            "활동내역": st.column_config.TextColumn("활동내역"),
            "_is_manual": None,
        },
    )
    _selected_indices = edited_with_sel[edited_with_sel["_sel"] == True].index.tolist() if "_sel" in edited_with_sel.columns else []
    edited_preview_df = edited_with_sel.drop(columns=["_sel"], errors="ignore")
    if active_filters:
        analysis_df = full_editor_df.copy()
        common_index = [idx for idx in edited_preview_df.index if idx in analysis_df.index]
        if common_index:
            analysis_df.loc[common_index, edited_preview_df.columns] = edited_preview_df.loc[common_index, edited_preview_df.columns]
        analysis_df = normalize_converted_history_df(analysis_df)
    else:
        analysis_df = normalize_converted_history_df(edited_preview_df)
    st.session_state[data_key] = analysis_df
    if st.session_state.get("user_excel_source") in (None, "bank"):
        st.session_state.user_excel_data = analysis_df
        st.session_state.user_excel_source = "bank"

    # ── 업체명↔사업자번호 매핑 (구글시트 + 기존 미리보기 데이터) ──────────
    _comp_biz_map = {}
    _cloud_raw = st.session_state.get("cloud_sheet_df")
    if _cloud_raw is not None and not _cloud_raw.empty:
        _cc = clean_header_logic(_cloud_raw.copy())
        _ow = find_col(_cc, ["담당자", "등록자", "성명"])
        _cm = find_col(_cc, ["업체명", "고객명", "상호"])
        _bm = find_col(_cc, ["사업자번호"])
        if _cm and _bm:
            _sub = _cc if not _ow else _cc[_cc[_ow].astype(str).str.strip() == str(st.session_state.user_name).strip()]
            for _, _r in _sub.iterrows():
                _c = str(_r.get(_cm, "")).strip()
                _b = str(_r.get(_bm, "")).strip()
                if _c and _c not in _comp_biz_map:
                    _comp_biz_map[_c] = _b
    # 기존 미리보기 데이터에서도 보완
    _pv = st.session_state.get(data_key)
    if isinstance(_pv, pd.DataFrame) and not _pv.empty:
        _pc = find_col(_pv, ["업체명", "상호", "고객명"])
        _pb = find_col(_pv, ["사업자번호"])
        if _pc and _pb:
            for _, _r in _pv.iterrows():
                _c = str(_r.get(_pc, "")).strip()
                _b = str(_r.get(_pb, "")).strip()
                if _c and _c not in _comp_biz_map:
                    _comp_biz_map[_c] = _b
    _company_list = ["-- 업체명 선택 --"] + sorted(_comp_biz_map.keys())

    # ── 하단 버튼 행: [🗑️ 삭제] [spacer] [이력 추가 ＋] ──────────────────────
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    _del_col, _, _add_col = st.columns([1, 8, 2])
    with _del_col:
        _has_sel = len(_selected_indices) > 0
        if st.button("🗑️ 삭제", disabled=not _has_sel, key="del_selected_btn", use_container_width=True):
            _new_df = analysis_df.drop(index=_selected_indices).reset_index(drop=True)
            st.session_state[data_key] = normalize_converted_history_df(_new_df)
            st.rerun()
    with _add_col:
        if st.button("이력 추가 ＋", use_container_width=True, key="add_history_row"):
            st.session_state["show_add_history_form"] = not st.session_state.get("show_add_history_form", False)
            st.rerun()

    # ── 이력 추가 폼 ──────────────────────────────────────────────────────────
    if st.session_state.get("show_add_history_form", False):
        with st.container(border=True):
            st.markdown("##### 이력 추가")
            _fc1, _fc2 = st.columns([0.6, 0.4])
            with _fc1:
                _sel_comp = st.selectbox("업체명", _company_list, key="add_form_company")
                _sel_comp = "" if _sel_comp == "-- 업체명 선택 --" else _sel_comp
            with _fc2:
                _auto_biz = _comp_biz_map.get(_sel_comp, "")
                st.text_input("사업자번호 (자동입력)", value=_auto_biz, disabled=True, key="add_form_biz_display")

            _fd1, _fd2, _fd3 = st.columns(3)
            with _fd1:
                _form_date = st.text_input(
                    "활동일자",
                    value=st.session_state.get("add_form_date_val", (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d")),
                    key="add_form_date",
                )
            with _fd2:
                _form_cat = st.selectbox("활동구분", ["방문", "상담", "원격"], key="add_form_cat")
            with _fd3:
                _form_det = st.selectbox("활동상세", ["운영", "개설", "연계"], key="add_form_det")

            _form_note = st.text_input("활동내역", key="add_form_note")

            _fb1, _fb2 = st.columns(2)
            with _fb1:
                if st.button("등록", type="primary", use_container_width=True, key="add_form_submit"):
                    if not _sel_comp:
                        st.warning("업체명을 선택해주세요.")
                    else:
                        _det_title = {"운영": "운영방문", "개설": "개설 방문", "연계": "연계 방문"}.get(_form_det, "운영방문")
                        _new_row = {col: "" for col in analysis_df.columns}
                        _new_row.update({
                            "지사": "HANA지사",
                            "상품": "통합CMS",
                            "등록자": st.session_state.user_name,
                            "업체명": _sel_comp,
                            "사업자번호": _auto_biz,
                            "활동일자": _form_date,
                            "활동구분": _form_cat,
                            "활동상세": _form_det,
                            "제목": _det_title,
                            "활동내역": _form_note,
                            "_is_manual": True,
                        })
                        st.session_state[data_key] = normalize_converted_history_df(
                            pd.concat([analysis_df, pd.DataFrame([_new_row])], ignore_index=True)
                        )
                        # 등록 후 업체명·활동일자·활동내역 초기화
                        st.session_state["add_form_date_val"] = ""
                        st.session_state["add_form_company"] = "-- 업체명 선택 --"
                        for _rk in ["add_form_date", "add_form_note", "add_form_biz_display"]:
                            st.session_state.pop(_rk, None)
                        st.rerun()
            with _fb2:
                if st.button("취소", use_container_width=True, key="add_form_cancel"):
                    st.session_state["show_add_history_form"] = False
                    st.rerun()

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    return analysis_df


def dataframe_to_excel_bytes(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = str(sheet_name)[:31] or "Sheet"
            df.to_excel(writer, index=False, sheet_name=safe_name)
    output.seek(0)
    return output.getvalue()


def _render_staff_admin():
    st.caption("직원 계정과 메뉴 접근 권한을 관리합니다.")

    # 파일 기준으로 재로드 (_deleted 목록 포함 반영)
    st.session_state.user_db = load_user_db()

    if st.session_state.pop("reset_staff_edit_sel", False):
        st.session_state.staff_edit_sel = "선택안함"

    if not st.session_state.get("_staff_admin_action_initialized"):
        st.session_state.staff_admin_action = ""
        st.session_state["_staff_admin_action_initialized"] = True

    workplace_data = _load_delegated_workplaces()
    workplace_names = sorted(
        {
            str(site.get("workplace_name", "")).strip()
            for site in workplace_data.get("workplaces", [])
            if str(site.get("workplace_name", "")).strip()
        }
    )

    staff_rows = []
    for uid, info in st.session_state.user_db.items():
        if uid == "1":
            continue
        staff_rows.append({
            "ID": uid,
            "성명": info.get("name", ""),
            "메일주소": info.get("email", ""),
            "핸드폰": info.get("phone", ""),
            "사업자 정보": info.get("dept_type", "사업부"),
            "마스터 구분": "관리자" if info.get("role") == "관리자" else "사용자",
        })

    if staff_rows:
        # ID 순서로 정렬
        staff_rows.sort(key=lambda x: x["ID"])

        # ── 사용자 목록 HTML 테이블 표시 (다크모드 호환, 가운데 정렬) ──
        render_plain_html_table(pd.DataFrame(staff_rows), center_align=True)
    else:
        st.info("등록된 직원이 없습니다.")

    st.markdown("---")

    action_cols = st.columns([0.15, 0.15, 0.15, 0.55])
    with action_cols[0]:
        if st.button("사용자 추가", key="staff_action_add", use_container_width=True, type="primary"):
            st.session_state.staff_admin_action = "add"
            st.rerun()
    with action_cols[1]:
        if st.button("수정", key="staff_action_edit", use_container_width=True):
            st.session_state.staff_admin_action = "edit"
            st.rerun()
    with action_cols[2]:
        if st.button("삭제", key="staff_action_delete", use_container_width=True):
            st.session_state.staff_admin_action = "delete"
            st.rerun()

    action = st.session_state.get("staff_admin_action", "")

    if action == "add":
        st.markdown("#### 사용자 추가")
        if not workplace_names:
            st.warning("[사업장 정보 관리]에서 사업장을 먼저 등록해주세요.")
            return

        with st.form("add_staff_form"):
            ac1, ac2, ac3 = st.columns(3)
            with ac1:
                new_uid = st.text_input("ID", key="add_staff_uid")
                new_pw = st.text_input("비밀번호", type="password", key="add_staff_pw")
                new_name = st.text_input("성명", key="add_staff_name")
            with ac2:
                new_email = st.text_input("메일주소", key="add_staff_email")
                new_phone = st.text_input("핸드폰", key="add_staff_phone")
            with ac3:
                new_business = st.selectbox("사업자 정보", workplace_names, key="add_staff_business")
                new_role = st.selectbox("마스터 구분", ["관리자", "사용자"], key="add_staff_role")
            add_submitted = st.form_submit_button("추가", use_container_width=True, type="primary")

        if add_submitted:
            uid_str = new_uid.strip()
            name_str = new_name.strip()
            pw_str = new_pw.strip()
            if not uid_str or not name_str or not pw_str:
                st.error("아이디, 성명, 비밀번호는 필수 입력입니다.")
            elif uid_str in st.session_state.user_db:
                st.error(f"이미 존재하는 아이디입니다: {uid_str}")
            else:
                st.session_state.user_db[uid_str] = {
                    "pw": pw_str,
                    "name": name_str,
                    "email": new_email.strip(),
                    "phone": new_phone.strip(),
                    "access": "허용",
                    "role": new_role,
                    "dept_type": new_business,
                    "staff_type": "정규직",
                    "outsource": "아니오",
                    "outsource_period": "해당없음",
                }
                # _deleted 목록 유지
                _cur_file = {}
                if os.path.exists(DB_FILE):
                    try:
                        with open(DB_FILE, "r", encoding="utf-8") as _f:
                            _cur_file = json.load(_f)
                    except Exception:
                        pass
                save_data = dict(st.session_state.user_db)
                if "_deleted" in _cur_file:
                    save_data["_deleted"] = _cur_file["_deleted"]
                save_db(DB_FILE, save_data)
                st.session_state.staff_admin_action = ""
                st.success(f"직원 '{name_str}({uid_str})'을(를) 추가했습니다.")
                st.rerun()

        cancel_col, _ = st.columns([0.15, 0.85])
        with cancel_col:
            if st.button("취소", key="cancel_add_staff", use_container_width=True):
                st.session_state.staff_admin_action = ""
                st.rerun()

        return

    if action not in {"edit", "delete"}:
        return

    st.markdown("#### 직원 정보 수정" if action == "edit" else "#### 직원 삭제")

    uid_options = [f"{r['ID']} — {r['성명']}" for r in staff_rows]
    sel_col, _ = st.columns([0.32, 0.68])
    with sel_col:
        sel = st.selectbox("대상 직원 선택", ["선택안함"] + uid_options, key="staff_edit_sel")

    if sel == "선택안함":
        return

    sel_uid = sel.split(" — ")[0].strip()
    info = st.session_state.user_db.get(sel_uid, {})

    if action == "delete":
        st.warning(f"{sel} 계정을 삭제합니다. 삭제 후 기본 계정에서도 복원되지 않도록 처리됩니다.")
        del_col, cancel_col, _ = st.columns([0.15, 0.15, 0.7])
        with del_col:
            if st.button("삭제 실행", type="primary", use_container_width=True):
                del st.session_state.user_db[sel_uid]
                # _deleted 목록을 파일에 함께 저장해 기본 계정에서도 복원되지 않도록 함
                _cur_file = {}
                if os.path.exists(DB_FILE):
                    try:
                        with open(DB_FILE, "r", encoding="utf-8") as _f:
                            _cur_file = json.load(_f)
                    except Exception:
                        pass
                _deleted_set = set(_cur_file.get("_deleted", []))
                _deleted_set.add(sel_uid)
                save_data = dict(st.session_state.user_db)
                save_data["_deleted"] = list(_deleted_set)
                save_db(DB_FILE, save_data, allow_shrink=True)
                st.session_state.reset_staff_edit_sel = True
                st.session_state.staff_admin_action = ""
                st.success(f"{sel} 삭제 완료")
                time.sleep(0.5)
                st.rerun()
        with cancel_col:
            if st.button("취소", use_container_width=True):
                st.session_state.staff_admin_action = ""
                st.rerun()
        return

    st.markdown(f"**메일주소:** {info.get('email', '—')}")

    if not workplace_names:
        st.warning("[사업장 정보 관리]에서 사업장을 먼저 등록해주세요.")
        return

    business_opts = workplace_names

    c1, c2, c3, c4, c5 = st.columns([1, 1.3, 1, 1.3, 1])
    with c1:
        new_name = st.text_input("성명", value=info.get("name", ""), key="edit_staff_name")
    with c2:
        new_email = st.text_input("메일주소", value=info.get("email", ""), key="edit_staff_email")
    with c3:
        new_phone = st.text_input("핸드폰", value=info.get("phone", ""), key="edit_staff_phone")
    with c4:
        new_dept_type = st.selectbox(
            "사업자 정보",
            business_opts,
            index=business_opts.index(info.get("dept_type", "")) if info.get("dept_type", "") in business_opts else 0,
            key="edit_dept_type",
        )
    with c5:
        new_access = st.selectbox("로그인 허용 여부", ["허용", "불가"],
                                  index=["허용", "불가"].index(info.get("access", "불가")),
                                  key="edit_access")
        new_role = st.selectbox("마스터 구분", ["관리자", "사용자"],
                                index=0 if info.get("role") == "관리자" else 1,
                                key="edit_role")

    bc1, bc2, _ = st.columns([0.15, 0.15, 0.7])
    with bc1:
        if st.button("저장", type="primary", use_container_width=True):
            st.session_state.user_db[sel_uid]["name"] = new_name.strip()
            st.session_state.user_db[sel_uid]["email"] = new_email.strip()
            st.session_state.user_db[sel_uid]["phone"] = new_phone.strip()
            st.session_state.user_db[sel_uid]["dept_type"] = new_dept_type
            st.session_state.user_db[sel_uid]["access"] = new_access
            st.session_state.user_db[sel_uid]["role"] = new_role
            # _deleted 목록 유지
            _cur_file = {}
            if os.path.exists(DB_FILE):
                try:
                    with open(DB_FILE, "r", encoding="utf-8") as _f:
                        _cur_file = json.load(_f)
                except Exception:
                    pass
            save_data = dict(st.session_state.user_db)
            if "_deleted" in _cur_file:
                save_data["_deleted"] = _cur_file["_deleted"]
            save_db(DB_FILE, save_data)
            st.session_state.reset_staff_edit_sel = True
            st.session_state.staff_admin_action = ""
            st.success("저장 완료")
            time.sleep(0.5)
            st.rerun()
    with bc2:
        if st.button("취소", use_container_width=True):
            st.session_state.staff_admin_action = ""
            st.session_state.reset_staff_edit_sel = True
            st.rerun()


def _load_server_connection():
    installed_drivers = _get_installed_odbc_drivers()
    default_data = {
        "server_host": "",
        "server_port": 1433,
        "driver_name": _preferred_mssql_driver(installed_drivers),
        "database_name": "",
        "auth_type": "SQL Server 인증",
        "username": "",
        "password": "",
        "encrypt": True,
        "trust_server_certificate": False,
        "connection_timeout": 30,
        "updated_at": "",
        "updated_by": "",
    }
    if not os.path.exists(SERVER_CONNECTION_FILE):
        return default_data
    try:
        with open(SERVER_CONNECTION_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return default_data
    if not isinstance(data, dict):
        return default_data
    merged = {**default_data, **data}
    if installed_drivers and merged.get("driver_name") not in installed_drivers:
        merged["driver_name"] = _preferred_mssql_driver(installed_drivers)
    return merged


def _save_server_connection(data):
    with open(SERVER_CONNECTION_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def _mask_secret(value):
    if not value:
        return "미등록"
    return "●" * min(max(len(str(value)), 6), 12)


def show_server_connection_info():
    st.markdown("### 서버 접속 정보")
    st.caption("내부 관리에서 사용할 MSSQL 서버 접속 정보를 관리합니다.")

    config = _load_server_connection()
    installed_drivers = _get_installed_odbc_drivers()
    driver_options = installed_drivers.copy()
    current_driver = config.get("driver_name", "")
    if current_driver and current_driver not in driver_options and not installed_drivers:
        driver_options.insert(0, current_driver)
    if not driver_options:
        driver_options = ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "FreeTDS"]

    with st.form("mssql_connection_form"):
        col_a, col_b = st.columns([3, 1])
        server_host = col_a.text_input(
            "MSSQL 서버 주소",
            value=config.get("server_host", ""),
            placeholder="예: 10.10.10.20 또는 db.company.local",
        )
        server_port = col_b.number_input(
            "포트",
            min_value=1,
            max_value=65535,
            value=int(config.get("server_port", 1433) or 1433),
            step=1,
        )

        database_name = st.text_input(
            "데이터베이스명",
            value=config.get("database_name", ""),
            placeholder="예: SalesManagement",
        )
        driver_name = st.selectbox(
            "ODBC 드라이버명",
            driver_options,
            index=driver_options.index(current_driver) if current_driver in driver_options else 0,
        )
        if installed_drivers:
            st.caption(f"현재 서버에서 감지된 ODBC 드라이버: {', '.join(installed_drivers)}")
        else:
            st.warning("현재 서버에서 감지된 ODBC 드라이버가 없습니다. MSSQL 접속 전 시스템 드라이버 설치가 필요합니다.")

        col_c, col_d = st.columns(2)
        auth_type = col_c.selectbox(
            "인증 방식",
            ["SQL Server 인증", "Windows 인증"],
            index=0 if config.get("auth_type", "SQL Server 인증") == "SQL Server 인증" else 1,
        )
        connection_timeout = col_d.number_input(
            "접속 제한 시간(초)",
            min_value=1,
            max_value=300,
            value=int(config.get("connection_timeout", 30) or 30),
            step=1,
        )

        username = st.text_input("계정", value=config.get("username", ""), placeholder="예: sales_app")
        password = st.text_input(
            "비밀번호",
            value=config.get("password", ""),
            type="password",
            placeholder="MSSQL 접속 비밀번호",
        )

        col_e, col_f = st.columns(2)
        encrypt = col_e.checkbox("암호화 연결 사용", value=bool(config.get("encrypt", True)))
        trust_server_certificate = col_f.checkbox(
            "서버 인증서 신뢰",
            value=bool(config.get("trust_server_certificate", False)),
        )

        saved = st.form_submit_button("서버 접속 정보 저장", type="primary")

    if saved:
        if not server_host.strip() or not database_name.strip():
            st.warning("MSSQL 서버 주소와 데이터베이스명을 입력해주세요.")
        elif auth_type == "SQL Server 인증" and (not username.strip() or not password):
            st.warning("SQL Server 인증을 사용하려면 계정과 비밀번호를 입력해주세요.")
        else:
            payload = {
                "server_host": server_host.strip(),
                "server_port": int(server_port),
                "driver_name": str(driver_name).strip() or _preferred_mssql_driver(installed_drivers),
                "database_name": database_name.strip(),
                "auth_type": auth_type,
                "username": username.strip(),
                "password": password,
                "encrypt": bool(encrypt),
                "trust_server_certificate": bool(trust_server_certificate),
                "connection_timeout": int(connection_timeout),
                "updated_at": (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S"),
                "updated_by": st.session_state.get("user_name", ""),
            }
            _save_server_connection(payload)
            try:
                with _connect_mssql(payload) as conn:
                    _ensure_sales_registration_table(conn)
                st.session_state["_server_connection_notice"] = (
                    "success",
                    "MSSQL 서버 접속 정보가 저장되었고, 접속 및 테이블 확인까지 완료되었습니다.",
                )
            except Exception as exc:
                st.session_state["_server_connection_notice"] = (
                    "warning",
                    f"MSSQL 서버 접속 정보는 저장되었지만 연결 확인에 실패했습니다: {exc}",
                )
            st.rerun()

    st.divider()
    notice = st.session_state.pop("_server_connection_notice", None)
    if notice:
        level, text = notice
        if level == "success":
            st.success(text)
        else:
            st.warning(text)

    test_col, _ = st.columns([1, 3])
    with test_col:
        if st.button("테이블 생성/접속 테스트", use_container_width=True):
            try:
                latest_for_test = _load_server_connection()
                with _connect_mssql(latest_for_test) as conn:
                    _ensure_sales_registration_table(conn)
                st.success("MSSQL 접속 및 dbo.sales_registrations 테이블 확인이 완료되었습니다.")
            except Exception as exc:
                st.error(f"MSSQL 접속 또는 테이블 생성에 실패했습니다: {exc}")

    latest = _load_server_connection()
    st.markdown("#### 현재 등록 정보")
    summary_df = pd.DataFrame(
        [
            {"항목": "서버 주소", "값": latest.get("server_host") or "미등록"},
            {"항목": "포트", "값": latest.get("server_port") or "미등록"},
            {"항목": "ODBC 드라이버명", "값": latest.get("driver_name") or "미등록"},
            {"항목": "데이터베이스명", "값": latest.get("database_name") or "미등록"},
            {"항목": "인증 방식", "값": latest.get("auth_type") or "미등록"},
            {"항목": "계정", "값": latest.get("username") or "미등록"},
            {"항목": "비밀번호", "값": _mask_secret(latest.get("password"))},
            {"항목": "암호화 연결", "값": "사용" if latest.get("encrypt") else "미사용"},
            {"항목": "서버 인증서 신뢰", "값": "사용" if latest.get("trust_server_certificate") else "미사용"},
            {"항목": "접속 제한 시간", "값": f"{latest.get('connection_timeout') or 30}초"},
            {"항목": "최종 수정", "값": latest.get("updated_at") or "미등록"},
            {"항목": "수정자", "값": latest.get("updated_by") or "미등록"},
        ]
    )
    st.dataframe(summary_df, use_container_width=True, hide_index=True)


def inject_theme_toggle():
    # JavaScript는 st.markdown innerHTML으로 실행 불가 → CSS :has() + radio 버튼 방식 사용
    st.markdown("""
    <style>
    .pms-sw-outer {
        position: fixed; top: 14px; right: 18px; z-index: 999999;
    }
    .pms-sw-track {
        display: inline-flex; align-items: center;
        background: #e5e7eb; border-radius: 22px; padding: 3px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.12); gap: 2px;
    }
    .pms-btn {
        position: relative; cursor: pointer;
        padding: 0 10px; height: 28px; border-radius: 18px;
        line-height: 28px; opacity: 0.45; color: #374151;
        display: flex; align-items: center; justify-content: center;
        transition: opacity 0.2s, background 0.15s, color 0.15s;
        user-select: none;
    }
    .pms-btn:hover { opacity: 0.75; }
    /* radio를 label 전체에 투명하게 덮어서 한 번 클릭에 즉시 체크 */
    .pms-theme-radio {
        position: absolute; inset: 0;
        opacity: 0; cursor: pointer; margin: 0;
    }
    .pms-btn:has(input:checked) {
        opacity: 1; background: #2F6FED; color: white;
    }

    /* ══ 다크 모드 컬러 팔레트 ══
       배경:    #1e1e2e  표면:   #252535  카드:    #2a2a3e
       테두리:  #45475a  텍스트: #cdd6f4  보조:    #a6adc8
       강조:    #89b4fa  성공:   #a6e3a1  경고:    #f9e2af
       오류:    #f38ba8  정보:   #89dceb  사이드:  #16162a  */

    /* 기본 배경·텍스트 — 모든 글씨 흰색 */
    body:has(#pms-d:checked) .stApp,
    body:has(#pms-d:checked) [data-testid="stAppViewContainer"],
    body:has(#pms-d:checked) [data-testid="stMain"],
    body:has(#pms-d:checked) section.main {
        background-color: #1e1e2e !important; color: #ffffff !important;
    }
    body:has(#pms-d:checked) .main .block-container { background-color: #1e1e2e !important; }

    /* 텍스트 — 전체 흰색 */
    body:has(#pms-d:checked) h1, body:has(#pms-d:checked) h2, body:has(#pms-d:checked) h3,
    body:has(#pms-d:checked) h4, body:has(#pms-d:checked) h5, body:has(#pms-d:checked) h6 {
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) p, body:has(#pms-d:checked) li,
    body:has(#pms-d:checked) [data-testid="stMarkdownContainer"],
    body:has(#pms-d:checked) [data-testid="stMarkdownContainer"] p {
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) label { color: #ffffff !important; }
    /* span: 인라인 style 색상이 있는 셀(대비/증감)은 제외하고 흰색 적용 */
    body:has(#pms-d:checked) span:not([style*="color"]) { color: #ffffff !important; }
    body:has(#pms-d:checked) [data-testid="stMarkdownContainer"] span:not([style*="color"]) { color: #ffffff !important; }
    body:has(#pms-d:checked) [data-testid="stSidebar"] span { color: #ffffff !important; }
    body:has(#pms-d:checked) .stButton span { color: #ffffff !important; }
    body:has(#pms-d:checked) [data-baseweb="tab"] span { color: inherit !important; }
    body:has(#pms-d:checked) [data-testid="stMetricValue"] span { color: #ffffff !important; }
    body:has(#pms-d:checked) [data-testid="stMetricLabel"] span { color: #ffffff !important; }

    /* 파일 업로더 다크모드 — 모든 자식 요소 포함 */
    body:has(#pms-d:checked) [data-testid="stFileUploader"],
    body:has(#pms-d:checked) [data-testid="stFileUploader"] > div,
    body:has(#pms-d:checked) [data-testid="stFileUploader"] > div > div {
        background-color: #252535 !important;
    }
    body:has(#pms-d:checked) [data-testid="stFileUploaderDropzone"],
    body:has(#pms-d:checked) [data-testid="stFileUploader"] section,
    body:has(#pms-d:checked) [data-testid="stFileUploader"] [class*="uploadDropzone"],
    body:has(#pms-d:checked) [data-testid="stFileUploader"] [class*="drop"] {
        background-color: #1e1e30 !important;
        border: 2px dashed #45475a !important;
        border-radius: 8px !important;
    }
    body:has(#pms-d:checked) [data-testid="stFileUploaderDropzone"] *,
    body:has(#pms-d:checked) [data-testid="stFileUploader"] section * {
        color: #ffffff !important;
        background-color: transparent !important;
    }
    body:has(#pms-d:checked) [data-testid="stFileUploaderDropzone"] button,
    body:has(#pms-d:checked) [data-testid="stFileUploader"] section button {
        background-color: #313244 !important; color: #ffffff !important;
        border: 1px solid #45475a !important; border-radius: 6px !important;
    }
    body:has(#pms-d:checked) [data-testid="stFileUploaderDropzone"] button:hover,
    body:has(#pms-d:checked) [data-testid="stFileUploader"] section button:hover {
        background-color: #45475a !important;
    }
    body:has(#pms-d:checked) [data-testid="stFileUploaderFileData"],
    body:has(#pms-d:checked) [data-testid="stFileUploader"] [data-testid="stFileUploaderFileData"] {
        background-color: #252535 !important; border-color: #45475a !important;
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) small, body:has(#pms-d:checked) [data-testid="stCaptionContainer"] {
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) hr { border-color: #45475a !important; }

    /* 사이드바 — Zoho 네이비 톤은 다크모드에서도 동일 유지 */
    body:has(#pms-d:checked) [data-testid="stSidebar"],
    body:has(#pms-d:checked) [data-testid="stSidebarContent"] {
        background-color: #16284A !important;
    }
    body:has(#pms-d:checked) [data-testid="stSidebar"] p,
    body:has(#pms-d:checked) [data-testid="stSidebar"] span,
    body:has(#pms-d:checked) [data-testid="stSidebar"] label {
        color: #E8EEF8 !important;
    }
    body:has(#pms-d:checked) [data-testid="stSidebar"] .stButton > button {
        background-color: transparent !important; color: #C9D6EC !important;
        border-color: transparent !important; box-shadow: none !important;
    }
    body:has(#pms-d:checked) [data-testid="stSidebar"] .stButton > button:hover {
        background-color: #233A61 !important;
    }

    /* 입력 컨트롤 */
    body:has(#pms-d:checked) input,
    body:has(#pms-d:checked) textarea,
    body:has(#pms-d:checked) [data-baseweb="input"] input,
    body:has(#pms-d:checked) [data-baseweb="textarea"] textarea {
        background-color: #252535 !important; color: #ffffff !important;
        border-color: #45475a !important;
    }
    body:has(#pms-d:checked) [data-baseweb="input"],
    body:has(#pms-d:checked) [data-baseweb="textarea"] {
        background-color: #252535 !important; border-color: #45475a !important;
    }
    body:has(#pms-d:checked) [data-baseweb="select"] > div,
    body:has(#pms-d:checked) [data-baseweb="select"] div[role="combobox"] {
        background-color: #252535 !important; border-color: #45475a !important; color: #ffffff !important;
    }
    /* 셀렉트박스 드롭다운 팝업 — 전체 컨테이너 및 모든 하위 요소 */
    body:has(#pms-d:checked) [data-baseweb="popover"],
    body:has(#pms-d:checked) [data-baseweb="popover"] > div,
    body:has(#pms-d:checked) [data-baseweb="popover"] > div > div,
    body:has(#pms-d:checked) [data-baseweb="popover"] > div > div > div {
        background-color: #252535 !important;
        border-color: #45475a !important;
    }
    body:has(#pms-d:checked) [data-baseweb="menu"],
    body:has(#pms-d:checked) [data-baseweb="menu"] ul,
    body:has(#pms-d:checked) [role="listbox"] {
        background-color: #252535 !important;
        border-color: #45475a !important;
    }
    body:has(#pms-d:checked) [data-baseweb="menu"] li,
    body:has(#pms-d:checked) [data-baseweb="menu"] [role="option"],
    body:has(#pms-d:checked) [role="option"],
    body:has(#pms-d:checked) [role="listbox"] > div {
        background-color: #252535 !important;
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-baseweb="menu"] li span,
    body:has(#pms-d:checked) [data-baseweb="menu"] [role="option"] span,
    body:has(#pms-d:checked) [role="option"] span {
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-baseweb="menu"] li:hover,
    body:has(#pms-d:checked) [data-baseweb="menu"] [aria-selected="true"],
    body:has(#pms-d:checked) [role="option"]:hover,
    body:has(#pms-d:checked) [role="option"][aria-selected="true"] {
        background-color: #3a3a5e !important;
    }
    body:has(#pms-d:checked) [data-testid="stNumberInput"] button {
        background-color: #313244 !important; color: #ffffff !important;
        border-color: #45475a !important;
    }

    /* 버튼 */
    body:has(#pms-d:checked) .stButton > button,
    body:has(#pms-d:checked) [data-testid="stButton"] button,
    body:has(#pms-d:checked) button[data-testid^="baseButton"] {
        background: #2a2a3e !important;
        background-color: #2a2a3e !important; color: #ffffff !important;
        border-color: #45475a !important;
    }
    body:has(#pms-d:checked) .stButton > button:hover,
    body:has(#pms-d:checked) [data-testid="stButton"] button:hover,
    body:has(#pms-d:checked) button[data-testid^="baseButton"]:hover {
        background: #313244 !important;
        background-color: #313244 !important; border-color: #ffffff !important;
    }
    body:has(#pms-d:checked) .stButton > button[kind="primary"],
    body:has(#pms-d:checked) .stButton > button[data-testid="baseButton-primary"] {
        background-color: #4F46E5 !important; color: #ffffff !important;
        border-color: #6366f1 !important;
    }
    body:has(#pms-d:checked) .stButton > button[kind="primary"]:hover {
        background-color: #6366f1 !important;
    }

    /* 홈 버튼 — 아이콘 전용: 배경·테두리 없음 */
    button.pms-home-btn,
    button.pms-home-btn:hover {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 82px !important;
        height: 82px !important;
        padding: 0 !important;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: transparent !important;
        font-size: 0 !important;
    }
    button.pms-home-btn *,
    button.pms-home-btn:hover * {
        color: transparent !important;
        display: none !important;
        font-size: 0 !important;
    }
    button.pms-home-btn::before {
        content: "";
        width: 75px;
        height: 75px;
        display: block;
        flex: 0 0 auto;
        background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 96 96'%3E%3Cpath fill='%23000' d='M32 8h56v56c0 13.255-10.745 24-24 24H8V32C8 18.745 18.745 8 32 8Z'/%3E%3Cpath fill='%23fff' d='M20 52 48 26l28 26h-10v26H30V52H20Z'/%3E%3Cpath fill='%23fff' d='M58 34h12v18H58z'/%3E%3Cpath fill='%23000' d='M41 52h8v8h-8zM53 52h8v8h-8zM41 64h8v8h-8zM53 64h8v8h-8z'/%3E%3C/svg%3E") center / contain no-repeat;
    }
    button.pms-home-btn:hover::before {
        filter: opacity(0.86);
    }
    body:has(#pms-d:checked) button.pms-home-btn {
        background: transparent !important;
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }
    body:has(#pms-d:checked) button.pms-home-btn::before {
        filter: none;
    }
    body:has(#pms-d:checked) button.pms-home-btn:hover {
        background: transparent !important;
        background-color: transparent !important;
        border: none !important;
    }
    body:has(#pms-d:checked) button.pms-home-btn:hover::before {
        filter: opacity(0.86);
    }

    /* 새로고침 버튼 다크모드 */
    body:has(#pms-d:checked) button.pms-refresh-btn {
        background: #2a2a3e !important;
        background-color: #2a2a3e !important;
        color: #cdd6f4 !important;
        border-color: #45475a !important;
    }
    body:has(#pms-d:checked) button.pms-refresh-btn:hover {
        background: #313244 !important;
        background-color: #313244 !important;
        color: #ffffff !important;
    }

    /* 탭 */
    body:has(#pms-d:checked) [data-baseweb="tab-list"] {
        background-color: #1e1e2e !important; border-bottom: 2px solid #45475a !important;
    }
    body:has(#pms-d:checked) [data-baseweb="tab"] {
        background-color: transparent !important; color: rgba(255,255,255,0.55) !important;
    }
    body:has(#pms-d:checked) [data-baseweb="tab"][aria-selected="true"] {
        color: #ffffff !important; font-weight: 700;
    }
    body:has(#pms-d:checked) [data-baseweb="tab-highlight"] {
        background-color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-baseweb="tab-border"] {
        background-color: #45475a !important;
    }

    /* 메트릭 카드 */
    body:has(#pms-d:checked) [data-testid="metric-container"] {
        background: #252535 !important; border: 1px solid #45475a !important; border-radius: 8px;
    }
    body:has(#pms-d:checked) [data-testid="stMetricValue"] {
        color: #ffffff !important; font-weight: 700;
    }
    body:has(#pms-d:checked) [data-testid="stMetricLabel"] { color: #ffffff !important; }
    body:has(#pms-d:checked) [data-testid="stMetricDeltaIcon-Up"]   { color: #a6e3a1 !important; }
    body:has(#pms-d:checked) [data-testid="stMetricDeltaIcon-Down"] { color: #f38ba8 !important; }

    /* Plotly 차트 */
    body:has(#pms-d:checked) .js-plotly-plot,
    body:has(#pms-d:checked) .js-plotly-plot .plot-container,
    body:has(#pms-d:checked) .js-plotly-plot .svg-container {
        background: transparent !important;
    }
    body:has(#pms-d:checked) .js-plotly-plot .bglayer rect.bg {
        fill: transparent !important;
        stroke: transparent !important;
    }
    body:has(#pms-d:checked) .js-plotly-plot .legend rect.bg {
        fill: #252535 !important;
        stroke: #45475a !important;
    }
    body:has(#pms-d:checked) .js-plotly-plot text {
        fill: #ffffff !important;
    }
    body:has(#pms-d:checked) .js-plotly-plot .gridlayer path {
        stroke: #45475a !important;
    }
    body:has(#pms-d:checked) .js-plotly-plot .zerolinelayer path,
    body:has(#pms-d:checked) .js-plotly-plot .xlines-below path,
    body:has(#pms-d:checked) .js-plotly-plot .ylines-below path {
        stroke: #585b70 !important;
    }
    body:has(#pms-d:checked) .js-plotly-plot .modebar-btn path {
        fill: #cdd6f4 !important;
    }

    /* 알림 박스 */
    body:has(#pms-d:checked) [data-testid="stNotification"],
    body:has(#pms-d:checked) .stAlert { border-radius: 6px !important; }
    body:has(#pms-d:checked) [data-testid="stNotification"][kind="info"],
    body:has(#pms-d:checked) .stAlert.stInfo {
        background-color: #1a2a3a !important; border-left: 4px solid #89dceb !important; color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-testid="stNotification"][kind="success"],
    body:has(#pms-d:checked) .stAlert.stSuccess {
        background-color: #1a2e1a !important; border-left: 4px solid #a6e3a1 !important; color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-testid="stNotification"][kind="warning"],
    body:has(#pms-d:checked) .stAlert.stWarning {
        background-color: #2e2a1a !important; border-left: 4px solid #f9e2af !important; color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-testid="stNotification"][kind="error"],
    body:has(#pms-d:checked) .stAlert.stError {
        background-color: #2e1a1a !important; border-left: 4px solid #f38ba8 !important; color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-testid="stNotification"] p,
    body:has(#pms-d:checked) .stAlert p { color: #ffffff !important; }

    /* 익스팬더 */
    body:has(#pms-d:checked) [data-testid="stExpander"] {
        background-color: #252535 !important; border: 1px solid #45475a !important; border-radius: 6px;
    }
    body:has(#pms-d:checked) [data-testid="stExpander"] summary {
        background-color: #252535 !important; color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-testid="stExpander"] summary:hover {
        background-color: #313244 !important;
    }

    /* ─── AG Grid 완전 다크모드 ─── */

    /* 1. CSS 변수 (AG Grid가 읽는 경우 적용) */
    body:has(#pms-d:checked) [class*="ag-theme"] {
        --ag-background-color: #252535;
        --ag-foreground-color: #ffffff;
        --ag-secondary-foreground-color: #ffffff;
        --ag-border-color: #45475a;
        --ag-secondary-border-color: #313244;
        --ag-row-border-color: #313244;
        --ag-header-background-color: #0f0f1f;
        --ag-header-foreground-color: #ffffff;
        --ag-header-column-separator-color: #45475a;
        --ag-odd-row-background-color: #1e1e30;
        --ag-row-hover-color: #2d2d45;
        --ag-selected-row-background-color: #3a3a5e;
        --ag-range-selection-background-color: rgba(255,255,255,0.1);
        --ag-input-focus-border-color: #ffffff;
        --ag-cell-horizontal-border: solid #313244;
        --ag-font-size: 13px;
        --ag-data-color: #ffffff;
        --ag-alpine-active-color: #4F46E5;
    }

    /* 2. 최외곽 래퍼 + 내부 모든 div 배경 강제 */
    body:has(#pms-d:checked) [data-testid="stDataFrame"],
    body:has(#pms-d:checked) [data-testid="stDataEditor"],
    body:has(#pms-d:checked) [data-testid="stDataFrame"] > div,
    body:has(#pms-d:checked) [data-testid="stDataFrame"] > div > div,
    body:has(#pms-d:checked) [data-testid="stDataFrame"] > div > div > div,
    body:has(#pms-d:checked) [data-testid="stDataEditor"] > div,
    body:has(#pms-d:checked) [data-testid="stDataEditor"] > div > div,
    body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="grid"] {
        background-color: #13131f !important;
        border-color: #2d2d4a !important;
        color: #e2e8f0 !important;
    }
    /* stDataEditor canvas — invert 필터로 라이트 렌더링을 다크로 반전 */
    body:has(#pms-d:checked) [data-testid="stDataEditor"] canvas {
        filter: invert(1) hue-rotate(180deg) brightness(0.92) contrast(0.88) saturate(0.9) !important;
    }
    /* stDataEditor 인라인 편집 입력창 */
    body:has(#pms-d:checked) [data-testid="stDataEditor"] input,
    body:has(#pms-d:checked) [data-testid="stDataEditor"] textarea,
    body:has(#pms-d:checked) [data-testid="stDataEditor"] [contenteditable="true"],
    body:has(#pms-d:checked) [data-testid="stDataEditor"] [data-baseweb="input"] input {
        background-color: #1e1e34 !important;
        color: #ffffff !important;
        caret-color: #ffffff !important;
        border-color: #6366f1 !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    /* 3. ag-theme 컨테이너 전체 */
    body:has(#pms-d:checked) [class*="ag-theme"],
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-root-wrapper,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-root-wrapper-body,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-root,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-body,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-body-viewport,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-body-horizontal-scroll,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-body-horizontal-scroll-viewport,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-center-cols-clipper,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-center-cols-viewport,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-center-cols-container,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-pinned-left-cols-container,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-pinned-right-cols-container,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-full-width-container,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-floating-top,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-floating-bottom,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-overlay {
        background-color: #252535 !important;
        border-color: #45475a !important;
    }

    /* 4. 헤더 */
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-row,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-viewport,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-container,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-pinned-left-header,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-pinned-right-header {
        background-color: #0f0f1f !important;
        border-bottom: 2px solid #45475a !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-cell,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-cell-comp-wrapper,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-group-cell {
        background-color: #0f0f1f !important;
        color: #ffffff !important;
        border-right: 1px solid #45475a !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-cell-text,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-cell-label {
        color: #ffffff !important;
        font-weight: 700 !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-sort-indicator-icon,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-header-icon {
        color: #ffffff !important;
    }

    /* 5. 행 — .ag-row 자체에 배경 지정 (ag-row-even 없을 때 대비) */
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-row {
        background-color: #252535 !important;
        border-bottom-color: #313244 !important;
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-row-odd {
        background-color: #1e1e30 !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-row:hover,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-row-hover {
        background-color: #2d2d45 !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-row-selected {
        background-color: #3a3a5e !important;
    }

    /* 6. 셀 */
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-cell {
        border-right-color: #313244 !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-cell:not([style*="color"]),
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-cell-value:not([style*="color"]) {
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-cell span:not([style*="color"]) {
        color: #ffffff !important;
    }

    /* 7. role 속성 기반 (columnheader / gridcell) */
    body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="columnheader"],
    body:has(#pms-d:checked) [data-testid="stDataFrame"]  [role="columnheader"] {
        background-color: #0f0f1f !important;
        color: #ffffff !important;
        border-bottom-color: #45475a !important;
        border-right-color: #45475a !important;
    }
    body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="columnheader"] *,
    body:has(#pms-d:checked) [data-testid="stDataFrame"]  [role="columnheader"] * {
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="gridcell"]:not([style*="color"]),
    body:has(#pms-d:checked) [data-testid="stDataFrame"]  [role="gridcell"]:not([style*="color"]),
    body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="gridcell"]:not([style*="color"]) *,
    body:has(#pms-d:checked) [data-testid="stDataFrame"]  [role="gridcell"]:not([style*="color"]) * {
        color: #ffffff !important;
    }

    /* 8. 셀 편집 팝업 (SelectboxColumn 등) */
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-popup,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-popup-editor,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-popup-child {
        background-color: #252535 !important; border-color: #45475a !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-rich-select,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-rich-select-list {
        background-color: #252535 !important; border-color: #45475a !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-rich-select-row,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-virtual-list-item {
        background-color: #252535 !important; color: #ffffff !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-rich-select-row:hover,
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-rich-select-row-selected {
        background-color: #3a3a5e !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-text-field-input {
        background-color: #252535 !important; color: #ffffff !important;
        border-color: #45475a !important;
    }
    body:has(#pms-d:checked) [class*="ag-theme"] .ag-checkbox-input-wrapper::after {
        color: #ffffff !important;
    }

    /* HTML 테이블 */
    body:has(#pms-d:checked) table { background-color: #252535 !important; border-color: #45475a !important; }
    body:has(#pms-d:checked) th {
        background-color: #0f0f1f !important; color: #ffffff !important;
        border-color: #45475a !important; font-weight: 700 !important;
    }
    body:has(#pms-d:checked) td {
        background-color: #252535 !important; color: #ffffff !important;
        border-color: #313244 !important;
    }
    body:has(#pms-d:checked) tr:nth-child(odd) td { background-color: #1e1e30 !important; }
    body:has(#pms-d:checked) tr:hover td { background-color: #2d2d45 !important; }

    /* 증감 색 — 다크모드에서 더 선명하게 */
    body:has(#pms-d:checked) [style*="color:#E53E3E"] { color: #ff7b7b !important; }
    body:has(#pms-d:checked) [style*="color:#3182CE"] { color: #74b9ff !important; }

    /* 체크박스 / 라디오 */
    body:has(#pms-d:checked) [data-baseweb="checkbox"] div,
    body:has(#pms-d:checked) [data-baseweb="radio"] div { border-color: #ffffff !important; }
    body:has(#pms-d:checked) [data-baseweb="checkbox"] [aria-checked="true"] div,
    body:has(#pms-d:checked) [data-baseweb="radio"] [aria-checked="true"] div {
        background-color: #ffffff !important;
    }

    /* 스크롤바 */
    body:has(#pms-d:checked) ::-webkit-scrollbar { width: 6px; height: 6px; }
    body:has(#pms-d:checked) ::-webkit-scrollbar-track { background: #252535; }
    body:has(#pms-d:checked) ::-webkit-scrollbar-thumb { background: #45475a; border-radius: 3px; }
    body:has(#pms-d:checked) ::-webkit-scrollbar-thumb:hover { background: #585b70; }

    /* 테마 토글 바 */
    body:has(#pms-d:checked) .pms-sw-track { background: #313244; }
    body:has(#pms-d:checked) .pms-btn { color: #ffffff; }

    /* Streamlit canvas-backed tables */
    html[data-pms-theme="d"] [data-testid="stDataEditor"],
    html[data-pms-theme="d"] [data-testid="stDataFrame"],
    body[data-pms-theme="d"] [data-testid="stDataEditor"],
    body[data-pms-theme="d"] [data-testid="stDataFrame"],
    .stApp:has(#pms-d:checked) [data-testid="stDataEditor"],
    .stApp:has(#pms-d:checked) [data-testid="stDataFrame"],
    body:has(#pms-d:checked) [data-testid="stDataEditor"],
    body:has(#pms-d:checked) [data-testid="stDataFrame"] {
        background-color: #1e1e2e !important;
        border-color: #45475a !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        overflow: hidden !important;
    }
    html[data-pms-theme="d"] [data-testid="stDataEditor"] canvas,
    html[data-pms-theme="d"] [data-testid="stDataFrame"] canvas,
    body[data-pms-theme="d"] [data-testid="stDataEditor"] canvas,
    body[data-pms-theme="d"] [data-testid="stDataFrame"] canvas,
    .stApp:has(#pms-d:checked) [data-testid="stDataEditor"] canvas,
    .stApp:has(#pms-d:checked) [data-testid="stDataFrame"] canvas,
    body:has(#pms-d:checked) [data-testid="stDataEditor"] canvas,
    body:has(#pms-d:checked) [data-testid="stDataFrame"] canvas {
        filter: invert(1) hue-rotate(180deg) brightness(0.92) contrast(0.88) saturate(0.9) !important;
    }
    html[data-pms-theme="d"] [data-testid="stDataEditor"] [role="columnheader"],
    html[data-pms-theme="d"] [data-testid="stDataEditor"] [role="gridcell"],
    html[data-pms-theme="d"] [data-testid="stDataFrame"] [role="columnheader"],
    html[data-pms-theme="d"] [data-testid="stDataFrame"] [role="gridcell"],
    .stApp:has(#pms-d:checked) [data-testid="stDataEditor"] [role="columnheader"],
    .stApp:has(#pms-d:checked) [data-testid="stDataEditor"] [role="gridcell"],
    .stApp:has(#pms-d:checked) [data-testid="stDataFrame"] [role="columnheader"],
    .stApp:has(#pms-d:checked) [data-testid="stDataFrame"] [role="gridcell"] {
        background-color: #1e1e2e !important;
        border-color: #45475a !important;
        color: #ffffff !important;
        font-weight: 700 !important;
    }

    @media (prefers-color-scheme: dark) {
        body:has(#pms-s:checked) .stApp,
        body:has(#pms-s:checked) [data-testid="stAppViewContainer"],
        body:has(#pms-s:checked) section.main {
            background-color: #1e1e2e !important; color: #ffffff !important;
        }
        body:has(#pms-s:checked) .main .block-container { background-color: #1e1e2e !important; }
        body:has(#pms-s:checked) h1, body:has(#pms-s:checked) h2, body:has(#pms-s:checked) h3,
        body:has(#pms-s:checked) h4, body:has(#pms-s:checked) h5, body:has(#pms-s:checked) h6 { color: #ffffff !important; }
        body:has(#pms-s:checked) p, body:has(#pms-s:checked) [data-testid="stMarkdownContainer"] p { color: #ffffff !important; }
        body:has(#pms-s:checked) label { color: #ffffff !important; }
        body:has(#pms-s:checked) span:not([style*="color"]) { color: #ffffff !important; }
        body:has(#pms-s:checked) input, body:has(#pms-s:checked) [data-baseweb="input"] input {
            background-color: #252535 !important; color: #ffffff !important; border-color: #45475a !important;
        }
        body:has(#pms-s:checked) [class*="ag-theme"] {
            --ag-background-color: #252535; --ag-foreground-color: #ffffff;
            --ag-secondary-foreground-color: #ffffff;
            --ag-header-background-color: #0f0f1f; --ag-header-foreground-color: #ffffff;
            --ag-border-color: #45475a; --ag-row-border-color: #313244;
            --ag-odd-row-background-color: #1e1e30; --ag-row-hover-color: #2d2d45;
        }
        body:has(#pms-s:checked) [class*="ag-theme"] .ag-header,
        body:has(#pms-s:checked) [class*="ag-theme"] .ag-header-row { background-color: #0f0f1f !important; }
        body:has(#pms-s:checked) [class*="ag-theme"] .ag-header-cell-text { color: #ffffff !important; font-weight: 700 !important; }
        body:has(#pms-s:checked) [class*="ag-theme"] .ag-row-even { background-color: #252535 !important; }
        body:has(#pms-s:checked) [class*="ag-theme"] .ag-row-odd  { background-color: #1e1e30 !important; }
        body:has(#pms-s:checked) [class*="ag-theme"] .ag-cell,
        body:has(#pms-s:checked) [class*="ag-theme"] .ag-cell-value { color: #ffffff !important; }
        body:has(#pms-s:checked) .js-plotly-plot,
        body:has(#pms-s:checked) .js-plotly-plot .plot-container,
        body:has(#pms-s:checked) .js-plotly-plot .svg-container {
            background: transparent !important;
        }
        body:has(#pms-s:checked) .js-plotly-plot .bglayer rect.bg {
            fill: transparent !important;
            stroke: transparent !important;
        }
        body:has(#pms-s:checked) .js-plotly-plot .legend rect.bg {
            fill: #252535 !important;
            stroke: #45475a !important;
        }
        body:has(#pms-s:checked) .js-plotly-plot text {
            fill: #ffffff !important;
        }
        body:has(#pms-s:checked) .js-plotly-plot .gridlayer path {
            stroke: #45475a !important;
        }
        body:has(#pms-s:checked) .js-plotly-plot .zerolinelayer path,
        body:has(#pms-s:checked) .js-plotly-plot .xlines-below path,
        body:has(#pms-s:checked) .js-plotly-plot .ylines-below path {
            stroke: #585b70 !important;
        }
        body:has(#pms-s:checked) .js-plotly-plot .modebar-btn path {
            fill: #cdd6f4 !important;
        }
        body:has(#pms-s:checked) [data-testid="stSidebar"],
        body:has(#pms-s:checked) [data-testid="stSidebarContent"] {
            background-color: #171717 !important;
        }
        body:has(#pms-s:checked) .stButton > button,
        body:has(#pms-s:checked) [data-testid="stButton"] button,
        body:has(#pms-s:checked) button[data-testid^="baseButton"] {
            background: #2a2a3e !important;
            background-color: #2a2a3e !important;
            color: #ffffff !important;
            border-color: #45475a !important;
        }
        body:has(#pms-s:checked) .stButton > button:hover,
        body:has(#pms-s:checked) [data-testid="stButton"] button:hover,
        body:has(#pms-s:checked) button[data-testid^="baseButton"]:hover {
            background: #313244 !important;
            background-color: #313244 !important;
            border-color: #ffffff !important;
        }
        body:has(#pms-s:checked) button.pms-home-btn {
            background: transparent !important;
            background-color: transparent !important;
            border: none !important;
            box-shadow: none !important;
        }
        body:has(#pms-s:checked) button.pms-home-btn::before {
            background-color: #cdd6f4;
        }
        body:has(#pms-s:checked) button.pms-home-btn:hover {
            background: transparent !important;
            background-color: transparent !important;
            border: none !important;
        }
        body:has(#pms-s:checked) button.pms-home-btn:hover::before {
            background-color: #818cf8;
        }
        body:has(#pms-s:checked) [data-baseweb="tab"][aria-selected="true"] { color: #ffffff !important; }
        body:has(#pms-s:checked) .pms-sw-track { background: #313244; }
    }
    </style>

    <div class="pms-sw-outer">
        <div class="pms-sw-track">
            <label class="pms-btn" title="라이트">
                <input type="radio" class="pms-theme-radio" name="pms-theme" id="pms-l" checked>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>
            </label>
            <label class="pms-btn" title="다크">
                <input type="radio" class="pms-theme-radio" name="pms-theme" id="pms-d">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>
            </label>
        </div>
    </div>

    <!-- localStorage 테마 영속성: onload 인라인 핸들러로 실행 (script 태그는 React innerHTML에서 미실행) -->
    <img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
         onload="(function(){
             function readCookie(name){
                 var found = document.cookie.split('; ').find(function(row){ return row.indexOf(name + '=') === 0; });
                 return found ? decodeURIComponent(found.split('=')[1]) : '';
             }
             function normalizeTheme(t){
                 return (t === 'd' || t === 'l') ? t : 'l';
             }
             function saveTheme(t){
                 t = normalizeTheme(t);
                 localStorage.setItem('pms-theme', t);
                 document.cookie = 'pms-theme=' + encodeURIComponent(t) + '; max-age=31536000; path=/; SameSite=Lax';
             }
             function applyTheme(){
                 var t = normalizeTheme(localStorage.getItem('pms-theme') || readCookie('pms-theme') || 'l');
                 saveTheme(t);
                 document.documentElement.setAttribute('data-pms-theme', t);
                 if(document.body){ document.body.setAttribute('data-pms-theme', t); }
                 var r = document.getElementById('pms-' + t);
                 if(r && !r.checked){ r.checked = true; }
             }
             function tagSpecialBtns(){
                 document.querySelectorAll('button').forEach(function(btn){
                     var t = btn.textContent.trim();
                     if(t === '🏠') btn.classList.add('pms-home-btn');
                     if(t === '↻') btn.classList.add('pms-refresh-btn');
                 });
             }
             applyTheme();
             tagSpecialBtns();
             if(!window._pmsThemeReady){
                 window._pmsThemeReady = true;
                 var debounce;
                 window._pmsThemeObs = new MutationObserver(function(){
                     clearTimeout(debounce);
                     debounce = setTimeout(function(){ applyTheme(); tagSpecialBtns(); }, 80);
                 });
                 window._pmsThemeObs.observe(document.body, {childList:true, subtree:true});
                 document.addEventListener('change', function(e){
                     if(e.target && e.target.name === 'pms-theme'){
                         saveTheme(e.target.id.replace('pms-',''));
                         applyTheme();
                     }
                 });
             }
         })()"
         style="display:none" alt="">
    """, unsafe_allow_html=True)

    # ── 모바일 반응형 CSS ──────────────────────────────────────────
    st.markdown("""
    <style>
    @media (max-width: 768px) {
        /* ① 사이드바: 모바일에서 전체 폭 오버레이 */
        [data-testid="stSidebar"] {
            width: 100vw !important;
            min-width: 100vw !important;
            max-width: 100vw !important;
        }
        section[data-testid="stSidebar"] > div:first-child {
            width: 100vw !important;
        }

        /* ② 메인 컨텐츠: 전체 폭, 패딩 최소화 */
        .block-container {
            padding: 0.5rem 0.75rem 2rem !important;
            max-width: 100% !important;
        }
        [data-testid="stAppViewContainer"] > section.main {
            width: 100% !important;
        }

        /* ③ 테마 토글 위치·크기 축소 */
        .pms-sw-outer { top: 6px; right: 6px; }
        .pms-btn { padding: 0 6px; height: 22px; font-size: 10px; }

        /* ④ 표: 가로 스크롤 */
        .pms-report-table {
            overflow-x: auto !important;
            -webkit-overflow-scrolling: touch !important;
            max-width: 100vw !important;
        }
        table { font-size: 11px !important; min-width: max-content; }
        table th, table td { padding: 4px 6px !important; white-space: nowrap !important; }

        /* ⑤ data_editor 가로 스크롤 */
        [data-testid="stDataEditor"] {
            overflow-x: auto !important;
            -webkit-overflow-scrolling: touch !important;
        }

        /* ⑥ 컬럼 레이아웃: 2개 이상이면 가로 스크롤 */
        [data-testid="stHorizontalBlock"] {
            overflow-x: auto !important;
            -webkit-overflow-scrolling: touch !important;
            flex-wrap: nowrap !important;
            gap: 4px !important;
        }
        [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
            min-width: 120px !important;
            flex-shrink: 0 !important;
        }

        /* ⑦ 버튼 터치 영역 */
        .stButton > button { min-height: 44px !important; font-size: 14px !important; }

        /* ⑧ 입력 필드 — iOS 자동 확대 방지 */
        input, textarea, select { font-size: 16px !important; }

        /* ⑨ 메트릭 */
        [data-testid="stMetricValue"] { font-size: 1.1rem !important; }
        [data-testid="stMetricLabel"] { font-size: 0.7rem !important; }
        [data-testid="metric-container"] { padding: 8px !important; }

        /* ⑩ 제목 */
        h1 { font-size: 1.3rem !important; }
        h2 { font-size: 1.15rem !important; }
        h3 { font-size: 1.05rem !important; }
        h4 { font-size: 0.95rem !important; }

        /* ⑪ 탭 */
        [data-baseweb="tab"] { padding: 6px 8px !important; font-size: 12px !important; }

        /* ⑫ 사이드바 햄버거 버튼 */
        [data-testid="stSidebarCollapsedControl"] button {
            width: 44px !important; height: 44px !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)


def read_billing_login_upload(uploaded_file):
    file_name = (uploaded_file.name or "").lower()
    raw = uploaded_file.getvalue()
    if file_name.endswith(".csv"):
        try:
            return pd.read_csv(BytesIO(raw), dtype=str).fillna("")
        except UnicodeDecodeError:
            return pd.read_csv(BytesIO(raw), dtype=str, encoding="cp949").fillna("")
    return pd.read_excel(BytesIO(raw), dtype=str).fillna("")


def read_billing_source_upload(uploaded_file):
    file_name = (uploaded_file.name or "").lower()
    raw = uploaded_file.getvalue()
    if file_name.endswith(".csv"):
        try:
            return {"업로드 파일": pd.read_csv(BytesIO(raw), dtype=str).fillna("")}
        except UnicodeDecodeError:
            return {"업로드 파일": pd.read_csv(BytesIO(raw), dtype=str, encoding="cp949").fillna("")}
    sheets = pd.read_excel(BytesIO(raw), sheet_name=None, dtype=str)
    return {str(name): df.fillna("") for name, df in sheets.items()}


def normalize_billing_login_df(df):
    source = clean_header_logic(df.copy()).replace({np.nan: ""})
    customer_col = find_col(source, ["고객번호"])
    company_col = find_col(source, ["고객명"])
    latest_login_col = find_col(source, ["최근로그인", "최종로그인일자", "최종로그인"])
    login_count_col = find_col(source, ["로그인", "로그인횟수"])

    missing = [
        label
        for label, col in {
            "고객번호": customer_col,
            "고객명": company_col,
            "최근로그인": latest_login_col,
            "로그인": login_count_col,
        }.items()
        if not col or col not in source.columns
    ]
    if missing:
        return pd.DataFrame(), missing

    result = pd.DataFrame(
        {
            "고객번호": source[customer_col].astype(str).str.strip(),
            "고객명": source[company_col].astype(str).str.strip(),
            "최근로그인": source[latest_login_col].astype(str).str.strip(),
            "로그인": pd.to_numeric(source[login_count_col].astype(str).str.replace(",", "", regex=False), errors="coerce")
            .fillna(0)
            .astype(int),
        }
    )
    result = result[result["고객번호"].astype(str).str.strip().ne("")]
    return result.reset_index(drop=True), []


def render_billing_source_tables(source_upload=None):
    open_columns = [
        "순번",
        "고객번호",
        "사업자번호",
        "업체명",
        "ERP연계 여부",
        "접수일자",
        "구축일자",
        "방문일자",
        "담당자",
        "비고",
        "최초로그인",
        "최종로그인일자",
        "로그인횟수",
        "청구원본 고객명",
        "실적파일 고객명",
    ]
    erp_columns = [
        "순서",
        "고객번호",
        "사업자번호",
        "업체명",
        "구분",
        "추가연계신청일자",
        "담당자",
        "구축일",
        "연계시작일자",
        "은행연계완료일자",
        "수령여부",
        "비고",
        "최초로그인",
        "최종로그인일자",
        "로그인횟수",
        "청구원본 고객명",
        "실적파일 고객명",
    ]

    st.markdown("#### 청구 원본 표")
    source_file = source_upload or st.file_uploader(
        "구축 및 연계 리스트 업로드",
        type=["xlsx", "xls", "csv"],
        key="billing_source_upload",
        help="2026년 6월 구축 및 연계 리스트_최종본.xlsx 파일을 업로드해주세요.",
    )
    open_tab, erp_tab = st.tabs(["청구원본(개설업로드)", "청구원본(연계업로드)"])
    with open_tab:
        st.caption("개설 청구자료 생성 화면 구성입니다. 실제 시트 데이터 연동 없이 표 영역만 표시합니다.")
        st.dataframe(pd.DataFrame(columns=open_columns), use_container_width=True, hide_index=True)
    with erp_tab:
        st.caption("연계 청구자료 생성 화면 구성입니다. 실제 시트 데이터 연동 없이 표 영역만 표시합니다.")
        st.dataframe(pd.DataFrame(columns=erp_columns), use_container_width=True, hide_index=True)

    if not source_file:
        st.warning("구축 및 연계 리스트를 업로드하면 아래에 파일 미리보기가 표시됩니다.")
        return

    try:
        source_sheets = read_billing_source_upload(source_file)
    except Exception as exc:
        st.error(f"구축 및 연계 리스트 파일을 읽을 수 없습니다: {exc}")
        return

    st.markdown("#### 구축 및 연계 리스트 미리보기")
    sheet_names = list(source_sheets.keys())
    if len(sheet_names) == 1:
        df = clean_header_logic(source_sheets[sheet_names[0]].copy()).replace({np.nan: ""})
        st.caption(f"{sheet_names[0]} · {len(df):,}행")
        st.dataframe(df.head(500), use_container_width=True, hide_index=True)
    else:
        tabs = st.tabs(sheet_names)
        for tab, sheet_name in zip(tabs, sheet_names):
            with tab:
                df = clean_header_logic(source_sheets[sheet_name].copy()).replace({np.nan: ""})
                st.caption(f"{sheet_name} · {len(df):,}행")
                st.dataframe(df.head(500), use_container_width=True, hide_index=True)


def show_billing_generation():
    st.markdown("### 청구자료 생성")
    st.caption("청구 원본과 은행 로그인 실적파일을 대사해 청구자료 생성 전 확인 목록을 만듭니다.")
    st.link_button(
        "원본 Google Sheet 열기",
        "https://docs.google.com/spreadsheets/d/12BeCTDegUWD-jomaG3WS75Jx1dJ9Lqjxn1hVs3FrpE4/edit?gid=1244892381#gid=1244892381",
    )
    st.info(
        "FastAPI/React 화면에서는 해당 시트의 개설·연계 청구원본과 은행 실적파일을 대사합니다. "
        "현재 Google Sheet가 비공개이면 CSV export 권한 안내가 표시됩니다."
    )
    uploaded_login = st.file_uploader(
        "은행로그인실적파일(은행) 엑셀 업로드",
        type=["xlsx", "xls", "csv"],
        help="고객번호, 고객명, 최근로그인, 로그인 컬럼이 포함된 파일을 업로드해주세요.",
    )

    if not uploaded_login:
        st.warning("은행로그인실적파일(은행)을 업로드하면 청구자료 생성용 실적 데이터를 확인할 수 있습니다.")
        render_billing_source_tables()
        return

    try:
        raw_df = read_billing_login_upload(uploaded_login)
        normalized_df, missing = normalize_billing_login_df(raw_df)
    except Exception as exc:
        st.error(f"파일을 읽을 수 없습니다: {exc}")
        return

    if missing:
        st.error(f"필수 컬럼을 찾을 수 없습니다: {', '.join(missing)}")
        st.caption("필요 컬럼: 고객번호, 고객명, 최근로그인, 로그인")
        st.dataframe(raw_df.head(20), use_container_width=True)
        render_billing_source_tables()
        return

    total_login = int(normalized_df["로그인"].sum()) if not normalized_df.empty else 0
    latest_login = "-"
    parsed_dates = pd.to_datetime(normalized_df["최근로그인"], errors="coerce")
    if parsed_dates.notna().any():
        latest_login = parsed_dates.max().strftime("%Y-%m-%d")

    col_total, col_customer, col_login, col_latest = st.columns(4)
    col_total.metric("업로드 행수", f"{len(normalized_df):,}건")
    col_customer.metric("고객 수", f"{normalized_df['고객번호'].nunique():,}곳")
    col_login.metric("로그인 합계", f"{total_login:,}회")
    col_latest.metric("최신 로그인일", latest_login)

    st.markdown("#### 업로드 실적 미리보기")
    st.dataframe(normalized_df.head(500), use_container_width=True, hide_index=True)

    excel_bytes = dataframe_to_excel_bytes({"은행로그인실적파일": normalized_df})
    st.download_button(
        "정규화된 은행로그인실적파일 다운로드",
        data=excel_bytes,
        file_name="은행로그인실적파일_정규화.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    render_billing_source_tables()


def show_main():
    apply_global_table_css()
    inject_theme_toggle()
    show_sidebar()
    persist_current_menu()

    menu = st.session_state.current_menu
    allowed_menus = {BILLING_MENU}
    if menu not in allowed_menus:
        st.session_state.current_menu = BILLING_MENU
        persist_current_menu()
        st.rerun()

    render_page_title(menu)

    if menu == BILLING_MENU:
        show_billing_generation()
        return
    else:
        st.session_state.current_menu = BILLING_MENU
        st.rerun()


show_main()
