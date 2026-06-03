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

st.set_page_config(page_title="실적관리 시스템", layout="wide", initial_sidebar_state="expanded")


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
SENT_FILE = "sent_results.json"
SENT_UPLOADS_FILE = "sent_uploads.json"
SAVED_STATE_FILE = "saved_state.json"
WEEKLY_REPORT_FILE = "weekly_reports.json"
WEEKLY_STATUS_SNAPSHOT_FILE = "weekly_status_snapshots.json"
PPT_TEMPLATE_FILE = resolve_template_file("LMB활동실적보고서_202605_하나지사.pptx")
WEEKLY_PPT_TEMPLATE_FILE = resolve_template_file("주간보고_통합CMS고객_개설운영_주간보고_템플릿.pptx")
EXCEL_SAMPLE_FILE = resolve_template_file("LMB월간 활동실적_000000(샘플).xlsx")

DEFAULT_URL_ANALYSIS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9XPHqrqcaFf9bCOVya7yHORr-c1R4KCF0eEpdE3ESn8qJELP0BkqTOslur9bsGcVabRUIcyOa877R/pub?output=csv"
DEFAULT_URL_SYNC = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9F7R7oLA2B02H-I25kVv2JeYHFgWQq0CT7TeW61hrNpJLdHWJFhFR_iDQGCFAW044o8rRwBDeovKG/pub?gid=1533424484&single=true&output=csv"
DEFAULT_URL_HANA = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQgRHnTZD4eDW2UeODQuGxmxFrflKpbQda3sBsVjj1s3qAFWMKcpke2U58UuT6VEDlkbXveZlaroTCr/pub?gid=0&single=true&output=csv"
DEFAULT_URL_HANA_BILLING = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQgRHnTZD4eDW2UeODQuGxmxFrflKpbQda3sBsVjj1s3qAFWMKcpke2U58UuT6VEDlkbXveZlaroTCr/pub?gid=1172734914&single=true&output=csv"


@st.cache_data(ttl=300, show_spinner=False)
def read_csv_cached(url, header=0, cache_buster=0):
    return pd.read_csv(url, header=header)


def read_google_csv(url, header=0, force_refresh=False):
    cache_buster = int(datetime.utcnow().timestamp()) if force_refresh else 0
    return read_csv_cached(url, header=header, cache_buster=cache_buster).copy()


GITHUB_REPO = "preciselee84-oss/LMB-C-S-PMS"
GITHUB_BRANCH = "main"
GITHUB_DATA_DIR = "data"
SESSION_UID_COOKIE = "auto_login_uid"
LAST_MENU_COOKIE = "last_menu"


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
        "current_menu": "대시보드",
        "url_analysis": DEFAULT_URL_ANALYSIS,
        "url_sync": DEFAULT_URL_SYNC,
        "url_hana": DEFAULT_URL_HANA,
        "url_hana_billing": DEFAULT_URL_HANA_BILLING,
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

    if st.session_state.current_menu == "이력확인 및 작성":
        st.session_state.current_menu = "업로드 및 실적 확인"
    if st.session_state.current_menu == "최종 실적 확인":
        st.session_state.current_menu = "업로드 및 실적 확인"
    if st.session_state.current_menu == "은행 이력 업로드":
        st.session_state.current_menu = "업로드 및 실적 확인"


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

    return series.apply(normalize_one)


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

    df = pd.merge(df, cloud_map, on="_biz_key", how="left")
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
        result["제목"] = result["활동상세"].map(title_from_activity_detail)
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
            p_sum = min(2800, min(1000, o_p + l_p) + min(1800, v_actual_p + manual_p))

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

            effective_manual_p = int(max(0, stats["p_sum"] - (stats["o_p"] + stats["l_p"] + stats["v_actual_p"])))

            if is_outsource:
                pay = int(max(0, (stats["p_sum"] - 1000) * 500))
            else:
                leader_bonus = 500 if rank == "팀장" else 0
                pay = int(max(0, (stats["p_sum"] + leader_bonus - 1000 + adj_per_regular) * 500))

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
                    "지급포인트": max(0, stats["p_sum"] - 1000),
                    "지급예상금액": pay,
                    "전월대비": 0,
                }
            )

        res_df = pd.DataFrame(rows)
        res_df = apply_rank_from_user_db(res_df)
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


def render_plain_html_table(df, max_rows=500, center_align=True, merge_cols=None):
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
            align = "left" if str(col).strip() == "활동내역" else ("center" if center_align else "left")
            td_align = f"text-align:{align};"
            rowspan = ""
            if col in rowspan_map and rowspan_map[col].get(i, 1) > 1:
                rowspan = f" rowspan='{rowspan_map[col][i]}'"
            tds += f"<td{rowspan} style='background:{bg};padding:5px 10px;border-bottom:1px solid #EDF2F7;font-size:12px;color:#2D3748;white-space:nowrap;vertical-align:middle;{td_align}'>{val}</td>"
        body += f"<tr>{tds}</tr>"
    st.markdown(
        f"""<div class="pms-report-table" style="overflow-x:auto;border:1px solid #E2E8F0;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,0.05);margin-bottom:1rem;">
        <table style="width:100%;border-collapse:collapse;">
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
            font-size: 50px;
            font-weight: 900;
            letter-spacing: -0.5px;
            margin-bottom: 5px;
        }
        .auth-logo-sub {
            color: #7B79AA;
            font-size: 20px;
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
                font-size: 32px !important;
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
                <div class="auth-logo-title">실적관리 시스템</div>
                <div class="auth-logo-sub">Performance Management System</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

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

                    st.session_state.current_menu = "대시보드"
                    persist_current_menu()
                    st.rerun()
                elif not u_id_str:
                    st.error("아이디를 입력해주세요.")
                elif not u_pw_str:
                    st.error("비밀번호를 입력해주세요.")
                else:
                    st.error("아이디 또는 비밀번호가 올바르지 않습니다.")

            st.divider()
            st.markdown("<div class='auth-small'>계정이 없으신가요?</div>", unsafe_allow_html=True)

            if st.button("회원가입", use_container_width=True):
                st.session_state.auth_mode = "signup"
                st.rerun()

            if st.button("비밀번호 찾기", use_container_width=True):
                st.session_state.auth_mode = "find_pw"
                st.rerun()

        elif st.session_state.auth_mode == "find_pw":
            st.markdown("#### 비밀번호 찾기")
            st.markdown("아이디와 가입 시 등록한 메일주소를 입력하세요.")

            with st.form("find_pw_form"):
                fp_id = st.text_input("아이디", key="fp_id")
                fp_email = st.text_input("메일주소", key="fp_email")
                fp_submitted = st.form_submit_button("확인", use_container_width=True, type="primary")

            if fp_submitted:
                fp_id_str = fp_id.strip() if fp_id else ""
                fp_email_str = fp_email.strip() if fp_email else ""
                if not fp_id_str or not fp_email_str:
                    st.error("아이디와 메일주소를 모두 입력해주세요.")
                else:
                    db = st.session_state.user_db
                    user_info = db.get(fp_id_str)
                    if user_info and user_info.get("email", "").strip() == fp_email_str:
                        st.success(f"비밀번호: **{user_info.get('pw', '')}**")
                    else:
                        st.error("일치하는 계정 정보가 없습니다.")

            if st.button("로그인으로 돌아가기", use_container_width=True, key="fp_back"):
                st.session_state.auth_mode = "login"
                st.rerun()

        else:
            # 아이디 실시간 중복 확인 (폼 바깥)
            r_id_check = st.text_input("아이디", key="r_id_check")
            if r_id_check:
                if r_id_check in st.session_state.user_db:
                    st.error("이미 사용 중인 아이디입니다.")
                else:
                    st.success("사용 가능한 아이디입니다.")

            with st.form("signup_form"):
                f_id = st.text_input("아이디 (확인용 재입력)", key="f_id")
                f_name = st.text_input("성명", key="f_name")
                f_email = st.text_input("메일주소", key="f_email")
                f_pw = st.text_input("비밀번호", type="password", key="f_pw")
                f_pw2 = st.text_input("비밀번호 확인", type="password", key="f_pw2")
                submitted = st.form_submit_button("회원가입", use_container_width=True, type="primary")

            if submitted:
                if not f_id or not f_name or not f_email or not f_pw or not f_pw2:
                    st.error("모든 항목을 입력해주세요.")
                elif f_pw != f_pw2:
                    st.error("비밀번호가 일치하지 않습니다.")
                elif f_id in st.session_state.user_db:
                    st.error("이미 존재하는 아이디입니다.")
                else:
                    st.session_state.user_db[f_id] = {
                        "pw": f_pw,
                        "name": f_name,
                        "email": f_email,
                        "access": "불가",
                        "role": "사용자",
                        "dept_type": "사업부",
                        "staff_type": "정규직",
                        "outsource": "아니오",
                        "outsource_period": "해당없음",
                    }
                    save_db(DB_FILE, st.session_state.user_db)
                    st.success("신청 완료. 관리자 승인 후 로그인 가능합니다.")
                    time.sleep(1.5)
                    st.session_state.auth_mode = "login"
                    st.rerun()

            if st.button("로그인으로 돌아가기", use_container_width=True):
                st.session_state.auth_mode = "login"
                st.rerun()


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
            [data-testid="stSidebar"] {
                background-color: #171717 !important;
                border-right: 1px solid #2f2f2f !important;
            }
            [data-testid="stSidebar"] > div:first-child {
                padding-top: 0 !important;
            }
            [data-testid="stSidebarContent"] {
                padding: 10px 10px 12px !important;
                background-color: #171717 !important;
            }
            [data-testid="stSidebar"] * {
                color: #ececec !important;
            }
            [data-testid="stSidebar"] .gpt-side-shell {
                display: flex;
                flex-direction: column;
                gap: 10px;
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
                width: 28px;
                height: 28px;
                border: 1px solid #3a3a3a;
                border-radius: 8px;
                color: #ffffff !important;
                font-size: 13px;
                font-weight: 900;
                letter-spacing: 0;
                background: #212121;
            }
            [data-testid="stSidebar"] .gpt-brand-text {
                display: flex;
                flex-direction: column;
                line-height: 1.2;
            }
            [data-testid="stSidebar"] .gpt-brand-title {
                font-size: 15px;
                font-weight: 800;
                color: #f5f5f5 !important;
            }
            [data-testid="stSidebar"] .gpt-brand-subtitle {
                margin-top: 2px;
                font-size: 12px;
                color: #a7a7a7 !important;
                font-weight: 500;
            }
            [data-testid="stSidebar"] .gpt-user-card {
                margin: 2px 2px 8px;
                padding: 10px;
                border: 1px solid #2f2f2f;
                border-radius: 12px;
                background: #212121;
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
                color: #9b9b9b !important;
                font-weight: 500;
                line-height: 1.35;
            }
            [data-testid="stSidebar"] .gpt-section {
                margin: 4px 8px 1px;
                font-size: 12px;
                color: #8e8e8e !important;
                font-weight: 700;
                letter-spacing: 0;
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
                background: #2f2f2f;
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
                border-color: #2f2f2f !important;
                margin: 10px 0 !important;
            }
            [data-testid="stSidebar"] div.stButton > button {
                background: transparent !important;
                border: none !important;
                box-shadow: none !important;
                color: #ececec !important;
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
                color: #ececec !important;
                white-space: normal !important;
                line-height: 1.25 !important;
            }
            [data-testid="stSidebar"] div.stButton > button:hover {
                background: #2f2f2f !important;
                color: #ffffff !important;
            }
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + div [data-testid="stButton"] button,
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + [data-testid="stElementContainer"] [data-testid="stButton"] button {
                background: #2f2f2f !important;
                color: #ffffff !important;
            }
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + div [data-testid="stButton"] button p,
            [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.gpt-nav-active) + [data-testid="stElementContainer"] [data-testid="stButton"] button p {
                color: #ffffff !important;
                font-weight: 700 !important;
            }
            [data-testid="stSidebar"] .gpt-logout-marker + div [data-testid="stButton"] button,
            [data-testid="stSidebar"] .gpt-logout-marker + [data-testid="stElementContainer"] [data-testid="stButton"] button {
                color: #d9d9d9 !important;
            }
            [data-testid="stSidebar"] .gpt-logout-marker + div [data-testid="stButton"] button:hover,
            [data-testid="stSidebar"] .gpt-logout-marker + [data-testid="stElementContainer"] [data-testid="stButton"] button:hover {
                background: #3a2424 !important;
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
                if menu_name == "업로드 및 실적 확인":
                    st.session_state["show_add_history_form"] = False
                st.session_state["_close_sidebar_mobile"] = True
                st.rerun()

        _lt = st.session_state.get("login_time", "")
        st.markdown(
            f"<div class='gpt-side-shell'>"
            f"<div class='gpt-brand'>"
            f"<div class='gpt-brand-mark'>P</div>"
            f"<div class='gpt-brand-text'>"
            f"<div class='gpt-brand-title'>실적관리 시스템</div>"
            f"<div class='gpt-brand-subtitle'>Performance workspace</div>"
            f"</div></div>"
            f"<div class='gpt-user-card'>"
            f"<div class='gpt-user-name'>{html.escape(st.session_state.user_name)}님</div>"
            f"<div class='gpt-user-meta'>반갑습니다.<br>접속 {_lt}</div>"
            f"</div>"
            f"</div>",
            unsafe_allow_html=True,
        )

        if st.session_state.user_role == "관리자":
            st.markdown("<div class='gpt-section'>관리자 메뉴</div>", unsafe_allow_html=True)
            for menu_name in ["관리자용 실적 확인", "실적 분석/계산", "실적 보고서", "청구자료 작성", "주간보고 취합", "운영계획"]:
                render_nav_button(menu_name)

        st.markdown("<div class='gpt-section'>사용자 메뉴</div>", unsafe_allow_html=True)
        for menu_name in ["업로드 및 실적 확인", "이번달 활동 대상고객 추천", "주간보고 이력 작성", "방문이력 작성"]:
            render_nav_button(menu_name)

        if st.session_state.user_role == "관리자":
            st.markdown("<div class='gpt-section'>설정</div>", unsafe_allow_html=True)
            for menu_name in ["직원 및 권한설정", "구글 스트레드시트 연동"]:
                render_nav_button(menu_name)

        st.markdown("<div class='gpt-sidebar-divider'></div><div class='gpt-logout-marker'></div>", unsafe_allow_html=True)
        if st.button("로그아웃", use_container_width=True, key="nav_logout"):
            st.session_state.logged_in = False
            st.session_state.auth_mode = "login"
            try:
                _cm_logout = safe_cookie_controller()
                cookie_remove(_cm_logout, SESSION_UID_COOKIE)
                cookie_remove(_cm_logout, LAST_MENU_COOKIE)
            except Exception:
                pass
            st.rerun()


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
        </style>
        """,
        unsafe_allow_html=True,
    )


MENU_GUIDES = {
    "대시보드": [
        "📊 로그인한 본인의 이번달 실적 추정치를 한눈에 확인할 수 있습니다.",
        "📅 전월 및 전년 동월과의 포인트 증감을 비교합니다.",
        "💡 현재 부족한 실적을 채울 수 있는 활동 방안을 안내합니다.",
        "🏦 이번달 추정 포인트는 하나은행 구글 시트에 입력된 수치를 기준으로 표시됩니다.",
        "📋 전월 대비 · 전년 동월 대비는 본사에 최종 제출하는 활동이력 데이터를 참조하여 표시됩니다.",
    ],
    "관리자용 실적 확인": [
        "👥 전체 직원의 실적 현황을 한눈에 확인할 수 있습니다.",
        "📊 담당자별 개설/연계/운영 실적과 포인트를 조회합니다.",
        "🔍 직원 검색, 필터링, 정렬 기능으로 빠르게 실적을 파악하세요.",
        "📥 전체 실적 데이터를 Excel로 다운로드할 수 있습니다.",
    ],
    "실적 분석/계산": [
        "📂 구글 스프레드시트에서 불러온 실적 데이터를 분석·계산합니다.",
        "📊 당월/전월 데이터를 비교하여 포인트 및 지급예상금액을 확인할 수 있습니다.",
        "✅ 검증 탭에서 중복/초과 방문, 누락 데이터 등 오류 여부를 확인하세요.",
        "🔒 검토 완료 후 [마감] 버튼을 눌러 실적을 확정합니다.",
    ],
    "실적 보고서": [
        "📋 마감된 실적 데이터를 기반으로 보고서를 조회합니다.",
        "📥 Excel 또는 PPT 형식으로 보고서를 다운로드할 수 있습니다.",
        "📅 당월·전월 비교 데이터가 함께 표시됩니다.",
    ],
    "직원 및 권한설정": [
        "👤 직원 계정의 접근 권한을 허용/불가로 설정합니다.",
        "🔑 비밀번호 초기화 및 역할(관리자/사용자) 변경이 가능합니다.",
        "➕ 신규 직원 계정을 직접 등록할 수 있습니다.",
    ],
    "구글 스트레드시트 연동": [
        "🔗 분석용·동기화용 구글 스프레드시트 URL을 설정합니다.",
        "🔄 [데이터 새로고침] 버튼으로 최신 데이터를 불러옵니다.",
        "⚠️ URL이 잘못된 경우 데이터를 불러오지 못할 수 있습니다.",
    ],
    "업로드 및 실적 확인": [
        "📝 본인의 활동 이력을 엑셀 파일로 업로드합니다.",
        "🔍 업로드 후 중복·초과 방문, 오류 데이터를 탭에서 확인하세요.",
        "💾 [저장] 버튼으로 임시 저장 후 [실적 결과 전송]으로 제출합니다.",
        "⚠️ 전송 전 검증 오류 항목을 반드시 확인하세요.",
    ],
}


def render_page_title(menu):
    if menu == "대시보드":
        return

    if menu != "대시보드":
        st.markdown(
            """
            <style>
            .block-container { padding-top: 1rem !important; }
            /* .home-btn div와 button은 DOM에서 형제(sibling) — :has()+인접형제로 타겟팅 */
            [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button {
                background: transparent !important;
                background-color: transparent !important;
                border: none !important;
                box-shadow: none !important;
                display: inline-flex !important;
                align-items: center !important;
                justify-content: center !important;
                font-size: 0 !important;
                padding: 0 !important;
                height: 82px !important;
                line-height: 1 !important;
                color: transparent !important;
                width: 82px !important;
            }
            [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button * {
                color: transparent !important;
                display: none !important;
                font-size: 0 !important;
            }
            [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button::before {
                content: "";
                width: 75px;
                height: 75px;
                display: block;
                flex: 0 0 auto;
                background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 96 96'%3E%3Cpath fill='%23000' d='M32 8h56v56c0 13.255-10.745 24-24 24H8V32C8 18.745 18.745 8 32 8Z'/%3E%3Cpath fill='%23fff' d='M20 52 48 26l28 26h-10v26H30V52H20Z'/%3E%3Cpath fill='%23fff' d='M58 34h12v18H58z'/%3E%3Cpath fill='%23000' d='M41 52h8v8h-8zM53 52h8v8h-8zM41 64h8v8h-8zM53 64h8v8h-8z'/%3E%3C/svg%3E") center / contain no-repeat;
            }
            [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button:hover {
                background: transparent !important;
                background-color: transparent !important;
            }
            [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button:hover::before {
                filter: opacity(0.86);
            }
            body:has(#pms-d:checked) [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button {
                background: transparent !important;
                background-color: transparent !important;
                border: none !important;
            }
            body:has(#pms-d:checked) [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button::before {
                filter: none;
            }
            body:has(#pms-d:checked) [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button:hover {
                background: transparent !important;
                background-color: transparent !important;
                border: none !important;
            }
            body:has(#pms-d:checked) [data-testid="stElementContainer"]:has(.home-btn) + [data-testid="stElementContainer"] [data-testid="stButton"] button:hover::before {
                filter: opacity(0.86);
            }
            </style>
            <div class="home-btn">
            """,
            unsafe_allow_html=True,
        )
        if st.button("🏠", key=f"home_btn_{menu}", help="대시보드로 이동"):
            st.session_state.current_menu = "대시보드"
            persist_current_menu()
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    if menu != "대시보드":
        st.markdown(f"## {menu}")
    settings_menus = ["직원 및 권한설정", "구글 스트레드시트 연동"]
    admin_menus = ["관리자용 실적 확인", "실적 분석/계산", "실적 보고서", "청구자료 작성", "주간보고 취합", "운영계획"]
    if menu in settings_menus:
        parent_nav = "설정"
    elif menu in admin_menus:
        parent_nav = "관리자 메뉴"
    else:
        parent_nav = "사용자 메뉴"

    if menu != "대시보드":
        st.markdown(
            f"<div style='color:#718096;font-size:14px;font-weight:600;margin-top:-8px;margin-bottom:12px;'>{parent_nav} &gt; {html.escape(menu)}</div>",
            unsafe_allow_html=True,
        )

    if menu in MENU_GUIDES and menu != "대시보드":
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


def show_all_staff_summary(staff_names):
    """전체 직원 실적 요약 표시"""
    # 본사 구글시트 로드
    if st.session_state.get("cloud_sheet_df") is None:
        try:
            with st.spinner("본사 구글시트를 불러오는 중..."):
                load_csv_to_state("url_sync", "cloud_sheet_df")
        except Exception as e:
            st.error("⚠️ 본사 구글시트를 불러오지 못했습니다.")
            with st.expander("📋 오류 상세 정보"):
                st.code(str(e))
            st.info("💡 [구글 스트레드시트 연동] 메뉴에서 '본사 구글 시트 CSV URL'을 확인해주세요.")

            if st.button("🔗 구글 스트레드시트 연동 메뉴로 이동", use_container_width=True):
                st.session_state.current_menu = "구글 스트레드시트 연동"
                st.rerun()
            return

    cloud_df = st.session_state.get("cloud_sheet_df")
    if cloud_df is None or cloud_df.empty:
        st.warning("📂 본사 구글시트 데이터가 없습니다.")
        st.info("💡 [구글 스트레드시트 연동] 메뉴에서 URL 설정 및 데이터 새로고침을 해주세요.")

        if st.button("🔗 구글 스트레드시트 연동 메뉴로 이동", use_container_width=True):
            st.session_state.current_menu = "구글 스트레드시트 연동"
            st.rerun()
        return

    # 운영 실적용 데이터: 업로드한 엑셀 파일만 사용 (없으면 None)
    analysis_df = st.session_state.get("admin_uploaded_excel")

    # 은행 구글시트 로드 (하나은행 시트)
    hana_df = None
    if st.session_state.get("hana_sheet_df") is None:
        try:
            load_csv_to_state("url_hana", "hana_sheet_df")
        except:
            pass
    hana_df = st.session_state.get("hana_sheet_df")

    # 데이터 정리
    cloud_df = clean_header_logic(cloud_df.copy())

    # 본사 시트 컬럼 찾기
    owner_col = find_col(cloud_df, ["담당자", "등록자", "성명"])
    open_date_col = find_col(cloud_df, ["개설완료일자", "개설일"])
    erp_date_col = find_col(cloud_df, ["ERP연계일자", "연계일자"])
    biz_num_col = find_col(cloud_df, ["사업자번호", "사업자등록번호"])
    company_col = find_col(cloud_df, ["고객명", "업체명", "상호"])

    if not owner_col or owner_col not in cloud_df.columns:
        st.warning("본사 시트에서 담당자 컬럼을 찾을 수 없습니다.")
        return

    # 은행 시트와 본사 시트 비교 - 누락 데이터 확인
    if hana_df is not None and not hana_df.empty:
        hana_clean = clean_header_logic(hana_df.copy())
        hana_biz_col = find_col(hana_clean, ["사업자번호", "사업자등록번호"])
        hana_company_col = find_col(hana_clean, ["고객명", "업체명", "상호"])
        hana_open_status_col = find_col(hana_clean, ["개설상태"])
        hana_erp_status_col = find_col(hana_clean, ["연계상태", "ERP연계상태"])
        hana_owner_col = find_col(hana_clean, ["담당자", "등록자"])

        cloud_status_col = find_col(cloud_df, ["상태"])

        missing_open = []
        missing_erp = []

        if hana_biz_col and biz_num_col:
            # 개설상태가 "개설완료"인데 본사 시트 상태가 "완료"가 아닌 경우
            if hana_open_status_col:
                hana_open_completed = hana_clean[
                    hana_clean[hana_open_status_col].astype(str).str.contains("개설완료", na=False)
                ].copy()
                for _, row in hana_open_completed.iterrows():
                    biz_num = str(row.get(hana_biz_col, "")).strip()
                    if not biz_num:
                        continue
                    # 본사 시트에서 해당 사업자번호 찾기
                    cloud_match = cloud_df[cloud_df[biz_num_col].astype(str).str.contains(biz_num, na=False, regex=False)]
                    if not cloud_match.empty:
                        # 상태가 "완료"가 아닌 경우
                        if cloud_status_col and cloud_status_col in cloud_match.columns:
                            cloud_status = cloud_match[cloud_status_col].astype(str).str.strip()
                            if not cloud_status.str.contains("완료", na=False).any():
                                missing_open.append({
                                    "사업자번호": biz_num,
                                    "고객명": row.get(hana_company_col, ""),
                                    "담당자": row.get(hana_owner_col, ""),
                                    "은행시트_개설상태": row.get(hana_open_status_col, ""),
                                    "본사시트_상태": cloud_match[cloud_status_col].iloc[0] if not cloud_match.empty else ""
                                })

            # 연계상태가 "연계완료" 또는 "청구완료"인데 본사 시트 ERP연계일자가 없는 경우
            if hana_erp_status_col:
                hana_erp_completed = hana_clean[
                    hana_clean[hana_erp_status_col].astype(str).str.contains("연계완료|청구완료", na=False, regex=True)
                ].copy()
                for _, row in hana_erp_completed.iterrows():
                    biz_num = str(row.get(hana_biz_col, "")).strip()
                    if not biz_num:
                        continue
                    # 본사 시트에서 해당 사업자번호 찾기
                    cloud_match = cloud_df[cloud_df[biz_num_col].astype(str).str.contains(biz_num, na=False, regex=False)]
                    if not cloud_match.empty:
                        # ERP연계일자가 없는 경우
                        if erp_date_col and erp_date_col in cloud_match.columns:
                            if cloud_match[erp_date_col].isna().all():
                                missing_erp.append({
                                    "사업자번호": biz_num,
                                    "고객명": row.get(hana_company_col, ""),
                                    "담당자": row.get(hana_owner_col, ""),
                                    "은행시트_연계상태": row.get(hana_erp_status_col, "")
                                })

        # 누락 데이터 표시
        if missing_open or missing_erp:
            with st.expander("⚠️ 본사 시트 누락 데이터", expanded=False):
                if missing_open:
                    st.markdown("#### 개설 상태 불일치")
                    st.write(f"은행 시트는 개설완료이나 본사 시트 상태가 완료가 아닌 고객사: **{len(missing_open)}건**")
                    st.dataframe(pd.DataFrame(missing_open), use_container_width=True)
                    st.markdown("---")

                if missing_erp:
                    st.markdown("#### ERP연계일자 누락")
                    st.write(f"은행 시트는 연계완료/청구완료이나 본사 시트에 ERP연계일자가 없는 고객사: **{len(missing_erp)}건**")
                    st.dataframe(pd.DataFrame(missing_erp), use_container_width=True)

    # 날짜 컬럼 변환
    if open_date_col and open_date_col in cloud_df.columns:
        cloud_df[open_date_col] = pd.to_datetime(cloud_df[open_date_col], errors="coerce")
    if erp_date_col and erp_date_col in cloud_df.columns:
        cloud_df[erp_date_col] = pd.to_datetime(cloud_df[erp_date_col], errors="coerce")

    # 대상 년월 (2026-05)
    target_year_month = "2026-05"

    # 직원 정보 로드 (직급 정보 가져오기)
    user_db = load_db(DB_FILE, {})

    # 실적 데이터 계산
    performance_data = []

    for staff_name in staff_names:
        # 본사 시트에서 담당자 데이터 필터링
        staff_cloud_df = cloud_df[cloud_df[owner_col] == staff_name].copy()

        # 개설건수: 2026년 5월 개설완료일자 카운트
        open_count = 0
        if open_date_col and open_date_col in staff_cloud_df.columns:
            may_2026_data = staff_cloud_df[
                staff_cloud_df[open_date_col].notna() &
                (staff_cloud_df[open_date_col].dt.strftime("%Y-%m") == target_year_month)
            ]
            open_count = len(may_2026_data)

        # 연계건수: 2026년 5월 ERP연계일자 카운트
        erp_count = 0
        if erp_date_col and erp_date_col in staff_cloud_df.columns:
            may_erp_data = staff_cloud_df[
                staff_cloud_df[erp_date_col].notna() &
                (staff_cloud_df[erp_date_col].dt.strftime("%Y-%m") == target_year_month)
            ]
            erp_count = len(may_erp_data)

        # 운영 실적: 업로드한 엑셀 파일에서 계산
        operation_count = 0
        operation_points = 0

        if analysis_df is not None and not analysis_df.empty:
            analysis_clean = clean_header_logic(analysis_df.copy())
            u_col = find_col(analysis_clean, ["등록자", "담당자", "성명"])
            d_col = find_col(analysis_clean, ["활동상세", "활동내용"])
            category_col = find_col(analysis_clean, ["활동구분", "접수유형"])

            if u_col and u_col in analysis_clean.columns:
                staff_activity = analysis_clean[analysis_clean[u_col] == staff_name].copy()

                if category_col and category_col in staff_activity.columns:
                    # 본사이력 업로드는 활동구분(방문/원격)을 우선 기준으로 운영 실적을 반영한다.
                    category_text = staff_activity[category_col].astype(str)
                    visit_count = int(category_text.str.contains("방문", na=False).sum())
                    remote_count = int(category_text.str.contains("원격", na=False).sum())
                    operation_count = visit_count + remote_count
                    operation_points = (visit_count * 30) + (remote_count * 10)
                elif d_col and d_col in staff_activity.columns:
                    # 활동구분이 없으면 기존처럼 활동상세 기준으로 운영 실적을 계산한다.
                    operation_count = int(
                        staff_activity[d_col].astype(str).str.contains("운영|방문|점검", na=False).sum()
                    )
                    operation_points = operation_count * 30

        # 포인트 계산
        open_points = open_count * 90  # 개설 1건당 90포인트
        erp_points = erp_count * 120  # 연계 1건당 120포인트

        # 합계 계산
        total_points = min(2800, min(1000, open_points + erp_points) + min(1800, operation_points))
        pay_points = max(0, total_points - 1000)

        # 지급예상금액 계산 (임시로 500원/포인트 적용, 실제로는 직급/외주 여부 확인 필요)
        pay_amount = pay_points * 500

        # 직급 정보 가져오기
        staff_rank = "직원"  # 기본값
        for uid, user_info in user_db.items():
            if isinstance(user_info, dict) and user_info.get("name") == staff_name:
                staff_rank = user_info.get("rank", "직원")
                break

        row_data = {
            "담당자": staff_name,
            "직급": staff_rank,
            "개설건수": open_count,
            "개설포인트": open_points,
            "연계건수": erp_count,
            "연계포인트": erp_points,
            "운영건수 (실제 활동)": operation_count,
            "운영포인트 (실제 활동)": operation_points,
            "합계포인트": total_points,
            "지급포인트": pay_points,
            "지급예상금액": pay_amount,
        }

        performance_data.append(row_data)

    if not performance_data:
        st.info("집계된 실적이 없습니다.")
        return

    # DataFrame 생성
    perf_df = pd.DataFrame(performance_data)

    # 직급 순서 정의 및 정렬용 컬럼 추가
    rank_order = {"부서장": 0, "팀장": 1, "과장": 2, "대리": 3, "주임": 4, "직원": 5}
    perf_df["_rank_order"] = perf_df["직급"].map(lambda x: rank_order.get(x, 99))

    # 필터 및 검색
    col1, col2, col3 = st.columns([2, 2, 1])

    with col1:
        search_name = st.text_input("🔍 담당자 검색", placeholder="이름 입력")

    with col2:
        sort_by = st.selectbox(
            "정렬 기준",
            ["직급", "합계포인트", "지급예상금액", "개설건수", "연계건수", "운영건수 (실제 활동)", "담당자"],
            index=0
        )

    with col3:
        sort_order = st.radio("정렬 순서", ["오름차순 ⬆", "내림차순 ⬇"], horizontal=True)

    # 검색 필터 적용
    if search_name:
        perf_df = perf_df[perf_df["담당자"].astype(str).str.contains(search_name, na=False)]

    if perf_df.empty:
        st.warning("검색 결과가 없습니다.")
        return

    # 정렬 적용
    ascending = sort_order == "오름차순 ⬆"
    if sort_by == "직급":
        perf_df = perf_df.sort_values(by="_rank_order", ascending=ascending)
    else:
        perf_df = perf_df.sort_values(by=sort_by, ascending=ascending)

    # 통계 요약
    st.markdown("#### 📊 전체 실적 요약")
    summary_cols = st.columns(5)

    with summary_cols[0]:
        st.metric("총 인원", f"{len(perf_df)}명")

    with summary_cols[1]:
        total_open = int(perf_df["개설건수"].sum()) if "개설건수" in perf_df.columns else 0
        st.metric("총 개설", f"{total_open}건")

    with summary_cols[2]:
        total_link = int(perf_df["연계건수"].sum()) if "연계건수" in perf_df.columns else 0
        st.metric("총 연계", f"{total_link}건")

    with summary_cols[3]:
        total_points = int(perf_df["합계포인트"].sum()) if "합계포인트" in perf_df.columns else 0
        st.metric("총 포인트", f"{total_points:,}pt")

    with summary_cols[4]:
        total_amount = int(perf_df["지급예상금액"].sum()) if "지급예상금액" in perf_df.columns else 0
        st.metric("총 지급예상", f"{total_amount:,}원")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # 표시할 컬럼 선택
    display_cols = [
        "담당자", "직급", "개설건수", "개설포인트", "연계건수", "연계포인트",
        "운영건수 (실제 활동)", "운영포인트 (실제 활동)",
        "합계포인트", "지급포인트", "지급예상금액", "오류건수", "중복건수"
    ]

    # 존재하는 컬럼만 선택
    display_cols = [col for col in display_cols if col in perf_df.columns]

    # 합계 행 추가
    total_row = {}
    for col in display_cols:
        if col in ["담당자"]:
            total_row[col] = "합계"
        elif col in ["직급"]:
            total_row[col] = ""
        elif col in perf_df.columns:
            if perf_df[col].dtype in ['int64', 'float64']:
                total_row[col] = perf_df[col].sum()
            else:
                total_row[col] = ""
        else:
            total_row[col] = ""

    perf_df_with_total = pd.concat([perf_df[display_cols], pd.DataFrame([total_row])], ignore_index=True)

    # Excel 다운로드 및 새로고침
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    col_dl1, col_dl2, col_dl3 = st.columns([1, 1, 2])

    with col_dl1:
        from datetime import datetime
        download_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_bytes = BytesIO()
        perf_df[display_cols].to_excel(excel_bytes, index=False, engine="openpyxl")
        excel_bytes.seek(0)

        st.download_button(
            "📥 Excel 다운로드",
            data=excel_bytes.getvalue(),
            file_name=f"전체실적_{download_time}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_dl2:
        if st.button("🔄 새로고침", use_container_width=True):
            # 본사 구글시트와 하나은행 시트 데이터 새로고침
            st.session_state.cloud_sheet_df = None
            st.session_state.analysis_lookup_df = None
            st.session_state.hana_sheet_df = None
            st.session_state.admin_uploaded_excel = None
            st.session_state.admin_uploaded_excel_display = None
            st.toast("데이터를 새로고침합니다.")
            st.rerun()

    # 엑셀 파일 업로드 기능
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    st.markdown("#### 📤 엑셀 파일 업로드")

    # 업로드된 데이터 확인
    if st.session_state.get("admin_uploaded_excel") is not None:
        uploaded_data = st.session_state.get("admin_uploaded_excel")
        with st.expander("🔍 업로드된 데이터 확인", expanded=False):
            st.write(f"**업로드 데이터 건수:** {len(uploaded_data)}")
            if not uploaded_data.empty:
                analysis_clean = clean_header_logic(uploaded_data.copy())
                u_col = find_col(analysis_clean, ["등록자", "담당자", "성명"])
                d_col = find_col(analysis_clean, ["활동상세", "활동내용"])
                st.write(f"**담당자 컬럼명:** {u_col}")
                st.write(f"**활동상세 컬럼명:** {d_col}")
                if u_col and u_col in analysis_clean.columns:
                    unique_users = analysis_clean[u_col].unique()
                    st.write(f"**담당자 목록:** {list(unique_users)}")
                st.write("**데이터 샘플:**")
                st.dataframe(uploaded_data.head(10))

    col1, col_convert, col_upload, col_sample, _ = st.columns([1, 1, 1, 1, 2])
    with col1:
        st.markdown("<div style='text-align:center;font-weight:700;margin-bottom:4px;'>은행 이력 업로드</div>", unsafe_allow_html=True)
        history_file = st.file_uploader("은행 이력 업로드", type=["xls", "xlsx"], key="admin_history_upload", label_visibility="collapsed")
        st.caption("은행에서 반출해준 이력파일 그대로 업로드 해주세요.")

    with col_convert:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if history_file is not None:
            try:
                with st.spinner("은행 이력을 샘플 양식으로 변환 중입니다."):
                    history_df = read_excel_history_file(history_file)
                    # 관리자 모드에서는 담당자명이 없으므로 빈 문자열 사용
                    converted_df, convert_info = convert_history_to_sample_df(history_df, "")
                if converted_df.empty:
                    st.button("변환파일 다운로드", use_container_width=True, disabled=True)
                    st.warning(convert_info.get("error", "변환할 데이터가 없습니다."))
                else:
                    converted_df = normalize_converted_history_df(converted_df)
                    # session_state에 저장하여 운영 카운트에 사용
                    st.session_state.admin_uploaded_excel = converted_df
                    st.session_state.admin_uploaded_excel_display = converted_df.copy()
                    converted_bytes = sample_format_excel_bytes(converted_df)
                    converted_ym = get_uploaded_month(converted_df).replace("-", "") or (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m")
                    st.download_button(
                        "변환파일 다운로드",
                        data=converted_bytes,
                        file_name=f"LMB월간 활동실적_{converted_ym}_관리자.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                    unmatched = int(convert_info.get("unmatched", 0))
                    if unmatched:
                        st.caption(f"고객번호 매핑 실패 {unmatched}건은 사업자번호가 공란으로 저장됩니다.")
                    st.success("✅ 업로드한 파일이 실적 계산에 반영됩니다.")
            except ImportError:
                st.button("변환파일 다운로드", use_container_width=True, disabled=True)
                st.error("xls 파일 변환을 위해 xlrd 패키지가 필요합니다.")
            except Exception as e:
                st.button("변환파일 다운로드", use_container_width=True, disabled=True)
                st.error(f"은행 이력 업로드 실패: {e}")
        else:
            st.button("변환파일 다운로드", use_container_width=True, disabled=True)

    with col_upload:
        st.markdown("<div style='text-align:center;font-weight:700;margin-bottom:4px;'>본사이력 업로드 (선택)</div>", unsafe_allow_html=True)
        u_file = st.file_uploader("본사이력 업로드 (선택)", type=["xlsx"], key="admin_office_upload", label_visibility="collapsed")
        if u_file is not None:
            try:
                office_df = pd.read_excel(u_file)
                st.session_state.admin_uploaded_excel_display = clean_header_logic(office_df.copy())
                office_df = normalize_converted_history_df(office_df)
                # session_state에 저장하여 운영 카운트에 사용
                st.session_state.admin_uploaded_excel = office_df
                st.success("✅ 본사이력 파일이 실적 계산에 반영됩니다.")
            except Exception as e:
                st.error(f"본사이력 업로드 실패: {e}")

    with col_sample:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if os.path.exists(EXCEL_SAMPLE_FILE):
            with open(EXCEL_SAMPLE_FILE, "rb") as sample_file:
                st.download_button(
                    "샘플파일 다운로드",
                    data=sample_file.read(),
                    file_name="LMB월간 활동실적_000000(샘플).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.button("샘플파일 다운로드", use_container_width=True, disabled=True)


    if (
        isinstance(st.session_state.get("admin_uploaded_excel"), pd.DataFrame)
        and not st.session_state.admin_uploaded_excel.empty
    ):
        display_source_df = st.session_state.get("admin_uploaded_excel_display")
        if not isinstance(display_source_df, pd.DataFrame) or display_source_df.empty:
            display_source_df = st.session_state.admin_uploaded_excel
        admin_uploaded_df = display_source_df.copy()
        admin_uploaded_df = admin_uploaded_df[
            [
                col for col in admin_uploaded_df.columns
                if not str(col).strip().lower().startswith("unnamed")
                and not str(col).strip().startswith("_")
            ]
        ]
        admin_uploaded_df = admin_uploaded_df.loc[
            :,
            admin_uploaded_df.apply(lambda col: col.astype(str).str.strip().replace("nan", "").ne("").any(), axis=0)
        ]
        st.markdown("#### 본사이력 업로드 데이터")
        st.caption(f"업로드 데이터 건수: {len(admin_uploaded_df):,}건")
        st.dataframe(admin_uploaded_df, use_container_width=True, hide_index=True)

    st.markdown("#### 담당자별 상세 실적")
    st.dataframe(
        perf_df_with_total,
        use_container_width=True,
        hide_index=True
    )


def show_user_history(is_admin_mode=False):
    # 관리자 모드: 담당자 선택
    selected_user = st.session_state.user_name  # 기본값: 본인

    if is_admin_mode:
        st.markdown("### 관리자용 실적 확인")

        # C&S 부서 직원 목록만 가져오기
        user_db = load_db(DB_FILE, {})
        staff_names = []
        for uid, user_info in user_db.items():
            if uid != "1" and isinstance(user_info, dict) and user_info.get("name") and user_info.get("dept_type") == "C&S":
                staff_names.append(user_info.get("name"))

        staff_names = sorted(staff_names)

        if not staff_names:
            st.warning("등록된 직원이 없습니다. [직원 및 권한설정] 메뉴에서 직원을 등록해주세요.")
            return

        # "전체" 옵션 추가
        staff_options = ["📊 전체"] + staff_names

        # 담당자 선택 드롭다운
        col_select1, col_select2 = st.columns([1, 3])
        with col_select1:
            selected_option = st.selectbox(
                "👤 담당자 선택",
                options=staff_options,
                index=0
            )

        # "전체" 선택 시 전체 실적 요약 표시
        if selected_option == "📊 전체":
            show_all_staff_summary(staff_names)
            return

        selected_user = selected_option
        st.markdown(f"**{selected_user}**님의 실적을 확인/수정합니다.")
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # 선택된 담당자로 작업 진행
    original_user_name = st.session_state.user_name
    if is_admin_mode:
        st.session_state.user_name = selected_user

    converted_preview_df = None
    analysis_df = None
    converted_ym = ""
    convert_info = {}
    col1, col_convert, col_upload, col_sample, _ = st.columns([1, 1, 1, 1, 2])
    with col1:
        st.markdown("<div style='text-align:center;font-weight:700;margin-bottom:4px;'>은행 이력 업로드</div>", unsafe_allow_html=True)
        history_file = st.file_uploader("은행 이력 업로드", type=["xls", "xlsx"], key="history_convert_upload", label_visibility="collapsed")
        st.caption("은행에서 반출해준 이력파일 그대로 업로드 해주세요.")
    with col_convert:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if history_file is not None:
            try:
                with st.spinner("은행 이력을 샘플 양식으로 변환 중입니다."):
                    history_df = read_excel_history_file(history_file)
                    converted_df, convert_info = convert_history_to_sample_df(history_df, st.session_state.user_name)
                if converted_df.empty:
                    st.button("변환파일 다운로드", use_container_width=True, disabled=True)
                    st.warning(convert_info.get("error", "변환할 데이터가 없습니다."))
                else:
                    converted_df = normalize_converted_history_df(converted_df)
                    converted_preview_df = converted_df
                    edited_key = "history_convert_preview_editor"
                    data_key = "history_convert_preview_data"
                    # 새 파일 업로드 여부 확인 → 파일이 바뀌면 preview 초기화
                    _file_id = f"{history_file.name}_{history_file.size}"
                    if st.session_state.get("_bank_file_id") != _file_id:
                        st.session_state["_bank_file_id"] = _file_id
                        st.session_state[data_key] = converted_df
                        edited_df = converted_df
                    else:
                        edited_df = st.session_state.get(data_key, converted_df)
                        if not isinstance(edited_df, pd.DataFrame) or list(edited_df.columns) != list(converted_df.columns):
                            edited_df = converted_df
                    edited_df = normalize_converted_history_df(edited_df)
                    st.session_state[data_key] = edited_df
                    st.session_state.user_excel_data = edited_df
                    st.session_state.user_excel_source = "bank"
                    analysis_df = edited_df
                    converted_bytes = sample_format_excel_bytes(edited_df)
                    converted_ym = get_uploaded_month(edited_df).replace("-", "") or (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m")
                    st.download_button(
                        "변환파일 다운로드",
                        data=converted_bytes,
                        file_name=f"LMB월간 활동실적_{converted_ym}_{st.session_state.user_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                    unmatched = int(convert_info.get("unmatched", 0))
                    if unmatched:
                        st.caption(f"고객번호 매핑 실패 {unmatched}건은 사업자번호가 공란으로 저장됩니다.")
            except ImportError:
                st.button("변환파일 다운로드", use_container_width=True, disabled=True)
                st.error("xls 파일 변환을 위해 xlrd 패키지가 필요합니다. 배포 후 requirements.txt 반영을 확인해주세요.")
            except Exception as e:
                st.button("변환파일 다운로드", use_container_width=True, disabled=True)
                st.error(f"은행 이력 업로드 실패: {e}")
        else:
            st.button("변환파일 다운로드", use_container_width=True, disabled=True)
    with col_upload:
        st.markdown("<div style='text-align:center;font-weight:700;margin-bottom:4px;'>본사이력 업로드 (선택)</div>", unsafe_allow_html=True)
        u_file = st.file_uploader("본사이력 업로드 (선택)", type=["xlsx"], label_visibility="collapsed")
    with col_sample:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if os.path.exists(EXCEL_SAMPLE_FILE):
            with open(EXCEL_SAMPLE_FILE, "rb") as sample_file:
                st.download_button(
                    "샘플파일 다운로드",
                    data=sample_file.read(),
                    file_name="LMB월간 활동실적_000000(샘플).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.button("샘플파일 다운로드", use_container_width=True, disabled=True)

    if False and converted_preview_df is not None and not converted_preview_df.empty:
        st.markdown("#### 변환파일 미리보기")
        st.markdown(
            """
            <style>
            body[data-pms-theme="d"] [data-testid="stDataEditor"],
            body[data-pms-theme="d"] [data-testid="stDataEditor"] > div,
            body[data-pms-theme="d"] [data-testid="stDataEditor"] [role="grid"],
            body[data-pms-theme="d"] [data-testid="stDataEditor"] canvas,
            body:has(#pms-d:checked) [data-testid="stDataEditor"],
            body:has(#pms-d:checked) [data-testid="stDataEditor"] > div,
            body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="grid"],
            body:has(#pms-d:checked) [data-testid="stDataEditor"] canvas {
                background-color: #252535 !important;
                color: #ffffff !important;
            }
            body[data-pms-theme="d"] [data-testid="stDataEditor"] canvas,
            body:has(#pms-d:checked) [data-testid="stDataEditor"] canvas {
                filter: invert(1) hue-rotate(180deg) brightness(0.72) contrast(1.18) saturate(0.85) !important;
            }
            body[data-pms-theme="d"] [data-testid="stDataEditor"] [role="columnheader"],
            body[data-pms-theme="d"] [data-testid="stDataEditor"] [role="gridcell"],
            body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="columnheader"],
            body:has(#pms-d:checked) [data-testid="stDataEditor"] [role="gridcell"] {
                background-color: #252535 !important;
                color: #ffffff !important;
                border-color: #45475a !important;
            }
            body[data-pms-theme="d"] [data-testid="stDataEditor"] input,
            body[data-pms-theme="d"] [data-testid="stDataEditor"] textarea,
            body[data-pms-theme="d"] [data-testid="stDataEditor"] [contenteditable="true"],
            body[data-pms-theme="d"] [data-testid="stDataEditor"] [data-baseweb="input"] input,
            body:has(#pms-d:checked) [data-testid="stDataEditor"] input,
            body:has(#pms-d:checked) [data-testid="stDataEditor"] textarea,
            body:has(#pms-d:checked) [data-testid="stDataEditor"] [contenteditable="true"],
            body:has(#pms-d:checked) [data-testid="stDataEditor"] [data-baseweb="input"] input {
                background-color: #0f0f1f !important;
                color: #ffffff !important;
                caret-color: #ffffff !important;
                border-color: #89b4fa !important;
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
                body:has(#pms-s:checked) [data-testid="stDataEditor"] [role="grid"],
                body:has(#pms-s:checked) [data-testid="stDataEditor"] canvas {
                    background-color: #252535 !important;
                    color: #ffffff !important;
                }
                body:has(#pms-s:checked) [data-testid="stDataEditor"] canvas {
                    filter: invert(1) hue-rotate(180deg) brightness(0.72) contrast(1.18) saturate(0.85) !important;
                }
                body:has(#pms-s:checked) [data-testid="stDataEditor"] input,
                body:has(#pms-s:checked) [data-testid="stDataEditor"] textarea,
                body:has(#pms-s:checked) [data-testid="stDataEditor"] [contenteditable="true"],
                body:has(#pms-s:checked) [data-testid="stDataEditor"] [data-baseweb="input"] input {
                    background-color: #0f0f1f !important;
                    color: #ffffff !important;
                    caret-color: #ffffff !important;
                    border-color: #89b4fa !important;
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
        edited_preview_df = st.data_editor(
            editor_source_df,
            key="history_convert_preview_editor",
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            disabled=[col for col in editor_source_df.columns if col not in ["활동일자", "활동구분", "활동상세"]],
            column_config={
                "지사": st.column_config.TextColumn("지사", disabled=True),
                "활동일자": st.column_config.TextColumn("활동일자"),
                "활동구분": st.column_config.SelectboxColumn(
                    "활동구분",
                    options=["방문", "상담", "원격"],
                    required=True,
                ),
                "활동상세": st.column_config.SelectboxColumn(
                    "활동상세",
                    options=["운영", "개설", "연계"],
                    required=True,
                ),
                "활동내역": st.column_config.TextColumn("활동내역"),
            },
        )
        analysis_df = normalize_converted_history_df(edited_preview_df)
        st.session_state[data_key] = analysis_df
        _, add_history_col = st.columns([0.88, 0.12])
        with add_history_col:
            if st.button("이력 추가", use_container_width=True, key="add_history_row"):
                new_row = {col: "" for col in analysis_df.columns}
                new_row.update({
                    "지사": "HANA지사",
                    "상품": "통합CMS",
                    "등록자": st.session_state.user_name,
                    "활동일자": (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d"),
                    "활동구분": "방문",
                    "활동상세": "운영",
                    "제목": "운영방문",
                })
                st.session_state[data_key] = normalize_converted_history_df(
                    pd.concat([analysis_df, pd.DataFrame([new_row])], ignore_index=True)
                )
                st.rerun()
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    if converted_preview_df is not None and not converted_preview_df.empty:
        _norm_cpdf = normalize_converted_history_df(converted_preview_df)
        _saved_preview = st.session_state.get("history_convert_preview_data")
        if isinstance(_saved_preview, pd.DataFrame) and list(_saved_preview.columns) == list(_norm_cpdf.columns):
            analysis_df = _saved_preview
        else:
            analysis_df = _norm_cpdf
        if st.session_state.get("user_excel_source") in (None, "bank"):
            st.session_state.user_excel_data = analysis_df
            st.session_state.user_excel_source = "bank"

    # 파일이 제거되면 데이터 초기화
    if u_file is None:
        if st.session_state.get("user_excel_source") == "hq" and st.session_state.get("user_excel_data") is not None:
            st.session_state.user_excel_data = None
            st.session_state.user_excel_source = None
            st.session_state.user_excel_file_key = None
            st.rerun()

    # 파일 업로드 시 자동 처리
    if u_file is not None:
        file_key = f"{u_file.name}_{u_file.size}"

        # 파일이 변경되었을 때만 처리
        if st.session_state.get("user_excel_file_key") != file_key:
            st.session_state.user_excel_file_key = file_key
            with st.spinner("본사이력 파일을 불러오는 중입니다..."):
                # cloud_sheet_df가 없을 때만 로드 시도
                if st.session_state.get("cloud_sheet_df") is None:
                    try:
                        load_csv_to_state("url_sync", "cloud_sheet_df")
                    except Exception as e:
                        st.warning(f"⚠️ 본사 구글시트를 불러오는데 실패했습니다: {str(e)}")

                uploaded_df = clean_header_logic(pd.read_excel(u_file, sheet_name=0))

                if not uploaded_df.empty:
                    st.session_state.user_excel_data = uploaded_df
                    st.session_state.user_excel_source = "hq"
                    st.toast("업로드 완료. 분석이 자동으로 시작됩니다.")
                    st.rerun()
                else:
                    st.error("업로드된 데이터가 없습니다. 파일을 확인해주세요.")
                    st.session_state.user_excel_data = None
                    st.session_state.user_excel_source = None

    if (
        st.session_state.get("user_excel_source") == "hq"
        and isinstance(st.session_state.get("user_excel_data"), pd.DataFrame)
        and not st.session_state.user_excel_data.empty
    ):
        hq_uploaded_df = st.session_state.user_excel_data.copy()
        hq_uploaded_df = hq_uploaded_df[
            [
                col for col in hq_uploaded_df.columns
                if not str(col).strip().lower().startswith("unnamed")
                and not str(col).strip().startswith("_")
            ]
        ]
        hq_uploaded_df = hq_uploaded_df.loc[
            :,
            hq_uploaded_df.apply(lambda col: col.astype(str).str.strip().replace("nan", "").ne("").any(), axis=0)
        ]
        st.markdown("#### 본사이력 업로드 데이터")
        st.caption(f"업로드 데이터 건수: {len(hq_uploaded_df):,}건")
        st.dataframe(hq_uploaded_df, use_container_width=True, hide_index=True)
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    current_df = current_history_analysis_df()
    if not current_df.empty:
        analysis_df = current_df
    elif st.session_state.user_excel_data is not None:
        analysis_df = prepare_history_analysis_df(st.session_state.user_excel_data)

    if analysis_df is None or analysis_df.empty:
        return

    if st.session_state.get("cloud_sheet_df") is None:
        try:
            load_csv_to_state("url_sync", "cloud_sheet_df")
        except Exception:
            pass

    df = analysis_df.copy()
    u_col = find_col(df, ["등록자", "담당자", "성명"], "등록자")
    d_col = find_col(df, ["활동상세", "활동내용"], "활동상세")

    df_user = df[df[u_col] == st.session_state.user_name].copy() if u_col in df.columns else df.iloc[0:0].copy()
    df_user = attach_cloud_dates(df_user)
    df_user_visit = filter_visit_rows(df_user)

    st.markdown("### 담당자 기본 활동 수치")
    res, err, dup = process_performance_analysis(df_user_visit, st.session_state.get("auto_prev_df"))

    if isinstance(res, pd.DataFrame) and not res.empty:
        my_res = res[res["담당자"] == st.session_state.user_name].copy()

        # 업로드 전 예상치 계산 (추가 활동 제외)
        if not my_res.empty:
            before_res = my_res.copy()
            o_p = int(float(before_res.iloc[0].get("개설포인트", 0)))
            l_p = int(float(before_res.iloc[0].get("연계포인트", 0)))
            v_p = int(float(before_res.iloc[0].get("운영포인트 (실제 활동)", 0)))

            # 합계포인트: 개설 + 연계 + 운영(실제)만
            before_total = min(2800, min(1000, o_p + l_p) + min(1800, v_p))
            before_res.loc[before_res.index[0], "합계포인트"] = before_total

            # 지급포인트
            before_pay_point = max(0, before_total - 1000)
            before_res.loc[before_res.index[0], "지급포인트"] = before_pay_point

            # 지급예상금액
            rank = before_res.iloc[0].get("직급", "")
            name = before_res.iloc[0].get("담당자", "")
            name_to_info = {
                info.get("name"): {
                    "staff_type": info.get("staff_type", "정규직"),
                }
                for uid, info in st.session_state.user_db.items()
                if uid != "1"
            }
            is_outsource = name_to_info.get(name, {}).get("staff_type", "정규직") == "외주"

            if is_outsource:
                before_pay = int(max(0, before_pay_point * 500))
            else:
                leader_bonus = 500 if rank == "팀장" else 0
                before_pay = int(max(0, (before_pay_point + leader_bonus) * 500))

            before_res.loc[before_res.index[0], "지급예상금액"] = before_pay

            # 전월대비 재계산
            if "전월대비" in before_res.columns:
                prev_df = st.session_state.get("auto_prev_df")
                if prev_df is not None:
                    try:
                        prev_res, _, _ = process_performance_analysis(prev_df)
                        if isinstance(prev_res, pd.DataFrame) and not prev_res.empty:
                            p_map = prev_res.set_index("담당자")["지급예상금액"].to_dict()
                            prev_pay = p_map.get(name, 0)
                            before_res.loc[before_res.index[0], "전월대비"] = int(before_pay - prev_pay)
                    except Exception:
                        pass

            hidden_cols = ["운영건수 (추가 활동)", "운영포인트(추가 활동)"] + [c for c in before_res.columns if "전월대비" in c]
            my_res_display = before_res.drop(columns=hidden_cols, errors="ignore")
        else:
            drop_cols = [c for c in my_res.columns if "전월대비" in c]
            my_res_display = my_res.drop(columns=drop_cols, errors="ignore")

        style_report_logic(my_res_display, compact=True)
    elif isinstance(res, str):
        st.error(res)

    # 탭 데이터 미리 계산 (경고 메시지 표시용)
    # 초과 방문 데이터 계산
    err_filtered = pd.DataFrame()
    if err is not None and not err.empty:
        if "담당자" in err.columns:
            err_my = err[err["담당자"] == st.session_state.user_name].copy()
        else:
            err_my = err.copy()
        err_filtered = err_my[err_my["일방문"] >= 6].copy()

    # 기타 오류 데이터 계산
    other_errors_df = build_other_validation_errors(df_user_visit)

    # 중복 이력 데이터
    dup_my = pd.DataFrame()
    if dup is not None and not dup.empty:
        u_col_dup = find_col(dup, ["등록자", "담당자", "성명"], "담당자")
        if u_col_dup and u_col_dup in dup.columns:
            dup_my = dup[dup[u_col_dup] == st.session_state.user_name]
        else:
            dup_my = dup

    # 개설완료일자 누락
    missing_open = pd.DataFrame()
    if "본사 개설완료일자" in df_user_visit.columns:
        missing_open = df_user_visit[
            pd.isna(df_user_visit["본사 개설완료일자"]) | (df_user_visit["본사 개설완료일자"].astype(str).str.strip() == "")
        ]
        if "본사 신규이행구분" in missing_open.columns:
            missing_open = missing_open[missing_open["본사 신규이행구분"].astype(str).str.strip() != "이행"]

    # ERP연계일자 누락
    missing_erp = pd.DataFrame()
    if "본사 ERP연계일자" in df_user_visit.columns:
        if d_col and d_col in df_user_visit.columns:
            target = df_user_visit[df_user_visit[d_col].astype(str).str.contains("연계", na=False)]
        else:
            target = df_user_visit
        missing_erp = target[
            pd.isna(target["본사 ERP연계일자"]) | (target["본사 ERP연계일자"].astype(str).str.strip() == "")
        ]
        # 신규/이행구분이 "이행"이고 이행추가연계가 비어있으면 ERP연계 불필요 → 제외
        if "본사 신규이행구분" in missing_erp.columns and "본사 이행추가연계" in missing_erp.columns:
            _is_ihang = missing_erp["본사 신규이행구분"].astype(str).str.strip() == "이행"
            _add_empty = pd.isna(missing_erp["본사 이행추가연계"]) | (missing_erp["본사 이행추가연계"].astype(str).str.strip() == "")
            missing_erp = missing_erp[~(_is_ihang & _add_empty)]

    # 검증 이슈 확인
    has_validation_issues = (
        not dup_my.empty or
        not err_filtered.empty or
        not missing_open.empty or
        not missing_erp.empty or
        not other_errors_df.empty
    )

    # 경고 메시지 표시
    error_tabs = []
    if not dup_my.empty:
        error_tabs.append("중복 이력")
    if not err_filtered.empty:
        error_tabs.append("초과 방문")
    if not missing_open.empty:
        error_tabs.append("본사 개설완료일자 누락")
    if not missing_erp.empty:
        error_tabs.append("본사 ERP연계일자 누락")
    if not other_errors_df.empty:
        error_tabs.append("기타 오류")

    st.divider()
    st.markdown("### 추가 실적 표")

    base = criteria_df()
    saved = load_db(PERF_FILE, {}).get(st.session_state.user_name, {})
    base["입력(건)"] = base["구분"].map(saved).fillna(0).astype(int)

    edited = render_manual_perf_input_table(base)

    edited["입력(건)"] = pd.to_numeric(edited["입력(건)"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    edited["계산점수"] = edited["단위 점수"] * edited["입력(건)"]
    manual_override = st.session_state.setdefault("manual_perf_preview_override", {})
    manual_override[st.session_state.user_name] = edited.set_index("구분")["입력(건)"].to_dict()

    total = calculate_manual_perf_total(edited)
    st.session_state["_manual_perf_total"] = total

    st.metric("추가 실적 합산 점수", f"{total:,} PT")

    st.divider()
    preview_source_df = st.session_state.get("history_convert_preview_data", converted_preview_df)
    if not isinstance(preview_source_df, pd.DataFrame):
        preview_source_df = converted_preview_df if isinstance(converted_preview_df, pd.DataFrame) else df_user_visit
    preview_source_df = normalize_converted_history_df(preview_source_df)
    search_source_df = preview_source_df if not preview_source_df.empty else df_user_visit
    search_filters = {
        "company": str(st.session_state.get("history_preview_search_company", "") or "").strip(),
        "date": st.session_state.get("history_preview_search_date", "전체"),
        "category": st.session_state.get("history_preview_search_category", "전체"),
        "detail": st.session_state.get("history_preview_search_detail", "전체"),
    }

    _render_preview = converted_preview_df
    if _render_preview is None or (isinstance(_render_preview, pd.DataFrame) and _render_preview.empty):
        _sess_prev = st.session_state.get("history_convert_preview_data")
        if isinstance(_sess_prev, pd.DataFrame) and not _sess_prev.empty:
            _render_preview = _sess_prev

    if _render_preview is not None and not _render_preview.empty:
        _pre_h = None
        _pd_before = st.session_state.get("history_convert_preview_data")
        if isinstance(_pd_before, pd.DataFrame) and not _pd_before.empty:
            try:
                _pre_h = int(pd.util.hash_pandas_object(_pd_before).sum())
            except Exception:
                pass
        render_converted_preview_editor(_render_preview, search_filters)
        _changed = False
        _pd_after = st.session_state.get("history_convert_preview_data")
        if isinstance(_pd_after, pd.DataFrame) and not _pd_after.empty:
            try:
                if _pre_h is not None and int(pd.util.hash_pandas_object(_pd_after).sum()) != _pre_h:
                    _changed = True
            except Exception:
                pass
        if _changed:
            st.rerun()

        # 추가 방문 운영 이력 등록 안내
        _add_visit_cnt = round(total / 30)
        if _add_visit_cnt > 0:
            st.markdown(
                f"<div style='margin-top:10px;padding:12px 16px;background:#EBF8FF;border-left:4px solid #4299E1;border-radius:6px;font-size:14px;color:#2B6CB0;'>"
                f"<b>📋 추가 이력 등록 안내</b><br>"
                f"추가 실적 합산 점수 <b>{total:,} PT</b> 기준으로 "
                f"<b style='font-size:16px;color:#C05621;'>{_add_visit_cnt}건</b>의 방문 운영 이력을 "
                f"우측 <b>이력 추가</b> 버튼을 통해 등록해주세요."
                f"</div>",
                unsafe_allow_html=True,
            )

    search_filters = render_history_search_filters(search_source_df, "history_preview_search")

    dup_my = apply_history_search_filters(dup_my, search_filters)
    err_filtered = apply_history_search_filters(err_filtered, search_filters)
    missing_open = apply_history_search_filters(missing_open, search_filters)
    missing_erp = apply_history_search_filters(missing_erp, search_filters)
    other_errors_df = apply_history_search_filters(other_errors_df, search_filters)

    t1, t2, t3, t4, t5 = validation_tabs_with_refresh("refresh_user_history_validation")
    with t1:
        if not dup_my.empty:
            _dc = next((c for c in dup_my.columns if "활동일자" in c), None) or next((c for c in dup_my.columns if "활동일" in c), None)
            dup_display = dup_my.copy()
            dup_display.insert(0, "_orig_idx", dup_my.index.astype(int))
            dup_display = dup_display.reset_index(drop=True)
            _ver = st.session_state.get("dup_editor_ver", 0)
            _dup_hidden = {c: None for c in dup_display.columns if c.startswith("_")}
            edited_dup = st.data_editor(
                dup_display,
                key=f"dup_date_editor_{_ver}",
                use_container_width=True,
                hide_index=True,
                disabled=[c for c in dup_display.columns if c != _dc],
                column_config=_dup_hidden,
            )
            if _dc and list(edited_dup[_dc]) != list(dup_display[_dc]):
                _preview = st.session_state.get("history_convert_preview_data")
                if _preview is not None:
                    _preview = _preview.copy()
                    _pd = find_col(_preview, ["활동일자", "활동일"])
                    if _pd:
                        for (_, o_row), (_, e_row) in zip(dup_display.iterrows(), edited_dup.iterrows()):
                            if str(o_row.get(_dc, "")) != str(e_row.get(_dc, "")):
                                try:
                                    _idx = int(o_row["_orig_idx"])
                                    _preview.loc[_idx, _pd] = e_row[_dc]
                                except (KeyError, IndexError, TypeError, ValueError):
                                    pass
                        st.session_state["history_convert_preview_data"] = _preview
                        st.session_state.pop("history_convert_preview_editor", None)
                        st.session_state["dup_editor_ver"] = _ver + 1
                        st.rerun()
        else:
            st.info("중복 이력이 없습니다.")
    with t2:
        if not err_filtered.empty:
            _disp = err_filtered.drop(columns=["월총방문"], errors="ignore").copy()
            _ver2 = st.session_state.get("err_editor_ver", 0)
            _err_hidden = {c: None for c in _disp.columns if c.startswith("_")}
            edited_err = st.data_editor(
                _disp,
                key=f"err_date_editor_{_ver2}",
                use_container_width=True,
                hide_index=True,
                disabled=[c for c in _disp.columns if c != "초과일자"],
                column_config=_err_hidden if _err_hidden else None,
            )
            if "초과일자" in _disp.columns and list(edited_err["초과일자"]) != list(_disp["초과일자"]):
                _preview = st.session_state.get("history_convert_preview_data")
                if _preview is not None:
                    _preview = _preview.copy()
                    _pu = find_col(_preview, ["등록자", "담당자", "성명"])
                    _pb = find_col(_preview, ["사업자번호"])
                    _pd = find_col(_preview, ["활동일자", "활동일"])
                    if _pu and _pb and _pd:
                        for (_, o_row), (_, e_row) in zip(_disp.iterrows(), edited_err.iterrows()):
                            if str(o_row["초과일자"]) != str(e_row["초과일자"]):
                                _mask = (
                                    (_preview[_pb].astype(str) == str(o_row["사업자번호"])) &
                                    (_preview[_pu].astype(str) == str(o_row["담당자"])) &
                                    (_preview[_pd].astype(str) == str(o_row["초과일자"]))
                                )
                                if _mask.any():
                                    _preview.loc[_mask, _pd] = e_row["초과일자"]
                        st.session_state["history_convert_preview_data"] = _preview
                        st.session_state.pop("history_convert_preview_editor", None)
                        st.session_state["err_editor_ver"] = _ver2 + 1
                        st.rerun()
        else:
            st.info("초과 방문 데이터가 없습니다.")
    with t3:
        if not missing_open.empty:
            style_report_logic(missing_open.drop(columns=["본사 ERP연계일자"], errors="ignore"))
        elif "본사 개설완료일자" not in df_user.columns:
            st.info("본사 구글시트에 개설완료일자 또는 사업자번호 컬럼이 없어 확인할 수 없습니다.")
    with t4:
        if not missing_erp.empty:
            style_report_logic(missing_erp.drop(columns=["본사 개설완료일자"], errors="ignore"))
        elif "본사 ERP연계일자" not in df_user.columns:
            st.info("본사 구글시트에 ERP연계일자 또는 사업자번호 컬럼이 없어 확인할 수 없습니다.")
    with t5:
        style_report_logic(other_errors_df, align_overrides={"오류 사유": "left"}, default_align="center")

    if error_tabs:
        error_list = ", ".join(error_tabs)
        st.markdown(
            f"<div style='margin-top:12px;margin-bottom:12px;padding:12px 16px;background:#FFF3CD;border-left:4px solid #FFC107;border-radius:4px;'>"
            f"<div style='font-size:14px;color:#856404;'><b>⚠️ 다음 항목을 수정해주세요:</b> {error_list}</div>"
            f"<div style='margin-top:4px;font-size:13px;color:#856404;'>위 탭에서 문제를 해결한 후 엑셀을 다시 업로드해주세요.</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"<div style='margin-top:12px;margin-bottom:12px;padding:12px 16px;background:#D4EDDA;border-left:4px solid #28A745;border-radius:4px;'>"
            f"<div style='font-size:14px;color:#155724;'><b>✓ 잘못된 데이터가 없습니다.</b></div>"
            f"</div>",
            unsafe_allow_html=True,
        )

    st.divider()
    st.markdown("### 최종 실적 확인")
    show_final_check()

    # 관리자 모드: user_name 복원
    if is_admin_mode:
        st.session_state.user_name = original_user_name


def show_final_check():
    # 저장된 데이터가 있으면 로드
    if current_history_analysis_df().empty and st.session_state.user_excel_data is None:
        saved_db = load_db(SAVED_STATE_FILE, {})
        user_saved = saved_db.get(st.session_state.user_name)
        if user_saved and user_saved.get("user_excel_data"):
            st.session_state.user_excel_data = pd.DataFrame.from_dict(user_saved["user_excel_data"])
            if user_saved.get("user_prev_month_sel"):
                st.session_state.user_prev_month_sel = user_saved["user_prev_month_sel"]

    select_prev_month("auto_prev_df", "user_prev_month_sel")

    original_df = current_history_analysis_df()
    if original_df.empty and st.session_state.user_excel_data is not None:
        original_df = prepare_history_analysis_df(st.session_state.user_excel_data)

    if original_df is None or original_df.empty:
        st.info("먼저 업로드 및 실적 확인 메뉴에서 엑셀을 업로드해주세요.")
        return

    # cloud_sheet_df가 없을 때만 로드 시도
    if st.session_state.get("cloud_sheet_df") is None:
        try:
            load_csv_to_state("url_sync", "cloud_sheet_df")
        except Exception as e:
            st.warning(f"⚠️ 본사 구글시트를 불러오는데 실패했습니다: {str(e)}")

    # Filter user data and attach cloud dates for tabs
    df_for_tabs = original_df.copy()
    u_col = find_col(df_for_tabs, ["등록자", "담당자", "성명"], "등록자")
    d_col = find_col(df_for_tabs, ["활동상세", "활동내용"], "활동상세")
    df_user = df_for_tabs[df_for_tabs[u_col] == st.session_state.user_name].copy() if u_col in df_for_tabs.columns else pd.DataFrame()
    df_user = attach_cloud_dates(df_user)
    df_user_visit = filter_visit_rows(df_user)

    res, err, dup = process_performance_analysis(filter_visit_rows(original_df), st.session_state.get("auto_prev_df"))

    if not isinstance(res, pd.DataFrame) or res.empty:
        st.error(res if isinstance(res, str) else "실적을 계산할 수 없습니다.")
        return

    my_res = res[res["담당자"] == st.session_state.user_name]
    uname = st.session_state.user_name

    # 업로드 전 예상치 계산 (추가 활동 제외)
    if not my_res.empty:
        before_res = my_res.copy()
        o_p = int(float(before_res.iloc[0].get("개설포인트", 0)))
        l_p = int(float(before_res.iloc[0].get("연계포인트", 0)))
        v_p = int(float(before_res.iloc[0].get("운영포인트 (실제 활동)", 0)))

        # 합계포인트: 개설 + 연계 + 운영(실제)만
        before_total = min(2800, min(1000, o_p + l_p) + min(1800, v_p))
        before_res.loc[before_res.index[0], "합계포인트"] = before_total

        # 지급포인트
        before_pay_point = max(0, before_total - 1000)
        before_res.loc[before_res.index[0], "지급포인트"] = before_pay_point

        # 지급예상금액
        rank = before_res.iloc[0].get("직급", "")
        name = before_res.iloc[0].get("담당자", "")
        name_to_info = {
            info.get("name"): {
                "staff_type": info.get("staff_type", "정규직"),
            }
            for uid, info in st.session_state.user_db.items()
            if uid != "1"
        }
        is_outsource = name_to_info.get(name, {}).get("staff_type", "정규직") == "외주"

        if is_outsource:
            before_pay = int(max(0, before_pay_point * 500))
        else:
            leader_bonus = 500 if rank == "팀장" else 0
            before_pay = int(max(0, (before_pay_point + leader_bonus) * 500))

        before_res.loc[before_res.index[0], "지급예상금액"] = before_pay

        # 전월대비 재계산 (필요시)
        if "전월대비" in before_res.columns:
            prev_df = st.session_state.get("auto_prev_df")
            if prev_df is not None:
                try:
                    prev_res, _, _ = process_performance_analysis(prev_df)
                    if isinstance(prev_res, pd.DataFrame) and not prev_res.empty:
                        p_map = prev_res.set_index("담당자")["지급예상금액"].to_dict()
                        prev_pay = p_map.get(name, 0)
                        before_res.loc[before_res.index[0], "전월대비"] = int(before_pay - prev_pay)
                except Exception:
                    pass

    if my_res.empty:
        return

    # 현재 실적 표시
    st.markdown("##### 실적 현황")
    display_res = my_res.drop(columns=["전송시각"], errors="ignore")
    style_report_logic(display_res, compact=True)

    # ── 중복방문/초과방문/누락 확인 탭 ──
    df_user_check = df_user_visit
    _, err_check, dup_check = process_performance_analysis(filter_visit_rows(original_df), st.session_state.get("auto_prev_df"))

    # 탭 데이터 미리 계산 (경고 메시지 표시용 - 재업로드한 경우만)
    # 중복 이력
    dup_my = pd.DataFrame()
    if dup_check is not None and not dup_check.empty:
        u_col_dup = find_col(dup_check, ["등록자", "담당자", "성명"], "담당자")
        if u_col_dup and u_col_dup in dup_check.columns:
            dup_my = dup_check[dup_check[u_col_dup] == st.session_state.user_name]
        else:
            dup_my = dup_check

    # 초과 방문
    err_filtered_final = pd.DataFrame()
    if err_check is not None and not err_check.empty:
        if "담당자" in err_check.columns:
            err_my_final = err_check[err_check["담당자"] == st.session_state.user_name].copy()
        else:
            err_my_final = err_check.copy()
        err_filtered_final = err_my_final[err_my_final["일방문"] >= 6].copy()

    # 개설완료일자 누락
    missing_open_final = pd.DataFrame()
    if "본사 개설완료일자" in df_user_check.columns:
        missing_open_final = df_user_check[
            pd.isna(df_user_check["본사 개설완료일자"]) | (df_user_check["본사 개설완료일자"].astype(str).str.strip() == "")
        ]
        if "본사 신규이행구분" in missing_open_final.columns:
            missing_open_final = missing_open_final[missing_open_final["본사 신규이행구분"].astype(str).str.strip() != "이행"]

    # ERP연계일자 누락
    missing_erp_final = pd.DataFrame()
    if "본사 ERP연계일자" in df_user_check.columns:
        d_col_check = find_col(df_user_check, ["활동상세", "활동내용"], "활동상세")
        if d_col_check and d_col_check in df_user_check.columns:
            target = df_user_check[df_user_check[d_col_check].astype(str).str.contains("연계", na=False)]
        else:
            target = df_user_check
        missing_erp_final = target[
            pd.isna(target["본사 ERP연계일자"]) | (target["본사 ERP연계일자"].astype(str).str.strip() == "")
        ]
        if "본사 신규이행구분" in missing_erp_final.columns and "본사 이행추가연계" in missing_erp_final.columns:
            _is_ihang = missing_erp_final["본사 신규이행구분"].astype(str).str.strip() == "이행"
            _add_empty = pd.isna(missing_erp_final["본사 이행추가연계"]) | (missing_erp_final["본사 이행추가연계"].astype(str).str.strip() == "")
            missing_erp_final = missing_erp_final[~(_is_ihang & _add_empty)]

    # 기타 오류
    other_errors_df_final = build_other_validation_errors(df_user_check)

    # 각 탭의 데이터 존재 여부 확인
    has_dup_data = not dup_my.empty
    has_err_data = not err_filtered_final.empty
    has_missing_open = not missing_open_final.empty
    has_missing_erp = not missing_erp_final.empty
    has_other_errors = not other_errors_df_final.empty

    # 검증 이슈 체크
    has_validation_issues = has_dup_data or has_err_data or has_missing_open or has_missing_erp or has_other_errors

    # 추가 이력 등록건수 불일치 체크
    total = st.session_state.get("_manual_perf_total", 0)
    _req_visit_cnt = round(total / 30) if total > 0 else 0
    _preview_df_check = st.session_state.get("history_convert_preview_data")
    if _req_visit_cnt > 0 and isinstance(_preview_df_check, pd.DataFrame) and not _preview_df_check.empty:
        _is_manual_mask = _preview_df_check.get("_is_manual", pd.Series([False] * len(_preview_df_check))).fillna(False).astype(bool)
        _is_visit_mask = _preview_df_check["활동구분"].astype(str).str.contains("방문", na=False) if "활동구분" in _preview_df_check.columns else pd.Series([False] * len(_preview_df_check))
        _is_oper_mask = _preview_df_check["활동상세"].astype(str).str.contains("운영", na=False) if "활동상세" in _preview_df_check.columns else pd.Series([False] * len(_preview_df_check))
        _actual_visit_cnt = int((_is_manual_mask & _is_visit_mask & _is_oper_mask).sum())
        _visit_count_mismatch = _actual_visit_cnt != _req_visit_cnt
    else:
        _actual_visit_cnt = 0
        _visit_count_mismatch = False

    # 전송 가능 여부
    can_send = not has_validation_issues and not _visit_count_mismatch

    if has_validation_issues:
        st.markdown(
            "<div style='margin-top:8px;padding:10px 16px;background:#FFF5F5;border:1px solid #FC8181;border-radius:8px;font-size:13px;color:#C53030;font-weight:700;'>"
            "❌ 검증 오류가 있습니다. 위 탭에서 문제를 해결한 후 실적을 전송해주세요.</div>",
            unsafe_allow_html=True,
        )
    if _visit_count_mismatch:
        st.markdown(
            f"<div style='margin-top:8px;padding:10px 16px;background:#FFFAF0;border:1px solid #F6AD55;border-radius:8px;font-size:13px;color:#C05621;font-weight:700;'>"
            f"⚠️ 추가 이력등록건수 불일치 — 안내 기준 <b>{_req_visit_cnt}건</b> 필요, 현재 등록 <b>{_actual_visit_cnt}건</b>.<br>"
            f"방문 운영 이력을 <b>{_req_visit_cnt - _actual_visit_cnt}건</b> 추가 등록해주세요.</div>",
            unsafe_allow_html=True,
        )
    if not has_validation_issues and not my_res.empty:
        개설건수 = int(my_res.iloc[0].get("개설건수", 0))
        연계건수 = int(my_res.iloc[0].get("연계건수", 0))
        운영건수_실제 = int(my_res.iloc[0].get("운영건수 (실제 활동)", 0))
        운영건수_추가 = int(my_res.iloc[0].get("운영건수 (추가 활동)", 0))
        운영건수 = min(60, 운영건수_실제 + 운영건수_추가)
        지급예상금액 = int(my_res.iloc[0].get("지급예상금액", 0))

        전월대비_text = ""
        전월대비_col = next((c for c in my_res.columns if "전월대비" in c), None)
        if 전월대비_col:
            전월대비 = int(my_res.iloc[0].get(전월대비_col, 0))
            user_sel = st.session_state.get("user_prev_month_sel", "선택안함")
            if user_sel and user_sel != "선택안함":
                prev_m = str(int(user_sel.split("-")[1])) + "월"
                비교월_label = f"{prev_m} 대비"
            else:
                비교월_label = "전월 대비"
            if 전월대비 > 0:
                전월대비_text = f" {비교월_label} <b>{전월대비:,}</b>원 증가하였습니다."
            elif 전월대비 < 0:
                전월대비_text = f" {비교월_label} <b>{abs(전월대비):,}</b>원 감소하였습니다."

        st.markdown(
            f"<div style='margin-top:8px;padding:10px 16px;background:#EBF8FF;border-radius:8px;font-size:13px;color:#2B6CB0;'>"
            f"<b>{html.escape(uname)}</b>님의 최종 실적은 개설 <b>{개설건수}</b>개, 연계 <b>{연계건수}</b>개, 운영 <b>{운영건수}</b>개 에 금액은 <b>{지급예상금액:,}</b>원 입니다."
            f"{전월대비_text}</div>",
            unsafe_allow_html=True,
        )
        if not _visit_count_mismatch:
            st.markdown(
                "<div style='margin-top:8px;padding:10px 16px;background:#F0FFF4;border:1px solid #9AE6B4;border-radius:8px;font-size:13px;color:#276749;font-weight:700;'>"
                "✅ 실적 결과를 전송할 수 있습니다.</div>",
                unsafe_allow_html=True,
            )

    st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)

    _, save_col, send_col = st.columns([0.8, 0.1, 0.1])

    with save_col:
        st.markdown('<div class="action-btn">', unsafe_allow_html=True)
        do_save = st.button("저장", use_container_width=True, type="secondary")
        st.markdown('</div>', unsafe_allow_html=True)

    with send_col:
        st.markdown('<div class="action-btn">', unsafe_allow_html=True)
        do_send = st.button("실적 결과 전송", use_container_width=True, type="primary", disabled=not can_send)
        st.markdown('</div>', unsafe_allow_html=True)

    if do_save:
        save_manual_perf_override_for_current_user()
        saved_db = load_db(SAVED_STATE_FILE, {})
        user_saved = {
            "user_excel_data": st.session_state.user_excel_data.to_dict() if st.session_state.user_excel_data is not None else None,
            "user_prev_month_sel": st.session_state.get("user_prev_month_sel", "선택안함"),
            "saved_at": (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S")
        }
        saved_db[st.session_state.user_name] = user_saved
        save_db(SAVED_STATE_FILE, saved_db)
        st.success("저장 완료")
        time.sleep(0.5)
        st.rerun()

    if do_send:
        save_manual_perf_override_for_current_user()
        sent_db = load_db(SENT_FILE, {})
        sent_uploads_db = load_db(SENT_UPLOADS_FILE, {})
        row_data = my_res.iloc[0].to_dict()
        row_data["전송시각"] = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S")

        ym = get_uploaded_month(original_df)
        if ym:
            row_data["이력작성월"] = ym

        user_sel = st.session_state.get("user_prev_month_sel", "선택안함")
        if "전월대비" in row_data and user_sel and user_sel != "선택안함":
            prev_m = str(int(user_sel.split("-")[1])) + "월"
            row_data[f"전월대비({prev_m})"] = row_data.pop("전월대비")

        sent_db[st.session_state.user_name] = row_data

        # 본인 이름에 맞는 이력만 저장
        u_col_send = find_col(original_df, ["등록자", "담당자", "성명"], "등록자")
        if u_col_send and u_col_send in original_df.columns:
            user_only_df = original_df[original_df[u_col_send] == st.session_state.user_name].copy()
            sent_uploads_db[st.session_state.user_name] = dataframe_to_upload_payload(user_only_df)
        else:
            sent_uploads_db[st.session_state.user_name] = dataframe_to_upload_payload(original_df)

        save_db(SENT_FILE, sent_db)
        save_db(SENT_UPLOADS_FILE, sent_uploads_db)

        st.success("전송 완료")
        time.sleep(3)
        st.rerun()


def sent_results_df():
    sent_db = load_db(SENT_FILE, {})
    if not sent_db:
        return pd.DataFrame()

    staff_names = get_staff_names()
    if staff_names:
        sent_db = {k: v for k, v in sent_db.items() if k in staff_names}
    if not sent_db:
        return pd.DataFrame()

    sent_df = pd.DataFrame(list(sent_db.values()))
    sent_df = sent_df.drop(columns=[c for c in sent_df.columns if "관리자전월대비" in c], errors="ignore")

    if "이력작성월" in sent_df.columns:
        idx = sent_df.columns.tolist().index(next((c for c in sent_df.columns if "전월대비" in c), sent_df.columns[-1]))
        sent_df.insert(idx, "등록월", sent_df["이력작성월"])
        sent_df = sent_df.drop(columns=["이력작성월"], errors="ignore")
    elif "전송시각" in sent_df.columns:
        idx = sent_df.columns.tolist().index(next((c for c in sent_df.columns if "전월대비" in c), sent_df.columns[-1]))
        sent_df.insert(idx, "등록월", sent_df["전송시각"].astype(str).str[:7])

    return sort_by_rank_name(apply_rank_from_user_db(sent_df))


def apply_admin_prev_diff(sent_df):
    if sent_df.empty:
        return sent_df

    result_df = sent_df.copy()
    result_df = result_df.drop(columns=[c for c in result_df.columns if "전월대비" in str(c)], errors="ignore")

    selected_month = st.session_state.get("adm_prev_month", "선택안함")
    prev_df = st.session_state.get("auto_prev_df")

    if selected_month == "선택안함" or prev_df is None or prev_df.empty:
        return result_df

    prev_res, _, _ = process_performance_analysis(prev_df)
    if not isinstance(prev_res, pd.DataFrame) or prev_res.empty:
        return result_df

    prev_pay_map = prev_res.set_index("담당자")["지급예상금액"].to_dict()

    if "담당자" in result_df.columns and "지급예상금액" in result_df.columns:
        diff_col = str(int(selected_month.split("-")[1])) + "월 대비"
        result_df[diff_col] = result_df.apply(
            lambda r: int(float(r.get("지급예상금액", 0))) - int(float(prev_pay_map.get(r.get("담당자"), 0))),
            axis=1,
        )

        cols = result_df.columns.tolist()
        cols.remove(diff_col)
        insert_at = cols.index("전송시각") if "전송시각" in cols else len(cols)
        cols.insert(insert_at, diff_col)
        result_df = result_df[cols]

    return result_df


def dataframe_to_excel_bytes(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = str(sheet_name)[:31] or "Sheet"
            df.to_excel(writer, index=False, sheet_name=safe_name)
    output.seek(0)
    return output.getvalue()


def sent_uploaded_files_excel_bytes():
    sent_uploads_db = load_db(SENT_UPLOADS_FILE, {})
    if not sent_uploads_db:
        return dataframe_to_excel_bytes({
            "안내": pd.DataFrame([{"내용": "아직 전송된 업로드 엑셀 파일이 없습니다. 직원이 실적 결과 전송을 다시 진행하면 이 파일에 포함됩니다."}])
        })

    staff_names = get_staff_names()
    if staff_names:
        sent_uploads_db = {k: v for k, v in sent_uploads_db.items() if k in staff_names}
    if not sent_uploads_db:
        return dataframe_to_excel_bytes({
            "안내": pd.DataFrame([{"내용": "직원 목록에 해당하는 전송 데이터가 없습니다."}])
        })

    analysis_result = st.session_state.get("analysis_result")
    if analysis_result is not None and not analysis_result.empty:
        if "담당자" in analysis_result.columns:
            report_names = set(analysis_result["담당자"].unique())
        else:
            report_names = set(sent_uploads_db.keys())
    else:
        report_names = set(sent_uploads_db.keys())

    all_data = []
    for name, payload in sent_uploads_db.items():
        if name in report_names:
            df = upload_payload_to_dataframe(payload)
            if not df.empty:
                all_data.append(df)

    if not all_data:
        return dataframe_to_excel_bytes({
            "안내": pd.DataFrame([{"내용": "저장된 업로드 파일 데이터가 비어 있습니다."}])
        })

    # 원본 업로드 양식과 동일하게 저장
    combined_df = pd.concat(all_data, ignore_index=True)
    return dataframe_to_excel_bytes({"전체 활동 이력": combined_df})


def report_month_info(report_df):
    ym = ""
    if "등록월" in report_df.columns:
        months = report_df["등록월"].dropna().astype(str)
        months = months[months.str.match(r"^\d{4}-\d{2}$", na=False)]
        if not months.empty:
            ym = months.value_counts().idxmax()
    if not ym:
        ym = get_uploaded_month(st.session_state.user_excel_data) if st.session_state.user_excel_data is not None else ""
    if not ym:
        ym = datetime.utcnow().strftime("%Y-%m")
    year = int(ym.split("-")[0])
    month = int(ym.split("-")[1])
    return ym, year, month


def as_date_series(series):
    return pd.to_datetime(series, errors="coerce")


def cloud_customer_counts(name=None):
    cloud = st.session_state.get("cloud_sheet_df")
    if cloud is None:
        try:
            load_csv_to_state("url_sync", "cloud_sheet_df")
            cloud = st.session_state.get("cloud_sheet_df")
        except Exception:
            cloud = None

    ym, year, _ = report_month_info(st.session_state.analysis_result if st.session_state.analysis_result is not None else pd.DataFrame())
    empty_counts = {
        "manage_total": 0, "manage_link": 0,
        "year_open": 0, "year_link": 0,
        "month_open": 0, "month_link": 0,
    }
    if cloud is None or cloud.empty:
        return empty_counts

    df = clean_header_logic(cloud.copy())
    owner_col = find_col(df, ["담당자", "등록자", "성명"])
    if name and owner_col and owner_col in df.columns:
        df = df[df[owner_col].astype(str).str.strip() == str(name).strip()].copy()

    status_col = find_col(df, ["상태항목", "상태"])
    open_col = find_col(df, ["개설완료일자", "개설일"])
    erp_col = find_col(df, ["ERP연계일자", "ERP연계", "연계일자"])

    if status_col and status_col in df.columns:
        status = df[status_col].astype(str).str.strip()
        manage_total = int(status.isin(["대기", "완료"]).sum())
        manage_link = int(((status == "완료") & as_date_series(df[erp_col]).notna()).sum()) if erp_col and erp_col in df.columns else 0
    else:
        manage_total = 0
        manage_link = int(as_date_series(df[erp_col]).notna().sum()) if erp_col and erp_col in df.columns else 0

    open_dates = as_date_series(df[open_col]) if open_col and open_col in df.columns else pd.Series(dtype="datetime64[ns]")
    erp_dates = as_date_series(df[erp_col]) if erp_col and erp_col in df.columns else pd.Series(dtype="datetime64[ns]")

    return {
        "manage_total": manage_total,
        "manage_link": manage_link,
        "year_open": int((open_dates.dt.year == year).sum()) if not open_dates.empty else 0,
        "year_link": int((erp_dates.dt.year == year).sum()) if not erp_dates.empty else 0,
        "month_open": int((open_dates.dt.strftime("%Y-%m") == ym).sum()) if not open_dates.empty else 0,
        "month_link": int((erp_dates.dt.strftime("%Y-%m") == ym).sum()) if not erp_dates.empty else 0,
    }


def sent_activity_counts(report_df, name=None):
    df = report_df.copy()
    if name and "담당자" in df.columns:
        df = df[df["담당자"].astype(str).str.strip() == str(name).strip()]

    def sum_col(col):
        return int(pd.to_numeric(df[col], errors="coerce").fillna(0).sum()) if col in df.columns else 0

    open_count = sum_col("개설건수")
    link_count = sum_col("연계건수")
    operation_col = "운영건수 (실제 활동)" if "운영건수 (실제 활동)" in df.columns else "운영건수"
    operation_count = sum_col(operation_col)
    total_count = open_count + link_count + operation_count

    unique_customers = uploaded_activity_customer_count(name)
    if unique_customers == 0:
        unique_customers = total_count

    return {
        "open": open_count,
        "link": link_count,
        "operation": operation_count,
        "total": total_count,
        "unique_customers": unique_customers,
    }


def uploaded_activity_customer_count(name=None):
    sent_uploads_db = load_db(SENT_UPLOADS_FILE, {})
    biz_keys = set()

    for saved_name, payload in sent_uploads_db.items():
        df = upload_payload_to_dataframe(payload)
        if df.empty:
            continue
        u_col = find_col(df, ["등록자", "담당자", "성명"])
        d_col = find_col(df, ["활동상세", "활동내용"])
        biz_col = find_col(df, ["사업자번호"])
        if name and str(saved_name).strip() != str(name).strip():
            if not u_col or u_col not in df.columns:
                continue
            df = df[df[u_col].astype(str).str.strip() == str(name).strip()].copy()
        if d_col and d_col in df.columns:
            df = df[df[d_col].astype(str).str.contains("개설|연계|운영|방문|점검", na=False)]
        if biz_col and biz_col in df.columns:
            biz_keys.update([v for v in normalize_biz(df[biz_col]).tolist() if v])

    return len(biz_keys)


def summarize_activity_note(value, max_lines=2, max_chars_per_line=34):
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text or text.lower() == "nan":
        return ""

    parts = [
        re.sub(r"\s+", " ", part).strip(" -ㆍ*•●○◦·.,;:/\\|")
        for part in re.split(r"(?:\n+|[.;。]|[.!?]\s+)", text)
        if part.strip(" -ㆍ*•●○◦·.,;:/\\|")
    ]
    if not parts:
        parts = [re.sub(r"^[^\w가-힣]+", "", text)]

    lines = []
    for part in parts:
        part = re.sub(r"^[^\w가-힣]+", "", part).strip()
        if not part:
            continue
        while len(part) > max_chars_per_line and len(lines) < max_lines:
            cut_at = part.rfind(" ", 0, max_chars_per_line + 1)
            if cut_at < max_chars_per_line // 2:
                cut_at = max_chars_per_line
            lines.append(part[:cut_at].strip())
            part = part[cut_at:].strip()
        if part and len(lines) < max_lines:
            lines.append(part[:max_chars_per_line].strip())
        if len(lines) >= max_lines:
            break

    if not lines:
        return ""
    lines[0] = f"● {lines[0]}"
    return "\n".join(lines[:max_lines])


def is_major_note_col(table, col_idx):
    try:
        header = table.cell(0, col_idx).text.replace(" ", "")
        return "비고" in header
    except Exception:
        return False


def uploaded_major_rows(name, keyword, row_type):
    payload = load_db(SENT_UPLOADS_FILE, {}).get(name)
    df = upload_payload_to_dataframe(payload)
    if df.empty:
        return []

    u_col = find_col(df, ["등록자", "담당자", "성명"])
    d_col = find_col(df, ["활동상세"]) or find_col(df, ["활동내용"])
    comp_col = find_col(df, ["업체명", "상호", "고객명"])
    product_col = find_col(df, ["상품"])
    system_col = find_col(df, ["내부시스템", "ERP", "시스템"])
    report_col = find_col(df, ["구축보고서"])
    note_col = find_col(df, ["활동내용"]) or find_col(df, ["비고", "제목"])

    if u_col and u_col in df.columns:
        df = df[df[u_col].astype(str).str.strip() == str(name).strip()]
    if d_col and d_col in df.columns:
        df = df[df[d_col].astype(str).str.contains(keyword, na=False)]

    rows = []
    for _, row in df.iterrows():
        rows.append([
            row.get(comp_col, "") if comp_col else "",
            row_type,
            row.get(product_col, "통합CMS") if product_col else "통합CMS",
            row.get(system_col, "") if system_col else "",
            row.get(report_col, "X") if report_col else "X",
            summarize_activity_note(row.get(note_col, "")) if note_col else "",
        ])
    return rows


def build_report_ppt_bytes(report_df, compare_df, curr_month_label, prev_month_label):
    from pptx import Presentation
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt
    from copy import deepcopy

    if not os.path.exists(PPT_TEMPLATE_FILE):
        raise FileNotFoundError(f"PPT 템플릿 파일을 찾을 수 없습니다: {PPT_TEMPLATE_FILE}")

    if report_df.empty:
        raise ValueError("실적 데이터가 비어있습니다. 마감 후 다시 시도해주세요.")

    prs = Presentation(PPT_TEMPLATE_FILE)
    ym, _, _ = report_month_info(report_df)

    def fmt(value):
        if value == "":
            return ""
        if value is None:
            return "-"
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        try:
            if isinstance(value, str) and value.strip() == "-":
                return "-"
            if isinstance(value, (int, float, np.integer, np.floating)):
                return f"{int(value):,}"
            if str(value).replace(",", "").replace("-", "").isdigit():
                return f"{int(str(value).replace(',', '')):,}"
        except Exception:
            pass
        return str(value)

    def set_cell_text(cell, value, align=None, font_size=14):
        cell.text = fmt(value)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = align if align is not None else PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.name = "맑은 고딕"
                run.font.size = Pt(font_size)

    def slide_tables(slide):
        return [shape.table for shape in slide.shapes if getattr(shape, "has_table", False)]

    def ensure_rows(table, required_rows):
        if len(table.rows) == 0 or len(table._tbl.tr_lst) == 0:
            return
        while len(table.rows) < required_rows:
            try:
                new_tr = deepcopy(table._tbl.tr_lst[-1])
                table._tbl.append(new_tr)
                for cell in table.rows[-1].cells:
                    cell.text = ""
            except (IndexError, AttributeError):
                break

    def fill_table_row(table, row_idx, values):
        ensure_rows(table, row_idx + 1)
        if row_idx >= len(table.rows):
            return
        for col_idx, value in enumerate(values):
            if col_idx < len(table.columns):
                set_cell_text(table.cell(row_idx, col_idx), value, PP_ALIGN.CENTER)

    def fill_stats_slide(slide, counts, cloud_counts):
        tables = slide_tables(slide)
        if len(tables) < 4:
            return

        fill_table_row(tables[0], 1, ["-", counts["open"], counts["link"], counts["operation"]])
        fill_table_row(tables[2], 1, [counts["unique_customers"], counts["total"], counts["total"], "-", "-", "-"])
        fill_table_row(tables[3], 2, [
            cloud_counts["manage_total"],
            cloud_counts["manage_link"],
            "-",
            cloud_counts["year_open"],
            cloud_counts["year_link"],
            "-",
            cloud_counts["month_open"],
            cloud_counts["month_link"],
        ])

    def fill_major_slide(slide, rows):
        tables = slide_tables(slide)
        if not tables:
            return
        table = tables[0]
        if len(table.rows) == 0 or len(table.columns) == 0:
            return

        ensure_rows(table, len(rows) + 1)

        # 기존 행 초기화 (헤더 제외)
        for r in range(1, len(table.rows)):
            for c in range(len(table.columns)):
                try:
                    set_cell_text(table.cell(r, c), "", font_size=10)
                except:
                    pass

        # 데이터 채우기
        for r_idx, row_values in enumerate(rows, start=1):
            if r_idx >= len(table.rows):
                break
            if not isinstance(row_values, (list, tuple)):
                continue
            for c_idx, value in enumerate(row_values[:len(table.columns)]):
                try:
                    align = PP_ALIGN.LEFT if is_major_note_col(table, c_idx) else PP_ALIGN.CENTER
                    set_cell_text(table.cell(r_idx, c_idx), value, align, font_size=10)
                except:
                    pass

    def delete_slide(index):
        slide_id_list = prs.slides._sldIdLst
        slide_id = slide_id_list[index]
        prs.part.drop_rel(slide_id.rId)
        del slide_id_list[index]

    # 전체 브랜치 서비스(BS성과)_운영: 4페이지
    try:
        if len(prs.slides) > 3:
            fill_stats_slide(prs.slides[3], sent_activity_counts(report_df), cloud_customer_counts())
    except Exception as e:
        raise Exception(f"전체 브랜치 슬라이드(4페이지) 처리 중 오류: {e}")

    sections = {
        "이성환": [5, 6, 7, 8],
        "임인지": [9, 10, 11, 12],
        "전준수": [13, 14, 15, 16],
        "이수현": [17, 18, 19, 20],
        "하성춘": [21, 22, 23, 24],
        "길민종": [25, 26, 27],
    }
    sent_names = set(report_df["담당자"].dropna().astype(str)) if "담당자" in report_df.columns else set()

    for name, idxs in sections.items():
        if name not in sent_names:
            continue
        stats_idx = idxs[1]
        open_idx = idxs[2] if len(idxs) > 2 else None
        link_idx = idxs[3] if len(idxs) > 3 else None
        try:
            if stats_idx < len(prs.slides):
                fill_stats_slide(prs.slides[stats_idx], sent_activity_counts(report_df, name), cloud_customer_counts(name))
            if open_idx is not None and open_idx < len(prs.slides):
                fill_major_slide(prs.slides[open_idx], uploaded_major_rows(name, "개설", "일반 개설"))
            if link_idx is not None and link_idx < len(prs.slides):
                fill_major_slide(prs.slides[link_idx], uploaded_major_rows(name, "연계", "추가 연계"))
        except Exception as e:
            raise Exception(f"{name} 슬라이드 처리 중 오류: {e}")

    # 전송된 담당자 외 개인성과 섹션은 제거
    remove_indices = []
    for name, idxs in sections.items():
        if name not in sent_names:
            remove_indices.extend(idxs)
    for idx in sorted(remove_indices, reverse=True):
        if idx < len(prs.slides):
            delete_slide(idx)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


def render_report_action_buttons(report_df, compare_df, curr_month_label, prev_month_label):
    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

    is_closed = bool(st.session_state.get("report_closed"))

    dc1, dc2 = st.columns([0.85, 0.15])
    with dc2:
        if not is_closed:
            if st.button("마감", use_container_width=True, type="primary"):
                close_time = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")
                st.session_state.report_closed = close_time
                saved_db = load_db(SAVED_STATE_FILE, {})
                saved_db["report_closed"] = {"time": close_time, "by": st.session_state.get("user_name", "")}
                save_db(SAVED_STATE_FILE, saved_db)
                st.success("실적 보고서 마감 완료")
                time.sleep(0.5)
                st.rerun()
        else:
            st.info(f"마감 완료: {st.session_state.report_closed}")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.button("실적파일 구글시트 전송", use_container_width=True, disabled=not is_closed)

    ym = get_uploaded_month(st.session_state.user_excel_data) if st.session_state.user_excel_data is not None else ""
    if ym:
        year_month = ym.replace("-", "")
    else:
        year_month = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m")

    download_time = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m%d_%H%M%S")
    file_name = f"LMB월간 활동실적__{year_month}_{download_time}.xlsx"

    with c2:
        if is_closed:
            excel_bytes = sent_uploaded_files_excel_bytes()
            st.download_button(
                "실적파일 엑셀 다운로드",
                data=excel_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("실적파일 엑셀 다운로드", use_container_width=True, disabled=True)

    with c3:
        _is_admin = st.session_state.get("user_role") == "관리자"
        if is_closed or _is_admin:
            try:
                ppt_bytes = build_report_ppt_bytes(report_df, compare_df, curr_month_label, prev_month_label)
                st.download_button(
                    "실적보고서 PPT 다운로드",
                    data=ppt_bytes,
                    file_name=f"LMB활동실적보고서_{year_month}_하나지사.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            except Exception as e:
                st.button("실적보고서 PPT 다운로드", use_container_width=True, disabled=True)
                st.caption(f"PPT 생성 준비 중 오류: {e}")
        else:
            st.button("실적보고서 PPT 다운로드", use_container_width=True, disabled=True)


def show_admin_performance():
    # 관리자 모드로 사용자 실적 확인 페이지 호출
    show_user_history(is_admin_mode=True)


def show_admin_analysis():
    select_prev_month("auto_prev_df", "adm_prev_month")

    st.markdown("### 직원 전송 실적 내역")
    sent_df = filter_by_staff(apply_admin_prev_diff(sent_results_df()))

    if sent_df.empty:
        st.info("아직 전송된 실적이 없습니다.")
        return

    style_report_logic(sent_df)

    c1, c2, c3 = st.columns([0.8, 0.1, 0.1])

    with c1:
        if st.button("전송 내역 초기화"):
            save_db(SENT_FILE, {})
            save_db(SENT_UPLOADS_FILE, {})
            st.success("초기화 완료")
            time.sleep(0.5)
            st.rerun()

    with c2:
        st.markdown('<div class="action-btn">', unsafe_allow_html=True)
        if st.button("저장", use_container_width=True, type="secondary"):
            admin_saved = {
                "sent_df": sent_df.to_dict(),
                "sent_uploads_db": load_db(SENT_UPLOADS_FILE, {}),
                "adm_prev_month": st.session_state.get("adm_prev_month", "선택안함"),
                "saved_at": (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S")
            }
            saved_db = load_db(SAVED_STATE_FILE, {})
            saved_db["admin_analysis"] = admin_saved
            save_db(SAVED_STATE_FILE, saved_db)
            st.success("저장 완료")
            time.sleep(0.5)
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with c3:
        st.markdown('<div class="action-btn">', unsafe_allow_html=True)
        if st.button("마감", use_container_width=True, type="primary"):
            deadline_time = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")
            st.session_state.deadline_time = deadline_time
            st.session_state.analysis_result = sent_df.copy()
            saved_db = load_db(SAVED_STATE_FILE, {})
            saved_db["deadline"] = {"time": deadline_time, "by": st.session_state.get("user_name", "")}
            save_db(SAVED_STATE_FILE, saved_db)
            send_kakao_notify(f"[LMB 실적관리] 실적 분석/계산 마감 완료\n마감시각: {deadline_time}\n마감자: {st.session_state.get('user_name', '')}")
            st.success("마감 처리 완료")
            time.sleep(0.5)
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.deadline_time:
        st.info(f"마감 완료: {st.session_state.deadline_time}")


def show_report():
    if st.session_state.analysis_result is None:
        st.info("실적 분석/계산 메뉴에서 마감 후 확인 가능합니다.")
        return

    drop_cols = ["전송시각"] + [c for c in st.session_state.analysis_result.columns if "관리자전월대비" in c]
    report_df = filter_by_staff(st.session_state.analysis_result.drop(columns=drop_cols, errors="ignore"))

    curr_month_label = "당월"
    prev_month_label = "전월"

    ym = get_uploaded_month(st.session_state.user_excel_data) if st.session_state.user_excel_data is not None else ""
    if ym:
        curr_month_label = str(int(ym.split("-")[1])) + "월"

    adm_sel = st.session_state.get("adm_prev_month", "선택안함")
    if adm_sel and adm_sel != "선택안함":
        prev_month_label = str(int(adm_sel.split("-")[1])) + "월"

    st.markdown(f"### {curr_month_label} 실적보고서")
    style_report_logic(report_df)

    prev_df = st.session_state.get("auto_prev_df")

    if prev_df is None or prev_df.empty:
        st.info("실적 분석/계산 메뉴에서 비교할 전월을 선택하면 비교 리포트가 표시됩니다.")
        return

    prev_res, _, _ = process_performance_analysis(prev_df)

    if not isinstance(prev_res, pd.DataFrame) or prev_res.empty:
        st.info("전월 데이터를 분석할 수 없습니다.")
        return

    prev_res = filter_by_staff(prev_res)

    st.markdown("### 비교 리포트")

    compare_cols = [
        "개설건수",
        "개설포인트",
        "연계건수",
        "연계포인트",
        "운영건수 (실제 활동)",
        "운영포인트 (실제 활동)",
        "운영건수 (추가 활동)",
        "운영포인트(추가 활동)",
        "합계포인트",
        "지급포인트",
        "지급예상금액",
    ]

    shared_cols = [c for c in compare_cols if c in report_df.columns and c in prev_res.columns]
    compare_all_rows = []

    names = report_df["담당자"].tolist()
    for start in range(0, len(names), 2):
        person_cols = st.columns(2)
        for person_col, name in zip(person_cols, names[start:start + 2]):
            curr_row = report_df[report_df["담당자"] == name]
            prev_row = prev_res[prev_res["담당자"] == name]

            if curr_row.empty:
                continue

            rank = curr_row.iloc[0].get("직급", "")
            rows = []
            for col in shared_cols:
                c_val = int(float(curr_row.iloc[0][col]))
                p_val = int(float(prev_row.iloc[0][col])) if not prev_row.empty else 0
                rows.append({"항목": col, prev_month_label: p_val, curr_month_label: c_val, "증감": c_val - p_val})
                compare_all_rows.append({
                    "담당자": name,
                    "직급": rank,
                    "항목": col,
                    prev_month_label: p_val,
                    curr_month_label: c_val,
                    "증감": c_val - p_val,
                })

            with person_col:
                st.markdown(f"#### {name} ({rank})")
                style_report_logic(pd.DataFrame(rows), compact=True)

    render_report_action_buttons(report_df, pd.DataFrame(compare_all_rows), curr_month_label, prev_month_label)


def show_staff_admin():
    st.markdown("### 직원 목록")

    # 파일 기준으로 재로드 (_deleted 목록 포함 반영)
    st.session_state.user_db = load_user_db()

    if st.session_state.pop("reset_staff_edit_sel", False):
        st.session_state.staff_edit_sel = "선택안함"

    # ── 직원 직접 추가 ──
    with st.expander("직원 직접 추가", expanded=False):
        with st.form("add_staff_form"):
            ac1, ac2 = st.columns(2)
            with ac1:
                new_uid = st.text_input("아이디", key="add_staff_uid")
                new_name = st.text_input("성명", key="add_staff_name")
                new_email = st.text_input("메일주소", key="add_staff_email")
            with ac2:
                new_pw = st.text_input("비밀번호", type="password", key="add_staff_pw")
                new_access = st.selectbox("로그인 허용 여부", ["허용", "불가"], key="add_staff_access")
                new_role = st.selectbox("메뉴 접근 권한", ["사용자 메뉴", "관리자 메뉴"], key="add_staff_role")
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
                    "access": new_access,
                    "role": "관리자" if new_role == "관리자 메뉴" else "사용자",
                    "dept_type": "사업부",
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
                st.success(f"직원 '{name_str}({uid_str})'을(를) 추가했습니다.")
                st.rerun()

    staff_rows = []
    for uid, info in st.session_state.user_db.items():
        if uid == "1":
            continue
        staff_rows.append({
            "ID": uid,
            "성명": info.get("name", ""),
            "직급": info.get("rank", "직원"),
            "메일주소": info.get("email", ""),
            "부서구분": info.get("dept_type", "사업부"),
            "직원구분": info.get("staff_type", "정규직"),
            "외주여부": info.get("outsource", "아니오"),
            "외주 근무기간": info.get("outsource_period", "해당없음"),
            "로그인 허용 여부": info.get("access", "불가"),
            "메뉴 접근 권한": "관리자 메뉴" if info.get("role") == "관리자" else "사용자 메뉴",
        })

    if not staff_rows:
        st.info("등록된 직원이 없습니다.")
        return

    # 직급 순서로 정렬
    rank_order = {"부서장": 0, "팀장": 1, "과장": 2, "대리": 3, "주임": 4, "직원": 5}
    staff_rows.sort(key=lambda x: rank_order.get(x["직급"], 99))

    # ── 직원 목록 HTML 테이블 표시 (다크모드 호환, 가운데 정렬) ──
    render_plain_html_table(pd.DataFrame(staff_rows), center_align=True)

    st.markdown("---")
    st.markdown("#### 직원 정보 수정")

    uid_options = [f"{r['ID']} — {r['성명']}" for r in staff_rows]
    sel_col, _ = st.columns([0.32, 0.68])
    with sel_col:
        sel = st.selectbox("수정할 직원 선택", ["선택안함"] + uid_options, key="staff_edit_sel")

    if sel == "선택안함":
        return

    sel_uid = sel.split(" — ")[0].strip()
    info = st.session_state.user_db.get(sel_uid, {})

    st.markdown(f"**메일주소:** {info.get('email', '—')}")

    c1, c2, c3, c4, c5, c6, c7 = st.columns([1, 1, 1, 1, 1.15, 1.15, 1.2])
    with c1:
        new_rank = st.selectbox("직급", ["부서장", "팀장", "과장", "대리", "주임", "직원"],
                                index=["부서장", "팀장", "과장", "대리", "주임", "직원"].index(info.get("rank", "직원")),
                                key="edit_rank")
    with c2:
        dept_opts = ["사업부", "C&S"]
        new_dept_type = st.selectbox("부서구분", dept_opts,
                                     index=dept_opts.index(info.get("dept_type", "사업부")) if info.get("dept_type", "사업부") in dept_opts else 0,
                                     key="edit_dept_type")
    with c3:
        new_staff_type = st.selectbox("직원구분", ["정규직", "계약직", "파견직", "외주"],
                                      index=["정규직", "계약직", "파견직", "외주"].index(info.get("staff_type", "정규직")),
                                      key="edit_staff_type")
    with c4:
        new_outsource = st.selectbox("외주여부", ["아니오", "예"],
                                     index=["아니오", "예"].index(info.get("outsource", "아니오")),
                                     key="edit_outsource")
    with c5:
        period_opts = ["해당없음", "1년 미만", "1년 이상", "2년 이상"]
        new_period = st.selectbox("외주 근무기간", period_opts,
                                  index=period_opts.index(info.get("outsource_period", "해당없음")),
                                  key="edit_period")
    with c6:
        new_access = st.selectbox("로그인 허용 여부", ["허용", "불가"],
                                  index=["허용", "불가"].index(info.get("access", "불가")),
                                  key="edit_access")
    with c7:
        new_role = st.selectbox("메뉴 접근 권한", ["사용자 메뉴", "관리자 메뉴"],
                                index=0 if info.get("role") != "관리자" else 1,
                                key="edit_role")

    bc1, bc2, _ = st.columns([0.15, 0.15, 0.7])
    with bc1:
        if st.button("저장", type="primary", use_container_width=True):
            st.session_state.user_db[sel_uid]["rank"] = new_rank
            st.session_state.user_db[sel_uid]["dept_type"] = new_dept_type
            st.session_state.user_db[sel_uid]["staff_type"] = new_staff_type
            st.session_state.user_db[sel_uid]["outsource"] = new_outsource
            st.session_state.user_db[sel_uid]["outsource_period"] = new_period
            st.session_state.user_db[sel_uid]["access"] = new_access
            st.session_state.user_db[sel_uid]["role"] = "관리자" if new_role == "관리자 메뉴" else "사용자"
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
            st.success("저장 완료")
            time.sleep(0.5)
            st.rerun()
    with bc2:
        if st.button("삭제", type="secondary", use_container_width=True):
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
            st.success(f"{sel} 삭제 완료")
            time.sleep(0.5)
            st.rerun()


KPI_MONTHLY_TARGETS = {
    "개설": 20,
    "연계": 12,
    "운영 순증": 13,
    "유통 활동": 40,
}


def _kpi_number(value):
    if is_blank_value(value):
        return 0
    parsed = pd.to_numeric(str(value).replace(",", "").strip(), errors="coerce")
    return int(parsed) if pd.notna(parsed) else 0


def build_kpi_activity_recommendations(hana_sheet, billing_sheet, user_name=None, use_dart=False):
    if hana_sheet is None or billing_sheet is None or hana_sheet.empty or billing_sheet.empty:
        return pd.DataFrame()

    hana = hana_sheet.copy()
    billing = billing_sheet.copy()
    hana.columns = [str(c).strip() for c in hana.columns]
    billing.columns = [str(c).strip() for c in billing.columns]

    hana_customer_col = find_col(hana, ["고객번호"])
    hana_owner_col = find_col(hana, ["담당자"])
    hana_company_col = find_col(hana, ["고객명", "업체명", "상호"])
    hana_biz_col = find_col(hana, ["사업자번호"])
    hana_build_type_col = find_col(hana, ["구축형"])
    hana_open_status_col = find_col(hana, ["개설상태"])
    hana_link_status_col = find_col(hana, ["연계상태", "ERP연계상태"])
    hana_manage_col = find_col(hana, ["관리구분"])
    hana_open_date_col = find_col(hana, ["개설/이행일", "개설일", "이행일"])

    billing_customer_col = find_col(billing, ["고객번호"])
    billing_company_col = find_col(billing, ["고객명", "업체명", "상호"])
    billing_biz_col = find_col(billing, ["사업자번호", "사업자등록번호"])
    billing_new_col = find_col(billing, ["신규일자", "최초신규일자"])
    billing_login_col = find_col(billing, ["최종로그인일자", "최종로그인", "로그인일자"])
    billing_transfer_col = find_col(billing, ["최종이체일자", "최종이체", "이체일자"])
    billing_end_col = find_col(billing, ["해지일자", "해지일", "해약일"])
    billing_login_count_col = find_col(billing, ["로그인건수", "로그인횟수"])
    billing_menu_col = exact_col(billing, ["메뉴사용"]) or find_col(billing, ["메뉴사용", "메뉴클릭수"])
    billing_service_detail_col = find_col(billing, ["서비스상세", "서비스 구분"])

    if not hana_customer_col or not hana_owner_col or not billing_customer_col:
        return pd.DataFrame()

    hana["_고객번호"] = hana[hana_customer_col].apply(normalize_billing_customer_no)
    hana = hana[hana["_고객번호"].ne("")]
    if user_name:
        hana = hana[hana[hana_owner_col].astype(str).str.strip() == str(user_name).strip()].copy()
    if hana_manage_col and hana_manage_col in hana.columns:
        hana = hana[~hana[hana_manage_col].astype(str).str.strip().str.contains("해지|취소", case=False, na=False)].copy()
    if hana.empty:
        return pd.DataFrame()

    base = pd.DataFrame()
    base["_고객번호"] = hana["_고객번호"]
    base["담당자"] = hana[hana_owner_col].fillna("").astype(str)
    base["고객명"] = hana[hana_company_col].fillna("").astype(str) if hana_company_col else ""
    base["사업자번호"] = hana[hana_biz_col].apply(normalize_biz_no) if hana_biz_col else ""
    base["구축형"] = hana[hana_build_type_col].fillna("").astype(str) if hana_build_type_col else ""
    base["개설상태"] = hana[hana_open_status_col].fillna("").astype(str) if hana_open_status_col else ""
    base["연계상태"] = hana[hana_link_status_col].fillna("").astype(str) if hana_link_status_col else ""
    base["개설/이행일"] = hana[hana_open_date_col].apply(parse_sheet_date) if hana_open_date_col else pd.NaT
    base = base.drop_duplicates("_고객번호", keep="first")

    billing["_고객번호"] = billing[billing_customer_col].apply(normalize_billing_customer_no)
    billing = billing[billing["_고객번호"].ne("")].copy()
    lookup = pd.DataFrame()
    lookup["_고객번호"] = billing["_고객번호"]
    lookup["청구고객명"] = billing[billing_company_col].fillna("").astype(str) if billing_company_col else ""
    lookup["청구사업자번호"] = billing[billing_biz_col].apply(normalize_biz_no) if billing_biz_col else ""
    lookup["신규일자"] = billing[billing_new_col].apply(parse_sheet_date) if billing_new_col else pd.NaT
    lookup["최종로그인일자"] = billing[billing_login_col].apply(parse_sheet_date) if billing_login_col else pd.NaT
    lookup["최종이체일자"] = billing[billing_transfer_col].apply(parse_sheet_date) if billing_transfer_col else pd.NaT
    lookup["해지일자"] = billing[billing_end_col].apply(parse_sheet_date) if billing_end_col else pd.NaT
    lookup["로그인건수"] = billing[billing_login_count_col].apply(_kpi_number) if billing_login_count_col else 0
    lookup["메뉴사용"] = billing[billing_menu_col].apply(_kpi_number) if billing_menu_col else 0
    lookup["서비스상세"] = billing[billing_service_detail_col].fillna("").astype(str) if billing_service_detail_col else ""
    lookup = lookup.drop_duplicates("_고객번호", keep="first")

    merged = base.merge(lookup, on="_고객번호", how="inner")
    merged = merged[merged["해지일자"].isna()].copy()
    if merged.empty:
        return pd.DataFrame()

    today = pd.Timestamp(datetime.utcnow() + timedelta(hours=9)).normalize()

    # DART 공시 기업코드 맵 로드 (use_dart=True 이고 API 키 있을 때만)
    dart_key = _get_dart_api_key() if use_dart else ""
    dart_corp_map = _load_dart_corp_map(dart_key) if dart_key else {}

    rows = []
    for _, row in merged.iterrows():
        score = 0
        areas = []
        reasons = []
        guide = []
        action = "해피콜"

        last_login = row.get("최종로그인일자")
        last_transfer = row.get("최종이체일자")
        new_date = row.get("신규일자")
        login_count = int(row.get("로그인건수", 0) or 0)
        menu_count = int(row.get("메뉴사용", 0) or 0)
        build_type = str(row.get("구축형", ""))
        link_status = str(row.get("연계상태", ""))
        service_detail = str(row.get("서비스상세", ""))

        days_from_login = None
        if pd.isna(last_login):
            score += 55
            areas.append("운영/MAU")
            reasons.append("청구 시트상 로그인 이력 없음")
            guide.append("담당자 변경, 사용 의사, 로그인 장애 여부 확인 후 매뉴얼/방문교육 제안")
        else:
            days_from_login = int((today - pd.Timestamp(last_login).normalize()).days)
            if days_from_login >= 90:
                score += 45
                areas.append("해지방어")
                reasons.append(f"최종 로그인 {days_from_login}일 경과")
                guide.append("미사용 사유를 유형화하고 해지 가능성 또는 잠재 활성화 여부를 비고에 기록")
            elif days_from_login >= 30:
                score += 25
                areas.append("MAU")
                reasons.append(f"최종 로그인 {days_from_login}일 경과")
                guide.append("최근 업무 일정 확인 후 이번 달 1회 사용 목표로 안내")

        if pd.isna(last_transfer):
            if login_count >= 100:
                score += 45
                areas.append("이체 활성화")
                reasons.append(f"로그인 {login_count}회 이상이나 이체 이력 없음")
                guide.append("자금이체/지급거래 흐름을 확인하고 지급거래 교육 또는 RM 연계 제안")
                action = "방문/원격교육"
            elif login_count > 0:
                score += 30
                areas.append("이체 활성화")
                reasons.append("로그인 이력은 있으나 이체 이력 없음")
                guide.append("조회만 사용하는 사유 확인 후 첫 이체까지 동행 지원")
                action = "해피콜+교육"

        if pd.notna(new_date) and (today - pd.Timestamp(new_date).normalize()).days <= 60:
            score += 25
            areas.append("초기정착")
            reasons.append("최근 2개월 내 신규 고객")
            guide.append("초기 2개월 집중관리 대상으로 방문 또는 사용자 교육 일정 확정")

        open_status = str(row.get("개설상태", ""))
        # 구축형이 연계형 + 개설상태가 이행완료인 경우 이미 연계 완료 → 연계전환 추천 불가
        already_linked = ("연계" in build_type) and ("이행완료" in open_status)
        is_not_linked = (not already_linked) and (("연계" not in build_type) or ("완료" not in link_status))
        if is_not_linked and (login_count >= 30 or menu_count >= 50 or "단독" in service_detail):
            score += 25
            areas.append("연계 전환")
            reasons.append("사용 흔적이 있어 ERP연계 전환 제안 가능")
            guide.append("사용 ERP, 고도화 여부, 연계 항목, 접속/DB 정보 사전 확인")
            action = "연계 니즈확인"

        if login_count >= 100 or menu_count >= 100:
            score += 20
            areas.append("유통")
            reasons.append(f"사용량 높음(로그인 {login_count}, 메뉴 {menu_count})")
            guide.append("대시보드/IHB/We-Hub/이음택스 중 재무업무 니즈 확인")

        # DART 공시 기반 유통활동 가산점
        if dart_corp_map:
            company_name = str(row.get("고객명", "") or row.get("청구고객명", "")).strip()
            dart_bonus, dart_reasons = _dart_enrich(dart_key, dart_corp_map, company_name)
            if dart_bonus > 0:
                score += dart_bonus
                if "유통" not in areas:
                    areas.append("유통")
                reasons.extend(dart_reasons)
                guide.append("DART 공시 내용 확인 후 재무 니즈에 맞춰 서비스 제안")

        if score < 35:
            continue

        priority = "상" if score >= 85 else "중" if score >= 60 else "하"
        due = "이번 주" if priority == "상" else "이번 달 2주 내" if priority == "중" else "이번 달 내"
        customer_name = str(row.get("고객명", "")).strip() or str(row.get("청구고객명", "")).strip()
        biz_no = str(row.get("사업자번호", "")).strip() or str(row.get("청구사업자번호", "")).strip()

        rows.append({
            "우선순위": priority,
            "추천점수": score,
            "KPI영역": " / ".join(dict.fromkeys(areas)),
            "담당자": str(row.get("담당자", "")).strip(),
            "고객명": customer_name,
            "사업자번호": biz_no,
            "구축형": build_type,
            "개설상태": row.get("개설상태", ""),
            "연계상태": link_status,
            "최종로그인일자": last_login.strftime("%Y-%m-%d") if pd.notna(last_login) else "없음",
            "최종이체일자": last_transfer.strftime("%Y-%m-%d") if pd.notna(last_transfer) else "없음",
            "로그인건수": login_count,
            "메뉴사용": menu_count,
            "추천활동": action,
            "활동사유": " / ".join(dict.fromkeys(reasons)),
            "활동가이드": " / ".join(dict.fromkeys(guide)),
            "권장기한": due,
        })

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(["추천점수", "로그인건수", "메뉴사용"], ascending=[False, False, False]).reset_index(drop=True)


def render_kpi_activity_recommendations(hana_sheet, billing_sheet, user_name=None, key_prefix="kpi_rec"):
    st.markdown("#### KPI 집중 활동 추천")
    remaining_months = 6
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("개설 잔여 목표", f"{KPI_MONTHLY_TARGETS['개설'] * remaining_months:,}건", "월 20건")
    c2.metric("연계 잔여 목표", f"{KPI_MONTHLY_TARGETS['연계'] * remaining_months:,}건", "월 12건")
    c3.metric("운영 순증 잔여 목표", f"{KPI_MONTHLY_TARGETS['운영 순증'] * remaining_months:,}개", "월 13개")
    c4.metric("유통 활동 잔여 목표", f"{KPI_MONTHLY_TARGETS['유통 활동'] * remaining_months:,}건", "월 40건")

    st.caption("2026년 KPI 계획 기준: 개설 20건/월, 연계 12건/월, 운영 순증 13개/월, 유통 활동 40건/월. 잔여 목표는 남은 6개월 기준입니다.")

    # DART 공시 연동 토글 (DART_API_KEY 설정된 경우에만 표시)
    _dart_available = bool(_get_dart_api_key())
    use_dart = False
    if _dart_available:
        use_dart = st.checkbox(
            "📰 DART 전자공시 기반 유통활동 추천 보정 (공시 있는 고객 점수 상향)",
            key=f"{key_prefix}_dart",
            value=False,
        )
        if use_dart:
            st.caption(
                "체크 시 고객명으로 DART 전자공시 시스템을 조회해 최근 6개월 내 아래 공시가 있는 고객의 "
                "유통활동 추천 점수를 자동으로 높입니다. 공시 내용은 활동사유 컬럼에 `[DART공시]` 형태로 표시됩니다.\n\n"
                "- **유상증자**: 회사가 새 주식을 발행해 외부에서 자금을 조달하는 것 → 현금 유입으로 이체·지급 거래 활성화 기대\n"
                "- **합병**: 두 회사가 하나로 합쳐지는 것 → 계좌 통합·정리 니즈 발생, 접점 확대 기회\n"
                "- **수주**: 기업이 대규모 계약·프로젝트를 따내는 것 → 매출 증가 예상, 자금 흐름 활성화\n"
                "- **신규사업**: 기존 사업 외 새로운 사업에 진출하는 것 → 향후 거래 규모 확대 가능성"
            )
        else:
            st.caption("체크하면 DART 전자공시(유상증자·합병·수주·신규사업 등) 기반으로 유통활동 대상 고객의 추천 점수를 자동 보정합니다.")
    else:
        st.caption("💡 DART 전자공시 연동을 활성화하려면 [설정 → 구글 스프레드시트 연동] 메뉴에서 DART API 키를 입력하세요.")

    rec_df = build_kpi_activity_recommendations(hana_sheet, billing_sheet, user_name, use_dart=use_dart)
    if rec_df.empty:
        st.info("청구 시트와 하나은행 시트 기준으로 추천 가능한 KPI 활동 고객이 없습니다.")
        return

    f1, f2, f3, f4 = st.columns([2, 2, 2, 4])
    with f1:
        area_options = ["전체"] + sorted({area for text in rec_df["KPI영역"].dropna() for area in str(text).split(" / ") if area})
        selected_area = st.selectbox("KPI영역", area_options, key=f"{key_prefix}_area")
    with f2:
        selected_priority = st.selectbox("우선순위", ["전체", "상", "중", "하"], key=f"{key_prefix}_priority")
    with f3:
        owner_options = ["전체"] + sorted(v for v in rec_df["담당자"].dropna().astype(str).str.strip().unique().tolist() if v)
        selected_owner = st.selectbox("담당자", owner_options, key=f"{key_prefix}_owner")

    filtered = rec_df.copy()
    if selected_area != "전체":
        filtered = filtered[filtered["KPI영역"].astype(str).str.contains(selected_area, regex=False, na=False)]
    if selected_priority != "전체":
        filtered = filtered[filtered["우선순위"] == selected_priority]
    if selected_owner != "전체":
        filtered = filtered[filtered["담당자"] == selected_owner]
    filtered = filtered.head(100).reset_index(drop=True)
    filtered.insert(0, "순번", range(1, len(filtered) + 1))

    st.metric("추천 고객", f"{len(filtered):,}건")
    render_plain_html_table(filtered, max_rows=100, center_align=False)
    today_str = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m%d")
    st.download_button(
        "KPI 추천 고객 다운로드",
        data=filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
        file_name=f"KPI추천고객_{selected_area}_{selected_priority}_{today_str}.csv",
        mime="text/csv",
        use_container_width=True,
    )


def show_target_customers():
    st.markdown("### 이번달 활동 대상고객 추천")

    # 하나은행 구글시트 로드 (개설 미완료용)
    hana_sheet = None
    try:
        hana_raw = read_google_csv(st.session_state.get("url_hana", DEFAULT_URL_HANA), header=2)
        hana_sheet = hana_raw
    except Exception as e:
        st.error(f"하나은행 구글시트를 불러오지 못했습니다: {e}")

    # 하나은행 청구 시트 로드 (운영 관리 대상용)
    hana_billing_sheet = None
    try:
        billing_raw = read_google_csv(st.session_state.get("url_hana_billing", DEFAULT_URL_HANA_BILLING))
        billing_raw = billing_raw.dropna(how="all").reset_index(drop=True)
        hana_billing_sheet = billing_raw
    except Exception as e:
        st.error(f"하나은행 청구 시트를 불러오지 못했습니다: {e}")

    # 본사 구글시트 로드 (ERP연계 미완료용)
    if st.session_state.get("cloud_sheet_df") is None:
        try:
            load_csv_to_state("url_sync", "cloud_sheet_df")
        except Exception as e:
            st.error(f"본사 구글시트를 불러오지 못했습니다: {e}")

    cloud = st.session_state.get("cloud_sheet_df")

    # 활동 이력 시트 로드 (최근 활동 정보용)
    if st.session_state.get("analysis_lookup_df") is None:
        try:
            load_csv_to_state("url_analysis", "analysis_lookup_df")
        except Exception:
            pass

    hana_df = st.session_state.get("analysis_lookup_df")

    user_name = st.session_state.user_name

    def _only_current_user_rows(result_df):
        if result_df is None or result_df.empty or "담당자" not in result_df.columns:
            return result_df
        return result_df[result_df["담당자"].astype(str).str.strip() == str(user_name).strip()].reset_index(drop=True)

    render_kpi_activity_recommendations(hana_sheet, hana_billing_sheet, user_name, key_prefix="target_kpi_rec")
    st.divider()

    cloud = clean_header_logic(cloud.copy())

    # 본사 시트 컬럼 탐색
    owner_col  = find_col(cloud, ["담당자", "등록자", "성명"])
    comp_col   = find_col(cloud, ["업체명", "고객명", "상호"])
    biz_col    = find_col(cloud, ["사업자번호"])
    open_col   = find_col(cloud, ["개설완료일자", "개설일"])
    erp_col    = find_col(cloud, ["ERP연계일자", "연계일자"])
    div_col    = find_col(cloud, ["신규/이행구분", "이행구분", "신규이행"])
    add_col    = find_col(cloud, ["이행추가연계"])
    status_col = find_col(cloud, ["상태항목", "상태"])
    open_cancel_col = find_col(cloud, ["개설취소"])
    erp_cancel_col = find_col(cloud, ["ERP연계취소", "연계취소"])
    manage_div_col = find_col(cloud, ["관리구분", "관리 구분"])
    open_status_col = find_col(cloud, ["개설상태"])
    erp_status_col = find_col(cloud, ["연계상태", "ERP연계상태"])

    if not owner_col:
        st.warning("구글시트에 담당자 컬럼을 찾을 수 없습니다.")
        return

    # 본인 담당 고객만 필터
    df = cloud[cloud[owner_col].astype(str).str.strip() == str(user_name).strip()].copy()

    # 관리구분이 "해지" 또는 "취소"를 포함하는 고객 제외
    if manage_div_col and manage_div_col in df.columns:
        manage_values = df[manage_div_col].astype(str).str.strip()
        # "해지" 또는 "취소"가 포함된 경우 모두 제외
        is_terminated = manage_values.str.contains("해지|취소", case=False, na=False)
        df = df[~is_terminated].copy()

    if df.empty:
        st.info(f"구글시트에 {user_name}님 담당 고객이 없습니다.")
        return

    # 날짜 정규화
    for col in [open_col, erp_col]:
        if col and col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # 하나은행 시트에서 활동 이력 가져오기
    def get_recent_activity(biz_no):
        if hana_df is None or hana_df.empty or not biz_no:
            return {"last_date": "", "last_detail": "", "activity_count": 0}

        hana_clean = clean_header_logic(hana_df.copy())
        hana_biz_col = find_col(hana_clean, ["사업자번호"])
        hana_date_col = find_col(hana_clean, ["활동일", "일자"])
        hana_detail_col = find_col(hana_clean, ["활동상세", "활동내용"])

        if not hana_biz_col or hana_biz_col not in hana_clean.columns:
            return {"last_date": "", "last_detail": "", "activity_count": 0}

        # 사업자번호 정규화하여 매칭
        biz_normalized = normalize_biz(pd.Series([biz_no])).iloc[0]
        hana_clean[hana_biz_col] = normalize_biz(hana_clean[hana_biz_col])

        matched = hana_clean[hana_clean[hana_biz_col] == biz_normalized].copy()

        if matched.empty:
            return {"last_date": "", "last_detail": "", "activity_count": 0}

        # 날짜 정렬
        if hana_date_col and hana_date_col in matched.columns:
            matched[hana_date_col] = pd.to_datetime(matched[hana_date_col], errors="coerce")
            matched = matched.sort_values(hana_date_col, ascending=False)
            last_date = matched.iloc[0][hana_date_col]
            last_date_str = last_date.strftime("%Y-%m-%d") if pd.notna(last_date) else ""
        else:
            last_date_str = ""

        last_detail = matched.iloc[0].get(hana_detail_col, "") if hana_detail_col else ""
        activity_count = len(matched)

        return {
            "last_date": last_date_str,
            "last_detail": str(last_detail),
            "activity_count": activity_count
        }

    def _empty(col):
        if not col or col not in df.columns:
            return pd.Series(True, index=df.index)
        return df[col].isna()

    def _is_ihang_no_add(row):
        if div_col and div_col in df.columns and add_col and add_col in df.columns:
            return (str(row.get(div_col, "")).strip() == "이행" and
                    str(row.get(add_col, "")).strip() == "")
        return False

    # 표시 컬럼 선택 및 활동 이력 추가
    base_show_cols = [c for c in [comp_col, biz_col, div_col, open_col, erp_col, add_col] if c and c in df.columns]

    def fmt_df(sub):
        if sub.empty:
            return sub
        d = sub[base_show_cols].copy().reset_index(drop=True)

        # 날짜 포맷팅
        for col in [open_col, erp_col]:
            if col and col in d.columns:
                d[col] = pd.to_datetime(d[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")

        # 활동 이력 추가
        if biz_col and biz_col in d.columns:
            activities = []
            for _, row in d.iterrows():
                biz_no = row.get(biz_col, "")
                activity = get_recent_activity(biz_no)
                activities.append(activity)

            d["최근활동일"] = [a["last_date"] for a in activities]
            d["최근활동내용"] = [a["last_detail"] for a in activities]
            d["총활동건수"] = [a["activity_count"] for a in activities]

        return d

    # ── 개설 미완료 (하나은행 시트 기반) ──────────────────────
    st.markdown("#### 🟠 개설 미완료 고객")

    if hana_sheet is not None and not hana_sheet.empty:
        hana_df_open = hana_sheet.copy()

        # 컬럼명 확인
        hana_owner_col = "담당자"
        hana_guchuk_col = "구축구분"
        hana_manage_col = "관리구분"
        hana_open_status_col = "개설상태"
        hana_open_date_col = "개설/이행일"
        hana_comp_col = find_col(hana_df_open, ["업체명", "고객명", "상호"])
        hana_biz_col = find_col(hana_df_open, ["사업자번호"])

        # 필터링 조건 적용
        # 1. 본인 담당자
        if hana_owner_col in hana_df_open.columns:
            hana_df_open = hana_df_open[hana_df_open[hana_owner_col].astype(str).str.strip() == str(user_name).strip()]

        # 2. 구축구분 = 신규
        if hana_guchuk_col in hana_df_open.columns:
            hana_df_open = hana_df_open[hana_df_open[hana_guchuk_col].astype(str).str.strip() == "신규"]

        # 3. 관리구분 = 정상
        if hana_manage_col in hana_df_open.columns:
            hana_df_open = hana_df_open[hana_df_open[hana_manage_col].astype(str).str.strip() == "정상"]

        # 4. 개설상태 = 개설대기
        if hana_open_status_col in hana_df_open.columns:
            hana_df_open = hana_df_open[hana_df_open[hana_open_status_col].astype(str).str.strip() == "개설대기"]

        # 5. 개설/이행일 = 공백
        if hana_open_date_col in hana_df_open.columns:
            hana_df_open[hana_open_date_col] = pd.to_datetime(hana_df_open[hana_open_date_col], errors="coerce")
            hana_df_open = hana_df_open[hana_df_open[hana_open_date_col].isna()]

        # 결과 표시
        if not hana_df_open.empty:
            # 표시할 컬럼 선택
            display_cols = [c for c in [hana_comp_col, hana_biz_col, hana_guchuk_col, hana_manage_col, hana_open_status_col, hana_open_date_col] if c and c in hana_df_open.columns]

            # 활동 이력 추가
            result_df = hana_df_open[display_cols].copy().reset_index(drop=True)

            if hana_biz_col and hana_biz_col in result_df.columns:
                activities = []
                for _, row in result_df.iterrows():
                    biz_no = row.get(hana_biz_col, "")
                    activity = get_recent_activity(biz_no)
                    activities.append(activity)

                result_df["최근활동일"] = [a["last_date"] for a in activities]
                result_df["최근활동내용"] = [a["last_detail"] for a in activities]
                result_df["총활동건수"] = [a["activity_count"] for a in activities]

            st.caption(f"{len(result_df)}건")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
        else:
            st.success("개설 미완료 고객 없음")
    else:
        st.warning("하나은행 시트 데이터를 불러올 수 없습니다.")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── ERP연계 미완료 (하나은행 시트 기반) ───────────────────
    st.markdown("#### 🔵 ERP연계 미완료 고객")

    if hana_sheet is not None and not hana_sheet.empty:
        hana_df_erp = hana_sheet.copy()

        # 컬럼명 확인
        hana_owner_col = "담당자"
        hana_manage_col = "관리구분"
        hana_guchuk_type_col = find_col(hana_df_erp, ["구축형", "구축 형"])
        hana_erp_status_col = find_col(hana_df_erp, ["연계상태", "ERP연계상태"])
        hana_comp_col = find_col(hana_df_erp, ["업체명", "고객명", "상호"])
        hana_biz_col = find_col(hana_df_erp, ["사업자번호"])

        # 필터링 조건 적용
        # 1. 본인 담당자
        if hana_owner_col in hana_df_erp.columns:
            hana_df_erp = hana_df_erp[hana_df_erp[hana_owner_col].astype(str).str.strip() == str(user_name).strip()]

        # 2. 관리구분 = 정상
        if hana_manage_col in hana_df_erp.columns:
            hana_df_erp = hana_df_erp[hana_df_erp[hana_manage_col].astype(str).str.strip() == "정상"]

        # 3. 구축형 = 연계형
        if hana_guchuk_type_col and hana_guchuk_type_col in hana_df_erp.columns:
            hana_df_erp = hana_df_erp[hana_df_erp[hana_guchuk_type_col].astype(str).str.strip() == "연계형"]

        # 4. 연계상태 = ERP연계대기 또는 ERP연계진행
        if hana_erp_status_col and hana_erp_status_col in hana_df_erp.columns:
            hana_df_erp = hana_df_erp[hana_df_erp[hana_erp_status_col].astype(str).str.strip().isin(["ERP연계대기", "ERP연계진행"])]

        # 결과 표시
        if not hana_df_erp.empty:
            # 표시할 컬럼 선택
            display_cols = [c for c in [hana_comp_col, hana_biz_col, hana_guchuk_type_col, hana_manage_col, hana_erp_status_col] if c and c in hana_df_erp.columns]

            # 활동 이력 추가
            result_df = hana_df_erp[display_cols].copy().reset_index(drop=True)

            if hana_biz_col and hana_biz_col in result_df.columns:
                activities = []
                for _, row in result_df.iterrows():
                    biz_no = row.get(hana_biz_col, "")
                    activity = get_recent_activity(biz_no)
                    activities.append(activity)

                result_df["최근활동일"] = [a["last_date"] for a in activities]
                result_df["최근활동내용"] = [a["last_detail"] for a in activities]
                result_df["총활동건수"] = [a["activity_count"] for a in activities]

            st.caption(f"{len(result_df)}건")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
        else:
            st.success("ERP연계 미완료 고객 없음")
    else:
        st.warning("하나은행 시트 데이터를 불러올 수 없습니다.")

    # ── 미로그인 고객 현황 ──────────────────────────────────────────
    _tg_force = st.session_state.pop("_nl_rf_tg_nl_refresh", False)
    if _tg_force:
        try:
            hana_sheet = read_google_csv(
                st.session_state.get("url_hana", DEFAULT_URL_HANA), header=2, force_refresh=True
            )
            billing_raw = read_google_csv(
                st.session_state.get("url_hana_billing", DEFAULT_URL_HANA_BILLING), force_refresh=True
            )
            hana_billing_sheet = billing_raw.dropna(how="all").reset_index(drop=True)
        except Exception as e:
            st.error(f"새로고침 실패: {e}")

    if hana_sheet is not None and not hana_sheet.empty and hana_billing_sheet is not None and not hana_billing_sheet.empty:
        _hana = hana_sheet.dropna(how="all").reset_index(drop=True)
        _billing = hana_billing_sheet

        st.divider()
        st.markdown("#### 개설/이행일 이후 미로그인 고객")
        st.caption("개설/이행일 기준으로 청구시트 최종로그인일자가 없거나 개설/이행일 이전인 고객")
        _render_no_login_section(
            _only_current_user_rows(build_no_login_after_open(_hana, _billing)),
            year_key="target_open_year", owner_key="target_open_owner",
            label="미로그인 고객", download_prefix="미로그인고객_개설이행일",
            refresh_key="tg_nl_refresh", exclude_key="tg_nl_open_exclude",
        )

        st.divider()
        st.markdown("#### 연계청구일자 이후 미로그인 고객")
        st.caption("연계청구일자 기준으로 청구시트 최종로그인일자가 없거나 연계청구일자 이전인 고객")
        _render_no_login_section(
            _only_current_user_rows(build_no_login_after_link_billing(_hana, _billing)),
            year_key="target_link_year", owner_key="target_link_owner",
            label="연계청구 미로그인 고객", download_prefix="미로그인고객_연계청구일자",
            exclude_key="tg_nl_link_exclude",
        )

        st.divider()
        st.markdown("#### 로그인 100회 이상 · 미이체 고객")
        st.caption("청구시트 기준 로그인 100회 이상이지만 이체 이력이 없는 고객 (해지 제외)")
        _render_high_login_no_transfer(
            _only_current_user_rows(build_high_login_no_transfer(_hana, _billing)),
            owner_key="tg_high_login_owner",
            download_prefix="미이체고객",
        )
    else:
        st.divider()
        st.warning("하나은행 구글시트 또는 청구시트 데이터를 불러올 수 없습니다.")


VISIT_HISTORY_FILE = "visit_history.json"


def show_visit_history():
    st.markdown("### 방문이력 작성")
    user_name = st.session_state.get("user_name", "")
    db = load_db(VISIT_HISTORY_FILE, {})
    my_entries = db.get(user_name, [])

    # 고객명 자동검색용 목록 — 하나은행 구글시트에서 추출
    customer_options = []
    hana_df = st.session_state.get("hana_sheet_df")
    if hana_df is not None and not hana_df.empty:
        comp_col = find_col(hana_df, ["고객명", "업체명", "상호"])
        owner_col = find_col(hana_df, ["담당자"])
        if comp_col:
            filtered_hana = hana_df
            if owner_col:
                filtered_hana = hana_df[
                    hana_df[owner_col].astype(str).str.strip() == str(user_name).strip()
                ]
            customer_options = sorted(
                filtered_hana[comp_col].dropna().astype(str).str.strip()
                .replace("", pd.NA).dropna().unique().tolist()
            )

    # ── 등록 폼 ──
    with st.expander("방문이력 등록", expanded=True):
        # 고객명 자동검색 (폼 바깥 — 실시간 필터링)
        if customer_options:
            v_customer_sel = st.selectbox(
                "고객명 검색",
                ["직접 입력"] + customer_options,
                key="visit_customer_sel",
            )
        else:
            v_customer_sel = "직접 입력"

        with st.form("visit_form"):
            vc1, vc2 = st.columns(2)
            with vc1:
                v_date = st.date_input("방문일자")
            with vc2:
                v_purpose = st.selectbox("방문목적", ["개설", "운영", "연계"])
            # 고객명: 자동검색 선택값 또는 직접 입력
            if v_customer_sel and v_customer_sel != "직접 입력":
                v_customer = st.text_input("고객명", value=v_customer_sel)
            else:
                v_customer = st.text_input("고객명", placeholder="고객명을 입력하세요")
            v_content = st.text_area("내용", height=120)
            v_submitted = st.form_submit_button("등록", use_container_width=True, type="primary")

        if v_submitted:
            if not v_customer.strip():
                st.error("고객명을 입력해주세요.")
            else:
                now_text = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S")
                entry = {
                    "id": hashlib.md5(f"{user_name}-{now_text}-{v_customer}".encode()).hexdigest(),
                    "방문일자": str(v_date),
                    "고객명": v_customer.strip(),
                    "방문목적": v_purpose,
                    "내용": v_content.strip(),
                    "담당자": user_name,
                    "작성시각": now_text,
                }
                db.setdefault(user_name, []).append(entry)
                save_db(VISIT_HISTORY_FILE, db)
                st.success("방문이력을 등록했습니다.")
                st.rerun()

    # ── 이력 목록 ──
    st.markdown("#### 내 방문이력")
    if not my_entries:
        st.info("등록된 방문이력이 없습니다.")
        return

    sorted_entries = sorted(my_entries, key=lambda x: x.get("방문일자", ""), reverse=True)
    show_cols = ["방문일자", "고객명", "방문목적", "내용", "작성시각"]
    display_df = pd.DataFrame(sorted_entries)
    display_df = display_df[[c for c in show_cols if c in display_df.columns]].copy()
    display_df.insert(0, "선택", False)

    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={"선택": st.column_config.CheckboxColumn("선택", default=False, width="small")},
        disabled=[c for c in display_df.columns if c != "선택"],
        key="visit_history_editor",
    )

    selected_ids = [sorted_entries[i].get("id") for i, v in enumerate(edited_df["선택"]) if v]
    _, del_col, _ = st.columns([15, 15, 70])
    with del_col:
        if st.button("삭제", use_container_width=True, disabled=not selected_ids, key="visit_del_btn"):
            db[user_name] = [r for r in my_entries if r.get("id") not in selected_ids]
            save_db(VISIT_HISTORY_FILE, db)
            st.success(f"{len(selected_ids)}건 삭제했습니다.")
            st.rerun()

    today_str = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m%d")
    dl_df = pd.DataFrame(sorted_entries)[[c for c in show_cols if c in pd.DataFrame(sorted_entries).columns]]
    st.download_button(
        "방문이력 다운로드",
        data=dl_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
        file_name=f"방문이력_{user_name}_{today_str}.csv",
        mime="text/csv",
        use_container_width=True,
    )


def show_weekly_report_user():
    opening_categories = ["개설대기", "개설진행", "개설완료"]
    link_categories = ["연계대기", "연계진행", "연계완료"]
    report_categories = opening_categories + link_categories
    categories = ["전체"] + report_categories + ["운영부문"]
    user_name = st.session_state.get("user_name", "")

    def load_hana_sheet_for_weekly(force_refresh=False):
        try:
            url = st.session_state.get("url_hana", DEFAULT_URL_HANA)
            df = read_google_csv(url, header=2, force_refresh=force_refresh)
            st.session_state.hana_sheet_df = df
            st.session_state.weekly_hana_loaded_at = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S")
        except Exception as e:
            df = st.session_state.get("hana_sheet_df")
            if df is None or df.empty:
                st.error(f"하나은행 구글 시트를 불러오지 못했습니다: {e}")
                return pd.DataFrame()
            st.warning("하나은행 구글 시트 최신 데이터를 불러오지 못해 세션에 저장된 데이터를 사용합니다.")
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(how="all").reset_index(drop=True)

    def weekly_category_mask(df, category, start_date=None, end_date=None):
        open_status_col = find_col(df, ["개설상태"])
        open_date_col = find_col(df, ["개설/이행일", "개설일", "이행일"])
        link_status_col = find_col(df, ["연계상태"])
        link_date_col = find_col(df, ["연계일자", "연계일"])
        build_type_col = find_col(df, ["구축형"])
        build_category_col = find_col(df, ["구축구분"])
        end_col = find_col(df, ["해지일자"])

        open_status = df[open_status_col].astype(str).str.strip() if open_status_col else pd.Series("", index=df.index)
        link_status = df[link_status_col].astype(str).str.strip() if link_status_col else pd.Series("", index=df.index)
        build_type = df[build_type_col].astype(str).str.strip() if build_type_col else pd.Series("", index=df.index)
        build_category = df[build_category_col].astype(str).str.strip() if build_category_col else pd.Series("", index=df.index)
        open_dates = pd.to_datetime(df[open_date_col].map(parse_sheet_date), errors="coerce") if open_date_col else pd.Series(pd.NaT, index=df.index)
        link_dates = pd.to_datetime(df[link_date_col].map(parse_sheet_date), errors="coerce") if link_date_col else pd.Series(pd.NaT, index=df.index)
        open_date = open_dates.notna()
        link_date = link_dates.notna()
        active = df[end_col].map(is_blank_value) if end_col else pd.Series(True, index=df.index)

        open_done = active & (open_status.str.contains("완료|이행완료", na=False) | open_date)
        link_done = active & (link_status.str.contains("완료", na=False) | link_date)
        open_progress = active & ~open_done & open_status.str.contains("진행|구축중|처리중", na=False)
        link_progress = active & ~link_done & link_status.str.contains("진행|연계중|처리중", na=False)
        link_target = active & (build_type.str.contains("연계|이행", na=False) | build_category.str.contains("연계|이행", na=False) | link_status.ne(""))

        if category == "개설완료":
            if start_date is not None and end_date is not None:
                return active & open_dates.between(pd.Timestamp(start_date), pd.Timestamp(end_date), inclusive="both")
            return active & open_date
        if category == "개설진행":
            return open_progress
        if category == "개설대기":
            return active & open_status.eq("개설대기")
        if category == "연계완료":
            link_complete = link_status.str.replace(r"\s+", "", regex=True).eq("ERP연계완료")
            if start_date is not None and end_date is not None:
                return active & link_complete & link_dates.between(pd.Timestamp(start_date), pd.Timestamp(end_date), inclusive="both")
            return active & link_complete & link_date
        if category == "연계진행":
            return link_progress
        if category == "연계대기":
            return active & link_status.str.contains("ERP연계대기", na=False)
        if category == "운영부문":
            return active & open_done & (~link_target | link_done)
        return pd.Series(True, index=df.index)

    def display_customer_df(df):
        wanted_keys = ["고객번호", "고객명", "구축구분", "구축형", "신규접수일", "추가연계접수일", "개설상태", "개설/이행일", "연계상태", "연계일자", "구축예정일", "담당자"]
        cols = []
        for key in wanted_keys:
            col = find_col(df, [key])
            if col and col in df.columns and col not in cols:
                cols.append(col)
        if not cols:
            cols = list(df.columns[:8])
        show = df[cols].copy()
        for col in show.columns:
            if any(k in str(col) for k in ["일자", "일", "예정"]):
                parsed = pd.to_datetime(show[col].map(parse_sheet_date), errors="coerce")
                show[col] = show[col].where(parsed.isna(), parsed.dt.strftime("%Y-%m-%d"))
        return show.reset_index(drop=True)

    def format_weekly_date(value):
        parsed = parse_sheet_date(value)
        if pd.isna(parsed):
            return "" if is_blank_value(value) else str(value).strip()
        return pd.Timestamp(parsed).strftime("%Y-%m-%d")

    def date_input_value(value, fallback):
        parsed = parse_sheet_date(value)
        if pd.isna(parsed):
            return fallback
        return pd.Timestamp(parsed).date()

    def render_weekly_detail_header(title, key, level=4):
        title_col, refresh_col = st.columns([0.82, 0.18])
        with title_col:
            st.markdown(f"{'#' * level} {title}")
        with refresh_col:
            if st.button("새로고침", key=f"weekly_refresh_{key}", use_container_width=True):
                st.session_state.weekly_force_hana_refresh = True
                st.rerun()

    title_col, refresh_col = st.columns([0.82, 0.18])
    with title_col:
        st.markdown("### 주간보고 작성")
    force_refresh = bool(st.session_state.pop("weekly_force_hana_refresh", False))
    with refresh_col:
        if st.button("새로고침", key="weekly_user_top_refresh", use_container_width=True):
            force_refresh = True
    st.caption("하나은행 구글 시트의 본인 담당 고객을 기준으로 카테고리별 현황을 선택해 주간보고 이력을 저장합니다.")

    hana = load_hana_sheet_for_weekly(force_refresh=force_refresh)
    if hana.empty:
        return

    staff_col = find_col(hana, ["담당자"])
    company_col = find_col(hana, ["고객명", "업체명", "상호"])
    branch_col = find_col(hana, ["신청점", "담당 부서", "담당부서"])
    service_col = find_col(hana, ["구축구분", "서비스"])
    build_type_col = find_col(hana, ["구축형"])
    customer_no_col = find_col(hana, ["고객번호"])
    open_status_col = find_col(hana, ["개설상태"])
    link_status_col = find_col(hana, ["연계상태"])
    link_date_col = find_col(hana, ["연계일자", "연계일"])

    if not staff_col or not company_col:
        st.error("하나은행 시트에서 담당자 또는 고객명 컬럼을 찾을 수 없습니다.")
        return

    mine = hana[hana[staff_col].astype(str).str.strip() == str(user_name).strip()].copy()
    if mine.empty:
        st.info(f"하나은행 시트에 {user_name}님 담당 고객이 없습니다.")
        return

    today = datetime.utcnow() + timedelta(hours=9)
    week_start_default = (today - timedelta(days=today.weekday() + 7)).date()
    week_end_default = week_start_default + timedelta(days=4)

    col_a, col_b, col_c = st.columns([1.1, 1.1, 1.4])
    with col_a:
        week_start = st.date_input("보고 시작일", value=week_start_default, key="weekly_report_start_v2")
    with col_b:
        week_end = st.date_input("보고 종료일", value=week_end_default, key="weekly_report_end_v2")
    with col_c:
        category = st.selectbox("카테고리", categories, key="weekly_report_category")
    st.caption(f"선택 기간: {week_start} ~ {week_end}")
    if st.session_state.get("weekly_hana_loaded_at"):
        st.caption(f"하나은행 구글 시트 새로고침: {st.session_state.weekly_hana_loaded_at}")

    if category == "전체":
        categorized_map = {
            cat: mine[weekly_category_mask(mine, cat, week_start, week_end)].copy()
            for cat in report_categories
        }
        st.markdown("#### 전체 현황")
        st.caption("개설 부문과 ERP연계 부문을 PPT 주간보고 구성에 맞춰 표시합니다.")
        st.markdown("##### 개설 부문")
        for cat in opening_categories:
            render_weekly_detail_header(cat, f"all_{cat}", level=5)
            st.caption(f"{len(categorized_map[cat])}건 · 하나은행 구글 시트 기준")
            render_plain_html_table(display_customer_df(categorized_map[cat]), max_rows=200)
        st.markdown("##### ERP연계 부문")
        for cat in link_categories:
            render_weekly_detail_header(cat, f"all_{cat}", level=5)
            st.caption(f"{len(categorized_map[cat])}건 · 하나은행 구글 시트 기준")
            render_plain_html_table(display_customer_df(categorized_map[cat]), max_rows=200)
    else:
        categorized = mine[weekly_category_mask(mine, category, week_start, week_end)].copy()
        render_weekly_detail_header(f"{category} 현황", f"single_{category}", level=4)
        st.caption(f"{len(categorized)}건 · 하나은행 구글 시트 기준")
        render_plain_html_table(display_customer_df(categorized), max_rows=200)

        if categorized.empty and category != "운영부문":
            st.info("선택한 카테고리에 작성할 고객이 없습니다.")
            if category == "연계완료" and link_status_col and link_date_col:
                status_clean = mine[link_status_col].astype(str).str.strip().str.replace(r"\s+", "", regex=True)
                candidates = mine[status_clean.eq("ERP연계완료")].copy()
                if not candidates.empty:
                    candidates["_연계일자변환"] = pd.to_datetime(candidates[link_date_col].map(parse_sheet_date), errors="coerce")
                    candidates["_선택기간포함"] = candidates["_연계일자변환"].between(pd.Timestamp(week_start), pd.Timestamp(week_end), inclusive="both")
                    with st.expander("ERP연계완료 후보 확인", expanded=False):
                        cols = [company_col, link_status_col, link_date_col, "_연계일자변환", "_선택기간포함"]
                        cols = [c for c in cols if c and c in candidates.columns]
                        show_candidates = candidates[cols].copy()
                        if "_연계일자변환" in show_candidates.columns:
                            show_candidates["_연계일자변환"] = show_candidates["_연계일자변환"].dt.strftime("%Y-%m-%d")
                        st.dataframe(show_candidates.reset_index(drop=True), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### 주간보고 이력 입력")
    if category == "전체":
        input_part = st.radio("입력 부문", ["개설 부문", "ERP연계 부문"], horizontal=True, key="weekly_input_part")
        input_category_options = opening_categories if input_part == "개설 부문" else link_categories
        input_category = st.selectbox("세부 카테고리", input_category_options, key="weekly_input_category")
    else:
        input_category = category
        input_part = "ERP연계 부문" if input_category in link_categories else "개설 부문"
        st.caption(f"{input_part} · {input_category}")

    if st.session_state.get("_weekly_prev_input_category") != input_category:
        st.session_state.pop("weekly_customer_combo", None)
        for _cat in ["개설대기", "개설진행", "연계대기", "연계진행"]:
            st.session_state.pop(f"weekly_plan_date_{_cat}", None)
        st.session_state["_weekly_prev_input_category"] = input_category

    categorized = mine[weekly_category_mask(mine, input_category, week_start, week_end)].copy()

    tmp = categorized.copy()
    customer_options = []
    if not tmp.empty:
        tmp["_company_label"] = tmp.apply(
            lambda r: f"{str(r.get(company_col, '')).strip()} | {str(r.get(customer_no_col, '')).strip() if customer_no_col else ''}".strip(" |"),
            axis=1,
        )
        customer_options = [v for v in tmp["_company_label"].tolist() if v]
    label_to_row = {str(row["_company_label"]): row for _, row in tmp.iterrows()} if not tmp.empty else {}
    if not customer_options:
        customer_options = ["선택 가능한 고객 없음"]

    if st.session_state.get("weekly_customer_combo") not in customer_options:
        st.session_state.pop("weekly_customer_combo", None)
    selected_company_label = st.selectbox("고객명", customer_options, key="weekly_customer_combo")
    selected_row = label_to_row.get(selected_company_label)
    default_company = "" if selected_row is None else str(selected_row.get(company_col, "")).strip()
    default_customer_no = "" if selected_row is None or not customer_no_col else str(selected_row.get(customer_no_col, "")).strip()
    default_service = "" if selected_row is None or not service_col else str(selected_row.get(service_col, "")).strip()
    default_build = "" if selected_row is None or not build_type_col else str(selected_row.get(build_type_col, "")).strip()
    default_branch = "" if selected_row is None or not branch_col else str(selected_row.get(branch_col, "")).strip()
    open_receipt_col = find_col(mine, ["신규접수일"])
    link_receipt_col = find_col(mine, ["추가연계접수일"])
    schedule_col = find_col(mine, ["구축예정일"])
    receipt_col = link_receipt_col if input_category in link_categories else open_receipt_col
    default_receipt_date = "" if selected_row is None or not receipt_col else format_weekly_date(selected_row.get(receipt_col, ""))
    default_plan_date = "" if selected_row is None or not schedule_col else format_weekly_date(selected_row.get(schedule_col, ""))
    plan_date_label = "구축/피드백 예정일"
    if input_category == "개설완료":
        plan_date_label = "개설완료일자"
        open_done_col = find_col(mine, ["개설/이행일", "개설일", "이행일"])
        default_plan_date = "" if selected_row is None or not open_done_col else format_weekly_date(selected_row.get(open_done_col, ""))
    elif input_category == "연계완료":
        plan_date_label = "연계완료일자"
        link_done_col = find_col(mine, ["연계일자", "연계일"])
        default_plan_date = "" if selected_row is None or not link_done_col else format_weekly_date(selected_row.get(link_done_col, ""))
    status_default = ""
    if selected_row is not None:
        if input_category in link_categories and link_status_col:
            status_default = str(selected_row.get(link_status_col, "")).strip()
        elif open_status_col:
            status_default = str(selected_row.get(open_status_col, "")).strip()

    with st.form("weekly_report_form", border=True):
        st.markdown(f"##### {input_part} - {input_category}")
        company = default_company
        f1, f2 = st.columns(2)
        with f1:
            service = st.text_input("서비스", value=default_service)
        with f2:
            build_type = st.text_input("구분", value=default_build)

        f4, f5, f6 = st.columns([1.2, 1.0, 1.0])
        with f4:
            branch = st.text_input("신청점/부서", value=default_branch)
        with f5:
            customer_no = st.text_input("고객번호", value=default_customer_no)
        with f6:
            status = st.text_input("상태", value=status_default)

        f7, f8 = st.columns(2)
        with f7:
            receipt_date = st.text_input("접수일자", value=default_receipt_date)
        with f8:
            if input_category in ["개설대기", "개설진행", "연계대기", "연계진행"]:
                plan_date = st.date_input(
                    plan_date_label,
                    value=date_input_value(default_plan_date, today.date()),
                    key=f"weekly_plan_date_{input_category}",
                ).strftime("%Y-%m-%d")
            else:
                plan_date = st.text_input(plan_date_label, value=default_plan_date)

        issue = st.text_area("이슈 / 진행내용", height=100, placeholder="PPT 주간보고의 이슈 항목처럼 고객별 진행상황을 작성")
        special_note = st.text_area("특이사항", height=80, placeholder="예정 작업, 고객 회신 필요사항, 은행 전달사항 등")

        submitted = st.form_submit_button("주간보고 저장", use_container_width=True, type="primary")

    if submitted:
        if not company.strip():
            st.warning("고객명을 입력해주세요.")
        else:
            db = load_db(WEEKLY_REPORT_FILE, {})
            entries = db.setdefault(user_name, [])
            now_text = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S")
            entries.append({
                "id": hashlib.md5(f"{user_name}-{now_text}-{company}-{input_category}".encode("utf-8")).hexdigest(),
                "보고시작일": str(week_start),
                "보고종료일": str(week_end),
                "부문": input_part,
                "카테고리": input_category,
                "서비스": service.strip(),
                "구분": build_type.strip(),
                "고객명": company.strip(),
                "고객번호": customer_no.strip(),
                "신청점": branch.strip(),
                "접수일자": receipt_date.strip(),
                "일자구분": plan_date_label,
                "일자": plan_date.strip(),
                "이슈": issue.strip(),
                "특이사항": special_note.strip(),
                "상태": status.strip(),
                "담당자": user_name,
                "작성시각": now_text,
            })
            save_db(WEEKLY_REPORT_FILE, db)
            st.success("주간보고 이력을 저장했습니다.")
            st.rerun()

    db = load_db(WEEKLY_REPORT_FILE, {})
    my_entries = db.get(user_name, [])
    st.markdown("#### 내 작성 이력")
    if not my_entries:
        st.info("아직 저장된 주간보고 이력이 없습니다.")
        return

    sorted_entries = sorted(my_entries, key=lambda x: x.get("작성시각", ""), reverse=True)
    history_df = pd.DataFrame(sorted_entries)
    show_cols = ["보고시작일", "보고종료일", "부문", "카테고리", "고객명", "접수일자", "일자구분", "일자", "이슈", "특이사항", "상태", "작성시각"]
    show_cols = [c for c in show_cols if c in history_df.columns]

    display_df = history_df[show_cols].copy()
    display_df.insert(0, "선택", False)

    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={"선택": st.column_config.CheckboxColumn("선택", default=False, width="small")},
        disabled=show_cols,
        key="weekly_history_editor",
    )

    selected_indices = [i for i, v in enumerate(edited_df["선택"]) if v]
    selected_ids = [sorted_entries[i].get("id") for i in selected_indices]

    btn_col1, btn_col2, _ = st.columns([15, 15, 70])
    with btn_col1:
        edit_clicked = st.button("수정", use_container_width=True, disabled=(len(selected_ids) != 1), key="weekly_edit_btn")
    with btn_col2:
        del_clicked = st.button("삭제", use_container_width=True, disabled=(len(selected_ids) == 0), key="weekly_del_btn")

    if edit_clicked and len(selected_ids) == 1:
        st.session_state["weekly_edit_id"] = selected_ids[0]
        st.session_state["weekly_edit_mode"] = True
        st.rerun()

    if del_clicked and selected_ids:
        db[user_name] = [row for row in my_entries if row.get("id") not in selected_ids]
        save_db(WEEKLY_REPORT_FILE, db)
        st.success(f"{len(selected_ids)}개 이력을 삭제했습니다.")
        st.rerun()

    if st.session_state.get("weekly_edit_mode") and st.session_state.get("weekly_edit_id"):
        edit_id = st.session_state["weekly_edit_id"]
        edit_entry = next((row for row in my_entries if row.get("id") == edit_id), None)
        if edit_entry:
            st.markdown("---")
            st.markdown("##### 선택 이력 수정")
            with st.form("weekly_edit_form"):
                ec1, ec2 = st.columns(2)
                with ec1:
                    e_start = st.text_input("보고시작일", value=edit_entry.get("보고시작일", ""))
                    e_category = st.text_input("카테고리", value=edit_entry.get("카테고리", ""))
                    e_receipt = st.text_input("접수일자", value=edit_entry.get("접수일자", ""))
                    e_date = st.text_input("일자", value=edit_entry.get("일자", ""))
                with ec2:
                    e_end = st.text_input("보고종료일", value=edit_entry.get("보고종료일", ""))
                    e_status = st.text_input("상태", value=edit_entry.get("상태", ""))
                    e_date_type = st.text_input("일자구분", value=edit_entry.get("일자구분", ""))
                e_issue = st.text_area("이슈 / 진행내용", value=edit_entry.get("이슈", ""), height=100)
                e_special = st.text_area("특이사항", value=edit_entry.get("특이사항", ""), height=80)
                sc1, sc2 = st.columns(2)
                with sc1:
                    save_edit = st.form_submit_button("저장", use_container_width=True, type="primary")
                with sc2:
                    cancel_edit = st.form_submit_button("취소", use_container_width=True)

            if save_edit:
                for i, row in enumerate(db.get(user_name, [])):
                    if row.get("id") == edit_id:
                        db[user_name][i].update({
                            "보고시작일": e_start.strip(),
                            "보고종료일": e_end.strip(),
                            "카테고리": e_category.strip(),
                            "접수일자": e_receipt.strip(),
                            "일자구분": e_date_type.strip(),
                            "일자": e_date.strip(),
                            "상태": e_status.strip(),
                            "이슈": e_issue.strip(),
                            "특이사항": e_special.strip(),
                        })
                        break
                save_db(WEEKLY_REPORT_FILE, db)
                st.session_state["weekly_edit_mode"] = False
                st.session_state["weekly_edit_id"] = None
                st.success("수정되었습니다.")
                st.rerun()

            if cancel_edit:
                st.session_state["weekly_edit_mode"] = False
                st.session_state["weekly_edit_id"] = None
                st.rerun()


def load_weekly_hana_for_status():
    try:
        df = read_google_csv(st.session_state.get("url_hana", DEFAULT_URL_HANA), header=2)
    except Exception:
        df = st.session_state.get("hana_sheet_df")
    if df is None:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)


def weekly_customer_status_tables(hana_df, year=2026, snapshot_end_date=None, save_snapshot=False):
    if hana_df is None or hana_df.empty:
        return {}

    df = hana_df.copy()
    open_receipt_col = find_col(df, ["신규접수일"])
    link_receipt_col = find_col(df, ["추가연계접수일"])
    open_date_col = find_col(df, ["개설/이행일", "개설일", "이행일"])
    link_date_col = find_col(df, ["연계일자", "연계일"])
    link_billing_date_col = find_col(df, ["연계청구일자", "연계청구일", "청구일자"])
    open_status_col = find_col(df, ["개설상태"])
    link_status_col = find_col(df, ["연계상태"])
    build_type_col = find_col(df, ["구축형"])
    manage_col = find_col(df, ["관리구분"])
    end_col = find_col(df, ["해지일자"])

    if not open_status_col:
        return {}

    def dates(col):
        if not col or col not in df.columns:
            return pd.Series(pd.NaT, index=df.index)
        return pd.to_datetime(df[col].map(parse_sheet_date), errors="coerce")

    def fmt(value):
        try:
            value = int(value)
        except Exception:
            value = 0
        return "-" if value == 0 else f"{value:,}"

    def count(mask):
        return int(mask.fillna(False).sum())

    open_receipt_dates = dates(open_receipt_col)
    link_receipt_dates = dates(link_receipt_col).fillna(open_receipt_dates)
    open_done_dates = dates(open_date_col)
    link_done_dates = dates(link_date_col)
    link_billing_dates = dates(link_billing_date_col)
    end_dates = dates(end_col)
    as_of_date = pd.to_datetime(snapshot_end_date, errors="coerce")
    if pd.isna(as_of_date):
        as_of_date = pd.Timestamp.max

    year_start = pd.Timestamp(f"{year}-01-01")
    active = end_dates.isna() | (end_dates > as_of_date)
    # 월별 표용: 전체 기간(연도 무관)
    open_received_by_asof = open_receipt_dates.notna() & (open_receipt_dates <= as_of_date)
    link_received_by_asof = link_receipt_dates.notna() & (link_receipt_dates <= as_of_date)
    # 상태 표용: year 기준(2026.01.01~) 필터
    open_received_in_year = open_receipt_dates.notna() & (open_receipt_dates >= year_start) & (open_receipt_dates <= as_of_date)
    link_received_in_year = link_receipt_dates.notna() & (link_receipt_dates >= year_start) & (link_receipt_dates <= as_of_date)
    open_done_by_asof = open_done_dates.notna() & (open_done_dates <= as_of_date)
    link_done_by_asof = link_done_dates.notna() & (link_done_dates <= as_of_date)
    link_billing_by_asof = link_billing_dates.notna() & (link_billing_dates <= as_of_date)
    open_status = df[open_status_col].astype(str).str.strip()
    link_status = df[link_status_col].astype(str).str.strip() if link_status_col else pd.Series("", index=df.index)
    link_clean = link_status.str.replace(r"\s+", "", regex=True)
    build_type = df[build_type_col].astype(str).str.strip() if build_type_col else pd.Series("", index=df.index)
    manage = df[manage_col].astype(str).str.strip() if manage_col else pd.Series("", index=df.index)

    basic = build_type.str.contains("기본", na=False)
    linked = build_type.str.contains("연계|이행", na=False)
    transfer = manage.str.contains("이관", na=False)

    # 개설상태 명시값 기준 분류 (소거법 사용 X, NaN 제외)
    open_cancel = open_received_by_asof & open_status.str.contains("취소|DROP|드랍", case=False, na=False)
    link_cancel = link_received_by_asof & link_status.str.contains("취소|DROP|드랍", case=False, na=False)
    open_done = active & open_received_by_asof & ~open_cancel & (open_status.str.contains("완료|이행완료", na=False) & open_done_by_asof)
    open_progress = active & open_received_by_asof & ~open_cancel & ~open_done & open_status.str.contains("진행|구축중|처리중", na=False)
    open_wait = active & open_received_by_asof & ~open_cancel & ~open_done & ~open_progress & open_status.str.contains("대기", na=False)
    # 연계완료: 연계일자 유무와 관계없이 상태값 "완료" 기준으로 판단
    link_done = active & link_received_by_asof & ~link_cancel & link_status.str.contains("완료", na=False)
    # 연계진행/대기: 패턴을 넓혀 "진행", "대기" 포함 변형 모두 매칭
    link_progress = active & link_received_by_asof & ~link_cancel & ~link_done & link_status.str.contains("진행", na=False)
    link_wait = active & link_received_by_asof & ~link_cancel & ~link_done & ~link_progress & link_status.str.contains("대기", na=False)
    link_receipt = active & (link_received_by_asof | link_wait | link_progress | link_done)
    terminated = end_dates.notna() & (end_dates <= as_of_date)

    # 상태 표: 구글시트 기준과 동일하게 연계상태 컬럼값만으로 단순 집계
    # active/link_received_by_asof 등 부가 조건 없이 연계상태 텍스트만 매칭
    st_link_wait     = link_status.str.contains("대기", na=False)
    st_link_progress = link_status.str.contains("진행", na=False)
    st_link_done     = link_status.str.contains("완료", na=False)
    status_rows = [
        ["전체", "", count(open_wait), count(open_progress), count(open_done), count(st_link_wait), count(st_link_progress), count(st_link_done), count(terminated)],
        ["", "기본형", count(open_wait & basic), count(open_progress & basic), count(open_done & basic), "-", "-", "-", count(terminated & basic)],
        ["", "연계형", count(open_wait & linked), count(open_progress & linked), count(open_done & linked), count(st_link_wait), count(st_link_progress), count(st_link_done), count(terminated & linked)],
    ]
    status_df = pd.DataFrame(status_rows, columns=["구분", "유형", "구축대기", "구축진행", "구축완료", "연계대기", "연계진행", "연계완료", "해지"])

    month_cols = ["2025년"] + [f"{year}{m:02d}" for m in range(1, 13)] + [f"{str(year)[-2:]}년합계", "합계"]

    def month_count_values(mask, event_date):
        values = []
        bounded_date = event_date.notna() & (event_date <= as_of_date)
        prev_year_value = count(mask & bounded_date & (event_date.dt.year == year - 1))
        values.append(prev_year_value)
        year_total = 0
        for month in range(1, 13):
            value = count(mask & bounded_date & (event_date.dt.year == year) & (event_date.dt.month == month))
            values.append(value)
            year_total += value
        values.append(year_total)                    # 26년합계 (2026년 월별 합산)
        values.append(prev_year_value + year_total)  # 합계 (2025년 + 2026년)
        return values

    def fmt_values(values):
        return [fmt(v) for v in values]

    def summed_values(*rows):
        return [sum(values) for values in zip(*rows)]

    def debug_rows(mask, event_date, reason):
        debug_cols = [
            find_col(df, ["업체명", "고객명", "상호", "회사명"]),
            find_col(df, ["담당자"]),
            find_col(df, ["구축형"]),
            find_col(df, ["관리구분"]),
            open_receipt_col,
            open_status_col,
            open_date_col,
            end_col,
        ]
        labels = ["고객명", "담당자", "구축형", "관리구분", "신규접수일", "개설상태", "개설/이행일", "해지일자"]
        rows = []
        for idx in df[mask].index:
            row = {}
            for label, col in zip(labels, debug_cols):
                row[label] = df.at[idx, col] if col and col in df.columns else ""
            row["신규접수일(파싱)"] = event_date.at[idx].strftime("%Y-%m-%d") if pd.notna(event_date.at[idx]) else ""
            row["산출사유"] = reason
            rows.append(row)
        return pd.DataFrame(rows, columns=labels + ["신규접수일(파싱)", "산출사유"])

    open_wait_values = month_count_values(open_wait, open_receipt_dates)
    open_progress_values = month_count_values(open_progress, open_receipt_dates)
    open_done_values = month_count_values(open_done, open_receipt_dates)
    open_cancel_values = month_count_values(open_cancel, open_receipt_dates)
    debug_open_wait_prev_year = debug_rows(
        open_wait & open_receipt_dates.notna() & (open_receipt_dates.dt.year == year - 1) & (open_receipt_dates <= as_of_date),
        open_receipt_dates,
        f"{year - 1}년 신규접수 + 기준일 현재 구축대기",
    )

    # 구축 접수: 신규접수일 기준 직접 카운트 (개설상태 패턴 불일치로 인한 누락 방지)
    open_receipt_count_values = month_count_values(open_received_by_asof, open_receipt_dates)
    open_month_rows = [
        ["구축", "접수"] + fmt_values(open_receipt_count_values),
        ["", "대기"] + fmt_values(open_wait_values),
        ["", "진행"] + fmt_values(open_progress_values),
        ["", "완료"] + fmt_values(open_done_values),
        ["", "취소"] + fmt_values(open_cancel_values),
    ]

    # 연계 월별 표: 연계상태 텍스트만 매칭 (구축형 필터 없음 → 전체 행 기준과 일치)
    _lr = link_receipt_dates.notna() & (link_receipt_dates <= as_of_date)
    link_received_mask   = _lr
    mt_link_wait         = _lr & link_status.str.contains("대기", na=False)
    mt_link_progress     = _lr & link_status.str.contains("진행", na=False)
    mt_link_done         = _lr & link_status.str.contains("완료", na=False)
    mt_link_cancel       = _lr & link_status.str.contains("취소|DROP|드랍", case=False, na=False)
    link_receipt_count_values = month_count_values(link_received_mask, link_receipt_dates)
    link_wait_values     = month_count_values(mt_link_wait,     link_receipt_dates)
    link_progress_values = month_count_values(mt_link_progress, link_receipt_dates)
    link_done_values     = month_count_values(mt_link_done,     link_receipt_dates)
    link_cancel_values   = month_count_values(mt_link_cancel,   link_receipt_dates)

    link_month_rows = [
        ["연계", "접수"] + fmt_values(summed_values(link_wait_values, link_progress_values, link_done_values, link_cancel_values)),
        ["", "대기"] + fmt_values(link_wait_values),
        ["", "진행"] + fmt_values(link_progress_values),
        ["", "완료"] + fmt_values(link_done_values),
        ["", "취소"] + fmt_values(link_cancel_values),
    ]
    monthly_df = pd.DataFrame(open_month_rows + link_month_rows, columns=["구분", "상태"] + month_cols)

    current_group = ""
    monthly_counts = {}
    for row in open_month_rows + link_month_rows:
        if row[0]:
            current_group = row[0]
        monthly_counts[f"{current_group}|{row[1]}"] = [int(str(v).replace(",", "").replace("-", "0")) for v in row[2:]]

    snapshot_key = str(snapshot_end_date or "").strip()
    snapshots = load_db(WEEKLY_STATUS_SNAPSHOT_FILE, {})
    prev_snapshot = None
    if snapshot_key and snapshots:
        prev_keys = sorted(k for k in snapshots.keys() if k < snapshot_key)
        if prev_keys:
            prev_snapshot = snapshots.get(prev_keys[-1], {}).get("monthly_counts", {})

    delta_rows = []
    current_group_for_delta = ""
    for _, row in monthly_df.iterrows():
        if str(row["구분"]).strip():
            current_group_for_delta = str(row["구분"]).strip()
        status = str(row["상태"]).strip()
        key = f"{current_group_for_delta}|{status}"
        current_values = monthly_counts.get(key, [0] * len(month_cols))
        previous_values = prev_snapshot.get(key, [0] * len(month_cols)) if prev_snapshot else None
        delta_values = []
        for idx, current_value in enumerate(current_values):
            if previous_values is None:
                delta_values.append("-")
                continue
            prev_value = previous_values[idx] if idx < len(previous_values) else 0
            delta = current_value - int(prev_value)
            if delta > 0:
                delta_values.append(f"▲{delta:,}")
            elif delta < 0:
                delta_values.append(f"▼{abs(delta):,}")
            else:
                delta_values.append("-")
        delta_rows.append(["", "전주대비"] + delta_values)
    expanded_rows = []
    for idx, row in monthly_df.iterrows():
        expanded_rows.append(row.tolist())
        expanded_rows.append(delta_rows[idx])
    monthly_display_df = pd.DataFrame(expanded_rows, columns=monthly_df.columns)

    if save_snapshot and snapshot_key:
        snapshots[snapshot_key] = {
            "saved_at": (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S"),
            "monthly_counts": monthly_counts,
        }
        save_db(WEEKLY_STATUS_SNAPSHOT_FILE, snapshots)

    complete_cols = ["구분", "-"] + [f"{year}{m:02d}" for m in range(1, 13)] + [f"{str(year)[-2:]}' 합계"]
    open_complete_values = []
    link_complete_values = []
    open_total = 0
    link_total = 0
    link_billing_done = link_clean.eq("ERP청구완료")
    for month in range(1, 13):
        ov = count(open_done_by_asof & (open_done_dates.dt.year == year) & (open_done_dates.dt.month == month))
        lv = count(link_billing_done & link_billing_by_asof & (link_billing_dates.dt.year == year) & (link_billing_dates.dt.month == month))
        open_complete_values.append(fmt(ov))
        link_complete_values.append(fmt(lv))
        open_total += ov
        link_total += lv
    complete_df = pd.DataFrame([
        ["구축", "-"] + open_complete_values + [fmt(open_total)],
        ["연계", "-"] + link_complete_values + [fmt(link_total)],
    ], columns=complete_cols)

    return {
        "status": status_df,
        "monthly": monthly_display_df,
        "monthly_base": monthly_df,
        "complete": complete_df,
        "debug_open_wait_prev_year": debug_open_wait_prev_year,
    }


def weekly_prev_friday():
    today = datetime.utcnow() + timedelta(hours=9)
    return (today - timedelta(days=(today.weekday() - 4) % 7 or 7)).date()


def save_weekly_status_snapshot(hana_df, snapshot_end_date):
    weekly_customer_status_tables(hana_df, snapshot_end_date=snapshot_end_date, save_snapshot=True)


def render_weekly_front_status_tables(hana_df, snapshot_end_date=None):
    tables = weekly_customer_status_tables(hana_df, snapshot_end_date=snapshot_end_date)
    if not tables:
        st.warning("하나은행 구글 시트 기준 고객 현황을 계산할 수 없습니다.")
        return

    prev_friday = pd.to_datetime(snapshot_end_date, errors="coerce")
    if pd.isna(prev_friday):
        prev_friday = pd.Timestamp(weekly_prev_friday())
    st.markdown(f"#### 2026년 기준 (2026.01.01 ~ {prev_friday.strftime('%Y.%m.%d')})")
    render_plain_html_table(tables["status"], max_rows=30, center_align=True, merge_cols=["구분"])
    st.markdown(f"#### ■ 26년 월별 접수고객 진행 현황 관리 (2026.01.01 ~ {prev_friday.strftime('%Y.%m.%d')})")
    render_plain_html_table(tables["monthly"], max_rows=30, center_align=True, merge_cols=["구분"])
    debug_open_wait_prev_year = tables.get("debug_open_wait_prev_year")
    if debug_open_wait_prev_year is not None and not debug_open_wait_prev_year.empty:
        with st.expander(f"2025년 구축대기 산출내역 ({len(debug_open_wait_prev_year)}건)"):
            render_plain_html_table(debug_open_wait_prev_year, max_rows=100)
    st.markdown("#### ■ 26년 월별 개설완료 현황 관리")
    render_plain_html_table(tables["complete"], max_rows=10, center_align=True, merge_cols=["구분"])


def build_weekly_report_ppt_bytes(report_df, week_start="", week_end="", hana_df=None, save_snapshot=False):
    from pptx import Presentation
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt
    from copy import deepcopy

    if not os.path.exists(WEEKLY_PPT_TEMPLATE_FILE):
        raise FileNotFoundError(f"주간보고 PPT 템플릿 파일을 찾을 수 없습니다: {WEEKLY_PPT_TEMPLATE_FILE}")

    prs = Presentation(WEEKLY_PPT_TEMPLATE_FILE)
    df = report_df.copy() if report_df is not None else pd.DataFrame()
    if df.empty:
        raise ValueError("다운로드할 주간보고 이력이 없습니다.")

    def clean(value):
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        text = str(value).strip()
        return "" if text.lower() in ["nan", "nat", "none"] else text

    def set_cell_text(cell, value, font_size=8, align=PP_ALIGN.CENTER):
        cell.text = clean(value)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = align
            for run in paragraph.runs:
                run.font.name = "맑은 고딕"
                run.font.size = Pt(font_size)

    def slide_tables(slide):
        return [shape.table for shape in slide.shapes if getattr(shape, "has_table", False)]

    def ensure_rows(table, required_rows):
        while len(table.rows) < required_rows:
            new_tr = deepcopy(table._tbl.tr_lst[-1])
            table._tbl.append(new_tr)
            for cell in table.rows[-1].cells:
                cell.text = ""

    def clear_body(table):
        for r in range(1, len(table.rows)):
            for c in range(len(table.columns)):
                set_cell_text(table.cell(r, c), "", font_size=8)

    def fill_table(table, rows, font_size=8, left_cols=None):
        left_cols = set(left_cols or [])
        required_rows = max(2, len(rows) + 1)
        ensure_rows(table, required_rows)
        clear_body(table)
        for r_idx, values in enumerate(rows, start=1):
            if r_idx >= len(table.rows):
                break
            for c_idx, value in enumerate(values[:len(table.columns)]):
                align = PP_ALIGN.LEFT if c_idx in left_cols else PP_ALIGN.CENTER
                set_cell_text(table.cell(r_idx, c_idx), value, font_size=font_size, align=align)

    def normalize_hana(df_hana):
        if df_hana is None or df_hana.empty:
            try:
                url = st.session_state.get("url_hana", DEFAULT_URL_HANA)
            except Exception:
                url = DEFAULT_URL_HANA
            df_hana = read_google_csv(url, header=2)
        df_hana = df_hana.copy()
        df_hana.columns = [str(c).strip() for c in df_hana.columns]
        return df_hana.dropna(how="all").reset_index(drop=True)

    def event_dates(series):
        return pd.to_datetime(series.map(parse_sheet_date), errors="coerce")

    def month_columns_for(table, start_col=3, end_col=None):
        end_col = end_col if end_col is not None else len(table.columns) - 1
        cols = []
        for c in range(start_col, end_col):
            label = clean(table.cell(0, c).text).replace("\n", "").replace(" ", "")
            m = re.search(r"(\d{2})\.(\d{2})", label)
            if m:
                cols.append((c, int(m.group(1)), int(m.group(2))))
        return cols

    def fmt_count(value):
        try:
            value = int(value)
        except Exception:
            value = 0
        return "-" if value == 0 else f"{value:,}"

    def fmt_delta(value):
        try:
            value = int(value)
        except Exception:
            value = 0
        if value > 0:
            return f"▲{value:,}"
        if value < 0:
            return f"▼{abs(value):,}"
        return "-"

    def count_by_month(mask, dates, year, month):
        return int((mask & dates.notna() & (dates.dt.year == year) & (dates.dt.month == month)).sum())

    def count_before_year(mask, dates, year):
        return int((mask & dates.notna() & (dates.dt.year < year)).sum())

    def week_mask(dates):
        if not week_start or not week_end:
            return pd.Series(False, index=dates.index)
        return dates.between(pd.Timestamp(week_start), pd.Timestamp(week_end), inclusive="both")

    def fill_monthly_table(table, metric_rows, year=2026, has_prev_year=True):
        start_col = 3 if has_prev_year else 2
        month_cols = month_columns_for(table, start_col=start_col, end_col=len(table.columns) - 1)
        for row_idx, metric in metric_rows.items():
            mask = metric["mask"]
            dates = metric["dates"]
            total = 0
            if has_prev_year and len(table.columns) > 2:
                prev_count = count_before_year(mask, dates, year)
                total += prev_count
                set_cell_text(table.cell(row_idx, 2), fmt_count(prev_count), font_size=8)
            for col_idx, yy, mm in month_cols:
                full_year = 2000 + yy
                count = count_by_month(mask, dates, full_year, mm)
                total += count
                set_cell_text(table.cell(row_idx, col_idx), fmt_count(count), font_size=8)
                delta_row = row_idx + 1
                if delta_row < len(table.rows):
                    delta = int((mask & week_mask(dates) & (dates.dt.year == full_year) & (dates.dt.month == mm)).sum())
                    set_cell_text(table.cell(delta_row, col_idx), fmt_delta(delta), font_size=8)
            set_cell_text(table.cell(row_idx, len(table.columns) - 1), fmt_count(total), font_size=8)

    def fill_overall_table(table, metrics):
        build_type = metrics["build_type"]
        rows = [
            ("전체", pd.Series(True, index=metrics["base"].index)),
            ("기본형", build_type.str.contains("기본", na=False)),
            ("연계형", build_type.str.contains("연계|이행", na=False)),
        ]
        keys = ["open_wait", "open_progress", "open_done", "link_wait", "link_progress", "link_done"]
        for r_idx, (_, row_mask) in enumerate(rows, start=1):
            if r_idx >= len(table.rows):
                break
            for c_idx, key in enumerate(keys, start=1):
                value = int((metrics[key] & row_mask).sum())
                set_cell_text(table.cell(r_idx, c_idx), fmt_count(value), font_size=8)

    def fill_customer_status_slides():
        try:
            hana = normalize_hana(hana_df)
        except Exception:
            return
        if hana.empty or len(prs.slides) < 4:
            return
        front_tables = weekly_customer_status_tables(hana, snapshot_end_date=week_end, save_snapshot=save_snapshot)

        open_receipt_col = find_col(hana, ["신규접수일"])
        link_receipt_col = find_col(hana, ["추가연계접수일"])
        open_date_col = find_col(hana, ["개설/이행일", "개설일", "이행일"])
        link_date_col = find_col(hana, ["연계일자", "연계일"])
        link_billing_date_col = find_col(hana, ["연계청구일자", "연계청구일", "청구일자"])
        open_status_col = find_col(hana, ["개설상태"])
        link_status_col = find_col(hana, ["연계상태"])
        build_type_col = find_col(hana, ["구축형"])
        end_col = find_col(hana, ["해지일자"])

        if not open_receipt_col or not open_status_col:
            return

        base = hana.copy()
        active = base[end_col].map(is_blank_value) if end_col else pd.Series(True, index=base.index)
        open_status = base[open_status_col].astype(str).str.strip() if open_status_col else pd.Series("", index=base.index)
        link_status = base[link_status_col].astype(str).str.strip() if link_status_col else pd.Series("", index=base.index)
        link_clean = link_status.str.replace(r"\s+", "", regex=True)
        build_type = base[build_type_col].astype(str).str.strip() if build_type_col else pd.Series("", index=base.index)

        open_receipt_dates = event_dates(base[open_receipt_col]) if open_receipt_col else pd.Series(pd.NaT, index=base.index)
        link_receipt_dates = event_dates(base[link_receipt_col]) if link_receipt_col else open_receipt_dates
        link_receipt_dates = link_receipt_dates.fillna(open_receipt_dates)
        open_done_dates = event_dates(base[open_date_col]) if open_date_col else pd.Series(pd.NaT, index=base.index)
        link_done_dates = event_dates(base[link_date_col]) if link_date_col else pd.Series(pd.NaT, index=base.index)
        link_billing_dates = event_dates(base[link_billing_date_col]) if link_billing_date_col else pd.Series(pd.NaT, index=base.index)

        open_done = active & (open_status.str.contains("완료|이행완료", na=False) | open_done_dates.notna())
        open_progress = active & ~open_done & open_status.str.contains("진행|구축중|처리중", na=False)
        open_wait = active & ~open_done & ~open_progress
        open_drop = open_status.str.contains("취소|DROP|드랍", case=False, na=False)

        link_done = active & link_status.str.contains("완료", na=False)
        link_billing_done = active & link_clean.eq("ERP청구완료")
        link_progress = active & ~link_done & link_status.str.contains("ERP연계진행|연계진행|진행", na=False)
        link_wait = active & link_status.str.contains("ERP연계대기", na=False)
        link_drop = link_status.str.contains("취소|DROP|드랍", case=False, na=False)
        link_receipt = active & (link_receipt_dates.notna() | link_wait | link_progress | link_done)

        metrics = {
            "base": base,
            "build_type": build_type,
            "type_col": build_type_col,
            "open_wait": open_wait,
            "open_progress": open_progress,
            "open_done": open_done,
            "link_wait": link_wait,
            "link_progress": link_progress,
            "link_done": link_done,
        }

        slide2_tables = slide_tables(prs.slides[1])
        if slide2_tables:
            fill_monthly_table(slide2_tables[0], {
                1: {"mask": active & open_receipt_dates.notna(), "dates": open_receipt_dates},
                3: {"mask": open_wait, "dates": open_receipt_dates},
                5: {"mask": open_progress, "dates": open_receipt_dates},
                7: {"mask": open_done, "dates": open_receipt_dates},
                9: {"mask": open_drop, "dates": open_receipt_dates},
            }, year=2026, has_prev_year=True)

        slide3_tables = slide_tables(prs.slides[2])
        if slide3_tables:
            fill_monthly_table(slide3_tables[0], {
                1: {"mask": link_receipt, "dates": link_receipt_dates},
                3: {"mask": link_wait, "dates": link_receipt_dates},
                5: {"mask": link_progress, "dates": link_receipt_dates},
                7: {"mask": link_done, "dates": link_receipt_dates},
                9: {"mask": link_drop, "dates": link_receipt_dates},
            }, year=2026, has_prev_year=True)

        slide4_tables = slide_tables(prs.slides[3])
        if slide4_tables:
            fill_monthly_table(slide4_tables[0], {
                1: {"mask": open_done, "dates": open_done_dates},
                3: {"mask": link_billing_done, "dates": link_billing_dates},
            }, year=2026, has_prev_year=False)
            if len(slide4_tables) > 1:
                fill_overall_table(slide4_tables[1], metrics)

    def update_title_dates():
        date_text = ""
        if week_start and week_end:
            date_text = f"{week_start} - {week_end}"
        elif "보고시작일" in df.columns and "보고종료일" in df.columns:
            starts = [clean(v) for v in df["보고시작일"].tolist() if clean(v)]
            ends = [clean(v) for v in df["보고종료일"].tolist() if clean(v)]
            if starts and ends:
                date_text = f"{min(starts)} - {max(ends)}"
        if not date_text or not prs.slides:
            return
        for shape in prs.slides[0].shapes:
            if getattr(shape, "has_text_frame", False) and "2026." in shape.text:
                shape.text = date_text
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "맑은 고딕"
                        run.font.size = Pt(18)
                break

    def rows_for(category, no_start=1):
        part = df[df["카테고리"].astype(str) == category].copy() if "카테고리" in df.columns else pd.DataFrame()
        rows = []
        for no, (_, row) in enumerate(part.iterrows(), start=no_start):
            service = clean(row.get("서비스", "")) or "일반"
            build_type = clean(row.get("구분", ""))
            company = clean(row.get("고객명", ""))
            branch = clean(row.get("신청점", ""))
            receipt = clean(row.get("접수일자", ""))
            issue = clean(row.get("이슈", ""))
            note = clean(row.get("특이사항", ""))
            detail = issue if not note else f"{issue}\n{note}" if issue else note
            staff = clean(row.get("담당자", ""))
            date_value = clean(row.get("일자", ""))
            status = clean(row.get("상태", ""))
            rows.append({
                "no": no,
                "service": service,
                "type": build_type,
                "company": company,
                "branch": branch,
                "receipt": receipt,
                "issue": detail,
                "staff": staff,
                "date": date_value,
                "status": status,
            })
        return rows

    def opening_wait_rows(rows):
        return [[r["no"], r["service"], r["type"], r["company"], r["branch"], r["receipt"], r["date"], r["issue"], r["staff"]] for r in rows]

    def opening_progress_rows(rows):
        return [[r["no"], r["type"], r["company"], r["branch"], r["receipt"], r["issue"], r["staff"], r["date"], r["status"]] for r in rows]

    def link_wait_rows(rows):
        return [[r["no"], r["company"], r["branch"], r["receipt"], r["date"], r["issue"], r["staff"]] for r in rows]

    def link_progress_rows(rows):
        return [[r["no"], r["company"], r["branch"], r["receipt"], "", r["status"], r["issue"], r["staff"], r["date"]] for r in rows]

    def operation_rows(rows):
        return [[r["no"], r["type"], r["company"], r["branch"], r["receipt"], r["issue"], r["staff"], r["date"], r["status"]] for r in rows]

    def fill_slide_table(slide_idx, values, left_cols):
        if slide_idx >= len(prs.slides):
            return
        tables = slide_tables(prs.slides[slide_idx])
        if tables:
            fill_table(tables[0], values, font_size=7, left_cols=left_cols)

    update_title_dates()
    for slide_idx in range(4, min(24, len(prs.slides))):
        for table in slide_tables(prs.slides[slide_idx]):
            clear_body(table)

    open_wait = rows_for("개설대기")
    open_progress = rows_for("개설진행")
    open_done = rows_for("개설완료")
    link_wait = rows_for("연계대기")
    link_progress = rows_for("연계진행")
    link_done = rows_for("연계완료")
    operation = rows_for("운영부문")

    fill_slide_table(4, opening_wait_rows(open_wait), left_cols={6, 7})
    fill_slide_table(7, opening_progress_rows(open_done), left_cols={5})
    fill_slide_table(9, opening_progress_rows(open_progress), left_cols={5})
    fill_slide_table(10, link_wait_rows(link_wait), left_cols={5})
    fill_slide_table(14, link_progress_rows(link_progress), left_cols={6})
    fill_slide_table(19, link_progress_rows(link_done), left_cols={6})
    fill_slide_table(21, operation_rows(operation), left_cols={5})

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


OPERATION_PLAN_FILE = "operation_plan.json"


def show_operation_plan():
    st.markdown("### 운영계획")
    tab_kpi_rec, tab_no_login = st.tabs(["KPI 집중 추천", "미로그인 고객 현황"])

    # ── 탭1: KPI 집중 추천 ──
    with tab_kpi_rec:
        _force_kpi = st.session_state.pop("_op_kpi_refresh", False)
        refresh_col, _ = st.columns([1, 5])
        with refresh_col:
            if st.button("새로고침", key="op_kpi_refresh_btn", use_container_width=True):
                st.session_state["_op_kpi_refresh"] = True
                st.rerun()
        try:
            hana_df_kpi = read_google_csv(
                st.session_state.get("url_hana", DEFAULT_URL_HANA),
                header=2, force_refresh=_force_kpi,
            )
            billing_df_kpi = load_hana_billing_df(force_refresh=_force_kpi)
        except Exception as e:
            st.error(f"데이터를 불러오지 못했습니다: {e}")
            hana_df_kpi = pd.DataFrame()
            billing_df_kpi = pd.DataFrame()

        if hana_df_kpi is None or hana_df_kpi.empty or billing_df_kpi is None or billing_df_kpi.empty:
            st.info("KPI 추천에 사용할 하나은행 구글시트 또는 청구시트 데이터가 없습니다.")
        else:
            hana_df_kpi = hana_df_kpi.dropna(how="all").reset_index(drop=True)
            billing_df_kpi = billing_df_kpi.dropna(how="all").reset_index(drop=True)
            render_kpi_activity_recommendations(
                hana_df_kpi,
                billing_df_kpi,
                user_name=None,
                key_prefix="op_kpi_rec",
            )

    # ── 탭2: 미로그인 고객 현황 ──
    with tab_no_login:
        _force_nl = st.session_state.pop("_nl_rf_op_nl_refresh", False)
        try:
            hana_df_nl    = read_google_csv(
                st.session_state.get("url_hana", DEFAULT_URL_HANA),
                header=2, force_refresh=_force_nl,
            )
            billing_df_nl = load_hana_billing_df(force_refresh=_force_nl)
        except Exception as e:
            st.error(f"데이터를 불러오지 못했습니다: {e}")
            return
        if hana_df_nl is None or hana_df_nl.empty or billing_df_nl is None or billing_df_nl.empty:
            st.info("구글시트 또는 청구시트 데이터가 없습니다.")
        else:
            hana_df_nl    = hana_df_nl.dropna(how="all").reset_index(drop=True)
            billing_df_nl = billing_df_nl.dropna(how="all").reset_index(drop=True)

            st.markdown("#### 개설/이행일 이후 미로그인 고객")
            st.caption("개설/이행일 기준으로 청구시트 최종로그인일자가 없거나 개설/이행일 이전인 고객")
            _render_no_login_section(
                build_no_login_after_open(hana_df_nl, billing_df_nl),
                year_key="no_login_open_year", owner_key="no_login_open_owner",
                label="미로그인 고객", download_prefix="미로그인고객_개설이행일",
                refresh_key="op_nl_refresh", exclude_key="op_nl_open_exclude",
            )

            st.divider()
            st.markdown("#### 연계청구일자 이후 미로그인 고객")
            st.caption("연계청구일자 기준으로 청구시트 최종로그인일자가 없거나 연계청구일자 이전인 고객")
            _render_no_login_section(
                build_no_login_after_link_billing(hana_df_nl, billing_df_nl),
                year_key="no_login_link_year", owner_key="no_login_link_owner",
                label="연계청구 미로그인 고객", download_prefix="미로그인고객_연계청구일자",
                exclude_key="op_nl_link_exclude",
            )

            st.divider()
            st.markdown("#### 로그인 100회 이상 · 미이체 고객")
            st.caption("청구시트 기준 로그인 100회 이상이지만 이체 이력이 없는 고객 (해지 제외)")
            _render_high_login_no_transfer(
                build_high_login_no_transfer(hana_df_nl, billing_df_nl),
                owner_key="op_high_login_owner",
                download_prefix="미이체고객",
            )


def show_weekly_report_admin():
    db = load_db(WEEKLY_REPORT_FILE, {})
    rows = []
    for entries in db.values():
        rows.extend(entries if isinstance(entries, list) else [])
    title_col, refresh_col = st.columns([5, 1])
    with title_col:
        st.markdown("### 주간보고 취합")
    with refresh_col:
        if st.button("새로고침", key="weekly_admin_refresh", use_container_width=True):
            try:
                st.session_state.hana_sheet_df = read_google_csv(st.session_state.get("url_hana", DEFAULT_URL_HANA), header=2, force_refresh=True)
                st.session_state.weekly_hana_loaded_at = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M:%S")
                st.success("새로고침 완료")
            except Exception as e:
                st.warning(f"새로고침 실패: {e}")
            st.rerun()
    raw_df_for_period = pd.DataFrame(rows) if rows else pd.DataFrame()
    default_snapshot_end = str(weekly_prev_friday())
    if "보고종료일" in raw_df_for_period.columns and not raw_df_for_period.empty:
        ends = [str(v) for v in raw_df_for_period["보고종료일"].dropna().tolist() if str(v).strip()]
        if ends:
            default_snapshot_end = max(ends)
    hana_status_df = load_weekly_hana_for_status()
    render_weekly_front_status_tables(hana_status_df, default_snapshot_end)
    st.markdown("---")

    if not rows:
        st.info("취합할 주간보고 이력이 없습니다.")
        return
    df = pd.DataFrame(rows)
    categories = ["전체", "개설대기", "개설진행", "개설완료", "연계대기", "연계진행", "연계완료", "운영부문"]
    c1, c2 = st.columns(2)
    with c1:
        category = st.selectbox("카테고리", categories, key="weekly_admin_category")
    with c2:
        staff_options = ["전체"] + sorted(df["담당자"].dropna().astype(str).unique().tolist()) if "담당자" in df.columns else ["전체"]
        staff = st.selectbox("담당자", staff_options, key="weekly_admin_staff")
    if category != "전체" and "카테고리" in df.columns:
        df = df[df["카테고리"] == category]
    if staff != "전체" and "담당자" in df.columns:
        df = df[df["담당자"] == staff]
    display_df = df.copy()
    if "일자" in display_df.columns:
        display_df["구축/피드백 예정일"] = display_df["일자"]
    show_cols = ["보고시작일", "보고종료일", "카테고리", "구분", "고객명", "접수일자", "이슈", "상태", "담당자", "구축/피드백 예정일", "작성시각"]
    show_cols = [c for c in show_cols if c in display_df.columns]
    render_plain_html_table(
        display_df[show_cols].sort_values("작성시각", ascending=False),
        max_rows=1000,
        center_align=False,
        merge_cols=["보고시작일", "보고종료일", "카테고리", "구분"],
    )

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    try:
        hana_for_ppt = None
        try:
            hana_for_ppt = read_google_csv(st.session_state.get("url_hana", DEFAULT_URL_HANA), header=2)
        except Exception:
            hana_for_ppt = st.session_state.get("hana_sheet_df")
        week_start = ""
        week_end = ""
        if "보고시작일" in df.columns and "보고종료일" in df.columns and not df.empty:
            starts = [str(v) for v in df["보고시작일"].dropna().tolist() if str(v).strip()]
            ends = [str(v) for v in df["보고종료일"].dropna().tolist() if str(v).strip()]
            week_start = min(starts) if starts else ""
            week_end = max(ends) if ends else ""
        ppt_bytes = build_weekly_report_ppt_bytes(df, week_start, week_end, hana_for_ppt)
        file_period = f"{week_start}_{week_end}".replace("-", "") if week_start and week_end else (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m%d")
        st.download_button(
            "주간보고 PPT 다운로드",
            data=ppt_bytes,
            file_name=f"주간보고_통합CMS고객_개설운영_주간보고_{file_period}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            on_click=save_weekly_status_snapshot,
            args=(hana_for_ppt, week_end or default_snapshot_end),
        )
    except Exception as e:
        st.button("주간보고 PPT 다운로드", use_container_width=True, disabled=True)
        st.caption(f"PPT 생성 준비 중 오류: {e}")


def show_dashboard():
    import plotly.graph_objects as go

    # ── 대시보드 진입 시 캐시된 구글시트 데이터를 재사용 ─────────────
    try:
        with st.spinner("데이터 불러오는 중..."):
            hana_raw = read_google_csv(st.session_state.get("url_hana", DEFAULT_URL_HANA), header=2)
            st.session_state.hana_sheet_df = hana_raw
            raw_act = read_google_csv(st.session_state.get("url_analysis", DEFAULT_URL_ANALYSIS))
            st.session_state.analysis_lookup_df = raw_act
            billing_raw = read_google_csv(st.session_state.get("url_hana_billing", DEFAULT_URL_HANA_BILLING))
            billing_raw = billing_raw.dropna(how="all").reset_index(drop=True)
            st.session_state.hana_billing_df = billing_raw
    except Exception as e:
        st.warning(f"데이터 로드 실패: {e}")

    df_hana = st.session_state.get("hana_sheet_df")
    user_name = st.session_state.user_name

    if df_hana is None or df_hana.empty:
        st.info("📂 데이터를 불러올 수 없습니다. [구글 스트레드시트 연동] 메뉴에서 URL을 확인해주세요.")
        return

    # ── 하나은행 데이터 전처리 (당월용) ───────────────
    df_hana = df_hana.copy()
    u_col = "담당자"
    gaeseol_date_col = "개설/이행일"
    yeonge_date_col = "연계일자"
    gubun_col = "구축구분"

    if u_col not in df_hana.columns or gaeseol_date_col not in df_hana.columns:
        st.error("데이터 컬럼을 찾을 수 없습니다. (담당자, 개설/이행일 필요)")
        return

    df_hana[gaeseol_date_col] = pd.to_datetime(
        df_hana[gaeseol_date_col].astype(str).str.strip().str[:8], format="%Y%m%d", errors="coerce"
    )
    if yeonge_date_col in df_hana.columns:
        df_hana[yeonge_date_col] = pd.to_datetime(
            df_hana[yeonge_date_col].astype(str).str.strip().str[:8], format="%Y%m%d", errors="coerce"
        )
    df_user_hana = df_hana[df_hana[u_col].astype(str).str.strip() == user_name].copy()

    # ── 하나지사 활동이력 전처리 (전월/전년동월용) ────
    df_user_act = None
    df_act = None
    act_u_col = None
    act_date_col = "활동일"
    act_d_col = "활동상세"
    df_activity = st.session_state.get("analysis_lookup_df")
    if df_activity is not None and not df_activity.empty:
        df_act = df_activity.copy()
        df_act.columns = [str(c).strip() for c in df_act.columns]
        act_u_col = next((c for c in df_act.columns if c.strip() in ["등록자", "담당자", "성명"]), None)
        act_date_col = next((c for c in df_act.columns if "활동일" in c.replace(" ", "") or "일자" in c.replace(" ", "")), None)
        act_d_col = next((c for c in df_act.columns if c.strip() in ["활동상세", "활동내용"]), None)
        if act_u_col and act_date_col and act_d_col:
            df_act[act_date_col] = pd.to_datetime(df_act[act_date_col], errors="coerce")
            df_user_act = df_act[df_act[act_u_col].astype(str).str.strip() == user_name].dropna(subset=[act_date_col]).copy()

    now = datetime.utcnow() + timedelta(hours=9)
    curr_ym = now.strftime("%Y-%m")
    prev_dt = (now.replace(day=1) - timedelta(days=1))
    prev_ym = prev_dt.strftime("%Y-%m")
    prev_year_ym = f"{now.year - 1}-{now.month:02d}"

    # ── 공통 계산 헬퍼 (유저 이름 인자) ──────────────────
    def calc_hana_for(uname, ym):
        df_u = df_hana[df_hana[u_col].astype(str).str.strip() == uname]
        gdf  = df_u[df_u[gaeseol_date_col].dt.strftime("%Y-%m") == ym]
        o = int((gdf[gubun_col].astype(str).str.strip() == "신규").sum()) if gubun_col in gdf.columns else len(gdf)
        ldf = df_u[df_u[yeonge_date_col].dt.strftime("%Y-%m") == ym] if yeonge_date_col in df_u.columns else df_u.iloc[0:0]
        l   = len(ldf)
        v = int((gdf[gubun_col].astype(str).str.strip() == "이행").sum()) if gubun_col in gdf.columns else 0
        o_p, l_p, v_p = o * 90, l * 120, v * 30
        total = min(2800, min(1000, o_p + l_p) + min(1800, v_p))
        return {"개설건수": o, "연계건수": l, "운영건수": v, "개설포인트": o_p, "연계포인트": l_p, "운영포인트": v_p, "합계포인트": total}

    def calc_act_for(uname, ym):
        empty = {"개설건수": 0, "연계건수": 0, "운영건수": 0, "합계포인트": 0}
        try:
            if df_activity is None or act_date_col is None or act_d_col is None or act_u_col is None:
                return empty
            df_u = df_act[df_act[act_u_col].astype(str).str.strip() == uname]
            df_m = df_u[df_u[act_date_col].dt.strftime("%Y-%m") == ym]
            if df_m.empty:
                return empty
            o = int(df_m[act_d_col].astype(str).str.contains("개설", na=False).sum())
            l = int(df_m[act_d_col].astype(str).str.contains("연계", na=False).sum())
            v = int(df_m[act_d_col].astype(str).str.contains("운영|방문|점검", na=False).sum())
            o_p, l_p, v_p = o * 90, l * 120, v * 30
            total = min(2800, min(1000, o_p + l_p) + min(1800, v_p))
            return {"개설건수": o, "연계건수": l, "운영건수": v, "합계포인트": total}
        except Exception:
            return empty

    def filter_month_gaeseol(df, ym):
        return df[df[gaeseol_date_col].dt.strftime("%Y-%m") == ym].copy()

    def filter_month_yeonge(df, ym):
        if yeonge_date_col not in df.columns:
            return df.iloc[0:0]
        return df[df[yeonge_date_col].dt.strftime("%Y-%m") == ym].copy()

    def calc_points_hana(ym):
        return calc_hana_for(user_name, ym)

    def calc_points_activity(ym):
        empty = {"개설건수": 0, "연계건수": 0, "운영건수": 0, "개설포인트": 0, "연계포인트": 0, "운영포인트": 0, "합계포인트": 0}
        try:
            if df_user_act is None or df_user_act.empty or act_date_col is None or act_d_col is None:
                return empty
            df_m = df_user_act[df_user_act[act_date_col].dt.strftime("%Y-%m") == ym].copy()
            if df_m.empty:
                return empty
            o = int(df_m[act_d_col].astype(str).str.contains("개설", na=False).sum())
            l = int(df_m[act_d_col].astype(str).str.contains("연계", na=False).sum())
            v = int(df_m[act_d_col].astype(str).str.contains("운영|방문|점검", na=False).sum())
            o_p, l_p, v_p = o * 90, l * 120, v * 30
            total = min(2800, min(1000, o_p + l_p) + min(1800, v_p))
            return {"개설건수": o, "연계건수": l, "운영건수": v, "개설포인트": o_p, "연계포인트": l_p, "운영포인트": v_p, "합계포인트": total}
        except Exception:
            return empty

    # ── 관리자: 전체 직원 현황 2열 그리드 ────────────────
    if st.session_state.user_role == "관리자":
        st.markdown(f"### 전체 직원 {curr_ym} 실적 현황")
        staff_items = [
            (info.get("name", ""), info.get("rank", "직원"))
            for uid, info in st.session_state.user_db.items()
            if uid != "1"
            and info.get("name")
            and info.get("access") == "허용"
            and info.get("rank") != "부서장"
            and info.get("dept_type", "사업부") == "C&S"
            and info.get("staff_type", "정규직") != "파견직"
        ]
        staff_items = sorted(staff_items, key=lambda item: (_RANK_ORDER.get(item[1], 9), item[0]))
        cards = []
        for uname, _rank in staff_items:
            c = calc_hana_for(uname, curr_ym)
            p = calc_act_for(uname, prev_ym)
            delta = c["합계포인트"] - p["합계포인트"]
            cards.append((uname, c, p, delta))

        for i in range(0, len(cards), 2):
            cols = st.columns(2)
            for j, col in enumerate(cols):
                if i + j >= len(cards):
                    break
                uname, c, p, delta = cards[i + j]
                with col:
                    delta_cls = "pms-delta-up" if delta >= 0 else "pms-delta-down"
                    delta_sym = "▲" if delta >= 0 else "▼"
                    st.markdown(f"""
                    <div class="pms-staff-card">
                        <div class="pms-card-name">👤 {uname}</div>
                        <div class="pms-card-stats">
                            <span>개설 <b>{c['개설건수']}건</b></span>
                            <span>연계 <b>{c['연계건수']}건</b></span>
                            <span>이행 <b>{c['운영건수']}건</b></span>
                        </div>
                        <div class="pms-card-points">
                            {c['합계포인트']:,} pt
                            <span class="pms-delta {delta_cls}">{delta_sym} {abs(delta):,}pt (전월비)</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown(f"### {user_name}님 상세 현황")

    curr = calc_points_hana(curr_ym)
    prev = calc_points_activity(prev_ym)
    py   = calc_points_activity(prev_year_ym)

    diff_prev = curr["합계포인트"] - prev["합계포인트"]
    diff_year = curr["합계포인트"] - py["합계포인트"]
    max_add = max(0, 2800 - curr["합계포인트"])

    # ── 요약 카드 ──────────────────────────────────────
    st.markdown(f"### {user_name}님의 {curr_ym} 실적 현황")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("이번달 추정 포인트", f"{curr['합계포인트']:,} pt")
    c2.metric("전월 대비", f"{prev['합계포인트']:,} pt", delta=f"{diff_prev:+,} pt")
    c3.metric("전년 동월 대비", f"{py['합계포인트']:,} pt", delta=f"{diff_year:+,} pt")
    c4.metric("최대 추가 가능", f"{max_add:,} pt")

    # ── 이번달 고객사 개설/운영 현황 ──────────────────────
    curr_g = filter_month_gaeseol(df_user_hana, curr_ym)
    curr_y = filter_month_yeonge(df_user_hana, curr_ym)
    if not curr_g.empty or not curr_y.empty:
        st.markdown(f"**{curr_ym} 고객사 진행 현황**")
        _comp_col = next((c for c in df_user_hana.columns if any(k in c for k in ["업체명", "고객명", "상호"])), None)
        _dg, _dy = st.columns(2)
        with _dg:
            st.caption(f"개설/이행 ({len(curr_g)}건)")
            if not curr_g.empty:
                _dcols = [c for c in [_comp_col, gubun_col, gaeseol_date_col] if c and c in curr_g.columns]
                _show = curr_g[_dcols].copy() if _dcols else curr_g.iloc[:, :3].copy()
                _show[gaeseol_date_col] = _show[gaeseol_date_col].dt.strftime("%Y-%m-%d") if gaeseol_date_col in _show.columns else _show.get(gaeseol_date_col, "")
                st.dataframe(_show.reset_index(drop=True), use_container_width=True, hide_index=True)
            else:
                st.info("이번달 개설/이행 없음")
        with _dy:
            st.caption(f"연계 ({len(curr_y)}건)")
            if not curr_y.empty:
                _dcols2 = [c for c in [_comp_col, yeonge_date_col] if c and c in curr_y.columns]
                _show2 = curr_y[_dcols2].copy() if _dcols2 else curr_y.iloc[:, :2].copy()
                if yeonge_date_col in _show2.columns:
                    _show2[yeonge_date_col] = _show2[yeonge_date_col].dt.strftime("%Y-%m-%d")
                st.dataframe(_show2.reset_index(drop=True), use_container_width=True, hide_index=True)
            else:
                st.info("이번달 연계 없음")

    st.markdown("---")

    # ── 활동 유형별 포인트 차트 ────────────────────────
    col_l, col_r = st.columns(2)

    with col_l:
        st.markdown("**활동 유형별 포인트 비교 (당월 / 전월 / 전년 동월)**")
        categories = ["개설포인트", "연계포인트", "운영포인트"]
        labels = ["개설", "연계", "운영"]
        fig = go.Figure(data=[
            go.Bar(name=f"당월 ({curr_ym})",   x=labels, y=[curr[c] for c in categories], marker_color="#6366f1"),
            go.Bar(name=f"전월 ({prev_ym})",   x=labels, y=[prev[c] for c in categories], marker_color="#a78bfa"),
            go.Bar(name=f"전년 동월 ({prev_year_ym})", x=labels, y=[py[c]   for c in categories], marker_color="#94a3b8"),
        ])
        _txt0 = "#1e293b"
        fig.update_layout(**_chart_layout(height=300, barmode="group",
                          legend=dict(orientation="h", y=-0.3, font=dict(color=_txt0))))
        st.plotly_chart(fig, use_container_width=True, theme=None)

    with col_r:
        st.markdown("**이번달 포인트 구성**")
        pie_labels = ["개설포인트", "연계포인트", "운영포인트"]
        pie_values = [curr["개설포인트"], curr["연계포인트"], curr["운영포인트"]]
        if sum(pie_values) > 0:
            _txt  = "#1e293b"
            fig2 = go.Figure(go.Pie(
                labels=["개설", "연계", "운영"],
                values=pie_values,
                hole=0.4,
                marker_colors=["#6366f1", "#7C3AED", "#a78bfa"],
                textfont=dict(color=_txt),
            ))
            fig2.update_layout(**_chart_layout(height=300,
                legend=dict(font=dict(color=_txt))))
            st.plotly_chart(fig2, use_container_width=True, theme=None)
        else:
            st.info("이번달 집계된 활동이 없습니다.")

    st.markdown("---")

    # ── 일별 개설 추이 ─────────────────────────────────
    curr_df = filter_month_gaeseol(df_user_hana, curr_ym)
    if not curr_df.empty:
        st.markdown("**이번달 일별 개설 건수 추이**")
        daily = curr_df.groupby(curr_df[gaeseol_date_col].dt.strftime("%Y-%m-%d")).size().reset_index()
        daily.columns = ["날짜", "건수"]
        fig3 = go.Figure(go.Scatter(x=daily["날짜"], y=daily["건수"], mode="lines+markers",
                                    line=dict(color="#6366f1", width=2), marker=dict(size=7, color="#a78bfa")))
        fig3.update_layout(**_chart_layout(height=220, xaxis_title="날짜", yaxis_title="건수"))
        st.plotly_chart(fig3, use_container_width=True, theme=None)

    st.markdown("---")

    # ── 전년도 / 올해 업무별 월평균 ──────────────────────
    if df_user_act is not None and not df_user_act.empty:
        st.markdown("**📊 업무별 월평균 비교**")

        this_year  = str(now.year)
        last_year  = str(now.year - 1)

        def monthly_avg(year_str):
            df_y = df_user_act[df_user_act[act_date_col].dt.strftime("%Y") == year_str].copy()
            if df_y.empty:
                return {"개설": 0.0, "연계": 0.0, "운영": 0.0}
            months = df_y[act_date_col].dt.strftime("%Y-%m").unique()
            totals = {"개설": 0, "연계": 0, "운영": 0}
            for ym in months:
                dm = df_y[df_y[act_date_col].dt.strftime("%Y-%m") == ym]
                totals["개설"] += int(dm[act_d_col].astype(str).str.contains("개설", na=False).sum())
                totals["연계"] += int(dm[act_d_col].astype(str).str.contains("연계", na=False).sum())
                totals["운영"] += int(dm[act_d_col].astype(str).str.contains("운영|방문|점검", na=False).sum())
            n = len(months)
            return {k: round(v / n, 1) for k, v in totals.items()}

        avg_last = monthly_avg(last_year)
        avg_this = monthly_avg(this_year)

        avg_labels = ["개설", "연계", "운영"]
        _txt2  = "#1e293b"
        fig_avg = go.Figure(data=[
            go.Bar(name=f"{last_year}년 월평균", x=avg_labels,
                   y=[avg_last[k] for k in avg_labels], marker_color="#a78bfa"),
            go.Bar(name=f"{this_year}년 월평균", x=avg_labels,
                   y=[avg_this[k] for k in avg_labels], marker_color="#6366f1"),
        ])
        fig_avg.update_layout(**_chart_layout(height=260, barmode="group",
                              legend=dict(orientation="h", y=-0.3, font=dict(color=_txt2))))
        st.plotly_chart(fig_avg, use_container_width=True, theme=None)

        ca, cb, cc = st.columns(3)
        for col_ui, key in zip([ca, cb, cc], avg_labels):
            col_ui.metric(
                f"{key} 월평균",
                f"{avg_this[key]}건 (올해)",
                delta=f"{avg_this[key] - avg_last[key]:+.1f}건 vs {last_year}년",
            )

    st.markdown("---")

    # ── 실적 보완 가이드 ───────────────────────────────
    st.markdown("**💡 실적 보완 가이드**")
    ol_used = min(1000, curr["개설포인트"] + curr["연계포인트"])
    ol_remain = max(0, 1000 - (curr["개설포인트"] + curr["연계포인트"]))
    v_remain = max(0, 1800 - curr["운영포인트"])

    guides = []
    if ol_remain > 0:
        add_o = ol_remain // 90
        add_l = ol_remain // 120
        guides.append(f"📌 개설 추가 시 최대 **{add_o}건** ({ol_remain:,}pt 확보 가능, 건당 90pt)")
        guides.append(f"📌 연계 추가 시 최대 **{add_l}건** ({ol_remain:,}pt 확보 가능, 건당 120pt)")
    else:
        guides.append("✅ 개설·연계 포인트 한도(1,000pt) 달성!")

    if v_remain > 0:
        add_v = v_remain // 30
        guides.append(f"📌 운영·방문 추가 시 최대 **{add_v}건** ({v_remain:,}pt 확보 가능, 건당 30pt)")
    else:
        guides.append("✅ 운영 포인트 한도(1,800pt) 달성!")

    if max_add == 0:
        guides.append("🎉 최대 포인트(2,800pt) 달성! 수고하셨습니다.")

    for g in guides:
        st.markdown(f"- {g}")

    if ol_remain > 0:
        st.markdown("")
        st.info(
            "💡 개설·연계 포인트 확보가 어렵다면 **교차판매·활성화·VOC 활동**을 통해 부족한 포인트를 채울 수 있습니다.\n\n"
            "- **교차판매**: 타겟고객 선별·메일발송·방문설명회 등 영업활동 — 건당 2\~60pt, 월 최대 100\~제한없음\n"
            "- **활성화**: 이체·집금 활성화, 조회업무 활성화 — 건당 10\~30pt, 월 최대 100\~150pt\n"
            "- **VOC**: 고객 개선 아이디어 제출 — 건당 10pt, 월 최대 50pt\n\n"
            "위 활동들은 개설·연계 실적 없이도 단독으로 포인트를 적립할 수 있는 가장 현실적인 방법입니다."
        )
        st.info("💡 위 활동들은 **[업로드 및 실적 확인] 메뉴 → 추가 실적 입력** 표에서 건수를 직접 입력하여 포인트를 적립할 수 있습니다.")

def load_hana_billing_df(force_refresh=False):
    billing_raw = read_google_csv(
        st.session_state.get("url_hana_billing", DEFAULT_URL_HANA_BILLING),
        force_refresh=force_refresh,
    )
    billing_raw = billing_raw.dropna(how="all").reset_index(drop=True)
    st.session_state.hana_billing_df = billing_raw
    return billing_raw


def normalize_billing_customer_no(value):
    if is_blank_value(value):
        return ""
    text = str(value).strip()
    if re.match(r"^\d+\.0$", text):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    return digits.zfill(9)


def normalize_biz_no(value):
    if is_blank_value(value):
        return ""
    text = str(value).strip()
    if re.match(r"^\d+\.0$", text):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits.zfill(10) if digits else ""


def format_yyyymmdd(value):
    parsed = parse_sheet_date(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%Y%m%d")


def series_yyyymm(series):
    parsed = pd.to_datetime(series.apply(parse_sheet_date), errors="coerce")
    return parsed.dt.strftime("%Y-%m")


def exact_col(df, names):
    normalized = {str(c).replace(" ", "").replace("\n", "").replace("\r", ""): c for c in df.columns}
    for name in names:
        found = normalized.get(str(name).replace(" ", "").replace("\n", "").replace("\r", ""))
        if found is not None:
            return found
    return None


def build_billing_lookup(billing_df):
    lookup = billing_df.copy()
    lookup.columns = [str(c).strip() for c in lookup.columns]

    customer_col = find_col(lookup, ["고객번호", "고객No", "고객no"])
    company_col = find_col(lookup, ["고객명", "업체명", "상호"])
    biz_col = find_col(lookup, ["사업자번호", "사업자등록번호"])
    first_login_col = exact_col(lookup, ["신규일자"]) or find_col(lookup, ["신규일자", "최초신규일자"])
    # "최종로그인" 단독 컬럼명도 커버 (일자 없는 변형 포함)
    last_login_col = find_col(lookup, ["최종로그인일자", "최근로그인일자", "최종로그인", "최근로그인", "로그인일자"])
    login_count_col  = find_col(lookup, ["로그인건수", "로그인횟수"])
    menu_click_col   = exact_col(lookup, ["메뉴사용"]) or find_col(lookup, ["메뉴사용", "메뉴클릭수"])
    transfer_col     = find_col(lookup, ["최종이체일자", "이체일자", "최종이체"])
    transfer_cnt_col = find_col(lookup, ["이체건수", "이체횟수"])

    if not customer_col:
        return pd.DataFrame(columns=["_고객번호", "_은행고객명", "_은행사업자번호", "_최초로그인", "_최종로그인", "_로그인횟수", "_메뉴클릭수", "_최종이체일자", "_이체건수"])

    rows = pd.DataFrame()
    rows["_고객번호"]     = lookup[customer_col].apply(normalize_billing_customer_no)
    rows["_은행고객명"]   = lookup[company_col].astype(str).str.strip() if company_col else ""
    rows["_은행사업자번호"] = lookup[biz_col].apply(normalize_biz_no) if biz_col else ""
    rows["_최초로그인"]   = lookup[first_login_col].apply(format_yyyymmdd) if first_login_col else ""
    rows["_최종로그인"]   = lookup[last_login_col].apply(format_yyyymmdd) if last_login_col else ""
    rows["_로그인횟수"]   = lookup[login_count_col].fillna("").astype(str).str.replace(r"\.0$", "", regex=True) if login_count_col else ""
    rows["_메뉴클릭수"]   = lookup[menu_click_col].fillna("").astype(str).str.replace(r"\.0$", "", regex=True) if menu_click_col else ""
    rows["_최종이체일자"] = lookup[transfer_col].apply(format_yyyymmdd) if transfer_col else ""
    rows["_이체건수"]     = lookup[transfer_cnt_col].fillna("").astype(str).str.replace(r"\.0$", "", regex=True) if transfer_cnt_col else ""
    rows = rows[rows["_고객번호"] != ""].drop_duplicates("_고객번호", keep="first")
    return rows


def add_bank_compare_columns(source_df, billing_lookup, source_company_col):
    result = source_df.merge(billing_lookup, on="_고객번호", how="left")
    result["청구시트 여부"] = np.where(result.get("_은행고객명", "").fillna("").astype(str).str.strip().ne(""), "있음", "없음")
    result["최초로그인"] = result.get("_최초로그인", "").fillna("")
    result["사업자번호_은행"] = result.get("_은행사업자번호", "").fillna("")
    result["최종로그인"] = result.get("_최종로그인", "").fillna("")
    result["로그인횟수"] = result.get("_로그인횟수", "").fillna("")
    result["메뉴클릭수"] = result.get("_메뉴클릭수", "").fillna("")
    result["청구원본 고객명"] = result[source_company_col].fillna("").astype(str) if source_company_col in result.columns else ""
    result["실적파일 고객명"] = result.get("_은행고객명", "").fillna("")
    return result


def build_open_billing_status(hana_df, billing_lookup, selected_month):
    hana = hana_df.copy()
    hana.columns = [str(c).strip() for c in hana.columns]

    customer_col = find_col(hana, ["고객번호"])
    biz_col = find_col(hana, ["사업자번호"])
    company_col = find_col(hana, ["고객명", "업체명", "상호"])
    build_type_col = find_col(hana, ["구축형"])
    receipt_col = find_col(hana, ["신규접수일"])
    open_date_col = find_col(hana, ["개설/이행일", "개설일", "이행일"])
    owner_col = find_col(hana, ["담당자"])
    manage_col = find_col(hana, ["관리구분"])
    open_status_col = find_col(hana, ["개설상태"])

    if not customer_col or not open_date_col:
        return pd.DataFrame()

    hana["_고객번호"] = hana[customer_col].apply(normalize_billing_customer_no)
    hana["_기준월"] = series_yyyymm(hana[open_date_col])
    mask = hana["_고객번호"].ne("") & hana["_기준월"].eq(selected_month)
    if manage_col and manage_col in hana.columns:
        mask &= ~hana[manage_col].astype(str).str.strip().isin(["해지", "취소"])
    if open_status_col and open_status_col in hana.columns:
        mask &= ~hana[open_status_col].astype(str).str.contains("취소|반려", na=False)

    base = hana[mask].copy().reset_index(drop=True)
    if base.empty:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["순번"] = range(1, len(base) + 1)
    out["고객번호"] = base["_고객번호"]
    out["사업자번호"] = base[biz_col].apply(normalize_biz_no) if biz_col else ""
    out["업체명"] = base[company_col].fillna("").astype(str) if company_col else ""
    out["ERP연계 여부"] = base[build_type_col].fillna("").astype(str) if build_type_col else ""
    out["접수일자"] = base[receipt_col].apply(format_yyyymmdd) if receipt_col else ""
    out["구축일자"] = base[open_date_col].apply(format_yyyymmdd)
    out["방문일자"] = base[open_date_col].apply(format_yyyymmdd)
    out["담당자"] = base[owner_col].fillna("").astype(str) if owner_col else ""
    out["비고"] = ""
    out["_고객번호"] = base["_고객번호"]

    out = add_bank_compare_columns(out, billing_lookup, "업체명")
    out = out.rename(columns={"최종로그인": "최종로그인일자"})
    return out[["순번", "고객번호", "사업자번호", "업체명", "ERP연계 여부", "접수일자", "구축일자", "방문일자", "담당자", "비고", "최초로그인", "최종로그인일자", "로그인횟수", "메뉴클릭수", "청구시트 여부"]]


def build_link_billing_status(hana_df, billing_lookup, selected_month):
    hana = hana_df.copy()
    hana.columns = [str(c).strip() for c in hana.columns]

    customer_col = find_col(hana, ["고객번호"])
    biz_col = find_col(hana, ["사업자번호"])
    company_col = find_col(hana, ["고객명", "업체명", "상호"])
    add_receipt_col = find_col(hana, ["추가연계접수일"])
    owner_col = find_col(hana, ["담당자"])
    open_date_col = find_col(hana, ["개설/이행일", "개설일", "이행일"])
    link_date_col = find_col(hana, ["연계일자"])
    link_billing_col = find_col(hana, ["연계청구일자", "연계청구일", "청구일자"])
    link_status_col = find_col(hana, ["연계상태"])

    if not customer_col:
        return pd.DataFrame()

    hana["_고객번호"] = hana[customer_col].apply(normalize_billing_customer_no)
    mask = hana["_고객번호"].ne("")
    if link_status_col and link_status_col in hana.columns:
        mask &= hana[link_status_col].astype(str).str.strip().str.contains("연계완료", na=False)
    else:
        return pd.DataFrame()

    base = hana[mask].copy().reset_index(drop=True)
    if base.empty:
        return pd.DataFrame()

    link_type = pd.Series("신규", index=base.index)
    if add_receipt_col and add_receipt_col in base.columns:
        link_type = np.where(base[add_receipt_col].apply(format_yyyymmdd).astype(str).ne(""), "추가", "신규")

    out = pd.DataFrame()
    out["순서"] = range(1, len(base) + 1)
    out["고객번호"] = base["_고객번호"]
    out["사업자번호"] = base[biz_col].apply(normalize_biz_no) if biz_col else ""
    out["업체명"] = base[company_col].fillna("").astype(str) if company_col else ""
    out["구분"] = link_type
    out["추가연계신청일자"] = base[add_receipt_col].apply(format_yyyymmdd) if add_receipt_col else ""
    out["담당자"] = base[owner_col].fillna("").astype(str) if owner_col else ""
    out["구축일"] = base[open_date_col].apply(format_yyyymmdd) if open_date_col else ""
    out["연계시작일자"] = base[link_date_col].apply(format_yyyymmdd) if link_date_col else ""
    out["은행연계완료일자"] = base[link_billing_col].apply(format_yyyymmdd) if link_billing_col else ""
    out["수령여부"] = ""
    out["비고"] = ""
    out["_고객번호"] = base["_고객번호"]

    out = add_bank_compare_columns(out, billing_lookup, "업체명")
    out = out.rename(columns={"사업자번호_은행": "은행 사업자번호"})
    return out[["순서", "고객번호", "사업자번호", "업체명", "구분", "추가연계신청일자", "담당자", "구축일", "연계시작일자", "은행연계완료일자", "수령여부", "비고", "청구시트 여부", "최초로그인", "은행 사업자번호", "최종로그인", "로그인횟수", "메뉴클릭수", "청구원본 고객명", "실적파일 고객명"]]


def build_billing_status_excel_bytes(open_df, link_df):
    return dataframe_to_excel_bytes({
        "개설현황": open_df,
        "연계현황": link_df,
    })


def _merge_with_billing_lookup(hana_df, billing_df, ref_date_col_keys, ref_date_label, extra_hana_col_keys=None):
    """
    기존 build_billing_lookup을 재사용해 hana_df와 billing_df를 고객번호로 대사한 뒤
    ref_date_label 이후 최종로그인일자가 없는 고객을 반환하는 공통 헬퍼.
    """
    hana = hana_df.copy()
    hana.columns = [str(c).strip() for c in hana.columns]

    hana_cust_col  = find_col(hana, ["고객번호"])
    ref_date_col   = find_col(hana, ref_date_col_keys)
    company_col    = find_col(hana, ["고객명", "업체명", "상호"])
    owner_col      = find_col(hana, ["담당자"])
    build_type_col = find_col(hana, ["구축형"])
    extra_col      = find_col(hana, extra_hana_col_keys) if extra_hana_col_keys else None
    end_col        = find_col(hana, ["해지일자", "해지일", "해약일"])

    if not all([hana_cust_col, ref_date_col]):
        return pd.DataFrame()

    hana["_고객번호"] = hana[hana_cust_col].apply(normalize_billing_customer_no)
    hana["_ref_dt"]   = pd.to_datetime(hana[ref_date_col].map(parse_sheet_date), errors="coerce")
    # 해지일자 있는 고객 제외
    if end_col:
        hana["_end_dt"] = pd.to_datetime(hana[end_col].map(parse_sheet_date), errors="coerce")
        hana = hana[hana["_end_dt"].isna()].copy()
    hana_valid = hana[hana["_고객번호"].ne("") & hana["_ref_dt"].notna()].copy()

    # 기존 검증된 build_billing_lookup 사용 (고객번호 매칭 동일 보장)
    bill_lkp = build_billing_lookup(billing_df)
    merged = hana_valid.merge(bill_lkp, on="_고객번호", how="left")

    # _최종로그인은 format_yyyymmdd 결과("YYYYMMDD" 또는 "") → datetime 변환
    merged["_last_dt"] = merged["_최종로그인"].apply(
        lambda v: pd.NaT if (v == "" or pd.isna(v))
        else pd.to_datetime(str(v).strip(), format="%Y%m%d", errors="coerce")
    )

    mask = merged["_last_dt"].isna() | (merged["_last_dt"] <= merged["_ref_dt"])
    result = merged[mask].copy().reset_index(drop=True)
    if result.empty:
        return pd.DataFrame()

    # 직전 3개월 구간 계산 (당월 제외 → 이전 3개월)
    today_kst   = datetime.utcnow() + timedelta(hours=9)
    month_start = today_kst.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    three_ago   = month_start - pd.DateOffset(months=3)
    m_labels    = [(three_ago + pd.DateOffset(months=i)).strftime("%m월") for i in range(3)]
    prev3_label = f"직전3개월({m_labels[0]}~{m_labels[2]})미로그인"

    out = pd.DataFrame()
    out["순번"]          = range(1, len(result) + 1)
    out["고객번호"]      = result["_고객번호"]
    out["고객명"]        = result[company_col].fillna("").astype(str)    if company_col    else ""
    out["담당자"]        = result[owner_col].fillna("").astype(str)      if owner_col      else ""
    out["구축형"]        = result[build_type_col].fillna("").astype(str) if build_type_col else ""
    out[ref_date_label]  = result["_ref_dt"].dt.strftime("%Y-%m-%d")
    if extra_col:
        out[extra_hana_col_keys[0]] = result[extra_col].fillna("").astype(str)
    out["최종로그인일자"] = result["_last_dt"].dt.strftime("%Y-%m-%d").fillna("없음")
    # 직전 3개월 내 로그인 이력 없음 여부 (개설/이행일 이후이면서 직전 3개월 내 미로그인)
    three_ago_ts = pd.Timestamp(three_ago)
    out[prev3_label] = result.apply(
        lambda r: "미로그인" if (
            pd.isna(r["_last_dt"]) or r["_last_dt"] < three_ago_ts
        ) and (pd.notna(r["_ref_dt"]) and r["_ref_dt"] < three_ago_ts)
        else "",
        axis=1,
    )
    return out


def build_no_login_after_open(hana_df, billing_df):
    """구글시트 개설/이행일 이후 청구시트에 최종로그인일자가 없는 고객 목록 반환."""
    return _merge_with_billing_lookup(
        hana_df, billing_df,
        ref_date_col_keys=["개설/이행일", "개설일", "이행일"],
        ref_date_label="개설/이행일",
        extra_hana_col_keys=["개설상태"],
    )


def build_no_login_after_link_billing(hana_df, billing_df):
    """구글시트 연계청구일자 이후 청구시트에 최종로그인일자가 없는 고객 목록 반환."""
    return _merge_with_billing_lookup(
        hana_df, billing_df,
        ref_date_col_keys=["연계청구일자", "연계청구일", "청구일자"],
        ref_date_label="연계청구일자",
        extra_hana_col_keys=["연계상태"],
    )


def build_high_login_no_transfer(hana_df, billing_df, login_threshold=100):
    """로그인 100회 이상이지만 이체 이력이 없는 고객 목록 반환."""
    hana = hana_df.copy()
    hana.columns = [str(c).strip() for c in hana.columns]

    hana_cust_col  = find_col(hana, ["고객번호"])
    company_col    = find_col(hana, ["고객명", "업체명", "상호"])
    owner_col      = find_col(hana, ["담당자"])
    build_type_col = find_col(hana, ["구축형"])
    open_date_col  = find_col(hana, ["개설/이행일", "개설일", "이행일"])
    open_status_col= find_col(hana, ["개설상태"])
    end_col        = find_col(hana, ["해지일자", "해지일", "해약일"])

    if not hana_cust_col:
        return pd.DataFrame()

    hana["_고객번호"] = hana[hana_cust_col].apply(normalize_billing_customer_no)

    # 해지 고객 제외
    if end_col:
        hana["_end_dt"] = pd.to_datetime(hana[end_col].map(parse_sheet_date), errors="coerce")
        hana = hana[hana["_end_dt"].isna()].copy()

    hana_valid = hana[hana["_고객번호"].ne("")].copy()

    bill_lkp = build_billing_lookup(billing_df)
    merged = hana_valid.merge(bill_lkp, on="_고객번호", how="inner")  # 청구시트에 있는 고객만

    # 로그인횟수 숫자 변환
    merged["_login_cnt"] = pd.to_numeric(
        merged["_로그인횟수"].replace("", "0"), errors="coerce"
    ).fillna(0)

    # 조건: 로그인 >= threshold AND 최종이체일자 없음
    mask = (merged["_login_cnt"] >= login_threshold) & (merged["_최종이체일자"].eq("") | merged["_최종이체일자"].isna())
    result = merged[mask].copy().reset_index(drop=True)
    if result.empty:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["순번"]        = range(1, len(result) + 1)
    out["고객번호"]    = result["_고객번호"]
    out["고객명"]      = result[company_col].fillna("").astype(str)    if company_col     else ""
    out["담당자"]      = result[owner_col].fillna("").astype(str)      if owner_col       else ""
    out["구축형"]      = result[build_type_col].fillna("").astype(str) if build_type_col  else ""
    out["개설상태"]    = result[open_status_col].fillna("").astype(str) if open_status_col else ""
    out["개설/이행일"] = result[open_date_col].apply(parse_sheet_date).apply(
        lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else ""
    ) if open_date_col else ""
    out["로그인횟수"]   = result["_로그인횟수"]
    out["최종로그인일자"] = result["_최종로그인"].apply(
        lambda v: pd.to_datetime(v, format="%Y%m%d", errors="coerce").strftime("%Y-%m-%d")
        if v and not pd.isna(v) else "없음"
    )
    out["최종이체일자"] = result["_최종이체일자"].apply(
        lambda v: pd.to_datetime(v, format="%Y%m%d", errors="coerce").strftime("%Y-%m-%d")
        if v and not pd.isna(v) else "없음"
    )
    return out


def _render_high_login_no_transfer(df, owner_key, download_prefix, year_key=None):
    """로그인 100회+ 미이체 고객 렌더링 헬퍼."""
    if df.empty:
        st.info("해당 조건의 고객이 없습니다.")
        return
    date_col = "개설/이행일" if "개설/이행일" in df.columns else None
    years  = sorted(df[date_col].str[:4].dropna().unique().tolist(), reverse=True) if date_col else []
    owners = sorted(df["담당자"].dropna().unique().tolist())
    fc1, fc2, _ = st.columns([2, 2, 6])
    with fc1:
        sel_year = st.selectbox("연도 조회", ["전체"] + years, key=year_key or f"{owner_key}_year")
    with fc2:
        sel_owner = st.selectbox("담당자 조회", ["전체"] + owners, key=owner_key)
    filtered = df.copy()
    if sel_year != "전체" and date_col:
        filtered = filtered[filtered[date_col].str.startswith(sel_year)]
    if sel_owner != "전체":
        filtered = filtered[filtered["담당자"] == sel_owner]
    filtered = filtered.reset_index(drop=True)
    filtered["순번"] = range(1, len(filtered) + 1)
    st.metric("미이체 고객 수", f"{len(filtered):,}건")
    render_plain_html_table(filtered, max_rows=500, center_align=False)
    today_str = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m%d")
    st.download_button(
        "미이체 고객 다운로드",
        data=filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
        file_name=f"{download_prefix}_{sel_year}_{today_str}.csv",
        mime="text/csv",
        use_container_width=True,
    )


def _render_no_login_section(df, year_key, owner_key, label, download_prefix,
                             refresh_key=None, exclude_key=None):
    """미로그인 고객 공통 렌더링 헬퍼 (연도·담당자 필터 + 표 + 다운로드)."""
    if df.empty:
        st.info("해당 조건의 고객이 없습니다.")
        return
    date_col = "개설/이행일" if "개설/이행일" in df.columns else "연계청구일자"
    years  = sorted(df[date_col].str[:4].dropna().unique().tolist(), reverse=True)
    owners = sorted(df["담당자"].dropna().unique().tolist())

    fc1, fc2, fc3, fc4 = st.columns([2, 2, 1, 5])
    with fc1:
        sel_year  = st.selectbox("연도 조회",  ["전체"] + years,  key=year_key)
    with fc2:
        sel_owner = st.selectbox("담당자 조회", ["전체"] + owners, key=owner_key)
    with fc3:
        if refresh_key and st.button("새로고침", key=refresh_key, use_container_width=True):
            st.session_state[f"_nl_rf_{refresh_key}"] = True
            st.rerun()
    exclude_current = False
    with fc4:
        if exclude_key:
            current_ym = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m")
            exclude_current = st.checkbox(f"당월({current_ym}) 개설 제외", key=exclude_key)

    filtered = df.copy()
    if sel_year  != "전체":
        filtered = filtered[filtered[date_col].str.startswith(sel_year)]
    if sel_owner != "전체":
        filtered = filtered[filtered["담당자"] == sel_owner]
    if exclude_current:
        filtered = filtered[~filtered[date_col].str.startswith(current_ym)]

    filtered = filtered.reset_index(drop=True)
    filtered["순번"] = range(1, len(filtered) + 1)
    st.metric(f"{label} 수", f"{len(filtered):,}건")
    render_plain_html_table(filtered, max_rows=500, center_align=False)
    today_str = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y%m%d")
    st.download_button(
        f"{label} 다운로드",
        data=filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
        file_name=f"{download_prefix}_{sel_year}_{today_str}.csv",
        mime="text/csv",
        use_container_width=True,
    )


def show_billing_materials():
    title_col, refresh_col = st.columns([5, 1])
    with title_col:
        st.markdown("### 청구자료 작성")
    with refresh_col:
        refresh = st.button("새로고침", key="billing_refresh", use_container_width=True)

    try:
        hana_df = read_google_csv(
            st.session_state.get("url_hana", DEFAULT_URL_HANA),
            header=2,
            force_refresh=refresh,
        )
        billing_df = load_hana_billing_df(force_refresh=refresh)
        if refresh:
            st.success("새로고침 완료")
    except Exception as e:
        st.error(f"청구자료 생성용 구글 시트를 불러오지 못했습니다: {e}")
        st.info("[구글 스트레드시트 연동] 메뉴에서 하나은행 구글 시트 CSV URL과 하나은행 청구 시트 CSV URL을 확인해주세요.")
        return

    if hana_df is None or hana_df.empty or billing_df is None or billing_df.empty:
        st.info("청구자료로 작성할 데이터가 없습니다.")
        return

    hana_df = hana_df.dropna(how="all").reset_index(drop=True)
    billing_df = billing_df.dropna(how="all").reset_index(drop=True)

    hana_cols = [str(c).strip() for c in hana_df.columns]
    hana_df.columns = hana_cols
    open_date_col = find_col(hana_df, ["개설/이행일", "개설일", "이행일"])
    link_billing_col = find_col(hana_df, ["연계청구일자", "연계청구일", "청구일자"])
    link_date_col = find_col(hana_df, ["연계일자"])

    month_candidates = []
    for date_col in [open_date_col, link_billing_col, link_date_col]:
        if date_col and date_col in hana_df.columns:
            month_candidates.extend(series_yyyymm(hana_df[date_col]).dropna().astype(str).tolist())
    month_values = sorted([m for m in set(month_candidates) if m and m != "NaT"], reverse=True)
    month_options = ["전체"] + month_values
    current_month = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m")
    default_month_index = month_options.index(current_month) if current_month in month_options else 0

    c1, c2 = st.columns([1, 2])
    with c1:
        selected_month = st.selectbox("청구 기준월", month_options, index=default_month_index, key="billing_month")
    with c2:
        st.caption("하나은행 구글 시트와 하나은행 청구 시트를 고객번호 기준으로 대사해 개설/연계현황을 생성합니다.")

    if selected_month == "전체":
        st.info("다운로드할 청구 기준월을 선택해주세요.")
        return

    billing_lookup = build_billing_lookup(billing_df)
    open_df = build_open_billing_status(hana_df, billing_lookup, selected_month)
    link_df = build_link_billing_status(hana_df, billing_lookup, selected_month)

    open_missing_billing = 0
    if "청구시트 여부" in open_df.columns:
        open_missing_billing = int(open_df["청구시트 여부"].astype(str).eq("없음").sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("개설현황", f"{len(open_df):,}건")
    m2.metric("연계현황", f"{len(link_df):,}건")
    m3.metric("개설 청구시트 없음", f"{open_missing_billing:,}건")
    m4.metric("기준월", selected_month)

    tab_open, tab_link = st.tabs(["개설현황", "연계현황"])
    with tab_open:
        selected_open_df = open_df.copy()
        if open_df.empty:
            render_plain_html_table(open_df, max_rows=1000, center_align=False)
        else:
            open_select_df = open_df.copy()
            open_select_df["다운로드"] = False
            edited_open_df = st.data_editor(
                open_select_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=[c for c in open_select_df.columns if c != "다운로드"],
                column_config={"다운로드": st.column_config.CheckboxColumn("다운로드", default=False, width="small")},
                key=f"billing_open_select_{selected_month}",
            )
            selected_open_df = edited_open_df[edited_open_df["다운로드"]].drop(columns=["다운로드"], errors="ignore")
            st.caption(f"다운로드 선택: {len(selected_open_df):,}건")
    with tab_link:
        selected_link_df = link_df.copy()
        if link_df.empty:
            render_plain_html_table(link_df, max_rows=1000, center_align=False)
        else:
            link_select_df = link_df.copy()
            link_select_df["다운로드"] = False
            edited_link_df = st.data_editor(
                link_select_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=[c for c in link_select_df.columns if c != "다운로드"],
                column_config={"다운로드": st.column_config.CheckboxColumn("다운로드", default=False, width="small")},
                key=f"billing_link_select_{selected_month}",
            )
            selected_link_df = edited_link_df[edited_link_df["다운로드"]].drop(columns=["다운로드"], errors="ignore")
            st.caption(f"다운로드 선택: {len(selected_link_df):,}건")

    file_month = selected_month.replace("-", "")
    st.download_button(
        "개설/연계현황 다운로드",
        data=build_billing_status_excel_bytes(selected_open_df, selected_link_df),
        file_name=f"개설_연계현황_{file_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=selected_open_df.empty and selected_link_df.empty,
    )


def show_google_sync():
    st.session_state.url_sync = st.text_input("본사 구글 시트 CSV URL", value=st.session_state.url_sync)
    st.session_state.url_analysis = st.text_input("하나지사 활동이력 구글 시트 CSV URL", value=st.session_state.url_analysis)
    st.session_state.url_hana = st.text_input("하나은행 구글 시트 CSV URL", value=st.session_state.url_hana)
    st.session_state.url_hana_billing = st.text_input("하나은행 청구 시트 CSV URL", value=st.session_state.url_hana_billing)

    st.divider()
    st.markdown("##### DART 전자공시 API 키")
    st.caption("opendart.fss.or.kr 에서 무료 발급 — 유통활동 추천 보정에 사용됩니다.")
    dart_input = st.text_input(
        "DART API 키",
        value=st.session_state.get("dart_api_key", ""),
        type="password",
        placeholder="발급받은 40자리 API 키를 입력하세요",
        key="dart_api_key_input",
    )
    if dart_input != st.session_state.get("dart_api_key", ""):
        st.session_state.dart_api_key = dart_input

    if st.button("데이터 저장", type="primary"):
        errors = []
        try:
            load_csv_to_state("url_sync", "temp_cloud_df", force_refresh=True)
            st.session_state.cloud_sheet_df = st.session_state.temp_cloud_df
        except Exception:
            errors.append("본사 구글 시트")
        try:
            load_csv_to_state("url_analysis", "analysis_lookup_df", force_refresh=True)
        except Exception:
            errors.append("하나지사 활동이력 구글 시트")
        try:
            hana_raw = read_google_csv(st.session_state.url_hana, header=2, force_refresh=True)
            hana_raw = hana_raw.dropna(how="all").reset_index(drop=True)
            st.session_state.hana_sheet_df = hana_raw
        except Exception:
            errors.append("하나은행 구글 시트")
        try:
            billing_raw = read_google_csv(st.session_state.url_hana_billing, force_refresh=True)
            billing_raw = billing_raw.dropna(how="all").reset_index(drop=True)
            st.session_state.hana_billing_df = billing_raw
        except Exception:
            errors.append("하나은행 청구 시트")
        if errors:
            st.error(f"불러오기 실패: {', '.join(errors)} — URL을 확인해주세요.")
        else:
            st.success("불러오기 및 저장 완료")

    if st.session_state.temp_cloud_df is not None:
        st.markdown("**본사 구글 시트 데이터** (상위 20행)")
        render_plain_html_table(strip_activity_time_columns(st.session_state.temp_cloud_df), max_rows=20)

    if st.session_state.analysis_lookup_df is not None:
        st.markdown("**하나지사 활동이력 구글 시트 데이터** (상위 20행)")
        render_plain_html_table(strip_activity_time_columns(st.session_state.analysis_lookup_df), max_rows=20)

    if st.session_state.hana_sheet_df is not None:
        st.markdown("**하나은행 구글 시트 데이터** (상위 20행)")
        render_plain_html_table(strip_activity_time_columns(st.session_state.hana_sheet_df), max_rows=20)

    if st.session_state.hana_billing_df is not None:
        st.markdown("**하나은행 청구 시트 데이터** (상위 20행)")
        render_plain_html_table(strip_activity_time_columns(st.session_state.hana_billing_df), max_rows=20)


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
        opacity: 1; background: #4F46E5; color: white;
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

    /* 사이드바 */
    body:has(#pms-d:checked) [data-testid="stSidebar"],
    body:has(#pms-d:checked) [data-testid="stSidebarContent"] {
        background-color: #171717 !important;
    }
    body:has(#pms-d:checked) [data-testid="stSidebar"] p,
    body:has(#pms-d:checked) [data-testid="stSidebar"] span,
    body:has(#pms-d:checked) [data-testid="stSidebar"] label {
        color: #ffffff !important;
    }
    body:has(#pms-d:checked) [data-testid="stSidebar"] .stButton > button {
        background-color: transparent !important; color: #ffffff !important;
        border-color: transparent !important; box-shadow: none !important;
    }
    body:has(#pms-d:checked) [data-testid="stSidebar"] .stButton > button:hover {
        background-color: #2f2f2f !important;
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


def show_main():
    apply_global_table_css()
    inject_theme_toggle()
    show_sidebar()
    persist_current_menu()

    menu = st.session_state.current_menu
    render_page_title(menu)

    if menu == "대시보드":
        show_dashboard()
    elif menu == "업로드 및 실적 확인":
        show_user_history()
    elif menu == "이번달 활동 대상고객 추천":
        show_target_customers()
    elif menu == "최종 실적 확인":
        st.session_state.current_menu = "업로드 및 실적 확인"
        st.rerun()
    elif menu == "관리자용 실적 확인":
        show_admin_performance()
    elif menu == "실적 분석/계산":
        show_admin_analysis()
    elif menu == "실적 보고서":
        show_report()
    elif menu == "청구자료 작성":
        show_billing_materials()
    elif menu == "직원 및 권한설정":
        show_staff_admin()
    elif menu == "구글 스트레드시트 연동":
        show_google_sync()
    elif menu == "주간보고 이력 작성":
        show_weekly_report_user()
    elif menu == "방문이력 작성":
        show_visit_history()
    elif menu == "주간보고 취합":
        show_weekly_report_admin()
    elif menu == "운영계획":
        show_operation_plan()


restore_login_from_cookie()

if st.session_state.logged_in:
    show_main()
else:
    show_auth_page()
