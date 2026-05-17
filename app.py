# VERSION: 20260516_colab_full_send_button_under_report
import streamlit as st
import pandas as pd
import numpy as np
import time
import json
import os
import html
from datetime import datetime, timedelta

st.set_page_config(page_title="실적관리 시스템", layout="wide", initial_sidebar_state="expanded")

DB_FILE = "users.json"
PERF_FILE = "manual_perf.json"
SENT_FILE = "sent_results.json"

DEFAULT_URL_ANALYSIS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9XPHqrqcaFf9bCOVya7yHORr-c1R4KCF0eEpdE3ESn8qJELP0BkqTOslur9bsGcVabRUIcyOa877R/pub?output=csv"
DEFAULT_URL_SYNC = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9F7R7oLA2B02H-I25kVv2JeYHFgWQq0CT7TeW61hrNpJLdHWJFhFR_iDQGCFAW044o8rRwBDeovKG/pub?gid=1533424484&single=true&output=csv"


def load_db(file_path, default_data):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return default_data


def save_db(file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def init_state():
    if "user_db" not in st.session_state:
        st.session_state.user_db = load_db(
            DB_FILE,
            {
                "1": {
                    "pw": "1",
                    "name": "최고관리자",
                    "email": "",
                    "access": "허용",
                    "role": "관리자",
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
                    "staff_type": "정규직",
                    "outsource": "아니오",
                    "outsource_period": "해당없음",
                },
            },
        )

    defaults = {
        "logged_in": False,
        "user_role": "사용자",
        "user_name": "",
        "auth_mode": "login",
        "current_menu": "실적 분석/계산",
        "url_analysis": DEFAULT_URL_ANALYSIS,
        "url_sync": DEFAULT_URL_SYNC,
        "analysis_lookup_df": None,
        "cloud_sheet_df": None,
        "analysis_result": None,
        "user_excel_data": None,
        "final_reupload_df": None,
        "final_reupload_key": "",
        "temp_cloud_df": None,
        "auto_prev_df": None,
        "deadline_time": "",
        "login_time": "",
        "_prev_menu": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()


def clean_header_logic(df):
    try:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        if len(df.columns) > 0 and str(df.columns[0]).startswith("Unnamed"):
            for i in range(min(len(df), 10)):
                row_text = " ".join(df.iloc[i].astype(str).tolist())
                if any(k in row_text for k in ["등록자", "담당자", "성명", "업체명", "사업자번호"]):
                    df.columns = [str(c).strip() for c in df.iloc[i]]
                    df = df.iloc[i + 1:].reset_index(drop=True)
                    break

        keep = ~pd.Series(df.columns).astype(str).str.contains("^Unnamed|^nan", case=False, na=False).values
        df = df.loc[:, keep]
        return df.dropna(how="all", axis=1).dropna(how="all", axis=0)
    except Exception:
        return df


def find_col(df, keys, fallback=None):
    for c in df.columns:
        if any(k in str(c) for k in keys):
            return c
    return fallback


def normalize_biz(series):
    return series.astype(str).str.replace(r"[^0-9]", "", regex=True)


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
        return df

    cloud = clean_header_logic(cloud.copy())

    biz_col_user = find_col(df, ["사업자번호"])
    biz_col_cloud = find_col(cloud, ["사업자번호"])
    open_col = find_col(cloud, ["개설완료일자", "개설일"])
    erp_col = find_col(cloud, ["ERP연계일자", "ERP", "연계일자"])

    if not biz_col_user or not biz_col_cloud:
        return df

    cloud_cols = [biz_col_cloud]
    rename_map = {}

    if open_col:
        cloud_cols.append(open_col)
        rename_map[open_col] = "본사 개설완료일자"
    if erp_col:
        cloud_cols.append(erp_col)
        rename_map[erp_col] = "본사 ERP연계일자"

    if len(cloud_cols) == 1:
        return df

    df["_biz_key"] = normalize_biz(df[biz_col_user])
    cloud["_biz_key"] = normalize_biz(cloud[biz_col_cloud])

    cloud_map = cloud[["_biz_key"] + [c for c in cloud_cols if c != biz_col_cloud]].rename(columns=rename_map)
    cloud_map = cloud_map.drop_duplicates("_biz_key")

    df = pd.merge(df, cloud_map, on="_biz_key", how="left")
    return df.drop(columns=["_biz_key"], errors="ignore")


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


def manual_points_for_user(name):
    perf_db = load_db(PERF_FILE, {})
    saved = perf_db.get(name, {})

    cl = {
        "타겟고객선별": 5,
        "메일발송": 2,
        "제안서전달": 5,
        "방문설명회&견적발송": 30,
        "계약진행 시": 50,
        "유선 (해피콜)": 5,
        "활성화 (조회업무)": 10,
        "이체, 집금 활성화": 30,
        "계열사 추가도입": 30,
        "신규연계도입": 60,
        "문서 작성 (본사)": 100,
        "문서 작성 (가이드)": 50,
        "문서 작성 (기타)": 20,
        "VOC (아이디어)": 10,
        "운영활동(원격)": 10,
    }

    ll = {
        "타겟고객선별": 100,
        "메일발송": 100,
        "제안서전달": 100,
        "유선 (해피콜)": 100,
        "활성화 (조회업무)": 100,
        "이체, 집금 활성화": 150,
        "문서 작성 (본사)": 100,
        "VOC (아이디어)": 50,
        "운영활동(원격)": 200,
    }

    total = 0
    for item, count in saved.items():
        try:
            score = cl.get(item, 0) * int(count)
            score = min(score, ll.get(item, 999999))
            total += score
        except Exception:
            pass
    return total


def process_performance_analysis(curr_df_raw, prev_df_raw=None):
    try:
        df = clean_header_logic(curr_df_raw)

        u_col = find_col(df, ["등록자", "담당자", "성명"], "등록자")
        d_col = find_col(df, ["활동상세", "활동내용"], "활동상세")
        date_col = find_col(df, ["활동일", "일자"], "활동일")
        biz_col = find_col(df, ["사업자번호"], "사업자번호")
        comp_col = find_col(df, ["업체명", "상호"], "업체명")

        missing = [c for c in [u_col, d_col, date_col] if c not in df.columns]
        if missing:
            return f"필수 컬럼이 없습니다: {', '.join(missing)}", None, None

        rank_order = {"팀장": 1, "과장": 2, "대리": 3, "주임": 4, "직원": 5}
        member_db = {
            "이성환": "팀장",
            "임인지": "과장",
            "전준수": "대리",
            "이수현": "대리",
            "하성춘": "대리",
            "길민종": "주임",
        }

        df_clean = df.dropna(subset=[u_col, date_col]).copy()
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

        if biz_col in df_clean.columns:
            dup_biz_df = df_clean[df_clean.duplicated(subset=[biz_col, u_col, date_col], keep=False)].sort_values(
                by=[date_col, biz_col, u_col]
            )
        else:
            dup_biz_df = pd.DataFrame()

        summary = (
            df.dropna(subset=[u_col, d_col])
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
            manual_p = manual_points_for_user(name)
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
                    "운영건수 (추가 활동)": min(int(np.ceil(stats["manual_p"] / 30.0)), int(stats["p_sum"] / 30), max(0, 60 - int(row["v"]))),
                    "운영포인트(추가 활동)": effective_manual_p,
                    "합계포인트": stats["p_sum"],
                    "지급포인트": max(0, stats["p_sum"] - 1000),
                    "지급예상금액": pay,
                    "전월대비": 0,
                }
            )

        res_df = pd.DataFrame(rows)
        res_df["_rank"] = res_df["직급"].map(rank_order).fillna(5)
        res_df = res_df.sort_values("_rank").drop(columns=["_rank"])

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


def style_report_logic(df, compact=False):
    if df is None or df.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    diff_cols = [c for c in df.columns if "전월대비" in str(c)]
    num_cols = [c for c in df.columns if c not in ["담당자", "직급", "전송시각", "등록월", "항목", "일치여부"] + diff_cols]

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

    th_pad = "5px 6px" if compact else "9px 12px"
    th_font = "12px" if compact else "13px"
    td_pad = "5px 6px" if compact else "8px 12px"
    td_font = "12px" if compact else "13px"
    th = f"background:#EDF2F7;color:#4A5568;font-weight:800;font-size:{th_font};padding:{th_pad};text-align:center;border-bottom:2px solid #E2E8F0;" + ("white-space:normal;word-break:keep-all;" if compact else "white-space:nowrap;")
    headers = "".join(f"<th style='{th}'>{html.escape(str(c))}</th>" for c in df.columns)

    body = ""
    for i, row in df.reset_index(drop=True).iterrows():
        tds = ""
        for col in df.columns:
            align = "right" if col in num_cols else "center"
            bg = "#FFFFFF" if i % 2 == 0 else "#F7FAFC"

            if col == "일치여부":
                value = fmt_match(row[col])
            elif col in diff_cols:
                value = fmt_diff(row[col])
            elif col in num_cols:
                value = fmt_num(row[col])
            else:
                value = "" if pd.isna(row[col]) else html.escape(str(row[col]))

            _ws = "white-space:normal;word-break:keep-all;" if compact else "white-space:nowrap;"
            tds += (
                f"<td style='background:{bg};padding:{td_pad};border-bottom:1px solid #EDF2F7;"
                f"font-size:{td_font};color:#2D3748;text-align:{align};{_ws}'>{value}</td>"
            )
        body += f"<tr>{tds}</tr>"

    _ov = "visible" if compact else "auto"
    _tbl_style = "width:100%;border-collapse:collapse;table-layout:fixed;" if compact else "width:100%;border-collapse:collapse;"
    st.markdown(
        f"""
        <div style="overflow-x:{_ov};border:1px solid #E2E8F0;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);margin-bottom:1rem;">
            <table style="{_tbl_style}">
                <thead><tr>{headers}</tr></thead>
                <tbody>{body}</tbody>
            </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_auth_page():
    st.markdown(
        """
        <style>
        .stAppHeader, header, [data-testid="stHeader"], .stDecoration {
            display: none !important;
            height: 0 !important;
            visibility: hidden !important;
        }
        [data-testid="stSidebar"] { display: none !important; }
        .stApp, [data-testid="stAppViewContainer"], .main {
            background: linear-gradient(135deg, #F0EAFF 0%, #FFFFFF 72%) !important;
        }
        .main .block-container {
            max-width: 1500px !important;
            padding: 46px 5vw 40px !important;
        }
        .auth-brand {
            color: #1E1A3A;
            font-size: 34px;
            font-weight: 900;
            line-height: 1.15;
            margin-bottom: 130px;
        }
        .auth-brand span {
            display: block;
            color: #6F6A98;
            font-size: 15px;
            font-weight: 600;
            margin-top: 10px;
        }
        .auth-copy-title {
            color: #1E1A3A;
            font-size: 54px;
            line-height: 1.18;
            font-weight: 900;
            letter-spacing: 0;
            margin-bottom: 28px;
        }
        .auth-copy-desc {
            color: #574FA0;
            font-size: 17px;
            line-height: 1.8;
            font-weight: 700;
            margin-bottom: 120px;
        }
        .auth-footer {
            color: #A19ACB;
            font-size: 12px;
            font-weight: 600;
        }
        .login-spacer { height: 170px; }
        .auth-small {
            text-align: center;
            color: #A09AC5;
            font-size: 14px;
            margin: 14px 0 8px;
        }
        div[data-testid="stTextInput"] { margin-bottom: 12px !important; }
        div[data-testid="stTextInput"] label {
            color: #242041 !important;
            font-size: 14px !important;
            font-weight: 800 !important;
        }
        div[data-testid="stTextInput"] > div[data-baseweb="input"] {
            min-height: 58px !important;
            border-radius: 12px !important;
            border: 1.5px solid #DFDAFF !important;
            background: #F8F7FF !important;
        }
        div[data-testid="stTextInput"] input {
            color: #1E1A3A !important;
            font-size: 16px !important;
        }
        [data-testid="InputInstructions"] { display: none !important; }
        div[data-testid="stCheckbox"] label {
            color: #4B466D !important;
            font-size: 15px !important;
            font-weight: 650 !important;
        }
        div.stButton > button {
            height: 58px !important;
            border-radius: 12px !important;
            font-size: 16px !important;
            font-weight: 800 !important;
            background: #FFFFFF !important;
            color: #5B51A8 !important;
            border: 1.5px solid #DDD8FF !important;
            box-shadow: 0 6px 18px rgba(104,91,190,0.14) !important;
        }
        div.stButton > button:hover {
            background: #F3F0FF !important;
            border-color: #7B6FD4 !important;
            color: #493E9A !important;
        }
        hr {
            border: 0 !important;
            border-top: 1px solid #E6E1FF !important;
            margin: 20px 0 0 !important;
        }
        @media (max-width: 900px) {
            .main .block-container { padding: 28px 24px 36px !important; }
            .auth-brand { margin-bottom: 70px; }
            .auth-copy-title, .auth-copy-desc, .auth-footer { display: none; }
            .login-spacer { height: 36px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    left, _, right = st.columns([0.85, 0.18, 0.72], gap="large")

    with left:
        st.markdown(
            """
            <div class="auth-brand">실적관리 시스템<span>Performance Management System</span></div>
            <div class="auth-copy-title">팀의 실적,<br>하나의<br>시스템에서.</div>
            <div class="auth-copy-desc">
                개설·연계·운영 실적을 통합 관리하고,<br>
                월별 포인트와 지급예상금액을<br>
                실시간으로 확인하세요.
            </div>
            <div class="auth-footer">© 2026 CMS · 실적관리 시스템</div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown("<div class='login-spacer'></div>", unsafe_allow_html=True)

        if st.session_state.auth_mode == "login":
            sid = ""
            if os.path.exists("saved_id.txt"):
                try:
                    with open("saved_id.txt", "r", encoding="utf-8") as f:
                        sid = f.read().strip()
                except Exception:
                    pass

            u_id = st.text_input("아이디", value=sid, placeholder="아이디를 입력하세요", key="l_id")
            u_pw = st.text_input("비밀번호", type="password", placeholder="비밀번호를 입력하세요", key="l_pw")
            save_id_cb = st.checkbox("아이디 저장", value=bool(sid))

            if st.button("로그인", use_container_width=True, type="primary"):
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
                        try:
                            with open("saved_id.txt", "w", encoding="utf-8") as f:
                                f.write(u_id_str)
                        except Exception:
                            pass
                    elif os.path.exists("saved_id.txt"):
                        try:
                            os.remove("saved_id.txt")
                        except Exception:
                            pass

                    user = db.get(u_id_str, {"role": "관리자", "name": "최고관리자"})
                    st.session_state.logged_in = True
                    st.session_state.user_role = user.get("role", "관리자")
                    st.session_state.user_name = user.get("name", "최고관리자")
                    st.session_state.login_time = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")
                    st.session_state.current_menu = "실적 분석/계산" if st.session_state.user_role == "관리자" else "이력확인 및 작성"
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

        else:
            st.markdown("<div style='height:80px'></div>", unsafe_allow_html=True)

            r_id = st.text_input("아이디", key="r_id")
            if r_id:
                if r_id in st.session_state.user_db:
                    st.error("이미 사용 중인 아이디입니다.")
                else:
                    st.success("사용 가능한 아이디입니다.")

            r_name = st.text_input("성명", key="r_name")
            r_email = st.text_input("메일주소", key="r_email")
            r_pw = st.text_input("비밀번호", type="password", key="r_pw")
            r_pw2 = st.text_input("비밀번호 확인", type="password", key="r_pw2")

            if st.button("회원가입", use_container_width=True, type="primary"):
                if not r_id or not r_name or not r_email or not r_pw or not r_pw2:
                    st.error("모든 항목을 입력해주세요.")
                elif r_pw != r_pw2:
                    st.error("비밀번호가 일치하지 않습니다.")
                elif r_id in st.session_state.user_db:
                    st.error("이미 존재하는 아이디입니다.")
                else:
                    st.session_state.user_db[r_id] = {
                        "pw": r_pw,
                        "name": r_name,
                        "email": r_email,
                        "access": "불가",
                        "role": "사용자",
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
    with st.sidebar:
        st.markdown(
            "<div style='font-size:26px;font-weight:900;color:#1A202C;letter-spacing:-0.5px;padding:8px 0 4px;'>실적관리 시스템</div>",
            unsafe_allow_html=True,
        )
        _lt = st.session_state.get("login_time", "")
        st.markdown(
            f"<div style='display:flex;justify-content:space-between;align-items:center;font-size:13px;color:#4A5568;padding:0 2px 8px;'>"
            f"<span style='font-weight:700;'>{st.session_state.user_name}</span>"
            f"<span style='color:#718096;font-size:12px;'>접속 {_lt}</span>"
            f"</div>",
            unsafe_allow_html=True,
        )
        st.divider()

        if st.session_state.user_role == "관리자":
            st.markdown("관리자 메뉴")
            for menu_name in ["실적 분석/계산", "실적 보고서", "직원 및 권한설정", "구글 스트레드시트 연동"]:
                if st.button(menu_name, use_container_width=True):
                    st.session_state.current_menu = menu_name
                    st.rerun()

        st.markdown("사용자 메뉴")
        for menu_name in ["이력확인 및 작성", "최종 실적 확인"]:
            if st.button(menu_name, use_container_width=True):
                st.session_state.current_menu = menu_name
                st.rerun()

        st.divider()
        if st.button("로그아웃", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.auth_mode = "login"
            st.rerun()


def render_page_title(menu):
    if st.session_state.get("_prev_menu") != menu:
        st.session_state._prev_menu = menu
        bar = st.progress(0)
        for pct in [25, 55, 85, 100]:
            time.sleep(0.03)
            bar.progress(pct)
        bar.empty()

    st.markdown(f"## {menu}")
    admin_menus = ["실적 분석/계산", "실적 보고서", "직원 및 권한설정", "구글 스트레드시트 연동"]
    parent_nav = "관리자 메뉴" if menu in admin_menus else "사용자 메뉴"

    st.markdown(
        f"<div style='color:#718096;font-size:14px;font-weight:600;margin-top:-8px;margin-bottom:26px;'>{parent_nav} &gt; {html.escape(menu)}</div>",
        unsafe_allow_html=True,
    )


def load_csv_to_state(url_key, state_key):
    st.session_state[state_key] = clean_header_logic(pd.read_csv(st.session_state[url_key]))


def select_prev_month(state_key, widget_key):
    if st.button("구글시트 최신 조회"):
        try:
            load_csv_to_state("url_analysis", "analysis_lookup_df")
            st.success("동기화 완료")
        except Exception:
            st.error("구글시트 URL을 확인해주세요.")

    if st.session_state.analysis_lookup_df is not None:
        c_df = st.session_state.analysis_lookup_df.copy()
        d_col = find_col(c_df, ["활동일", "일자"])
        if d_col and d_col in c_df.columns:
            c_df[d_col] = pd.to_datetime(c_df[d_col], errors="coerce")
            opts = sorted(c_df[d_col].dropna().dt.strftime("%Y-%m").unique(), reverse=True)
            sel = st.selectbox("비교할 전월 선택", ["선택안함"] + list(opts), key=widget_key)
            st.session_state[state_key] = c_df[c_df[d_col].dt.strftime("%Y-%m") == sel] if sel != "선택안함" else None


def show_user_history():
    u_file = st.file_uploader("활동실적 엑셀 업로드", type=["xlsx"])

    if st.button("조회"):
        if not u_file:
            st.warning("엑셀 파일을 먼저 업로드해주세요.")
        else:
            with st.spinner("분석 중입니다."):
                try:
                    load_csv_to_state("url_sync", "cloud_sheet_df")
                except Exception:
                    pass
                st.session_state.user_excel_data = clean_header_logic(pd.read_excel(u_file, sheet_name=0))
                st.session_state.final_reupload_df = None
                st.session_state.final_reupload_key = ""
                st.success("분석 완료")

    if st.session_state.user_excel_data is None:
        return

    df = st.session_state.user_excel_data.copy()
    u_col = find_col(df, ["등록자", "담당자", "성명"], "등록자")
    d_col = find_col(df, ["활동상세", "활동내용"], "활동상세")

    df_user = df[df[u_col] == st.session_state.user_name].copy() if u_col in df.columns else pd.DataFrame()
    df_user = attach_cloud_dates(df_user)

    st.markdown("### 담당자 활동 수치")
    res, err, dup = process_performance_analysis(df_user, st.session_state.get("auto_prev_df"))

    if isinstance(res, pd.DataFrame) and not res.empty:
        my_res = res[res["담당자"] == st.session_state.user_name].copy()
        drop_cols = [c for c in my_res.columns if "전월대비" in c]
        style_report_logic(my_res.drop(columns=drop_cols, errors="ignore"), compact=True)
    elif isinstance(res, str):
        st.error(res)

    t1, t2, t3, t4 = st.tabs(["중복 방문", "초과 방문", "본사 개설완료일자 누락", "본사 ERP연계일자 누락"])

    with t1:
        style_report_logic(dup)

    with t2:
        if err is not None and not err.empty:
            err_filtered = err[(err["일방문"] >= 5) | (err["월총방문"] >= 60)].copy()
            monthly_map = err_filtered.groupby("담당자")["월총방문"].first().to_dict() if "담당자" in err_filtered.columns else {}

            style_report_logic(err_filtered.drop(columns=["월총방문"], errors="ignore"))

            if monthly_map:
                parts = "　｜　".join([f"<b>{name}</b>: {cnt:,}회" for name, cnt in monthly_map.items()])
                st.markdown(
                    f"<div style='margin-top:8px;padding:10px 16px;background:#EBF8FF;border-radius:8px;font-size:13px;color:#2B6CB0;'>월총방문　{parts}</div>",
                    unsafe_allow_html=True,
                )
        else:
            st.info("초과 방문 데이터가 없습니다.")

    with t3:
        if "본사 개설완료일자" in df_user.columns:
            missing_open = df_user[
                pd.isna(df_user["본사 개설완료일자"]) | (df_user["본사 개설완료일자"].astype(str).str.strip() == "")
            ]
            style_report_logic(missing_open.drop(columns=["본사 ERP연계일자"], errors="ignore"))
        else:
            st.info("본사 구글시트에 개설완료일자 또는 사업자번호 컬럼이 없어 확인할 수 없습니다.")

    with t4:
        if "본사 ERP연계일자" in df_user.columns:
            if d_col and d_col in df_user.columns:
                target = df_user[df_user[d_col].astype(str).str.contains("연계", na=False)]
            else:
                target = df_user

            missing_erp = target[
                pd.isna(target["본사 ERP연계일자"]) | (target["본사 ERP연계일자"].astype(str).str.strip() == "")
            ]
            style_report_logic(missing_erp.drop(columns=["본사 개설완료일자"], errors="ignore"))
        else:
            st.info("본사 구글시트에 ERP연계일자 또는 사업자번호 컬럼이 없어 확인할 수 없습니다.")

    st.divider()
    st.markdown("### 추가 실적 입력")

    base = criteria_df()
    saved = load_db(PERF_FILE, {}).get(st.session_state.user_name, {})
    base["입력(건)"] = base["구분"].map(saved).fillna(0).astype(int)

    edited = st.data_editor(
        base,
        use_container_width=True,
        hide_index=True,
        height=520,
        column_config={"입력(건)": st.column_config.NumberColumn("입력(건)", min_value=0, step=1, default=0)},
        disabled=["활동구분", "구분", "단위 점수", "월 최대점수"],
    )

    edited["입력(건)"] = edited["입력(건)"].fillna(0).astype(int)
    edited["계산점수"] = edited["단위 점수"] * edited["입력(건)"]

    limit_map = {
        "타겟고객선별": 100,
        "메일발송": 100,
        "제안서전달": 100,
        "유선 (해피콜)": 100,
        "활성화 (조회업무)": 100,
        "이체, 집금 활성화": 150,
        "문서 작성 (본사)": 100,
        "VOC (아이디어)": 50,
        "운영활동(원격)": 200,
    }

    total = 0
    for _, row in edited.iterrows():
        item = row["구분"]
        score = int(row["계산점수"])
        total += min(score, limit_map[item]) if item in limit_map else score

    st.metric("추가 실적 합산 점수", f"{total:,} PT")

    if st.button("저장 후 최종 실적 확인", type="primary"):
        db = load_db(PERF_FILE, {})
        db[st.session_state.user_name] = edited.set_index("구분")["입력(건)"].to_dict()
        save_db(PERF_FILE, db)
        st.session_state.current_menu = "최종 실적 확인"
        st.success("저장 완료")
        time.sleep(0.5)
        st.rerun()


def show_final_check():
    select_prev_month("auto_prev_df", "user_prev_month_sel")

    if st.session_state.user_excel_data is None:
        st.info("먼저 이력확인 및 작성 메뉴에서 엑셀을 업로드해주세요.")
        return

    original_df = st.session_state.user_excel_data

    # Filter user data and attach cloud dates for tabs
    df_for_tabs = original_df.copy()
    u_col = find_col(df_for_tabs, ["등록자", "담당자", "성명"], "등록자")
    d_col = find_col(df_for_tabs, ["활동상세", "활동내용"], "활동상세")
    df_user = df_for_tabs[df_for_tabs[u_col] == st.session_state.user_name].copy() if u_col in df_for_tabs.columns else pd.DataFrame()
    df_user = attach_cloud_dates(df_user)

    res, err, dup = process_performance_analysis(original_df, st.session_state.get("auto_prev_df"))

    if not isinstance(res, pd.DataFrame) or res.empty:
        st.error(res if isinstance(res, str) else "실적을 계산할 수 없습니다.")
        return

    my_res = res[res["담당자"] == st.session_state.user_name]
    uname = st.session_state.user_name

    hidden_final_cols = ["운영건수 (추가 활동)", "운영포인트(추가 활동)"]
    my_res_display = my_res.drop(columns=hidden_final_cols, errors="ignore")

    st.markdown(
        f"<div style='padding:10px 16px;background:#EBF8FF;border-left:4px solid #4299E1;border-radius:6px;font-size:15px;font-weight:700;color:#2B6CB0;margin-bottom:12px;'>"
        f"{html.escape(uname)}님의 실제 활동 실적 입니다.</div>",
        unsafe_allow_html=True,
    )

    st.dataframe(my_res_display.reset_index(drop=True), use_container_width=True, hide_index=True)

    if my_res.empty:
        return

    add_cnt = int(my_res.iloc[0].get("운영건수 (추가 활동)", 0))
    st.markdown(
        f"<div style='margin-top:8px;padding:10px 16px;background:#FFFBEB;border:1px solid #F6AD55;border-radius:8px;font-size:14px;color:#92400E;font-weight:600;'>"
        f"{html.escape(uname)}님은 <span style='color:#C05621;font-size:16px;font-weight:800;'>{add_cnt}건</span>의 활동이력을 추가로 등록 후 엑셀업로드 해주세요.</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    new_file = st.file_uploader("엑셀 재업로드", type=["xlsx"], key="final_reupload")

    if new_file is not None:
        file_key = f"{new_file.name}_{new_file.size}"

        if st.session_state.get("final_reupload_key") != file_key:
            st.session_state.final_reupload_key = file_key
            uploaded_df = clean_header_logic(pd.read_excel(new_file, sheet_name=0))
            st.session_state.final_reupload_df = uploaded_df
            st.toast("업로드 완료. 검증 리포트에 반영됩니다.")
            st.rerun()

    uploaded_df = st.session_state.get("final_reupload_df")
    uploaded_my_res = pd.DataFrame()

    if uploaded_df is not None:
        uploaded_res, _, _ = process_performance_analysis(uploaded_df, st.session_state.get("auto_prev_df"))
        if isinstance(uploaded_res, pd.DataFrame) and not uploaded_res.empty:
            uploaded_my_res = uploaded_res[uploaded_res["담당자"] == st.session_state.user_name]

    st.markdown("##### 실적 예상치 검증 리포트")

    uploaded_exists = uploaded_df is not None and not uploaded_my_res.empty

    report_items = [
        ("개설건수", "개설건수", "compare"),
        ("개설포인트", "개설포인트", "compare"),
        ("연계건수", "연계건수", "compare"),
        ("연계포인트", "연계포인트", "compare"),
        ("운영건수 (실제 활동)", "운영건수 (실제 활동)", "compare"),
        ("운영포인트 (실제 활동)", "운영포인트 (실제 활동)", "compare"),
        ("추가 등록건수", "운영건수 (추가 활동)", "compare"),
        ("최종 운영건수", None, "final_operation_count"),
        ("합계포인트", "합계포인트", "compare"),
        ("지급포인트", "지급포인트", "compare"),
        ("지급예상금액", "지급예상금액", "compare"),
    ]

    cmp_rows = []

    for label, source_col, mode in report_items:
        first_value = ""
        uploaded_value = ""
        match_value = ""

        if mode == "final_operation_count":
            first_actual_count = int(float(my_res.iloc[0].get("운영건수 (실제 활동)", 0)))
            first_extra_count = int(float(my_res.iloc[0].get("운영건수 (추가 활동)", 0)))
            first_compare_value = min(60, first_actual_count + first_extra_count)
            first_value = first_compare_value

            if uploaded_exists:
                uploaded_actual_count = int(float(uploaded_my_res.iloc[0].get("운영건수 (실제 활동)", 0)))
                # 재업로드 Excel의 실제 활동건수만 사용 (추가 등록분이 실제 활동으로 반영된 결과)
                uploaded_value = min(60, uploaded_actual_count)
                match_value = "일치" if first_compare_value == uploaded_value else "불일치"
        elif mode == "compare_upload_only":
            first_value = ""

            if uploaded_exists and source_col in uploaded_my_res.columns:
                uploaded_value = int(float(uploaded_my_res.iloc[0].get(source_col, 0)))
        else:
            if source_col in my_res.columns:
                first_value = int(float(my_res.iloc[0].get(source_col, 0)))

            if uploaded_exists and source_col in uploaded_my_res.columns:
                uploaded_value = int(float(uploaded_my_res.iloc[0].get(source_col, 0)))
                match_value = "일치" if first_value == uploaded_value else "불일치"

        cmp_rows.append(
            {
                "항목": label,
                "업로드 후 예상치": first_value,
                "업로드 후 결과": uploaded_value,
                "일치여부": match_value,
            }
        )

    cmp_df = pd.DataFrame(cmp_rows)
    style_report_logic(cmp_df)

    # ── 중복방문/초과방문/누락 확인 탭 ──
    t1, t2, t3, t4 = st.tabs(["중복 방문", "초과 방문", "본사 개설완료일자 누락", "본사 ERP연계일자 누락"])

    with t1:
        style_report_logic(dup)

    with t2:
        if err is not None and not err.empty:
            err_filtered = err[(err["일방문"] >= 5) | (err["월총방문"] >= 60)].copy()
            monthly_map = err_filtered.groupby("담당자")["월총방문"].first().to_dict() if "담당자" in err_filtered.columns else {}

            style_report_logic(err_filtered.drop(columns=["월총방문"], errors="ignore"))

            if monthly_map:
                parts = "　｜　".join([f"<b>{name}</b>: {cnt:,}회" for name, cnt in monthly_map.items()])
                st.markdown(
                    f"<div style='margin-top:8px;padding:10px 16px;background:#EBF8FF;border-radius:8px;font-size:13px;color:#2B6CB0;'>월총방문　{parts}</div>",
                    unsafe_allow_html=True,
                )
        else:
            st.info("초과 방문 데이터가 없습니다.")

    with t3:
        if "본사 개설완료일자" in df_user.columns:
            missing_open = df_user[
                pd.isna(df_user["본사 개설완료일자"]) | (df_user["본사 개설완료일자"].astype(str).str.strip() == "")
            ]
            style_report_logic(missing_open.drop(columns=["본사 ERP연계일자"], errors="ignore"))
        else:
            st.info("본사 구글시트에 개설완료일자 또는 사업자번호 컬럼이 없어 확인할 수 없습니다.")

    with t4:
        if "본사 ERP연계일자" in df_user.columns:
            if d_col and d_col in df_user.columns:
                target = df_user[df_user[d_col].astype(str).str.contains("연계", na=False)]
            else:
                target = df_user

            missing_erp = target[
                pd.isna(target["본사 ERP연계일자"]) | (target["본사 ERP연계일자"].astype(str).str.strip() == "")
            ]
            style_report_logic(missing_erp.drop(columns=["본사 개설완료일자"], errors="ignore"))
        else:
            st.info("본사 구글시트에 ERP연계일자 또는 사업자번호 컬럼이 없어 확인할 수 없습니다.")

    # ── 모든 항목 일치 여부 체크 ──
    all_match = uploaded_exists and all(r.get("일치여부") == "일치" for r in cmp_rows)

    if not uploaded_exists:
        st.markdown(
            "<div style='margin-top:8px;padding:10px 16px;background:#FFFBEB;border:1px solid #F6AD55;border-radius:8px;font-size:13px;color:#92400E;font-weight:700;'>"
            "📂 엑셀을 재업로드하여 검증을 완료한 후 실적을 전송할 수 있습니다.</div>",
            unsafe_allow_html=True,
        )
    elif not all_match:
        st.markdown(
            "<div style='margin-top:8px;padding:10px 16px;background:#FFF5F5;border:1px solid #FC8181;border-radius:8px;font-size:13px;color:#C53030;font-weight:700;'>"
            "❌ 일치하지 않는 항목이 있습니다. 모든 항목이 일치해야 실적 결과를 전송할 수 있습니다.</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<div style='margin-top:8px;padding:10px 16px;background:#F0FFF4;border:1px solid #9AE6B4;border-radius:8px;font-size:13px;color:#276749;font-weight:700;'>"
            "✅ 모든 항목이 일치합니다. 실적 결과를 전송할 수 있습니다.</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)

    _, send_col = st.columns([0.78, 0.22])
    with send_col:
        do_send = st.button("실적 결과 전송", use_container_width=True, type="primary", disabled=not all_match)

    if do_send:
        sent_db = load_db(SENT_FILE, {})
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
        save_db(SENT_FILE, sent_db)

        st.success("전송 완료")
        time.sleep(0.5)
        st.rerun()


def sent_results_df():
    sent_db = load_db(SENT_FILE, {})
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

    return sent_df


def show_admin_analysis():
    select_prev_month("auto_prev_df", "adm_prev_month")

    st.markdown("### 직원 전송 실적 내역")
    sent_df = sent_results_df()

    if sent_df.empty:
        st.info("아직 전송된 실적이 없습니다.")
        return

    style_report_logic(sent_df)

    c1, c2 = st.columns([1, 1])

    with c1:
        if st.button("전송 내역 초기화"):
            save_db(SENT_FILE, {})
            st.success("초기화 완료")
            time.sleep(0.5)
            st.rerun()

    with c2:
        if st.button("마감", type="primary"):
            st.session_state.deadline_time = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")
            st.session_state.analysis_result = sent_df.copy()
            st.success("마감 처리 완료")
            time.sleep(0.5)
            st.rerun()

    if st.session_state.deadline_time:
        st.info(f"마감 완료: {st.session_state.deadline_time}")


def show_report():
    if st.session_state.analysis_result is None:
        st.info("실적 분석/계산 메뉴에서 마감 후 확인 가능합니다.")
        return

    drop_cols = ["전송시각"] + [c for c in st.session_state.analysis_result.columns if "관리자전월대비" in c]
    report_df = st.session_state.analysis_result.drop(columns=drop_cols, errors="ignore")

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

    for name in report_df["담당자"].tolist():
        curr_row = report_df[report_df["담당자"] == name]
        prev_row = prev_res[prev_res["담당자"] == name]

        if curr_row.empty:
            continue

        rank = curr_row.iloc[0].get("직급", "")
        st.markdown(f"#### {name} ({rank})")

        rows = []
        for col in shared_cols:
            c_val = int(float(curr_row.iloc[0][col]))
            p_val = int(float(prev_row.iloc[0][col])) if not prev_row.empty else 0
            rows.append({"항목": col, prev_month_label: p_val, curr_month_label: c_val, "증감": c_val - p_val})

        style_report_logic(pd.DataFrame(rows))


def show_staff_admin():
    st.markdown("### 직원 목록")

    staff_rows = []
    for uid, info in st.session_state.user_db.items():
        if uid == "1":
            continue

        staff_rows.append(
            {
                "ID": uid,
                "성명": info.get("name", ""),
                "메일주소": info.get("email", ""),
                "직원구분": info.get("staff_type", "정규직"),
                "외주여부": info.get("outsource", "아니오"),
                "외주 근무기간": info.get("outsource_period", "해당없음"),
                "로그인 허용 여부": info.get("access", "불가"),
                "메뉴 접근 권한": "관리자 메뉴" if info.get("role") == "관리자" else "사용자 메뉴",
                "삭제": False,
            }
        )

    if not staff_rows:
        st.info("등록된 직원이 없습니다.")
        return

    edited = st.data_editor(
        pd.DataFrame(staff_rows),
        use_container_width=True,
        hide_index=True,
        column_config={
            "직원구분": st.column_config.SelectboxColumn("직원구분", options=["정규직", "계약직", "파견직", "외주"]),
            "외주여부": st.column_config.SelectboxColumn("외주여부", options=["예", "아니오"]),
            "외주 근무기간": st.column_config.SelectboxColumn("외주 근무기간", options=["해당없음", "1년 미만", "1년 이상", "2년 이상"]),
            "로그인 허용 여부": st.column_config.SelectboxColumn("로그인 허용 여부", options=["허용", "불가"]),
            "메뉴 접근 권한": st.column_config.SelectboxColumn("메뉴 접근 권한", options=["사용자 메뉴", "관리자 메뉴"]),
            "삭제": st.column_config.CheckboxColumn("삭제"),
        },
    )

    if st.button("저장", type="primary"):
        for _, row in edited.iterrows():
            uid = row["ID"]

            if uid not in st.session_state.user_db:
                continue

            if row.get("삭제", False):
                del st.session_state.user_db[uid]
            else:
                st.session_state.user_db[uid]["email"] = row.get("메일주소", "")
                st.session_state.user_db[uid]["staff_type"] = row.get("직원구분", "정규직")
                st.session_state.user_db[uid]["outsource"] = row.get("외주여부", "아니오")
                st.session_state.user_db[uid]["outsource_period"] = row.get("외주 근무기간", "해당없음")
                st.session_state.user_db[uid]["access"] = row.get("로그인 허용 여부", "불가")
                st.session_state.user_db[uid]["role"] = "관리자" if row.get("메뉴 접근 권한") == "관리자 메뉴" else "사용자"

        save_db(DB_FILE, st.session_state.user_db)
        st.success("저장 완료")
        time.sleep(0.5)
        st.rerun()


def show_google_sync():
    st.session_state.url_sync = st.text_input("본사 구글 시트 CSV URL", value=st.session_state.url_sync)
    st.session_state.url_analysis = st.text_input("하나지사 활동이력 구글 시트 CSV URL", value=st.session_state.url_analysis)

    if st.button("데이터 저장", type="primary"):
        try:
            load_csv_to_state("url_sync", "temp_cloud_df")
            st.session_state.cloud_sheet_df = st.session_state.temp_cloud_df
            st.success("불러오기 및 저장 완료")
        except Exception:
            st.error("불러오기 실패. URL을 확인해주세요.")

    if st.session_state.temp_cloud_df is not None:
        st.dataframe(st.session_state.temp_cloud_df, use_container_width=True, hide_index=True)


def show_main():
    show_sidebar()

    menu = st.session_state.current_menu
    render_page_title(menu)

    if menu == "이력확인 및 작성":
        show_user_history()
    elif menu == "최종 실적 확인":
        show_final_check()
    elif menu == "실적 분석/계산":
        show_admin_analysis()
    elif menu == "실적 보고서":
        show_report()
    elif menu == "직원 및 권한설정":
        show_staff_admin()
    elif menu == "구글 스트레드시트 연동":
        show_google_sync()


if st.session_state.logged_in:
    show_main()
else:
    show_auth_page()
