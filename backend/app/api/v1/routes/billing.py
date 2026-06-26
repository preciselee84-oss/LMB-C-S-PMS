from io import BytesIO
import re

import httpx
import pandas as pd
from fastapi import APIRouter, File, HTTPException, UploadFile, status

from app.schemas.billing import BillingPreview, BillingPreviewRow, BillingPreviewSummary

router = APIRouter()

SPREADSHEET_ID = "12BeCTDegUWD-jomaG3WS75Jx1dJ9Lqjxn1hVs3FrpE4"
SPREADSHEET_TITLE = "청구자료대사"
SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"

SHEETS = {
    "open": {"gid": "1244892381", "title": "청구원본(개설업로드)", "header": 2},
    "erp": {"gid": "1551524823", "title": "청구원본(연계업로드)", "header": 2},
    "login": {"gid": "0", "title": "은행로그인실적파일(은행)", "header": 0},
    "customer": {"gid": "1769977491", "title": "고객정보파일(은행)", "header": 3},
}


def _digits(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _text(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"nan", "none", "nat"} else text


def _int(value: object) -> int:
    try:
        return int(float(str(value or "0").replace(",", "")))
    except ValueError:
        return 0


def _company_key(value: object) -> str:
    text = re.sub(r"\s+", "", str(value or "").lower())
    for token in ["주식회사", "(주)", "㈜"]:
        text = text.replace(token, "")
    return re.sub(r"[^\w가-힣]", "", text)


async def _read_sheet_csv(client: httpx.AsyncClient, gid: str, header: int) -> pd.DataFrame:
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export"
    try:
        response = await client.get(url, params={"format": "csv", "gid": gid})
    except httpx.RequestError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Google Sheet CSV export에 연결할 수 없습니다. 서버 네트워크 접근을 확인해주세요.",
        ) from exc
    if response.status_code in {401, 403}:
        raise HTTPException(
            status_code=status.HTTP_424_FAILED_DEPENDENCY,
            detail="Google Sheet CSV export 권한이 없습니다. 시트를 공개 또는 서버 접근 가능 상태로 전환해주세요.",
        )
    if response.status_code >= 400:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Google Sheet CSV export 실패: HTTP {response.status_code}",
        )
    return pd.read_csv(BytesIO(response.content), header=header, dtype=str).fillna("")


async def _read_upload_file(file: UploadFile) -> pd.DataFrame:
    content = await file.read()
    file_name = (file.filename or "").lower()
    try:
        if file_name.endswith(".csv"):
            try:
                return pd.read_csv(BytesIO(content), dtype=str).fillna("")
            except UnicodeDecodeError:
                return pd.read_csv(BytesIO(content), dtype=str, encoding="cp949").fillna("")
        if file_name.endswith((".xlsx", ".xls")):
            return pd.read_excel(BytesIO(content), dtype=str).fillna("")
    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="은행 로그인 실적파일을 읽을 수 없습니다. xlsx, xls, csv 형식을 확인해주세요.",
        ) from exc
    raise HTTPException(
        status_code=status.HTTP_400_BAD_REQUEST,
        detail="지원하지 않는 파일 형식입니다. xlsx, xls, csv 파일을 업로드해주세요.",
    )


def _row_value(row: pd.Series, candidates: list[str]) -> str:
    for candidate in candidates:
        if candidate in row.index:
            return _text(row.get(candidate))
    return ""


def _build_login_map(login_df: pd.DataFrame) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for _, row in login_df.iterrows():
        customer_number = _digits(row.get("고객번호"))
        if not customer_number:
            continue
        result[customer_number] = {
            "customer_number": customer_number,
            "company_name": _text(row.get("고객명")),
            "latest_login": _text(row.get("최근로그인")),
            "login_count": str(_int(row.get("로그인"))),
        }
    return result


def _build_customer_maps(customer_df: pd.DataFrame) -> tuple[dict[str, dict[str, str]], dict[str, dict[str, str]]]:
    by_biz: dict[str, dict[str, str]] = {}
    by_customer: dict[str, dict[str, str]] = {}
    for _, row in customer_df.iterrows():
        customer_number = _digits(row.get("고객번호"))
        business_number = _digits(row.get("사업자등록번호"))
        item = {
            "customer_number": customer_number,
            "business_number": business_number,
            "company_name": _text(row.get("고객명")),
            "first_login": _text(row.get("최초신규일자")),
            "latest_login": _text(row.get("최종로그인일자")),
        }
        if business_number:
            by_biz.setdefault(business_number, item)
        if customer_number:
            by_customer.setdefault(customer_number, item)
    return by_biz, by_customer


def _billing_rows(
    billing_df: pd.DataFrame,
    source_type: str,
    login_by_customer: dict[str, dict[str, str]],
    customer_by_biz: dict[str, dict[str, str]],
    customer_by_customer: dict[str, dict[str, str]],
) -> list[BillingPreviewRow]:
    rows: list[BillingPreviewRow] = []
    for _, row in billing_df.iterrows():
        sequence = _row_value(row, ["순번", "순서"])
        customer_number = _digits(_row_value(row, ["고객번호"]))
        business_number = _digits(_row_value(row, ["사업자번호", "사업자등록번호"]))
        company_name = _row_value(row, ["업체명", "고객명"])
        if not any([sequence, customer_number, business_number, company_name]):
            continue
        manager_name = _row_value(row, ["담당자"])
        base_date = _row_value(row, ["구축일자", "구축일", "은행연계완료일자", "연계시작일자"])

        customer_info = customer_by_biz.get(business_number) or customer_by_customer.get(customer_number) or {}
        matched_customer_number = customer_info.get("customer_number") or customer_number
        login_info = login_by_customer.get(matched_customer_number, {})

        bank_company_name = login_info.get("company_name") or customer_info.get("company_name", "")
        first_login = customer_info.get("first_login") or _row_value(row, ["최초로그인"])
        latest_login = (
            login_info.get("latest_login")
            or customer_info.get("latest_login")
            or _row_value(row, ["최종로그인일자"])
        )
        login_count = _int(login_info.get("login_count") or _row_value(row, ["로그인횟수"]))

        if not business_number:
            match_status = "사업자번호 없음"
        elif not customer_info and not login_info:
            match_status = "실적 없음"
        elif (
            _company_key(company_name)
            and _company_key(bank_company_name)
            and _company_key(company_name) != _company_key(bank_company_name)
        ):
            match_status = "고객명 상이"
        else:
            match_status = "일치"

        rows.append(
            BillingPreviewRow(
                source_type=source_type,
                sequence=sequence,
                customer_number=matched_customer_number,
                business_number=business_number or customer_info.get("business_number", ""),
                company_name=company_name,
                manager_name=manager_name,
                base_date=base_date,
                first_login=first_login,
                latest_login=latest_login,
                login_count=login_count,
                billing_company_name=company_name,
                bank_company_name=bank_company_name,
                match_status=match_status,
                note="은행 로그인 실적과 대사 완료" if match_status == "일치" else "확인 필요",
            )
        )
    return rows


def _build_preview(
    open_df: pd.DataFrame,
    erp_df: pd.DataFrame,
    login_df: pd.DataFrame,
    customer_df: pd.DataFrame,
    login_source_title: str,
) -> BillingPreview:
    login_by_customer = _build_login_map(login_df)
    customer_by_biz, customer_by_customer = _build_customer_maps(customer_df)

    open_rows = _billing_rows(open_df, "개설", login_by_customer, customer_by_biz, customer_by_customer)
    erp_rows = _billing_rows(erp_df, "연계", login_by_customer, customer_by_biz, customer_by_customer)
    rows = open_rows + erp_rows

    matched_count = sum(1 for row in rows if row.match_status == "일치")
    name_mismatch_count = sum(1 for row in rows if row.match_status == "고객명 상이")
    missing_count = sum(1 for row in rows if row.match_status in {"실적 없음", "사업자번호 없음"})

    return BillingPreview(
        spreadsheet_title=SPREADSHEET_TITLE,
        spreadsheet_url=SPREADSHEET_URL,
        generated_from=[SHEETS["open"]["title"], SHEETS["erp"]["title"], login_source_title, SHEETS["customer"]["title"]],
        rows=rows,
        summary=BillingPreviewSummary(
            total_count=len(rows),
            matched_count=matched_count,
            name_mismatch_count=name_mismatch_count,
            missing_count=missing_count,
            open_count=len(open_rows),
            erp_count=len(erp_rows),
        ),
    )


async def _read_sheet_sources(include_login: bool) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame | None, pd.DataFrame]:
    async with httpx.AsyncClient(timeout=30.0, follow_redirects=True) as client:
        open_df = await _read_sheet_csv(client, SHEETS["open"]["gid"], SHEETS["open"]["header"])
        erp_df = await _read_sheet_csv(client, SHEETS["erp"]["gid"], SHEETS["erp"]["header"])
        login_df = await _read_sheet_csv(client, SHEETS["login"]["gid"], SHEETS["login"]["header"]) if include_login else None
        customer_df = await _read_sheet_csv(client, SHEETS["customer"]["gid"], SHEETS["customer"]["header"])
    return open_df, erp_df, login_df, customer_df


@router.get("/preview", response_model=BillingPreview)
async def preview_billing() -> BillingPreview:
    open_df, erp_df, login_df, customer_df = await _read_sheet_sources(include_login=True)
    if login_df is None:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail="은행 로그인 실적파일을 불러오지 못했습니다.")
    return _build_preview(open_df, erp_df, login_df, customer_df, SHEETS["login"]["title"])


@router.post("/preview", response_model=BillingPreview)
async def preview_billing_with_uploaded_login_file(
    login_file: UploadFile = File(...),
) -> BillingPreview:
    login_df = await _read_upload_file(login_file)
    open_df, erp_df, _, customer_df = await _read_sheet_sources(include_login=False)
    return _build_preview(open_df, erp_df, login_df, customer_df, f"업로드: {login_file.filename or '은행로그인실적파일'}")
