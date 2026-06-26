from pydantic import BaseModel


class BillingPreviewRow(BaseModel):
    source_type: str
    sequence: str
    customer_number: str
    business_number: str
    company_name: str
    manager_name: str
    base_date: str
    first_login: str
    latest_login: str
    login_count: int
    billing_company_name: str
    bank_company_name: str
    match_status: str
    note: str


class BillingPreviewSummary(BaseModel):
    total_count: int
    matched_count: int
    name_mismatch_count: int
    missing_count: int
    open_count: int
    erp_count: int


class BillingPreview(BaseModel):
    spreadsheet_title: str
    spreadsheet_url: str
    generated_from: list[str]
    rows: list[BillingPreviewRow]
    summary: BillingPreviewSummary
