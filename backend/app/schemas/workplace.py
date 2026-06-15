from datetime import datetime

from pydantic import BaseModel, Field


class CompanyProfileBase(BaseModel):
    name: str = Field(min_length=1, max_length=160)
    business_number: str | None = Field(default=None, max_length=30)
    ceo_name: str | None = Field(default=None, max_length=100)
    address: str | None = Field(default=None, max_length=255)
    contact: str | None = Field(default=None, max_length=100)
    memo: str | None = None


class CompanyProfileRead(CompanyProfileBase):
    id: int
    updated_at: datetime

    model_config = {"from_attributes": True}


class WorkplaceBase(BaseModel):
    workplace_name: str = Field(min_length=1, max_length=160)
    business_number: str | None = Field(default=None, max_length=30)
    business_alias: str | None = Field(default=None, max_length=100)
    regular_payment_day: int = Field(default=0, ge=0, le=31)
    manager_name: str | None = Field(default=None, max_length=100)
    manager_contact: str | None = Field(default=None, max_length=100)
    memo: str | None = None


class WorkplaceCreate(WorkplaceBase):
    pass


class WorkplaceUpdate(BaseModel):
    workplace_name: str | None = Field(default=None, min_length=1, max_length=160)
    business_number: str | None = Field(default=None, max_length=30)
    business_alias: str | None = Field(default=None, max_length=100)
    regular_payment_day: int | None = Field(default=None, ge=0, le=31)
    manager_name: str | None = Field(default=None, max_length=100)
    manager_contact: str | None = Field(default=None, max_length=100)
    memo: str | None = None


class WorkplaceRead(WorkplaceBase):
    id: int
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}


class BankAccountBase(BaseModel):
    account_type: str = Field(default="위탁사업장", max_length=50)
    account_name: str = Field(min_length=1, max_length=160)
    bank_name: str = Field(min_length=1, max_length=80)
    account_number: str = Field(min_length=1, max_length=80)
    holder_name: str | None = Field(default=None, max_length=100)
    linked_workplace_id: int | None = None
    balance: int = Field(default=0, ge=0)
    memo: str | None = None


class BankAccountCreate(BankAccountBase):
    pass


class BankAccountUpdate(BaseModel):
    account_type: str | None = Field(default=None, max_length=50)
    account_name: str | None = Field(default=None, min_length=1, max_length=160)
    bank_name: str | None = Field(default=None, min_length=1, max_length=80)
    account_number: str | None = Field(default=None, min_length=1, max_length=80)
    holder_name: str | None = Field(default=None, max_length=100)
    linked_workplace_id: int | None = None
    balance: int | None = Field(default=None, ge=0)
    memo: str | None = None


class BankAccountRead(BankAccountBase):
    id: int
    balance_updated_at: datetime | None
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}


class AdvanceRequestCreate(BaseModel):
    workplace_id: int
    request_amount: int = Field(gt=0)
    requested_by: str = Field(min_length=1, max_length=100)
    request_reason: str | None = None


class AdvanceRequestDecision(BaseModel):
    processed_by: str = Field(min_length=1, max_length=100)
    reject_reason: str | None = None


class AdvanceRequestRead(BaseModel):
    id: int
    workplace_id: int
    workplace_name: str
    request_amount: int
    requested_by: str
    request_reason: str | None
    status: str
    reject_reason: str | None
    approved_at: datetime | None
    processed_by: str | None
    transfer_generated_at: datetime | None
    paid_at: datetime | None
    requested_at: datetime


class TransferRow(BaseModel):
    request_id: int
    workplace_name: str
    bank_name: str
    account_number: str
    holder_name: str
    amount: int
    memo: str


class WorkplaceForecast(BaseModel):
    workplace_id: int
    workplace_name: str
    average_monthly_amount: int
    suggested_amount: int
    guide: str


class WorkplaceSummary(BaseModel):
    workplace_count: int
    request_count: int
    pending_count: int
    approved_count: int
    paid_count: int
    paid_amount: int
    month_request_count: int
    accounts_balance: int
    forecasts: list[WorkplaceForecast]
