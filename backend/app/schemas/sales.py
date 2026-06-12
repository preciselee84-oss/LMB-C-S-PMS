from datetime import datetime

from pydantic import BaseModel, Field


class SalesLeadCreate(BaseModel):
    customer_name: str = Field(min_length=1, max_length=160)
    owner_name: str = Field(min_length=1, max_length=100)
    owner_contact: str | None = Field(default=None, max_length=100)
    meeting_note: str | None = None
    expected_amount: int = Field(gt=0)


class SalesLeadRead(BaseModel):
    id: int
    customer_name: str
    owner_name: str
    owner_contact: str | None
    meeting_note: str | None
    expected_amount: int
    status: str
    claimed_at: datetime

    model_config = {"from_attributes": True}


class BankTransactionIn(BaseModel):
    depositor_name: str = Field(min_length=1, max_length=160)
    amount: int = Field(gt=0)
    transaction_at: datetime | None = None


class PaymentMatchRead(BaseModel):
    id: int
    sales_lead_id: int
    customer_name: str
    owner_name: str
    depositor_name: str
    amount: int
    matched_rule: str
    confidence: float
    created_at: datetime


class PipelineSummary(BaseModel):
    total_leads: int
    waiting_payment: int
    paid: int
    total_expected_amount: int
    total_paid_amount: int
    overdue_risk: list[SalesLeadRead]
