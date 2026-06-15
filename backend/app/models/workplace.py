from datetime import datetime

from sqlalchemy import BigInteger, DateTime, ForeignKey, Integer, String, Text, UniqueConstraint, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.models.base import Base


class CompanyProfile(Base):
    __tablename__ = "company_profiles"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(160))
    business_number: Mapped[str | None] = mapped_column(String(30), nullable=True)
    ceo_name: Mapped[str | None] = mapped_column(String(100), nullable=True)
    address: Mapped[str | None] = mapped_column(String(255), nullable=True)
    contact: Mapped[str | None] = mapped_column(String(100), nullable=True)
    memo: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )


class DelegatedWorkplace(Base):
    __tablename__ = "delegated_workplaces"
    __table_args__ = (UniqueConstraint("workplace_name", name="uq_delegated_workplaces_workplace_name"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    workplace_name: Mapped[str] = mapped_column(String(160), index=True)
    business_number: Mapped[str | None] = mapped_column(String(30), nullable=True)
    business_alias: Mapped[str | None] = mapped_column(String(100), nullable=True)
    regular_payment_day: Mapped[int] = mapped_column(Integer, default=0)
    manager_name: Mapped[str | None] = mapped_column(String(100), nullable=True)
    manager_contact: Mapped[str | None] = mapped_column(String(100), nullable=True)
    memo: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )

    accounts: Mapped[list["BankAccount"]] = relationship(
        back_populates="workplace", cascade="all, delete-orphan"
    )
    advance_requests: Mapped[list["AdvancePaymentRequest"]] = relationship(
        back_populates="workplace", cascade="all, delete-orphan"
    )


class BankAccount(Base):
    __tablename__ = "bank_accounts"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    account_type: Mapped[str] = mapped_column(String(50), default="위탁사업장")
    account_name: Mapped[str] = mapped_column(String(160))
    bank_name: Mapped[str] = mapped_column(String(80))
    account_number: Mapped[str] = mapped_column(String(80))
    holder_name: Mapped[str | None] = mapped_column(String(100), nullable=True)
    linked_workplace_id: Mapped[int | None] = mapped_column(
        ForeignKey("delegated_workplaces.id", ondelete="SET NULL"), nullable=True, index=True
    )
    balance: Mapped[int] = mapped_column(BigInteger, default=0)
    memo: Mapped[str | None] = mapped_column(Text, nullable=True)
    balance_updated_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )

    workplace: Mapped[DelegatedWorkplace | None] = relationship(back_populates="accounts")


class AdvancePaymentRequest(Base):
    __tablename__ = "advance_payment_requests"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    workplace_id: Mapped[int] = mapped_column(ForeignKey("delegated_workplaces.id"), index=True)
    request_amount: Mapped[int] = mapped_column(BigInteger)
    requested_by: Mapped[str] = mapped_column(String(100), index=True)
    request_reason: Mapped[str | None] = mapped_column(Text, nullable=True)
    status: Mapped[str] = mapped_column(String(30), default="요청", index=True)
    reject_reason: Mapped[str | None] = mapped_column(Text, nullable=True)
    approved_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    processed_by: Mapped[str | None] = mapped_column(String(100), nullable=True)
    transfer_generated_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    paid_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    requested_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )

    workplace: Mapped[DelegatedWorkplace] = relationship(back_populates="advance_requests")
