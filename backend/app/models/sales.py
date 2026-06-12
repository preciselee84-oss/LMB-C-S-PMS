from datetime import datetime

from sqlalchemy import BigInteger, DateTime, ForeignKey, Integer, Numeric, String, Text, UniqueConstraint, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.models.base import Base


class SalesLead(Base):
    __tablename__ = "sales_leads"
    __table_args__ = (UniqueConstraint("customer_name", name="uq_sales_leads_customer_name"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    customer_name: Mapped[str] = mapped_column(String(160), index=True)
    owner_name: Mapped[str] = mapped_column(String(100), index=True)
    owner_contact: Mapped[str | None] = mapped_column(String(100), nullable=True)
    meeting_note: Mapped[str | None] = mapped_column(Text, nullable=True)
    expected_amount: Mapped[int] = mapped_column(BigInteger)
    status: Mapped[str] = mapped_column(String(30), default="lead", index=True)
    claimed_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )

    matches: Mapped[list["PaymentMatch"]] = relationship(back_populates="lead")


class PaymentMatch(Base):
    __tablename__ = "payment_matches"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    sales_lead_id: Mapped[int] = mapped_column(ForeignKey("sales_leads.id"), index=True)
    depositor_name: Mapped[str] = mapped_column(String(160), index=True)
    amount: Mapped[int] = mapped_column(BigInteger)
    matched_rule: Mapped[str] = mapped_column(String(30))
    confidence: Mapped[float] = mapped_column(Numeric(5, 2), default=100)
    notified_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    transaction_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())

    lead: Mapped[SalesLead] = relationship(back_populates="matches")
