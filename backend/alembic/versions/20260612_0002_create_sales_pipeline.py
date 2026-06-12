"""create sales pipeline tables

Revision ID: 20260612_0002
Revises: 20260608_0001
Create Date: 2026-06-12
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "20260612_0002"
down_revision: str | None = "20260608_0001"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "sales_leads",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("customer_name", sa.String(length=160), nullable=False),
        sa.Column("owner_name", sa.String(length=100), nullable=False),
        sa.Column("owner_contact", sa.String(length=100), nullable=True),
        sa.Column("meeting_note", sa.Text(), nullable=True),
        sa.Column("expected_amount", sa.BigInteger(), nullable=False),
        sa.Column("status", sa.String(length=30), nullable=False, server_default="lead"),
        sa.Column("claimed_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.UniqueConstraint("customer_name", name="uq_sales_leads_customer_name"),
    )
    op.create_index("ix_sales_leads_customer_name", "sales_leads", ["customer_name"])
    op.create_index("ix_sales_leads_owner_name", "sales_leads", ["owner_name"])
    op.create_index("ix_sales_leads_status", "sales_leads", ["status"])

    op.create_table(
        "payment_matches",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("sales_lead_id", sa.Integer(), sa.ForeignKey("sales_leads.id"), nullable=False),
        sa.Column("depositor_name", sa.String(length=160), nullable=False),
        sa.Column("amount", sa.BigInteger(), nullable=False),
        sa.Column("matched_rule", sa.String(length=30), nullable=False),
        sa.Column("confidence", sa.Numeric(5, 2), nullable=False, server_default="100"),
        sa.Column("notified_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("transaction_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
    )
    op.create_index("ix_payment_matches_sales_lead_id", "payment_matches", ["sales_lead_id"])
    op.create_index("ix_payment_matches_depositor_name", "payment_matches", ["depositor_name"])


def downgrade() -> None:
    op.drop_index("ix_payment_matches_depositor_name", table_name="payment_matches")
    op.drop_index("ix_payment_matches_sales_lead_id", table_name="payment_matches")
    op.drop_table("payment_matches")
    op.drop_index("ix_sales_leads_status", table_name="sales_leads")
    op.drop_index("ix_sales_leads_owner_name", table_name="sales_leads")
    op.drop_index("ix_sales_leads_customer_name", table_name="sales_leads")
    op.drop_table("sales_leads")
