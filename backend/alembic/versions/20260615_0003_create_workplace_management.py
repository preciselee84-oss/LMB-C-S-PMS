"""create workplace management tables

Revision ID: 20260615_0003
Revises: 20260612_0002
Create Date: 2026-06-15
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "20260615_0003"
down_revision: str | None = "20260612_0002"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "company_profiles",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("name", sa.String(length=160), nullable=False),
        sa.Column("business_number", sa.String(length=30), nullable=True),
        sa.Column("ceo_name", sa.String(length=100), nullable=True),
        sa.Column("address", sa.String(length=255), nullable=True),
        sa.Column("contact", sa.String(length=100), nullable=True),
        sa.Column("memo", sa.Text(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
    )

    op.create_table(
        "delegated_workplaces",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("workplace_name", sa.String(length=160), nullable=False),
        sa.Column("business_number", sa.String(length=30), nullable=True),
        sa.Column("business_alias", sa.String(length=100), nullable=True),
        sa.Column("regular_payment_day", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("manager_name", sa.String(length=100), nullable=True),
        sa.Column("manager_contact", sa.String(length=100), nullable=True),
        sa.Column("memo", sa.Text(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.UniqueConstraint("workplace_name", name="uq_delegated_workplaces_workplace_name"),
    )
    op.create_index("ix_delegated_workplaces_workplace_name", "delegated_workplaces", ["workplace_name"])

    op.create_table(
        "bank_accounts",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("account_type", sa.String(length=50), nullable=False, server_default="위탁사업장"),
        sa.Column("account_name", sa.String(length=160), nullable=False),
        sa.Column("bank_name", sa.String(length=80), nullable=False),
        sa.Column("account_number", sa.String(length=80), nullable=False),
        sa.Column("holder_name", sa.String(length=100), nullable=True),
        sa.Column(
            "linked_workplace_id",
            sa.Integer(),
            sa.ForeignKey("delegated_workplaces.id", ondelete="SET NULL"),
            nullable=True,
        ),
        sa.Column("balance", sa.BigInteger(), nullable=False, server_default="0"),
        sa.Column("memo", sa.Text(), nullable=True),
        sa.Column("balance_updated_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
    )
    op.create_index("ix_bank_accounts_linked_workplace_id", "bank_accounts", ["linked_workplace_id"])

    op.create_table(
        "advance_payment_requests",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("workplace_id", sa.Integer(), sa.ForeignKey("delegated_workplaces.id"), nullable=False),
        sa.Column("request_amount", sa.BigInteger(), nullable=False),
        sa.Column("requested_by", sa.String(length=100), nullable=False),
        sa.Column("request_reason", sa.Text(), nullable=True),
        sa.Column("status", sa.String(length=30), nullable=False, server_default="요청"),
        sa.Column("reject_reason", sa.Text(), nullable=True),
        sa.Column("approved_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("processed_by", sa.String(length=100), nullable=True),
        sa.Column("transfer_generated_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("paid_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("requested_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
    )
    op.create_index("ix_advance_payment_requests_workplace_id", "advance_payment_requests", ["workplace_id"])
    op.create_index("ix_advance_payment_requests_requested_by", "advance_payment_requests", ["requested_by"])
    op.create_index("ix_advance_payment_requests_status", "advance_payment_requests", ["status"])


def downgrade() -> None:
    op.drop_index("ix_advance_payment_requests_status", table_name="advance_payment_requests")
    op.drop_index("ix_advance_payment_requests_requested_by", table_name="advance_payment_requests")
    op.drop_index("ix_advance_payment_requests_workplace_id", table_name="advance_payment_requests")
    op.drop_table("advance_payment_requests")
    op.drop_index("ix_bank_accounts_linked_workplace_id", table_name="bank_accounts")
    op.drop_table("bank_accounts")
    op.drop_index("ix_delegated_workplaces_workplace_name", table_name="delegated_workplaces")
    op.drop_table("delegated_workplaces")
    op.drop_table("company_profiles")
