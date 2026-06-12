from decimal import Decimal

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models.sales import PaymentMatch, SalesLead
from app.schemas.sales import BankTransactionIn


def normalize_name(value: str) -> str:
    return "".join(value.lower().split())


def amount_matches(expected_amount: int, paid_amount: int) -> tuple[bool, str | None]:
    if paid_amount == expected_amount:
        return True, "exact"

    vat_amount = int(Decimal(expected_amount) * Decimal("1.1"))
    if paid_amount == vat_amount:
        return True, "vat_included"

    return False, None


async def match_bank_transaction(
    db: AsyncSession, transaction: BankTransactionIn
) -> PaymentMatch | None:
    result = await db.execute(select(SalesLead).where(SalesLead.status != "paid"))
    leads = result.scalars().all()
    depositor = normalize_name(transaction.depositor_name)

    for lead in leads:
        customer = normalize_name(lead.customer_name)
        name_match = customer in depositor or depositor in customer
        amount_match, rule = amount_matches(lead.expected_amount, transaction.amount)

        if name_match and amount_match and rule:
            match = PaymentMatch(
                sales_lead_id=lead.id,
                depositor_name=transaction.depositor_name,
                amount=transaction.amount,
                matched_rule=rule,
                confidence=100 if rule == "exact" else 95,
                transaction_at=transaction.transaction_at,
            )
            lead.status = "paid"
            db.add(match)
            await db.commit()
            await db.refresh(match)
            return match

    return None
