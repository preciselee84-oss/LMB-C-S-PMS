from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import func, select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.api.deps import get_db
from app.models.sales import PaymentMatch, SalesLead
from app.schemas.sales import (
    BankTransactionIn,
    PaymentMatchRead,
    PipelineSummary,
    SalesLeadCreate,
    SalesLeadRead,
)
from app.services.sales_matching import match_bank_transaction

router = APIRouter()


@router.post("/leads", response_model=SalesLeadRead, status_code=status.HTTP_201_CREATED)
async def create_lead(payload: SalesLeadCreate, db: AsyncSession = Depends(get_db)) -> SalesLead:
    lead = SalesLead(**payload.model_dump())
    db.add(lead)
    try:
        await db.commit()
    except IntegrityError as exc:
        await db.rollback()
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="Customer is already claimed",
        ) from exc
    await db.refresh(lead)
    return lead


@router.get("/leads", response_model=list[SalesLeadRead])
async def list_leads(db: AsyncSession = Depends(get_db)) -> list[SalesLead]:
    result = await db.execute(select(SalesLead).order_by(SalesLead.claimed_at.desc()))
    return list(result.scalars().all())


@router.post("/transactions/match", response_model=PaymentMatchRead | None)
async def match_transaction(
    payload: BankTransactionIn, db: AsyncSession = Depends(get_db)
) -> PaymentMatchRead | None:
    match = await match_bank_transaction(db, payload)
    if not match:
        return None

    result = await db.execute(
        select(PaymentMatch)
        .options(selectinload(PaymentMatch.lead))
        .where(PaymentMatch.id == match.id)
    )
    hydrated = result.scalar_one()
    return PaymentMatchRead(
        id=hydrated.id,
        sales_lead_id=hydrated.sales_lead_id,
        customer_name=hydrated.lead.customer_name,
        owner_name=hydrated.lead.owner_name,
        depositor_name=hydrated.depositor_name,
        amount=hydrated.amount,
        matched_rule=hydrated.matched_rule,
        confidence=float(hydrated.confidence),
        created_at=hydrated.created_at,
    )


@router.get("/matches", response_model=list[PaymentMatchRead])
async def list_matches(db: AsyncSession = Depends(get_db)) -> list[PaymentMatchRead]:
    result = await db.execute(
        select(PaymentMatch).options(selectinload(PaymentMatch.lead)).order_by(PaymentMatch.created_at.desc())
    )
    matches = result.scalars().all()
    return [
        PaymentMatchRead(
            id=match.id,
            sales_lead_id=match.sales_lead_id,
            customer_name=match.lead.customer_name,
            owner_name=match.lead.owner_name,
            depositor_name=match.depositor_name,
            amount=match.amount,
            matched_rule=match.matched_rule,
            confidence=float(match.confidence),
            created_at=match.created_at,
        )
        for match in matches
    ]


@router.get("/summary", response_model=PipelineSummary)
async def get_summary(db: AsyncSession = Depends(get_db)) -> PipelineSummary:
    total_leads = await db.scalar(select(func.count()).select_from(SalesLead))
    waiting_payment = await db.scalar(select(func.count()).select_from(SalesLead).where(SalesLead.status != "paid"))
    paid = await db.scalar(select(func.count()).select_from(SalesLead).where(SalesLead.status == "paid"))
    total_expected = await db.scalar(select(func.coalesce(func.sum(SalesLead.expected_amount), 0)))
    total_paid = await db.scalar(select(func.coalesce(func.sum(PaymentMatch.amount), 0)))
    risk_result = await db.execute(
        select(SalesLead)
        .where(SalesLead.status != "paid")
        .order_by(SalesLead.claimed_at.asc())
        .limit(5)
    )

    return PipelineSummary(
        total_leads=total_leads or 0,
        waiting_payment=waiting_payment or 0,
        paid=paid or 0,
        total_expected_amount=total_expected or 0,
        total_paid_amount=total_paid or 0,
        overdue_risk=list(risk_result.scalars().all()),
    )
