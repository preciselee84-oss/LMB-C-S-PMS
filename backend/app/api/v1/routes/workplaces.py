from collections import defaultdict
from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import func, select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.api.deps import get_db
from app.models.workplace import AdvancePaymentRequest, BankAccount, CompanyProfile, DelegatedWorkplace
from app.schemas.workplace import (
    AdvanceRequestCreate,
    AdvanceRequestDecision,
    AdvanceRequestRead,
    BankAccountCreate,
    BankAccountRead,
    BankAccountUpdate,
    CompanyProfileBase,
    CompanyProfileRead,
    TransferRow,
    WorkplaceCreate,
    WorkplaceForecast,
    WorkplaceRead,
    WorkplaceSummary,
    WorkplaceUpdate,
)

router = APIRouter()


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _request_read(row: AdvancePaymentRequest) -> AdvanceRequestRead:
    return AdvanceRequestRead(
        id=row.id,
        workplace_id=row.workplace_id,
        workplace_name=row.workplace.workplace_name,
        request_amount=row.request_amount,
        requested_by=row.requested_by,
        request_reason=row.request_reason,
        status=row.status,
        reject_reason=row.reject_reason,
        approved_at=row.approved_at,
        processed_by=row.processed_by,
        transfer_generated_at=row.transfer_generated_at,
        paid_at=row.paid_at,
        requested_at=row.requested_at,
    )


async def _get_workplace_or_404(db: AsyncSession, workplace_id: int) -> DelegatedWorkplace:
    workplace = await db.get(DelegatedWorkplace, workplace_id)
    if not workplace:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Workplace not found")
    return workplace


async def _get_request_or_404(db: AsyncSession, request_id: int) -> AdvancePaymentRequest:
    result = await db.execute(
        select(AdvancePaymentRequest)
        .options(selectinload(AdvancePaymentRequest.workplace))
        .where(AdvancePaymentRequest.id == request_id)
    )
    request = result.scalar_one_or_none()
    if not request:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Advance payment request not found")
    return request


@router.get("/company", response_model=CompanyProfileRead | None)
async def get_company_profile(db: AsyncSession = Depends(get_db)) -> CompanyProfile | None:
    result = await db.execute(select(CompanyProfile).order_by(CompanyProfile.id.asc()).limit(1))
    return result.scalar_one_or_none()


@router.put("/company", response_model=CompanyProfileRead)
async def upsert_company_profile(payload: CompanyProfileBase, db: AsyncSession = Depends(get_db)) -> CompanyProfile:
    result = await db.execute(select(CompanyProfile).order_by(CompanyProfile.id.asc()).limit(1))
    profile = result.scalar_one_or_none()
    if profile is None:
        profile = CompanyProfile(**payload.model_dump())
        db.add(profile)
    else:
        for key, value in payload.model_dump().items():
            setattr(profile, key, value)
    await db.commit()
    await db.refresh(profile)
    return profile


@router.get("", response_model=list[WorkplaceRead])
async def list_workplaces(db: AsyncSession = Depends(get_db)) -> list[DelegatedWorkplace]:
    result = await db.execute(select(DelegatedWorkplace).order_by(DelegatedWorkplace.workplace_name.asc()))
    return list(result.scalars().all())


@router.post("", response_model=WorkplaceRead, status_code=status.HTTP_201_CREATED)
async def create_workplace(payload: WorkplaceCreate, db: AsyncSession = Depends(get_db)) -> DelegatedWorkplace:
    workplace = DelegatedWorkplace(**payload.model_dump())
    db.add(workplace)
    try:
        await db.commit()
    except IntegrityError as exc:
        await db.rollback()
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Workplace name already exists") from exc
    await db.refresh(workplace)
    return workplace


@router.get("/accounts", response_model=list[BankAccountRead])
async def list_accounts(db: AsyncSession = Depends(get_db)) -> list[BankAccount]:
    result = await db.execute(select(BankAccount).order_by(BankAccount.created_at.desc()))
    return list(result.scalars().all())


@router.post("/accounts", response_model=BankAccountRead, status_code=status.HTTP_201_CREATED)
async def create_account(payload: BankAccountCreate, db: AsyncSession = Depends(get_db)) -> BankAccount:
    if payload.linked_workplace_id:
        await _get_workplace_or_404(db, payload.linked_workplace_id)
    account = BankAccount(**payload.model_dump())
    if payload.balance:
        account.balance_updated_at = _now()
    db.add(account)
    await db.commit()
    await db.refresh(account)
    return account


@router.patch("/accounts/{account_id}", response_model=BankAccountRead)
async def update_account(
    account_id: int, payload: BankAccountUpdate, db: AsyncSession = Depends(get_db)
) -> BankAccount:
    account = await db.get(BankAccount, account_id)
    if not account:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Bank account not found")
    values = payload.model_dump(exclude_unset=True)
    if values.get("linked_workplace_id"):
        await _get_workplace_or_404(db, values["linked_workplace_id"])
    if "balance" in values:
        account.balance_updated_at = _now()
    for key, value in values.items():
        setattr(account, key, value)
    await db.commit()
    await db.refresh(account)
    return account


@router.delete("/accounts/{account_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_account(account_id: int, db: AsyncSession = Depends(get_db)) -> None:
    account = await db.get(BankAccount, account_id)
    if not account:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Bank account not found")
    await db.delete(account)
    await db.commit()


@router.get("/requests", response_model=list[AdvanceRequestRead])
async def list_advance_requests(db: AsyncSession = Depends(get_db)) -> list[AdvanceRequestRead]:
    result = await db.execute(
        select(AdvancePaymentRequest)
        .options(selectinload(AdvancePaymentRequest.workplace))
        .order_by(AdvancePaymentRequest.requested_at.desc())
    )
    return [_request_read(row) for row in result.scalars().all()]


@router.post("/requests", response_model=AdvanceRequestRead, status_code=status.HTTP_201_CREATED)
async def create_advance_request(
    payload: AdvanceRequestCreate, db: AsyncSession = Depends(get_db)
) -> AdvanceRequestRead:
    await _get_workplace_or_404(db, payload.workplace_id)
    request = AdvancePaymentRequest(**payload.model_dump())
    db.add(request)
    await db.commit()
    result = await db.execute(
        select(AdvancePaymentRequest)
        .options(selectinload(AdvancePaymentRequest.workplace))
        .where(AdvancePaymentRequest.id == request.id)
    )
    return _request_read(result.scalar_one())


@router.post("/requests/{request_id}/approve", response_model=AdvanceRequestRead)
async def approve_advance_request(
    request_id: int, payload: AdvanceRequestDecision, db: AsyncSession = Depends(get_db)
) -> AdvanceRequestRead:
    request = await _get_request_or_404(db, request_id)
    if request.status != "요청":
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Only pending requests can be approved")
    request.status = "품의 확정"
    request.processed_by = payload.processed_by
    request.approved_at = _now()
    request.reject_reason = None
    await db.commit()
    await db.refresh(request)
    return _request_read(request)


@router.post("/requests/{request_id}/reject", response_model=AdvanceRequestRead)
async def reject_advance_request(
    request_id: int, payload: AdvanceRequestDecision, db: AsyncSession = Depends(get_db)
) -> AdvanceRequestRead:
    request = await _get_request_or_404(db, request_id)
    if request.status != "요청":
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Only pending requests can be rejected")
    request.status = "반려"
    request.processed_by = payload.processed_by
    request.reject_reason = payload.reject_reason
    await db.commit()
    await db.refresh(request)
    return _request_read(request)


@router.post("/requests/{request_id}/transfer", response_model=TransferRow)
async def generate_transfer_row(request_id: int, db: AsyncSession = Depends(get_db)) -> TransferRow:
    request = await _get_request_or_404(db, request_id)
    if request.status != "품의 확정":
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Only approved requests can be transferred")
    account_result = await db.execute(
        select(BankAccount)
        .where(BankAccount.linked_workplace_id == request.workplace_id)
        .order_by(BankAccount.created_at.desc())
        .limit(1)
    )
    account = account_result.scalar_one_or_none()
    if not account:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Linked workplace account is required")

    request.status = "이체 대상"
    request.transfer_generated_at = _now()
    await db.commit()
    return TransferRow(
        request_id=request.id,
        workplace_name=request.workplace.workplace_name,
        bank_name=account.bank_name,
        account_number=account.account_number,
        holder_name=account.holder_name or request.workplace.workplace_name,
        amount=request.request_amount,
        memo="전도금 지급",
    )


@router.post("/requests/{request_id}/paid", response_model=AdvanceRequestRead)
async def mark_request_paid(request_id: int, db: AsyncSession = Depends(get_db)) -> AdvanceRequestRead:
    request = await _get_request_or_404(db, request_id)
    if request.status not in {"품의 확정", "이체 대상"}:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Only approved or transfer-ready requests can be paid")
    request.status = "이체 완료"
    request.paid_at = _now()
    await db.commit()
    await db.refresh(request)
    return _request_read(request)


@router.get("/summary", response_model=WorkplaceSummary)
async def get_workplace_summary(db: AsyncSession = Depends(get_db)) -> WorkplaceSummary:
    now = _now()
    workplace_count = await db.scalar(select(func.count()).select_from(DelegatedWorkplace))
    request_count = await db.scalar(select(func.count()).select_from(AdvancePaymentRequest))
    pending_count = await db.scalar(
        select(func.count()).select_from(AdvancePaymentRequest).where(AdvancePaymentRequest.status == "요청")
    )
    approved_count = await db.scalar(
        select(func.count()).select_from(AdvancePaymentRequest).where(AdvancePaymentRequest.status == "품의 확정")
    )
    paid_count = await db.scalar(
        select(func.count()).select_from(AdvancePaymentRequest).where(AdvancePaymentRequest.status == "이체 완료")
    )
    paid_amount = await db.scalar(
        select(func.coalesce(func.sum(AdvancePaymentRequest.request_amount), 0)).where(
            AdvancePaymentRequest.status == "이체 완료"
        )
    )
    month_request_count = await db.scalar(
        select(func.count()).select_from(AdvancePaymentRequest).where(
            func.extract("year", AdvancePaymentRequest.requested_at) == now.year,
            func.extract("month", AdvancePaymentRequest.requested_at) == now.month,
        )
    )
    accounts_balance = await db.scalar(select(func.coalesce(func.sum(BankAccount.balance), 0)))

    workplaces_result = await db.execute(select(DelegatedWorkplace).order_by(DelegatedWorkplace.workplace_name.asc()))
    requests_result = await db.execute(
        select(AdvancePaymentRequest)
        .where(AdvancePaymentRequest.status == "이체 완료")
        .order_by(AdvancePaymentRequest.paid_at.desc())
    )
    month_totals: dict[int, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for row in requests_result.scalars().all():
        key = row.paid_at.strftime("%Y-%m") if row.paid_at else row.requested_at.strftime("%Y-%m")
        month_totals[row.workplace_id][key] += row.request_amount

    forecasts: list[WorkplaceForecast] = []
    for workplace in workplaces_result.scalars().all():
        totals = list(month_totals[workplace.id].values())
        average = int(sum(totals) / len(totals)) if totals else 0
        suggested = int(average * 1.1) if average else 0
        guide = "지급 이력이 부족합니다. 초기 기준금액을 등록해 추적하세요."
        if average:
            guide = "최근 지급 평균 기준으로 10% 여유분을 포함한 추천 금액입니다."
        forecasts.append(
            WorkplaceForecast(
                workplace_id=workplace.id,
                workplace_name=workplace.workplace_name,
                average_monthly_amount=average,
                suggested_amount=suggested,
                guide=guide,
            )
        )

    return WorkplaceSummary(
        workplace_count=workplace_count or 0,
        request_count=request_count or 0,
        pending_count=pending_count or 0,
        approved_count=approved_count or 0,
        paid_count=paid_count or 0,
        paid_amount=paid_amount or 0,
        month_request_count=month_request_count or 0,
        accounts_balance=accounts_balance or 0,
        forecasts=forecasts[:5],
    )


@router.patch("/{workplace_id}", response_model=WorkplaceRead)
async def update_workplace(
    workplace_id: int, payload: WorkplaceUpdate, db: AsyncSession = Depends(get_db)
) -> DelegatedWorkplace:
    workplace = await _get_workplace_or_404(db, workplace_id)
    for key, value in payload.model_dump(exclude_unset=True).items():
        setattr(workplace, key, value)
    try:
        await db.commit()
    except IntegrityError as exc:
        await db.rollback()
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Workplace name already exists") from exc
    await db.refresh(workplace)
    return workplace


@router.delete("/{workplace_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_workplace(workplace_id: int, db: AsyncSession = Depends(get_db)) -> None:
    workplace = await _get_workplace_or_404(db, workplace_id)
    await db.delete(workplace)
    await db.commit()
