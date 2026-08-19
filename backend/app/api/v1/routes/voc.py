from datetime import datetime, timezone
from uuid import uuid4

from fastapi import APIRouter, File, Form, HTTPException, UploadFile, status

from app.schemas.voc import MeetingMinutesCreate, MeetingMinutesRead, VisitVocCreate, VisitVocRead

router = APIRouter()

TEXT_EXTENSIONS = {".txt", ".md", ".srt", ".vtt", ".csv"}


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _split_lines(text: str) -> list[str]:
    return [line.strip(" -\t") for line in text.splitlines() if line.strip(" -\t")]


def _extension(filename: str | None) -> str:
    if not filename or "." not in filename:
        return ""
    return f".{filename.rsplit('.', 1)[1].lower()}"


def _compact_sentence(line: str, limit: int = 90) -> str:
    normalized = " ".join(line.split())
    return normalized if len(normalized) <= limit else f"{normalized[:limit].rstrip()}..."


def _pick_by_keywords(lines: list[str], keywords: tuple[str, ...], fallback: list[str]) -> list[str]:
    picked = [
        _compact_sentence(line)
        for line in lines
        if any(keyword in line.lower() for keyword in keywords)
    ]
    return (picked or fallback)[:5]


def _build_minutes(payload: MeetingMinutesCreate, source_file_name: str | None = None) -> MeetingMinutesRead:
    lines = _split_lines(payload.transcript_text)
    if not lines:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="회의록으로 변환할 텍스트가 비어 있습니다.")

    participants = [
        item.strip()
        for item in (payload.participants or "").replace("\n", ",").split(",")
        if item.strip()
    ]
    overview_lines = lines[:3]
    summary = " ".join(_compact_sentence(line, 120) for line in overview_lines)

    key_topics = _pick_by_keywords(
        lines,
        ("문의", "요청", "불편", "개선", "필요", "voc", "이슈", "문제", "확인"),
        [_compact_sentence(line) for line in lines[:4]],
    )
    decisions = _pick_by_keywords(
        lines,
        ("결정", "합의", "진행", "확정", "도입", "반영"),
        ["회의 중 명시적 결정사항은 별도 확인이 필요합니다."],
    )
    action_items = _pick_by_keywords(
        lines,
        ("할 일", "조치", "전달", "검토", "공유", "회신", "업로드", "등록", "추가"),
        ["담당자가 주요 VOC를 CRM에 등록하고 후속 조치 담당자를 지정합니다."],
    )
    risks = _pick_by_keywords(
        lines,
        ("리스크", "지연", "장애", "오류", "불가", "어려움", "불만"),
        ["중요 리스크가 없으면 고객 요청사항의 처리 기한만 추적합니다."],
    )

    return MeetingMinutesRead(
        title=payload.title,
        company_name=payload.company_name,
        meeting_date=payload.meeting_date,
        participants=participants,
        source_file_name=source_file_name,
        summary=summary,
        key_topics=key_topics,
        decisions=decisions,
        action_items=action_items,
        risks=risks,
        original_transcript=payload.transcript_text,
    )


@router.post("/entries", response_model=VisitVocRead, status_code=status.HTTP_201_CREATED)
async def create_visit_voc(payload: VisitVocCreate) -> VisitVocRead:
    return VisitVocRead(
        id=str(uuid4()),
        created_at=_now(),
        status="수집 완료",
        **payload.model_dump(),
    )


@router.post("/minutes", response_model=MeetingMinutesRead)
async def create_minutes(payload: MeetingMinutesCreate) -> MeetingMinutesRead:
    return _build_minutes(payload)


@router.post("/minutes/upload", response_model=MeetingMinutesRead)
async def create_minutes_from_upload(
    title: str = Form(...),
    company_name: str | None = Form(None),
    meeting_date: str | None = Form(None),
    participants: str | None = Form(None),
    transcript_text: str | None = Form(None),
    recording_file: UploadFile | None = File(None),
) -> MeetingMinutesRead:
    text = (transcript_text or "").strip()

    if recording_file is not None:
        content = await recording_file.read()
        ext = _extension(recording_file.filename)
        if ext in TEXT_EXTENSIONS:
            for encoding in ("utf-8-sig", "utf-8", "cp949"):
                try:
                    text = content.decode(encoding).strip()
                    break
                except UnicodeDecodeError:
                    continue
        elif not text:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="음성/영상 파일은 사내 STT 또는 Whisper 연동 후 생성된 텍스트를 함께 입력해야 합니다.",
            )

    payload = MeetingMinutesCreate(
        title=title,
        company_name=company_name,
        meeting_date=meeting_date,
        participants=participants,
        transcript_text=text,
    )
    return _build_minutes(payload, recording_file.filename if recording_file else None)
