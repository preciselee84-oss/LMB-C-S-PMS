from datetime import datetime

from pydantic import BaseModel, Field


class VisitVocCreate(BaseModel):
    company_name: str = Field(min_length=1, max_length=160)
    visit_date: str = Field(min_length=1, max_length=20)
    visitor_name: str = Field(min_length=1, max_length=100)
    contact_name: str | None = Field(default=None, max_length=100)
    channel: str = Field(default="방문", max_length=30)
    sentiment: str = Field(default="보통", max_length=30)
    product_area: str | None = Field(default=None, max_length=80)
    voc_text: str = Field(min_length=1)
    next_action: str | None = None


class VisitVocRead(VisitVocCreate):
    id: str
    created_at: datetime
    status: str


class MeetingMinutesCreate(BaseModel):
    title: str = Field(min_length=1, max_length=160)
    company_name: str | None = Field(default=None, max_length=160)
    meeting_date: str | None = Field(default=None, max_length=20)
    participants: str | None = Field(default=None, max_length=255)
    transcript_text: str = Field(min_length=1)


class MeetingMinutesRead(BaseModel):
    title: str
    company_name: str | None
    meeting_date: str | None
    participants: list[str]
    source_file_name: str | None
    summary: str
    key_topics: list[str]
    decisions: list[str]
    action_items: list[str]
    risks: list[str]
    original_transcript: str
