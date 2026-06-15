from datetime import datetime

from pydantic import BaseModel, Field


class UserCreate(BaseModel):
    username: str = Field(min_length=1, max_length=50)
    password: str = Field(min_length=4, max_length=100)
    name: str = Field(min_length=1, max_length=100)
    email: str | None = Field(default=None, max_length=255)
    role: str = Field(default="staff", max_length=30)
    is_active: bool = True


class UserUpdate(BaseModel):
    password: str | None = Field(default=None, min_length=4, max_length=100)
    name: str | None = Field(default=None, min_length=1, max_length=100)
    email: str | None = Field(default=None, max_length=255)
    role: str | None = Field(default=None, max_length=30)
    is_active: bool | None = None


class UserRead(BaseModel):
    id: int
    username: str
    name: str
    email: str | None
    role: str
    is_active: bool
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}
