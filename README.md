# LMB Performance System

정식 실적관리시스템 전환 프로젝트입니다.

기존 Streamlit 단일 앱을 FastAPI 백엔드, React 프론트엔드, PostgreSQL 데이터베이스 구조로 재구축합니다.

## Tech Stack

| Layer | Stack |
| --- | --- |
| Frontend | React 19, TypeScript, Vite, Ant Design 5, Zustand, React Router, @ant-design/charts |
| Backend | FastAPI, Python 3.12, SQLAlchemy 2.0, Alembic, asyncpg |
| Database | PostgreSQL |
| Auth | JWT, python-jose, bcrypt |
| Excel / PPT | openpyxl, python-pptx |
| External APIs | httpx.AsyncClient |

## Structure

```txt
lmb-performance-system/
  backend/      FastAPI API server
  frontend/     React SPA
  docs/         migration and architecture notes
  docker-compose.yml
```

## Local Development

1. PostgreSQL 실행

```bash
docker compose up -d db
```

2. Backend 실행

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[dev]"
copy .env.example .env
alembic upgrade head
python -m app.cli.create_admin --username admin --password admin1234 --name 관리자
uvicorn app.main:app --reload
```

3. Frontend 실행

```bash
cd frontend
npm install
npm run dev
```

## Migration Direction

기존 `LMB-C-S-PMS/app.py`의 pandas, Excel, PPT, 외부 연동 로직은 `backend/app/services`, `backend/app/reports`, `backend/app/integrations`로 순차 분리합니다.
