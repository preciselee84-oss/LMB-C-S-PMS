# Architecture

## 목표

Streamlit에 섞여 있던 UI, 상태, 데이터 처리, 저장 로직을 다음 계층으로 분리합니다.

- Frontend: 화면, 사용자 입력, 라우팅, 클라이언트 상태
- Backend: 인증, 권한, 업무 API, 파일 생성, 외부 연동
- Database: 사용자, 실적, 이력, 보고서, 설정 저장
- Integrations: Google Sheets, OpenDART, Kakao API

## Backend Modules

```txt
app/
  api/          HTTP route definitions
  core/         settings, security, database
  models/       SQLAlchemy ORM models
  schemas/      Pydantic request/response schemas
  services/     business logic migrated from Streamlit
  reports/      Excel/PPT generation
  integrations/ external API clients
```

## Migration First Targets

1. 사용자/권한 모델
2. GitHub JSON 저장 데이터의 PostgreSQL 이전
3. Google Sheets CSV 조회 로직
4. 실적 집계 순수 함수 분리
5. PPT/Excel 다운로드 API

