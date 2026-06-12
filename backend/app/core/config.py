from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    app_name: str = "LMB Performance System API"
    app_env: str = "local"
    api_v1_prefix: str = "/api/v1"
    backend_cors_origins: str = Field(default="http://localhost:5173")

    database_url: str = "postgresql+asyncpg://lmb_admin:lmb_password@localhost:5432/lmb_performance"
    jwt_secret_key: str = "local-dev-secret-change-before-deploy"
    jwt_algorithm: str = "HS256"
    access_token_expire_minutes: int = 720

    github_token: str = ""
    dart_api_key: str = ""
    kakao_token: str = ""

    @property
    def cors_origins(self) -> list[str]:
        return [origin.strip() for origin in self.backend_cors_origins.split(",") if origin.strip()]


settings = Settings()
