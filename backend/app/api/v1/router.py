from fastapi import APIRouter

from app.api.v1.routes import auth, health, sales, workplaces

api_router = APIRouter()
api_router.include_router(health.router, prefix="/health", tags=["health"])
api_router.include_router(auth.router, prefix="/auth", tags=["auth"])
api_router.include_router(sales.router, prefix="/sales", tags=["sales"])
api_router.include_router(workplaces.router, prefix="/workplaces", tags=["workplaces"])
