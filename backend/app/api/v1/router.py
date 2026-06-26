from fastapi import APIRouter

from app.api.v1.routes import billing, health, sales, users, workplaces

api_router = APIRouter()
api_router.include_router(health.router, prefix="/health", tags=["health"])
api_router.include_router(sales.router, prefix="/sales", tags=["sales"])
api_router.include_router(users.router, prefix="/users", tags=["users"])
api_router.include_router(workplaces.router, prefix="/workplaces", tags=["workplaces"])
api_router.include_router(billing.router, prefix="/billing", tags=["billing"])
