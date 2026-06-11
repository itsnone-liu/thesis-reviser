"""FastAPI application entry point for OpenWorld Agent."""
from __future__ import annotations

import os

from fastapi import FastAPI

from app.api.routes import router as api_router
from app.core.scheduler import Scheduler
from app.storage.db import init_db

AGENT_ID = os.getenv("OWAGENT_ID", "default")


app = FastAPI(
    title="OpenWorld Agent API",
    description="OpenWorld Agent REST API — 状态查询、对话、目标与记忆管理",
    version="0.1.0",
)


@app.on_event("startup")
async def startup_event() -> None:
    """启动时初始化数据库并启动调度器。"""
    init_db()

    scheduler = Scheduler()
    scheduler.start(agent_id=AGENT_ID)
    app.state.scheduler = scheduler


@app.on_event("shutdown")
async def shutdown_event() -> None:
    """关闭时停止调度器。"""
    scheduler: Scheduler | None = getattr(app.state, "scheduler", None)
    if scheduler:
        scheduler.stop()


# 挂载 API router
app.include_router(api_router)
