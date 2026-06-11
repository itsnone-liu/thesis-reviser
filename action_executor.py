"""Shim module: re-export action executor components from app.modules.action_executor."""
from app.modules.action_executor import ActionExecutor, action_executor

__all__ = ["ActionExecutor", "action_executor"]
