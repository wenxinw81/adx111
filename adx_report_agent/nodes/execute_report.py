from __future__ import annotations

from ..state import AgentState
from ..tools import ToolRegistry


def execute_report_node(state: AgentState, tool_registry: ToolRegistry) -> dict:
    """Execute planned report scripts through the registered tool."""

    try:
        output = tool_registry.execute(
            "execute_report_pipeline",
            request=state["request"],
            config=state["config"],
            scripts=state["scripts"],
            output_path=state["output_path"],
        )
        return {"output_path": output, "status": "executed", "errors": []}
    except Exception as exc:
        message = str(exc)
        if "未投放" not in message:
            message = f"execute_report failed: {message}"
        return {"status": "failed", "errors": [message]}
