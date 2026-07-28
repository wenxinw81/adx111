from __future__ import annotations

from ..state import AgentState
from ..tools import ToolRegistry


def validate_report_node(state: AgentState, tool_registry: ToolRegistry) -> dict:
    """Validate generated workbook with report-specific checks."""

    if state.get("errors"):
        return {"status": "failed"}
    validation = tool_registry.execute(
        "validate_report_artifact",
        output_path=state["output_path"],
        analysis_type=state["request"].analysis_type,
    )
    patch = {"validation": validation, "status": "validated" if validation.ok else "failed"}
    if not validation.ok:
        patch["errors"] = validation.errors
    return patch
