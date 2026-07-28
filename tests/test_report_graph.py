from __future__ import annotations

from datetime import date
from pathlib import Path

from adx_report_agent.config import load_agent_config
from adx_report_agent.graph import build_report_graph
from adx_report_agent.state import ValidationResult
from adx_report_agent.tools.registry import ToolRegistry, ToolSpec


def test_graph_invokes_bidding_flow_without_database() -> None:
    def fake_execute_report_pipeline(request, config, scripts, output_path):
        return Path(output_path)

    def fake_validate_report_artifact(output_path, analysis_type):
        return ValidationResult(ok=True, checks=["fake ok"], errors=[])

    registry = ToolRegistry()
    registry.register(ToolSpec("execute_report_pipeline", "fake", {"type": "object"}, fake_execute_report_pipeline))
    registry.register(ToolSpec("validate_report_artifact", "fake", {"type": "object"}, fake_validate_report_artifact))

    app = build_report_graph(registry)
    state = app.invoke({"raw_text": "看下 2026-07-25 竞价", "today": date(2026, 7, 28), "config": load_agent_config(None)})

    assert state["request"].analysis_type == "bidding"
    assert state["status"] == "validated"
    assert state["validation"].ok is True
    assert "竞价分析报表.xlsx" in state["output_path"].as_posix()
