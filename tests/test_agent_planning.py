from __future__ import annotations

from datetime import date

from adx_report_agent.config import load_agent_config
from adx_report_agent.nodes.parse_request import parse_request_node
from adx_report_agent.nodes.plan_report import plan_report_node
from adx_report_agent.parser import parse_request


def test_parse_bidding_request() -> None:
    request = parse_request("看下 2026-07-25 竞价", today=date(2026, 7, 28))
    assert request.report_date == date(2026, 7, 25)
    assert request.analysis_type == "bidding"


def test_plan_bidding_report() -> None:
    config = load_agent_config(None)
    parsed = parse_request_node({"raw_text": "看下 2026-07-25 竞价", "today": date(2026, 7, 28)})
    planned = plan_report_node({"request": parsed["request"], "config": config})
    assert planned["standard_path"].as_posix() == "configs/bidding_report.json"
    assert planned["scripts"] == ["scripts/build_725_bidding_report.py"]
    assert "竞价分析报表.xlsx" in planned["output_path"].as_posix()
