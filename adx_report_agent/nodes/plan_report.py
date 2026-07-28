from __future__ import annotations

from pathlib import Path

from ..config import load_report_standard
from ..parser import date_to_output_path
from ..state import AgentState


SCRIPT_PLAN = {
    "basic": [
        "scripts/build_725_basic_report.py",
        "scripts/restyle_725_report.py",
        "scripts/fix_gantt_sheet.py",
    ],
    "spend": ["scripts/build_725_spend_report.py"],
    "bidding": ["scripts/build_725_bidding_report.py"],
}


STANDARD_PLAN = {
    "basic": Path("configs/basic_report.json"),
    "spend": Path("configs/spend_report.json"),
    "bidding": Path("configs/bidding_report.json"),
}


def plan_report_node(state: AgentState) -> dict:
    """Plan scripts, standard, and output path for a parsed request."""

    request = state["request"]
    config = state["config"]
    standard_path = config.standard_path
    if str(standard_path).endswith("basic_report.json"):
        standard_path = STANDARD_PLAN[request.analysis_type]

    standard = load_report_standard(standard_path)
    if standard.get("analysis_type") != request.analysis_type:
        raise RuntimeError(f"Unsupported report standard for {request.analysis_type}: {standard_path}")

    return {
        "standard_path": standard_path,
        "standard": standard,
        "scripts": SCRIPT_PLAN[request.analysis_type],
        "output_path": Path(date_to_output_path(request.report_date, request.analysis_type, request.order_id)),
        "status": "planned",
    }
