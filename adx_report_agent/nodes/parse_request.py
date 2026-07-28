from __future__ import annotations

from ..parser import parse_request
from ..state import AgentState


def parse_request_node(state: AgentState) -> dict:
    """Parse natural-language input into a structured report request."""

    request = state.get("request") or parse_request(state["raw_text"], today=state.get("today"))
    return {"request": request, "status": "initialized"}
