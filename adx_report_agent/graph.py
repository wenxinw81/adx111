from __future__ import annotations

from functools import partial

from langgraph.graph import END, START, StateGraph

from .nodes import (
    execute_report_node,
    parse_request_node,
    plan_report_node,
    reflect_node,
    route_after_validation,
    validate_report_node,
)
from .state import AgentState
from .tools import ToolRegistry, build_default_registry


def build_report_graph(tool_registry: ToolRegistry | None = None):
    """Build the LangGraph StateGraph for ADX report generation."""

    registry = tool_registry or build_default_registry()
    graph = StateGraph(AgentState)
    graph.add_node("parse_request", parse_request_node)
    graph.add_node("plan_report", plan_report_node)
    graph.add_node("execute_report", partial(execute_report_node, tool_registry=registry))
    graph.add_node("validate_report", partial(validate_report_node, tool_registry=registry))
    graph.add_node("reflect", reflect_node)

    graph.add_edge(START, "parse_request")
    graph.add_edge("parse_request", "plan_report")
    graph.add_edge("plan_report", "execute_report")
    graph.add_edge("execute_report", "validate_report")
    graph.add_conditional_edges("validate_report", route_after_validation, {"reflect": "reflect", "end": END})
    graph.add_edge("reflect", END)
    return graph.compile()


def run_report_graph(initial_state: AgentState) -> AgentState:
    """Invoke the compiled report graph."""

    app = build_report_graph()
    return app.invoke(initial_state)
