from __future__ import annotations

from ..state import AgentState, ReflectionDecision


def route_after_validation(state: AgentState) -> str:
    """Route successful validations to END and failures to reflection."""

    validation = state.get("validation")
    if state.get("errors") or (validation is not None and not validation.ok):
        return "reflect"
    return "end"


def reflect_node(state: AgentState) -> dict:
    """Centralized failure diagnosis node.

    The current report pipelines can be expensive, so reflection records the
    failure and stops instead of retrying blindly.
    """

    errors = state.get("errors", [])
    validation = state.get("validation")
    if validation and validation.errors:
        errors = [*errors, *validation.errors]
    decision = ReflectionDecision(
        should_retry=False,
        reason="; ".join(errors) if errors else "unknown failure",
        max_retries=0,
    )
    return {"reflection": decision, "status": "failed", "errors": errors}
