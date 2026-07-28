from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any, Literal, TypedDict

from pydantic import BaseModel, Field

from .models import AgentConfig, ReportRequest


AgentStatus = Literal["initialized", "planned", "executed", "validated", "failed"]


class ValidationResult(BaseModel):
    """Validation outcome for a generated report artifact."""

    ok: bool
    checks: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)


class ReflectionDecision(BaseModel):
    """Reflection node decision after an execution or validation failure."""

    should_retry: bool = False
    reason: str = ""
    max_retries: int = 0


class AgentState(TypedDict, total=False):
    """Serializable graph state for ADX report generation."""

    raw_text: str
    today: date | None
    request: ReportRequest
    config: AgentConfig
    standard_path: Path
    standard: dict[str, Any]
    scripts: list[str]
    output_path: Path
    status: AgentStatus
    errors: list[str]
    validation: ValidationResult
    reflection: ReflectionDecision
