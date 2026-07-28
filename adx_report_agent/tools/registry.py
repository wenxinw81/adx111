from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import Any


ToolCallable = Callable[..., Any]


@dataclass(frozen=True)
class ToolSpec:
    """Function-calling style tool specification."""

    name: str
    description: str
    parameters: dict[str, Any]
    func: ToolCallable


class ToolRegistry:
    """Small registry that keeps tool names, schemas, and callables aligned."""

    def __init__(self) -> None:
        self._tools: dict[str, ToolSpec] = {}

    def register(self, spec: ToolSpec) -> None:
        if spec.name in self._tools:
            raise ValueError(f"Tool already registered: {spec.name}")
        self._tools[spec.name] = spec

    def get(self, name: str) -> ToolSpec:
        try:
            return self._tools[name]
        except KeyError as exc:
            raise KeyError(f"Unknown tool: {name}") from exc

    def execute(self, name: str, **kwargs: Any) -> Any:
        return self.get(name).func(**kwargs)

    def schemas(self) -> list[dict[str, Any]]:
        return [
            {
                "type": "function",
                "function": {
                    "name": spec.name,
                    "description": spec.description,
                    "parameters": spec.parameters,
                },
            }
            for spec in self._tools.values()
        ]


def build_default_registry() -> ToolRegistry:
    from .report_pipeline import execute_report_pipeline
    from .report_validator import validate_report_artifact

    registry = ToolRegistry()
    registry.register(
        ToolSpec(
            name="execute_report_pipeline",
            description="Run the report scripts for one parsed report request.",
            parameters={
                "type": "object",
                "properties": {
                    "request": {"type": "object"},
                    "config": {"type": "object"},
                    "scripts": {"type": "array", "items": {"type": "string"}},
                    "output_path": {"type": "string"},
                },
                "required": ["request", "config", "scripts", "output_path"],
            },
            func=execute_report_pipeline,
        )
    )
    registry.register(
        ToolSpec(
            name="validate_report_artifact",
            description="Validate generated workbook sheets and key report rules.",
            parameters={
                "type": "object",
                "properties": {
                    "output_path": {"type": "string"},
                    "analysis_type": {"type": "string"},
                },
                "required": ["output_path", "analysis_type"],
            },
            func=validate_report_artifact,
        )
    )
    return registry
