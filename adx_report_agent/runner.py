from pathlib import Path

from .graph import run_report_graph
from .models import AgentConfig, ReportRequest


def run_report(request: ReportRequest, config: AgentConfig) -> Path:
    """Run the LangGraph-backed report pipeline."""

    final_state = run_report_graph({"raw_text": request.raw_text, "request": request, "config": config})
    if final_state.get("errors"):
        raise RuntimeError("; ".join(final_state["errors"]))
    output = final_state.get("output_path")
    if output is None:
        raise RuntimeError("Report graph completed without output_path.")
    return Path(output).resolve()


def run_basic_report(request: ReportRequest, config: AgentConfig) -> Path:
    """Backward-compatible wrapper for older entrypoints."""

    return run_report(request, config)
