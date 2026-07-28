from __future__ import annotations

import argparse
import os
import uuid
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from threading import Lock
from typing import Literal

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from .config import load_agent_config
from .models import AgentConfig, ReportRequest
from .parser import parse_request
from .resources import resource_path
from .runner import run_report


ReportStatus = Literal["queued", "running", "success", "failed"]


class GenerateReportPayload(BaseModel):
    text: str = Field(..., min_length=1, max_length=300)
    today: date | None = None


class ParsedRequestPayload(BaseModel):
    raw_text: str
    report_date: date
    analysis_type: str
    analysis_name: str
    order_id: int | None = None


class ReportJobPayload(BaseModel):
    job_id: str
    status: ReportStatus
    request: ParsedRequestPayload
    created_at: datetime
    updated_at: datetime
    output_path: str | None = None
    download_url: str | None = None
    error: str | None = None


@dataclass
class ReportJob:
    job_id: str
    request: ReportRequest
    status: ReportStatus
    created_at: datetime
    updated_at: datetime
    output_path: Path | None = None
    error: str | None = None


def analysis_name(analysis_type: str) -> str:
    if analysis_type == "spend":
        return "花销专门分析"
    if analysis_type == "bidding":
        return "竞价专门分析"
    return "基础分析"


class JobStore:
    def __init__(self) -> None:
        self._jobs: dict[str, ReportJob] = {}
        self._lock = Lock()

    def create(self, request: ReportRequest) -> ReportJob:
        now = datetime.now()
        job = ReportJob(
            job_id=uuid.uuid4().hex,
            request=request,
            status="queued",
            created_at=now,
            updated_at=now,
        )
        with self._lock:
            self._jobs[job.job_id] = job
        return job

    def get(self, job_id: str) -> ReportJob:
        with self._lock:
            job = self._jobs.get(job_id)
        if job is None:
            raise KeyError(job_id)
        return job

    def update(self, job_id: str, **patch: object) -> ReportJob:
        with self._lock:
            job = self._jobs[job_id]
            for key, value in patch.items():
                setattr(job, key, value)
            job.updated_at = datetime.now()
            return job


def serialize_job(job: ReportJob) -> ReportJobPayload:
    output_path = str(job.output_path) if job.output_path else None
    return ReportJobPayload(
        job_id=job.job_id,
        status=job.status,
        request=ParsedRequestPayload(
            raw_text=job.request.raw_text,
            report_date=job.request.report_date,
            analysis_type=job.request.analysis_type,
            analysis_name=analysis_name(job.request.analysis_type),
            order_id=job.request.order_id,
        ),
        created_at=job.created_at,
        updated_at=job.updated_at,
        output_path=output_path,
        download_url=f"/api/reports/{job.job_id}/download" if job.status == "success" else None,
        error=job.error,
    )


def create_app(config: AgentConfig | None = None, today: date | None = None) -> FastAPI:
    app = FastAPI(title="ADX Report Agent", version="0.1.0")
    app.state.agent_config = config or load_agent_config()
    app.state.today = today
    app.state.jobs = JobStore()
    app.state.executor = ThreadPoolExecutor(max_workers=int(os.getenv("ADX_WEB_WORKERS", "2")))

    static_dir = resource_path("adx_report_agent/web_static")

    def run_job(job_id: str) -> None:
        job = app.state.jobs.update(job_id, status="running", error=None)
        try:
            output = run_report(job.request, app.state.agent_config)
        except Exception as exc:  # noqa: BLE001 - surfaced to the web UI for operator diagnosis.
            app.state.jobs.update(job_id, status="failed", error=str(exc))
            return
        app.state.jobs.update(job_id, status="success", output_path=output, error=None)

    @app.post("/api/parse", response_model=ParsedRequestPayload)
    def parse_text(payload: GenerateReportPayload) -> ParsedRequestPayload:
        request = parse_request(payload.text, today=payload.today or app.state.today)
        return ParsedRequestPayload(
            raw_text=request.raw_text,
            report_date=request.report_date,
            analysis_type=request.analysis_type,
            analysis_name=analysis_name(request.analysis_type),
            order_id=request.order_id,
        )

    @app.post("/api/reports", response_model=ReportJobPayload)
    def generate_report(payload: GenerateReportPayload) -> ReportJobPayload:
        request = parse_request(payload.text, today=payload.today or app.state.today)
        job = app.state.jobs.create(request)
        future: Future[None] = app.state.executor.submit(run_job, job.job_id)
        app.state.jobs.update(job.job_id, future=future)
        return serialize_job(job)

    @app.get("/api/reports/{job_id}", response_model=ReportJobPayload)
    def get_report(job_id: str) -> ReportJobPayload:
        try:
            job = app.state.jobs.get(job_id)
        except KeyError as exc:
            raise HTTPException(status_code=404, detail="report job not found") from exc
        return serialize_job(job)

    @app.get("/api/reports/{job_id}/download")
    def download_report(job_id: str) -> FileResponse:
        try:
            job = app.state.jobs.get(job_id)
        except KeyError as exc:
            raise HTTPException(status_code=404, detail="report job not found") from exc
        if job.status != "success" or job.output_path is None:
            raise HTTPException(status_code=409, detail="report is not ready")
        if not job.output_path.exists():
            raise HTTPException(status_code=404, detail="report file not found")
        return FileResponse(
            job.output_path,
            filename=job.output_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")
    return app


def main() -> None:
    parser = argparse.ArgumentParser(description="ADX report web agent")
    parser.add_argument("--config", default=None, help="agent runtime config JSON")
    parser.add_argument("--host", default="127.0.0.1", help="bind host")
    parser.add_argument("--port", type=int, default=8787, help="bind port")
    parser.add_argument("--today", default=None, help="测试用：指定今天，格式 YYYY-MM-DD")
    args = parser.parse_args()

    import uvicorn

    config = load_agent_config(args.config)
    today = date.fromisoformat(args.today) if args.today else None
    uvicorn.run(create_app(config=config, today=today), host=args.host, port=args.port)


if __name__ == "__main__":
    main()
