from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field


class ReportRequest(BaseModel):
    """A parsed user request for report generation."""

    raw_text: str
    report_date: date
    analysis_type: Literal["basic", "bidding", "spend"] = "basic"
    order_id: int | None = None


class DatabaseConfig(BaseModel):
    """Doris/MySQL-compatible connection settings."""

    host: str = "192.168.100.23"
    port: int = 29030
    database: str = "ads"
    user: str = "WishFox"


class SshTunnelConfig(BaseModel):
    """SSH tunnel settings for reaching Doris from this Mac."""

    enabled: bool = True
    ssh_user: str = "noel"
    ssh_host: str = "192.168.110.139"
    local_port: int = 19030


class AgentConfig(BaseModel):
    """Runtime configuration for the report agent."""

    database: DatabaseConfig = Field(default_factory=DatabaseConfig)
    ssh_tunnel: SshTunnelConfig = Field(default_factory=SshTunnelConfig)
    standard_path: Path = Path("configs/basic_report.json")
    output_root: Path = Path("outputs")
