from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from .models import AgentConfig
from .resources import resource_path, runtime_root


def load_dotenv(path: str | Path | None = None) -> None:
    """Load simple KEY=VALUE lines without requiring python-dotenv."""

    env_path = Path(path) if path is not None else runtime_root() / ".env"
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'").strip('"')
        os.environ.setdefault(key, value)


def _env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def apply_env_overrides(config: AgentConfig) -> AgentConfig:
    """Let .env/environment values override JSON runtime config."""

    data = config.model_dump(mode="json")
    database = data.setdefault("database", {})
    ssh_tunnel = data.setdefault("ssh_tunnel", {})

    if os.getenv("DORIS_HOST"):
        database["host"] = os.environ["DORIS_HOST"]
    if os.getenv("DORIS_PORT"):
        database["port"] = int(os.environ["DORIS_PORT"])
    if os.getenv("DORIS_DATABASE"):
        database["database"] = os.environ["DORIS_DATABASE"]
    if os.getenv("DORIS_USER"):
        database["user"] = os.environ["DORIS_USER"]

    if os.getenv("ADX_SSH_ENABLED"):
        ssh_tunnel["enabled"] = _env_bool("ADX_SSH_ENABLED", ssh_tunnel.get("enabled", False))
    if os.getenv("ADX_SSH_USER"):
        ssh_tunnel["ssh_user"] = os.environ["ADX_SSH_USER"]
    if os.getenv("ADX_SSH_HOST"):
        ssh_tunnel["ssh_host"] = os.environ["ADX_SSH_HOST"]
    if os.getenv("ADX_SSH_LOCAL_PORT"):
        ssh_tunnel["local_port"] = int(os.environ["ADX_SSH_LOCAL_PORT"])

    if os.getenv("ADX_OUTPUT_ROOT"):
        data["output_root"] = os.environ["ADX_OUTPUT_ROOT"]
    return AgentConfig.model_validate(data)


def load_agent_config(path: str | Path | None = None) -> AgentConfig:
    """Load optional JSON config and merge it over defaults."""

    load_dotenv()
    if path is None:
        return apply_env_overrides(AgentConfig())
    config_path = resource_path(path)
    if not config_path.exists():
        return apply_env_overrides(AgentConfig(standard_path=Path("configs/basic_report.json")))
    data = json.loads(config_path.read_text(encoding="utf-8"))
    return apply_env_overrides(AgentConfig.model_validate(data))


def load_report_standard(path: str | Path) -> dict[str, Any]:
    """Load the editable report standard."""

    return json.loads(resource_path(path).read_text(encoding="utf-8"))
