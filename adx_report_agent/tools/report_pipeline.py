from __future__ import annotations

import os
import runpy
import subprocess
import sys
from datetime import timedelta
from pathlib import Path

from ..models import AgentConfig, ReportRequest
from ..resources import resource_path
from ..tunnel import ensure_tunnel

sys.path.insert(0, str(resource_path(".vendor_pymysql")))
import pymysql  # noqa: E402


def assert_order_was_delivered(request: ReportRequest, env: dict[str, str]) -> None:
    """Stop early when an explicitly requested order had no delivery that day."""

    if request.order_id is None:
        return
    password = env.get("DORIS_PASSWORD")
    if not password:
        raise RuntimeError("DORIS_PASSWORD is required and must be provided via environment variable.")
    with pymysql.connect(
        host=env["DORIS_HOST"],
        port=int(env["DORIS_PORT"]),
        user=env["DORIS_USER"],
        password=password,
        database=env["DORIS_DATABASE"],
        connect_timeout=10,
        read_timeout=120,
        write_timeout=120,
        charset="utf8mb4",
    ) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                select count(*) bid_count
                from ods.ods_ad_adx_bid_response_rt
                where response_time >= %s
                  and response_time < date_add(%s, interval 1 day)
                  and order_id = %s
                """,
                (request.report_date.isoformat(), request.report_date.isoformat(), request.order_id),
            )
            row = cur.fetchone()
    bid_count = int(row[0] or 0) if row else 0
    if bid_count == 0:
        raise RuntimeError(
            f"{request.report_date.isoformat()} 订单{request.order_id}未投放："
            "该日期在出价响应表无记录，未生成报表。"
        )


def execute_report_pipeline(
    request: ReportRequest,
    config: AgentConfig,
    scripts: list[str],
    output_path: str | Path,
) -> Path:
    """Execute report scripts under the configured Doris connection."""

    output = Path(output_path)
    env = os.environ.copy()
    env["REPORT_DAY"] = request.report_date.isoformat()
    env["REPORT_PREV_DAY"] = (request.report_date - timedelta(days=1)).isoformat()
    if request.order_id is not None:
        env["REPORT_ORDER_ID"] = str(request.order_id)
    # The caller passes the exact output path from the planning node.
    env["REPORT_OUTPUT"] = str(output)

    if config.ssh_tunnel.enabled:
        env["DORIS_HOST"] = "127.0.0.1"
        env["DORIS_PORT"] = str(config.ssh_tunnel.local_port)
    else:
        env["DORIS_HOST"] = config.database.host
        env["DORIS_PORT"] = str(config.database.port)
    env["DORIS_USER"] = config.database.user
    env["DORIS_DATABASE"] = config.database.database
    if "DORIS_PASSWORD" not in env:
        raise RuntimeError("DORIS_PASSWORD is required and must be provided via environment variable.")

    with ensure_tunnel(config):
        assert_order_was_delivered(request, env)
        if getattr(sys, "frozen", False):
            old_env = os.environ.copy()
            root = str(resource_path("."))
            old_sys_path = list(sys.path)
            try:
                os.environ.clear()
                os.environ.update(env)
                if root not in sys.path:
                    sys.path.insert(0, root)
                for script in scripts:
                    runpy.run_path(str(resource_path(script)), run_name="__main__")
            finally:
                os.environ.clear()
                os.environ.update(old_env)
                sys.path[:] = old_sys_path
        else:
            for script in scripts:
                subprocess.run([sys.executable, str(resource_path(script))], check=True, env=env)
    return output.resolve()
