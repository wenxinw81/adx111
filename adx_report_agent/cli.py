from __future__ import annotations

import argparse
import os
import sys
from datetime import date

from .config import load_agent_config
from .parser import parse_request
from .runner import run_report


def main() -> None:
    parser = argparse.ArgumentParser(description="ADX daily report agent")
    parser.add_argument("text", help="例如：看下昨天数据 / 看下 2026-07-25 数据")
    parser.add_argument("--config", default=None, help="agent runtime config JSON")
    parser.add_argument("--today", default=None, help="测试用：指定今天，格式 YYYY-MM-DD")
    args = parser.parse_args()

    today = date.fromisoformat(args.today) if args.today else None
    request = parse_request(args.text, today=today)
    config = load_agent_config(args.config)
    try:
        output = run_report(request, config)
    except RuntimeError as exc:
        message = str(exc)
        if "未投放" in message:
            print(message)
            sys.exit(0)
        raise
    report_name = (
        "花销专门分析"
        if request.analysis_type == "spend"
        else "竞价专门分析"
        if request.analysis_type == "bidding"
        else "基础分析"
    )
    order_text = f"订单{request.order_id}" if request.order_id is not None else ""
    print(f"已生成{request.report_date.isoformat()}{order_text}{report_name}报表：{output}")


if __name__ == "__main__":
    main()
