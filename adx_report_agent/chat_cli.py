from __future__ import annotations

import argparse
from datetime import date

from .chat import run_chat


def main() -> None:
    parser = argparse.ArgumentParser(description="ADX report conversational agent")
    parser.add_argument("--config", default=None, help="agent runtime config JSON")
    parser.add_argument("--today", default=None, help="测试用：指定今天，格式 YYYY-MM-DD")
    args = parser.parse_args()
    today = date.fromisoformat(args.today) if args.today else None
    run_chat(config_path=args.config, today=today)


if __name__ == "__main__":
    main()
