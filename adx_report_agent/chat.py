from __future__ import annotations

from datetime import date

from .config import load_agent_config
from .models import ReportRequest
from .parser import parse_request
from .runner import run_report
from .standard_editor import load_standard, propose_standard_update, save_standard


HELP = """可输入：
- 看下昨天数据
- 看下 2026-07-25 数据
- 看下 2026-07-25 花销
- 看下 2026-07-25 竞价
- 改标准：小时分析不要展示无出价小时
- 改标准：APP Top50
- 显示标准
- 退出
"""


def run_chat(config_path: str | None = None, today: date | None = None) -> None:
    config = load_agent_config(config_path)
    print("ADX 报表 Agent 已启动。输入“帮助”查看示例，输入“退出”结束。")
    pending_standard: dict | None = None
    pending_summary: str | None = None

    while True:
        try:
            text = input("你：").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nAgent：已退出。")
            return

        if not text:
            continue
        if text in {"退出", "quit", "exit", "q"}:
            print("Agent：已退出。")
            return
        if text in {"帮助", "help", "?"}:
            print("Agent：" + HELP)
            continue

        if pending_standard is not None:
            if text in {"确认", "保存", "是", "yes", "y"}:
                save_standard(config.standard_path, pending_standard)
                print(f"Agent：已保存标准。{pending_summary}")
                pending_standard = None
                pending_summary = None
                continue
            if text in {"取消", "否", "no", "n"}:
                print("Agent：已取消修改标准。")
                pending_standard = None
                pending_summary = None
                continue
            print("Agent：上一条标准修改还没确认。请输入“确认”保存，或“取消”。")
            continue

        if text in {"显示标准", "查看标准", "标准"}:
            standard = load_standard(config.standard_path)
            rules = standard.get("rules", {})
            print("Agent：当前基础分析标准：")
            for key, value in rules.items():
                print(f"  - {key}: {value}")
            continue

        if any(x in text for x in ("改标准", "修改标准", "以后", "后续", "标准")):
            standard = load_standard(config.standard_path)
            proposal = propose_standard_update(text, standard)
            if proposal is None:
                print("Agent：这条标准修改我还没识别出来。可以说得更具体一点，比如“改标准：APP Top50”。")
                continue
            pending_standard, pending_summary = proposal
            print(f"Agent：{pending_summary} 输入“确认”保存为默认标准，或“取消”。")
            continue

        try:
            request = parse_request(text, today=today)
            output = run_report(request, config)
            report_name = "花销专门分析" if request.analysis_type == "spend" else "竞价专门分析" if request.analysis_type == "bidding" else "基础分析"
            print(f"Agent：已生成 {request.report_date.isoformat()} {report_name}报表：{output}")
        except Exception as exc:
            print(f"Agent：执行失败：{exc}")
