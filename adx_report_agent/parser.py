from __future__ import annotations

import re
from datetime import date, datetime, timedelta

from .models import ReportRequest


def parse_request(text: str, today: date | None = None) -> ReportRequest:
    """Parse Chinese one-line report requests into a dated report request."""

    base = today or date.today()
    report_date = base
    if "昨天" in text:
        report_date = base - timedelta(days=1)
    elif "前天" in text:
        report_date = base - timedelta(days=2)
    else:
        match = re.search(r"(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})", text)
        if match:
            y, m, d = map(int, match.groups())
            report_date = date(y, m, d)
        else:
            match = re.search(r"(\d{1,2})[./月-](\d{1,2})", text)
            if match:
                m, d = map(int, match.groups())
                report_date = date(base.year, m, d)

    analysis_type = "basic"
    if "花销" in text or "花费" in text or "消耗" in text:
        analysis_type = "spend"
    if "竞价" in text or "出价" in text:
        analysis_type = "bidding"
    if "基础" in text or (analysis_type == "basic" and ("大盘" in text or "昨天数据" in text)):
        analysis_type = "basic"

    order_id = None
    order_match = re.search(r"(?:订单\s*(?:号|id|ID)?\s*[:：=]?\s*|order[_\s-]*id\s*[:：=]?\s*)(\d{2,})", text, re.I)
    if not order_match:
        order_match = re.search(r"(\d{2,})\s*(?:订单|order)", text, re.I)
    if order_match:
        order_id = int(order_match.group(1))

    return ReportRequest(raw_text=text, report_date=report_date, analysis_type=analysis_type, order_id=order_id)


def date_to_output_path(report_date: date, analysis_type: str = "basic", order_id: int | None = None) -> str:
    """Return the default Excel output path for a report date."""

    order_prefix = f"订单{order_id}" if order_id is not None else ""
    if analysis_type == "spend":
        return (
            f"outputs/adx_spend_report_{report_date:%Y%m%d}/"
            f"{report_date.month}.{report_date.day}{order_prefix}广告投放花费分析报表.xlsx"
        )
    if analysis_type == "bidding":
        return (
            f"outputs/adx_bidding_report_{report_date:%Y%m%d}/"
            f"{report_date.month}.{report_date.day}{order_prefix}广告投放竞价分析报表.xlsx"
        )
    return f"outputs/adx_basic_report_{report_date:%Y%m%d}/{report_date.month}.{report_date.day}{order_prefix}广告投放基础分析.xlsx"
