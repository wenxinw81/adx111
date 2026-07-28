from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any


def load_standard(path: str | Path) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def save_standard(path: str | Path, data: dict[str, Any]) -> None:
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def propose_standard_update(text: str, standard: dict[str, Any]) -> tuple[dict[str, Any], str] | None:
    """Return an updated standard and a human-readable summary, if text matches."""

    updated = json.loads(json.dumps(standard, ensure_ascii=False))
    rules = updated.setdefault("rules", {})

    if "小时" in text and ("无出价" in text or "0出价" in text or "出价为0" in text):
        if any(x in text for x in ("删除", "删掉", "不展示", "不要", "只保留")):
            rules["drop_zero_bid_hours"] = True
            return updated, "已设置：小时分析删除无出价小时。"
        if any(x in text for x in ("保留", "展示")):
            rules["drop_zero_bid_hours"] = False
            return updated, "已设置：小时分析保留无出价小时。"

    top_match = re.search(r"APP\s*Top\s*(\d+)|app\s*top\s*(\d+)|APP.*?(\d+)", text, re.I)
    if top_match and ("top" in text.lower() or "Top" in text or "前" in text):
        value = next(int(g) for g in top_match.groups() if g)
        rules["app_top_n"] = value
        return updated, f"已设置：APP 分析展示 Top{value}。"

    if "落地页" in text and "click" in text.lower():
        if any(x in text for x in ("必须", "先经过", "关联", "需要")):
            rules["landing_page_requires_click"] = True
            return updated, "已设置：落地页曝光/跳转必须先经过 click 表。"
        if any(x in text for x in ("不需要", "不用", "直接")):
            rules["landing_page_requires_click"] = False
            return updated, "已设置：落地页曝光/跳转不强制经过 click 表。"

    limit_match = re.search(r"交叉.*?(\d+)|cross.*?(\d+)", text, re.I)
    if limit_match:
        value = next(int(g) for g in limit_match.groups() if g)
        rules["cross_table_limit"] = value
        return updated, f"已设置：交叉表最多展示 {value} 行。"

    return None
