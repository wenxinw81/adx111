from __future__ import annotations

from pathlib import Path
from urllib.parse import parse_qs, urlparse

import requests
from requests_toolbelt.multipart.encoder import MultipartEncoder


class WeComBotError(RuntimeError):
    """Raised when WeCom bot API returns an error."""


def _check_response(payload: dict) -> dict:
    if payload.get("errcode") not in (0, None):
        raise WeComBotError(f"WeCom API error {payload.get('errcode')}: {payload.get('errmsg')}")
    return payload


def extract_webhook_key(webhook_url: str) -> str:
    key = parse_qs(urlparse(webhook_url).query).get("key", [""])[0]
    if not key:
        raise ValueError("WECOM_WEBHOOK_URL must contain a key query parameter.")
    return key


def send_markdown(webhook_url: str, content: str, timeout: int = 30) -> dict:
    response = requests.post(
        webhook_url,
        json={"msgtype": "markdown", "markdown": {"content": content}},
        timeout=timeout,
    )
    response.raise_for_status()
    return _check_response(response.json())


def upload_file(webhook_url: str, file_path: str | Path, timeout: int = 120) -> str:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(path)
    key = extract_webhook_key(webhook_url)
    upload_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={key}&type=file"
    with path.open("rb") as fp:
        multipart = MultipartEncoder(
            fields={
                "media": (
                    path.name,
                    fp,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            }
        )
        response = requests.post(upload_url, data=multipart, headers={"Content-Type": multipart.content_type}, timeout=timeout)
    response.raise_for_status()
    payload = _check_response(response.json())
    media_id = payload.get("media_id")
    if not media_id:
        raise WeComBotError(f"WeCom upload did not return media_id: {payload}")
    return media_id


def send_file(webhook_url: str, file_path: str | Path, timeout: int = 120) -> dict:
    media_id = upload_file(webhook_url, file_path, timeout=timeout)
    response = requests.post(
        webhook_url,
        json={"msgtype": "file", "file": {"media_id": media_id}},
        timeout=timeout,
    )
    response.raise_for_status()
    return _check_response(response.json())
