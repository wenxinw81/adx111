from __future__ import annotations

import os
import socket
import threading
import time
import webbrowser
from pathlib import Path

import uvicorn

from .config import load_agent_config
from .resources import resource_path, runtime_root
from .web_app import create_app


def ensure_runtime_files() -> None:
    root = runtime_root()
    root.mkdir(parents=True, exist_ok=True)
    env_path = root / ".env"
    if not env_path.exists():
        for template_name in ("env.example", ".env.example"):
            template = resource_path(template_name)
            if template.exists():
                env_path.write_text(template.read_text(encoding="utf-8"), encoding="utf-8")
                break


def open_browser_later(url: str) -> None:
    def worker() -> None:
        time.sleep(1.2)
        webbrowser.open(url)

    threading.Thread(target=worker, daemon=True).start()


def port_is_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.3)
        return sock.connect_ex(("127.0.0.1", port)) == 0


def choose_port(preferred: int) -> int:
    for port in range(preferred, preferred + 20):
        if not port_is_open(port):
            return port
    raise RuntimeError(f"No available local port from {preferred} to {preferred + 19}.")


def main() -> None:
    ensure_runtime_files()
    os.chdir(runtime_root())
    host = os.getenv("ADX_WEB_HOST", "127.0.0.1")
    port = choose_port(int(os.getenv("ADX_WEB_PORT", "8787")))
    url = f"http://127.0.0.1:{port}"
    config_path = os.getenv("ADX_AGENT_CONFIG", "configs/agent.direct.example.json")
    config = load_agent_config(resource_path(config_path))
    open_browser_later(url)
    uvicorn.run(create_app(config=config), host=host, port=port)


if __name__ == "__main__":
    main()
