from __future__ import annotations

import os
import shutil
import socket
import subprocess
from contextlib import contextmanager, suppress
from typing import Iterator

from .models import AgentConfig


def _is_port_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex(("127.0.0.1", port)) == 0


def _start_sshtunnel_forwarder(config: AgentConfig, password: str):
    try:
        from sshtunnel import SSHTunnelForwarder
    except ImportError as exc:
        raise RuntimeError("sshtunnel is not installed") from exc

    tunnel = config.ssh_tunnel
    server = SSHTunnelForwarder(
        (tunnel.ssh_host, 22),
        ssh_username=tunnel.ssh_user,
        ssh_password=password,
        remote_bind_address=(config.database.host, config.database.port),
        local_bind_address=("127.0.0.1", tunnel.local_port),
    )
    server.start()
    return server


def _open_expect_tunnel(config: AgentConfig, password: str) -> None:
    tunnel = config.ssh_tunnel
    cmd = f"""
set timeout 30
spawn ssh -fN -o StrictHostKeyChecking=no -o ExitOnForwardFailure=yes -L {tunnel.local_port}:{config.database.host}:{config.database.port} {tunnel.ssh_user}@{tunnel.ssh_host}
expect {{
  "*Password:*" {{ send "{password}\\r"; exp_continue }}
  "*password:*" {{ send "{password}\\r"; exp_continue }}
  eof
}}
"""
    subprocess.run(["expect", "-c", cmd], check=True)
    if not _is_port_open(tunnel.local_port):
        raise RuntimeError(f"SSH tunnel did not open local port {tunnel.local_port}.")


@contextmanager
def ensure_tunnel(config: AgentConfig) -> Iterator[None]:
    """Ensure an SSH tunnel exists for Doris access.

    Passwords are read from ADX_SSH_PASSWORD and are not stored.
    If the local port is already open, the tunnel is reused.
    """

    tunnel = config.ssh_tunnel
    if not tunnel.enabled or _is_port_open(tunnel.local_port):
        yield
        return

    password = os.getenv("ADX_SSH_PASSWORD")
    if not password:
        raise RuntimeError(
            "SSH tunnel is required but ADX_SSH_PASSWORD is not set. "
            "Set it or open the tunnel manually."
        )
    server = None
    try:
        server = _start_sshtunnel_forwarder(config, password)
    except Exception:
        if not shutil.which("expect"):
            raise
        _open_expect_tunnel(config, password)
        yield
        return
    try:
        yield
    finally:
        with suppress(Exception):
            server.stop()
