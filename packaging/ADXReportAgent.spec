# -*- mode: python ; coding: utf-8 -*-

import os
import sys

from PyInstaller.utils.hooks import collect_submodules


project_root = os.path.dirname(SPECPATH)
variant = os.environ.get("ADX_PACKAGE_VARIANT", "direct").lower()
if variant not in {"direct", "ssh"}:
    raise ValueError(f"Unsupported ADX_PACKAGE_VARIANT={variant!r}")

entry_script = os.path.join(
    SPECPATH,
    "desktop_launcher_ssh_entry.py" if variant == "ssh" else "desktop_launcher_direct_entry.py",
)
app_suffix = "SSH" if variant == "ssh" else "Direct"
app_name = f"ADXReportAgent-{app_suffix}"
bundle_identifier = f"com.local.adx-report-agent.{variant}"

datas = [
    (os.path.join(project_root, "adx_report_agent/web_static"), "adx_report_agent/web_static"),
    (os.path.join(project_root, "configs"), "configs"),
    (os.path.join(project_root, "scripts"), "scripts"),
    (os.path.join(project_root, "data"), "data"),
    (os.path.join(project_root, "env.example"), "."),
]

hiddenimports = (
    collect_submodules("langgraph")
    + collect_submodules("langchain_core")
    + collect_submodules("pymysql")
    + collect_submodules("sshtunnel")
    + collect_submodules("paramiko")
)

a = Analysis(
    [entry_script],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pandas.tests", "numpy.tests", "pytest"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=app_name,
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name=f"{app_name}.app",
        icon=None,
        bundle_identifier=bundle_identifier,
    )
