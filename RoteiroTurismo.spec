# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

block_cipher = None
# SPECPATH = pasta onde está o .spec (PyInstaller)
project = Path(SPECPATH).resolve()
_app_icon = project / "data" / "logo_cuesta.ico"
_version_info = project / "build" / "version_info.txt"

a = Analysis(
    [str(project / "desktop_main.py")],
    pathex=[str(project)],
    binaries=[],
    datas=[(str(project / "data"), "data")],
    hiddenimports=["openpyxl.cell._writer"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["streamlit"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="RoteiroTurismo",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(_app_icon) if _app_icon.is_file() else None,
    version=str(_version_info) if _version_info.is_file() else None,
)
