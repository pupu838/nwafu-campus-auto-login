# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path(__file__).resolve().parent
bundled_driver = project_root / "drivers" / "msedgedriver.exe"

datas = []
if bundled_driver.exists():
    datas.append((str(bundled_driver), "drivers"))

hiddenimports = [
    "keyring.backends.Windows",
    "keyring.backends.win32ctypes",
    "pystray._win32",
    "PIL.ImageTk",
]

a = Analysis(
    ["nwafu_login.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    exclude_binaries=False,
    name="NWAFUAutoLogin",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
