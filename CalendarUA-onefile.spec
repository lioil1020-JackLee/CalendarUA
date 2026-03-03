# -*- mode: python ; coding: utf-8 -*-
import os

datas = []
icon = None

if os.name == "nt" and os.path.exists("lioil.ico"):
    datas.append(("lioil.ico", "."))
    icon = "lioil.ico"
elif os.path.exists("lioil.icns"):
    datas.append(("lioil.icns", "."))
    icon = "lioil.icns"

a = Analysis(
    ["CalendarUA.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
        "qasync",
        "dateutil.rrule",
    ],
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
    name="CalendarUA-onefile",
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
    icon=icon,
)