# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_dir = Path(SPECPATH)
application_data = [
    (str(project_dir / "templates"), "templates"),
    (str(project_dir / "assets"), "assets"),
]

a = Analysis(
    ['desktop_app.py'],
    pathex=[],
    binaries=[],
    datas=application_data,
    hiddenimports=[],
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
    name='desktop_app',
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
)
