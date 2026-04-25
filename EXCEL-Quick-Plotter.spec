# -*- mode: python ; coding: utf-8 -*-

import os

project_dir = os.path.abspath(SPECPATH)


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[(os.path.join(project_dir, 'icon.ico'), '.'), (os.path.join(project_dir, 'style.qss'), '.')],
    hiddenimports=['pynput', 'pynput.keyboard', 'pynput.keyboard._win32', 'pynput._util.win32', 'keyboard', 'mplcursors'],
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
    name='EXCEL-Quick-Plotter',
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
    icon=[os.path.join(project_dir, 'icon.ico')],
)
