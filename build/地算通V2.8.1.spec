# -*- mode: python ; coding: utf-8 -*-

ICO = 'C:/Users/82235/WorkBuddy/2026-11-04-11-10-29/地算通.ico'
CAT = 'C:/Users/82235/WorkBuddy/2026-11-04-11-10-29/cat_64.png'
VER = 'C:/Users/82235/WorkBuddy/2026-11-04-11-10-29/build/version_info.txt'


a = Analysis(
    ['../地算通.py'],
    pathex=[],
    binaries=[],
    datas=[(ICO, '.'), (CAT, '.')],
    hiddenimports=['openpyxl'],
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
    name='地算通V2.8.1',
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
    icon=[ICO],
    version=VER,
)
