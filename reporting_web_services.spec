# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['reporting_web_services.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\TEMP\\Reporting Web Services\\Logo_RWS.ico', '.'), ('C:\\TEMP\\Reporting Web Services\\Logo_RWS.png', '.')],
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
    name='reporting_web_services',
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
    icon=['C:\\TEMP\\Reporting Web Services\\Logo_RWS.ico'],
)
