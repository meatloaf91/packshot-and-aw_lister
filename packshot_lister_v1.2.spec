# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:/Users/vestim/OneDrive - Henkel/02_Extras/36_Temporary_Automations/01_packshot_lister/main.py'],
    pathex=[],
    binaries=[],
    datas=[],
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
    name='packshot_lister_v1.2',
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
    icon=['C:/Users/vestim/OneDrive - Henkel/02_Extras/36_Temporary_Automations/01_packshot_lister/icon/icon_packshot_lister.ico'],
)
