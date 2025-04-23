# -*- coding: utf-8 -*-
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT, BUNDLE

a = Analysis(
    ['bscscan_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BSCScan Automation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
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
    name='BSCScan Automation',
)

app = BUNDLE(
    coll,
    name='BSCScan Automation.app',
    bundle_identifier='com.bscscan.automation',
    info_plist={
        'NSHighResolutionCapable': 'True',
    },
)