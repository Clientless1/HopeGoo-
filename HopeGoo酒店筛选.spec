# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['hopegoo_gcoins.py'],
    pathex=[],
    binaries=[],
    datas=[('hopegoo_api.py', '.'), ('capture_addon.py', '.'), ('cities.json', '.'), ('/opt/homebrew/bin/mitmdump', '.'), ('/opt/homebrew/bin/opencc', '.')],
    hiddenimports=['requests', 'urllib3', 'docx', 'tkinter'],
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
    [],
    exclude_binaries=True,
    name='HopeGoo酒店筛选',
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
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='HopeGoo酒店筛选',
)
app = BUNDLE(
    coll,
    name='HopeGoo酒店筛选.app',
    icon=None,
    bundle_identifier=None,
)
