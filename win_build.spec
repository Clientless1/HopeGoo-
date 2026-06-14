# -*- mode: python -*-
import sys,os
a = Analysis(
    ['hopegoo_gcoins.py'],
    pathex=[],
    binaries=[],
    datas=[('capture_addon.py','.'),('cities.json','.'),('t2s_lite.py','.'),('accounts.json','.')],
    hiddenimports=['requests','tkinter','docx','urllib3'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name='HopeGoo酒店筛选', debug=False, strip=False, upx=True, console=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=True, name='HopeGoo酒店筛选')
app = BUNDLE(coll, name='HopeGoo酒店筛选.app')
