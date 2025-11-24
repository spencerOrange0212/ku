# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from config.settings import APP_NAME, VERSION

# 使用 sys.argv[0] 來獲取正在執行的指令碼路徑 (即 .spec 檔案的路徑)
SPEC_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
ICON_PATH = os.path.join(SPEC_DIR, "ai.ico")

# 把 theme 資料夾一起打包進 exe
datas = [
    ('theme/orange.json', 'theme'),    # ← 這行確保 theme/orange.json 打包後會出現在 exe 裡的 theme 資料夾
 ('ai.ico', '.'),   # ★ 把 ai.ico 放進 exe 同層
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    datas=datas,        # ★ 把 theme 資料夾整個打包進去
    binaries=[],
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
    name="{APP_NAME}",
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
    icon=ICON_PATH,   # ★ 用上面算好的絕對路徑
)

