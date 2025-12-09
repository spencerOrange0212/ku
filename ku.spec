# -*- mode: python ; coding: utf-8 -*-

import os
import sys
import datetime # 引入 datetime 模組

# ----------------------------------------------------
# ⭐️ 核心修正：使用當前日期與時間（YYYYMMDD_HHMMSS） ⭐️
# ----------------------------------------------------
# 取得執行 PyInstaller 當下的日期和時間，並格式化為 YYYYMMDD_HHMMSS 字串
current_datetime_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
base_name = '科餘自動化工具' # 基礎名稱

# 組合最終的 exe 名稱
app_name = f"{base_name}_{current_datetime_str}"
# 範例結果：科餘自動化工具_20251209_165215

# ----------------------------------------------------
# 檔案路徑與打包資料 (其餘部分不變)
# ----------------------------------------------------
SPEC_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
ICON_PATH = os.path.join(SPEC_DIR, "ai.ico")

datas = [
    ('theme/orange.json', 'theme'),
    ('ai.ico', '.'),
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
    name=app_name,
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

