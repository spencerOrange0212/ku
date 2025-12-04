# -*- mode: python ; coding: utf-8 -*-

import os
import sys
import json

# ----------------------------------------------------
# 1. 讀取設定檔並取得應用程式名稱
# ----------------------------------------------------
# 假設 config.json 在您的專案根目錄下的 config 資料夾中
# 如果 config.json 在專案根目錄，請改為 'config.json'
CONFIG_FILE = 'config/config.json'

def load_config(file_path):
    """安全地讀取 JSON 設定檔並返回其內容。"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"錯誤: 找不到設定檔 {file_path}")
        return None
    except json.JSONDecodeError:
        print(f"錯誤: 設定檔 {file_path} 格式不正確")
        return None
    except Exception as e:
        print(f"讀取設定檔時發生未知錯誤: {e}")
        return None

# 讀取設定檔
config_data = load_config(CONFIG_FILE)

# 預設名稱
app_name = '科餘自動化工具_Default'

if config_data and 'app_settings' in config_data:
     version_str = config_data['app_settings'].get('version', 'V')
     app_name = f"{config_data['app_settings'].get('author', '工具')}_{version_str.replace('.', '_')}"

# 取得路徑
SPEC_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
ICON_PATH = os.path.join(SPEC_DIR, "ai.ico")

# ----------------------------------------------------
# 2. 設定打包資源 (datas)
# ----------------------------------------------------
datas = [
    ('theme/orange.json', 'theme'),
    # ★ 關鍵：將 config/config.json 複製到輸出目錄下的 config 資料夾
    (CONFIG_FILE, 'config'),
    ('ai.ico', '.'),
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    datas=datas,
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

# ----------------------------------------------------
# 3. 建立 EXE (僅包含執行邏輯，不包含檔案)
# ----------------------------------------------------
exe = EXE(
    pyz,
    a.scripts,
    [],             # ★ 注意：這裡是空的，因為檔案要放到資料夾模式中
    exclude_binaries=True, # ★ 設為 True，表示不把二進制檔塞進 EXE
    name=app_name,
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
    icon=ICON_PATH,
)

# ----------------------------------------------------
# 4. 新增 COLLECT (建立資料夾並放入所有檔案)
# ----------------------------------------------------
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,       # ★ 這裡會把 config.json 和 theme 放進資料夾
    strip=False,
    upx=True,
    upx_exclude=[],
    name=app_name, # 這會是 dist 下面的資料夾名稱
)