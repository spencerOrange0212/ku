# -*- mode: python ; coding: utf-8 -*-

import os
import sys
import json # ★ 新增：匯入 json 模組

# ----------------------------------------------------
# 讀取設定檔並取得應用程式名稱
# ----------------------------------------------------
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
# 預設名稱，以防讀取失敗
app_name = '科餘自動化工具_Default'

if config_data and 'app_settings' in config_data:
    # 假設您希望使用 "科餘自動化工具" + "版本號" 作為名稱，但如果只想使用作者名稱，請調整路徑

    # 方式二: 使用版本號組合名稱 (如果需要版本資訊)
     version_str = config_data['app_settings'].get('version', 'V')
     app_name = f"{config_data['app_settings'].get('author', '工具')}_{version_str.replace('.', '_')}"


# 使用 sys.argv[0] 來獲取正在執行的指令碼路徑 (即 .spec 檔案的路徑)
SPEC_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
ICON_PATH = os.path.join(SPEC_DIR, "ai.ico")

# 把 theme 資料夾一起打包進 exe
datas = [
    ('theme/orange.json', 'theme'),    # ← 這行確保 theme/orange.json 打包後會出現在 exe 裡的 theme 資料夾
    (CONFIG_FILE, 'config'),
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

