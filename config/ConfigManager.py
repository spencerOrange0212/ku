# 檔案: config_manager.py
import json
import os  # ★ 新增：用於檢查檔案是否存在
from typing import Dict, Any


class ConfigManager:
    """
    設定管理類別：負責載入 JSON 設定檔並提供設定數值。
    """

    def __init__(self, config_file: str = 'config/config.json'):
        self.config_file = config_file
        self._config_data: Dict[str, Any] = {}
        self._load_config()

    def _load_config(self) -> None:
        """
        私有方法，用於從 JSON 檔案載入設定。
        包含：檔案存在檢查、JSON 格式錯誤處理。
        """

        # -----------------------------------------------------------------
        # ★ 檔案存在檢查
        if not os.path.exists(self.config_file):
            print(f"⚠️ 錯誤：找不到設定檔 '{self.config_file}'。")
            print("將使用空設定。請檢查檔案路徑。")
            return
        # -----------------------------------------------------------------

        try:
            # (2) 實際連接並開啟檔案
            with open(self.config_file, 'r', encoding='utf-8') as f:
                # (3) 讀取 JSON 內容
                self._config_data = json.load(f)
            print(f"✅ 設定檔 '{self.config_file}' 成功載入。")

        except json.JSONDecodeError:
            # 處理 JSON 格式錯誤
            print(f"❌ 錯誤：設定檔 '{self.config_file}' 格式不正確，無法解析 JSON。")
            self._config_data = {}

        except Exception as e:
            # 處理其他可能的 I/O 錯誤
            print(f"❌ 載入設定檔時發生未知錯誤: {e}")
            self._config_data = {}

    def get(self, key_path: str, default: Any = None) -> Any:
        """
        依據路徑獲取設定數值。路徑使用點號 '.' 分隔。
        """
        keys = key_path.split('.')
        current = self._config_data

        try:
            for key in keys:
                if isinstance(current, dict) and key in current:
                    current = current[key]
                else:
                    raise KeyError(f"找不到路徑中的鍵: {key}")
            return current
        except KeyError:
            return default

    def reload(self) -> None:
        """重新載入設定檔。"""
        self._load_config()


# 在模組層級實例化一個全域的設定管理器 (供其他模組導入)
CONFIG = ConfigManager()