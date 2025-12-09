# 檔案: config_manager.py
import json
import os
from typing import Dict, Any, Union  # 調整 Union 導入


class ConfigManager:
    """
    設定管理類別：負責載入 JSON 設定檔並提供設定數值。
    當檔案不存在時，會自動建立預設配置。
    """

    # -----------------------------------------------------------------
    # ⭐️ 新增 1: 預設配置結構
    # -----------------------------------------------------------------
    DEFAULT_CONFIG = {
        "module_management": {
            "1_download": False,
            "2_insert": False,
            "3_update": True,
            "4_delete": True
        },
        "file_handling": {
            "overwrite": True
        }
    }

    # -----------------------------------------------------------------

    def __init__(self, config_file: str = 'config/config.json'):
        self.config_file = config_file
        self._config_data: Dict[str, Any] = {}
        self._load_config()

    def _load_config(self) -> None:
        """
        私有方法，用於從 JSON 檔案載入設定。
        若檔案不存在或損壞，則建立並載入預設設定。
        """

        # -----------------------------------------------------------------
        # ⭐️ 修正 2: 處理檔案不存在或損壞的邏輯
        # -----------------------------------------------------------------
        if not os.path.exists(self.config_file):
            print(f"⚠️ 找不到設定檔 '{self.config_file}'。將自動建立預設設定。")
            self._create_default_config()
            return

        try:
            # (2) 實際連接並開啟檔案
            with open(self.config_file, 'r', encoding='utf-8') as f:
                # (3) 讀取 JSON 內容
                self._config_data = json.load(f)
            print(f"✅ 設定檔 '{self.config_file}' 成功載入。")

        except json.JSONDecodeError:
            # 處理 JSON 格式錯誤
            print(f"❌ 錯誤：設定檔 '{self.config_file}' 格式不正確。將自動使用預設設定覆蓋。")
            self._create_default_config()

        except Exception as e:
            # 處理其他可能的 I/O 錯誤
            print(f"❌ 載入設定檔時發生未知錯誤: {e}。將使用空設定。")
            self._config_data = {}

    def _create_default_config(self) -> None:
        """
        建立配置檔案的路徑，並將 DEFAULT_CONFIG 寫入檔案。
        """
        # 確保配置檔案的目錄存在
        config_dir = os.path.dirname(self.config_file)
        if config_dir and not os.path.exists(config_dir):
            os.makedirs(config_dir)

        self._config_data = self.DEFAULT_CONFIG.copy()  # 使用 copy 避免後續修改影響 DEFAULT_CONFIG
        self.save()  # 寫入磁碟
        print(f"📝 預設設定已建立並寫入至 '{self.config_file}'。")

    def save(self) -> None:
        """將當前設定寫入 JSON 檔案。"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                # 使用 indent=2 提高檔案的可讀性
                json.dump(self._config_data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"❌ 儲存設定檔時發生錯誤: {e}")

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

    def set(self, key_path: str, value: Any) -> None:
        """
        依據路徑設定數值。路徑使用點號 '.' 分隔。
        注意：這會修改記憶體中的配置，但不會自動儲存到磁碟。
        """
        keys = key_path.split('.')
        node = self._config_data

        # 遍歷到倒數第二層
        for key in keys[:-1]:
            if key not in node or not isinstance(node[key], dict):
                node[key] = {}
            node = node[key]

        # 設定最終值
        node[keys[-1]] = value

    def reload(self) -> None:
        """重新載入設定檔。"""
        self._load_config()


# 在模組層級實例化一個全域的設定管理器 (供其他模組導入)
CONFIG = ConfigManager()