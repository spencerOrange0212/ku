
import customtkinter as ctk
import logging
from gui.main_app import ExcelToolApp
from utils.gui_logger import setup_logging
import os, sys  # 導入 os 和 sys 模組，用於檔案路徑操作和退出程序

# 取得根 Logger
logger = logging.getLogger()

if __name__ == "__main__":

    # 1. 初始化應用程式實例
    # 先初始化 app 實例，才能將其實例傳給 Log Handler
    try:
        def resource_path(relative_path):
            if hasattr(sys, "_MEIPASS"):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)
        ctk.set_default_color_theme(resource_path("theme/orange.json"))
        app = ExcelToolApp()
    except Exception as e:
        # 如果 ExcelToolApp 啟動失敗，則無法使用 GUI Log Handler
        print(f"致命錯誤：ExcelToolApp 初始化失敗。無法啟動 Log 系統。錯誤: {e}")
        # 在此處可以嘗試將錯誤寫入一個基本的txt檔案
        sys.exit(1)

    # 2. 設定 Log 系統
    # 傳入 app 實例，以便 GuiHandler 可以將 Log 導入 Textbox
    # 設置 Log 層級為 INFO，在發布版本中使用
    setup_logging(app, log_level=logging.INFO)

    # 3. 設定 CustomTkinter 基礎外觀
    # 這裡使用 app.resource_path 確保主題在 exe 模式下能被找到
    ctk.set_appearance_mode("dark")
    # 這裡假設您的 resource_path 在 ExcelToolApp 中，為了運行，我們簡化處理
    # ctk.set_default_color_theme(app.resource_path("theme/orange.json"))

    # 假設 orange.json 位於 CWD 或 _MEIPASS 根目錄，使用 os.path.join 尋找
    try:
        theme_path = os.path.join(os.path.dirname(sys.argv[0]), "theme/orange.json")
        if not os.path.exists(theme_path):
            theme_path = os.path.join(getattr(sys, '_MEIPASS', '.'), "theme/orange.json")

    except Exception as e:
        logger.warning(f"無法載入自訂主題，使用預設主題。錯誤: {e}")

    # 4. 運行應用程式主迴圈
    try:
        logger.info("啟動 CustomTkinter 主事件迴圈...")
        app.mainloop()
        logger.info("應用程式正常關閉。")
    except Exception as e:
        # 捕捉主迴圈期間的任何致命錯誤
        logger.critical("應用程式主迴圈發生致命錯誤。", exc_info=True)