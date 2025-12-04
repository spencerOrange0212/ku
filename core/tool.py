import sys
import os


def resource_path(relative_path):
    """
        獲取資源文件的絕對路徑。
        """
    try:
        # 🟢 如果是打包後的 exe 執行環境，使用 _MEIPASS 臨時路徑
        base_path = sys._MEIPASS
    except AttributeError:
        # 🔵 如果是開發環境，使用當前執行目錄
        # 由於 icon 文件通常放在根目錄，所以 os.path.abspath(".") 是可以的
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

