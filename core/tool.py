import sys
import os


def resource_path(relative_path):
    """
    獲取資源文件的絕對路徑。

    用於解決 PyInstaller 打包成 exe 後，
    資源檔案會被解壓到臨時目錄 (_MEIPASS) 的路徑問題。
    """
    if hasattr(sys, "_MEIPASS"):
        # 如果是打包後的 exe 執行環境
        return os.path.join(sys._MEIPASS, relative_path)

    # 如果是開發環境 (直接執行 .py)
    return os.path.join(os.path.abspath("."), relative_path)


