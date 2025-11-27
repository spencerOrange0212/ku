import logging
from logging.handlers import RotatingFileHandler
import os
import sys
from datetime import datetime, timedelta
import re  # 需要用於解析日誌文件名中的日期


def get_log_file_dir_and_path(app_name="ExcelToolApp"):
    """
    根據作業系統返回 Log 檔案的儲存路徑，並以當前日期命名日誌檔。

    Log 檔案會放在應用程式執行檔或腳本所在的目錄下的 'logs' 子資料夾中。
    例如: C:\\YourAppFolder\\logs\\2025-11-27.log
    """
    # 判斷是 PyInstaller 打包模式還是開發模式
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包模式: 使用執行檔所在的目錄作為基底路徑
        base_path = os.path.dirname(sys.executable)
    else:
        # 開發模式: 使用腳本所在的目錄
        base_path = os.path.abspath(os.path.dirname(sys.argv[0]))

    # 將 Log 資料夾放在基底路徑下
    log_dir = os.path.join(base_path, "logs")

    # 確保 logs 目錄存在
    os.makedirs(log_dir, exist_ok=True)

    # 設置 Log 檔案名稱為今天的日期 (YYYY-MM-DD.log)
    today_str = datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"{today_str}.log")

    return log_dir, log_file


def cleanup_old_logs(log_dir, retention_days=182):
    """
    刪除超過指定天數的舊日誌檔案。
    Log 檔案格式必須是 YYYY-MM-DD.log。
    """
    cutoff_date = datetime.now() - timedelta(days=retention_days)
    # 正則表達式匹配 YYYY-MM-DD.log
    date_pattern = re.compile(r'(\d{4}-\d{2}-\d{2})\.log$')

    deleted_count = 0

    # 檢查目錄是否存在且可讀
    if not os.path.isdir(log_dir):
        return

    for filename in os.listdir(log_dir):
        match = date_pattern.match(filename)
        if match:
            try:
                # 解析日誌文件名中的日期
                log_date = datetime.strptime(match.group(1), '%Y-%m-%d')

                # 如果日誌日期早於截止日期，則刪除
                if log_date < cutoff_date:
                    file_path = os.path.join(log_dir, filename)
                    os.remove(file_path)
                    deleted_count += 1
            except ValueError:
                # 忽略文件名格式錯誤的檔案
                continue
            except OSError as e:
                # 忽略刪除失敗的檔案（例如權限問題或檔案被佔用）
                logging.warning(f"無法刪除舊日誌檔案 {filename} (可能被佔用): {e}")
                continue

    if deleted_count > 0:
        logging.info(f"已清理並刪除 {deleted_count} 個超過 {retention_days} 天的舊日誌檔案。")


class GuiHandler(logging.Handler):
    """
    自訂 Log Handler，將 Log 紀錄導向 CustomTkinter 應用程式的 append_log 方法。
    """

    def __init__(self, app_instance):
        super().__init__()
        # 儲存應用程式實例，以便呼叫其 append_log 方法
        self.app_instance = app_instance

    def emit(self, record):
        """處理並發送 Log 紀錄到 GUI"""
        try:
            # 格式化 Log 紀錄
            msg = self.format(record)

            # 使用 app_instance.append_log 安全地更新 GUI
            self.app_instance.append_log(msg)
        except Exception:
            # 如果 GUI 更新失敗，則將錯誤傳遞給預設錯誤處理機制
            self.handleError(record)


def setup_logging(app_instance, log_level=logging.INFO):
    """
    設定整個應用程式的 Log 紀錄系統。
    包含檔案輪換、控制台輸出，以及 GUI 顯示。
    """
    # 取得 Log 目錄和今天的 Log 檔案路徑
    log_dir, log_file = get_log_file_dir_and_path()

    # === 新增步驟：清理舊檔案 ===
    # 在日誌系統啟動前，先執行清理工作
    cleanup_old_logs(log_dir, retention_days=182)

    # 1. 根紀錄器設定
    logger = logging.getLogger()
    logger.setLevel(log_level)

    # 確保不會重複添加 handler，如果已經有 handler 則直接返回
    if logger.handlers:
        return

    # 2. 統一的格式器
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(name)s - %(message)s"
    )

    # 3. 檔案處理器 (FileHandler) - 穩定紀錄到磁碟
    # 使用 FileHandler 寫入今天的檔案。模式 'a' (append) 確保在今天內多次啟動會寫入同一個檔案。
    file_handler = logging.FileHandler(
        log_file,
        mode="a",
        encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # 4. 控制台處理器 (StreamHandler) - 僅在開發/除錯時輸出到終端
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    # 只有在非打包環境下才輸出到終端機，避免在 EXE 模式下彈出控制台
    if not getattr(sys, 'frozen', False):
        logger.addHandler(stream_handler)

    # 5. GUI 處理器 (GuiHandler) - 輸出到應用程式的 Textbox
    gui_handler = GuiHandler(app_instance)
    # GUI 顯示層級設為 INFO
    gui_handler.setLevel(logging.INFO)
    gui_handler.setFormatter(formatter)
    logger.addHandler(gui_handler)

    logging.info(f"Log 紀錄系統已啟動。今天的 Log 檔案路徑: {log_file}")

    # 嘗試獲取應用程式標題，並紀錄
    try:
        logging.info(f"應用程式版本: {app_instance.title()}")
    except Exception:
        # 在極少數情況下，如果 title 呼叫失敗，則靜默處理
        pass