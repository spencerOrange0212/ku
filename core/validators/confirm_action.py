from tkinter import messagebox
from core.services.date_service import DateService
import re

# ✅ validators/precheck.py
def validate_before_action(file_path, tax_id, make_month, latest_month):
    """
    驗證執行前輸入是否正確。
    回傳：
        (True, "通過檢查") 或 (False, "錯誤訊息")
    """

    # ✅ 1️⃣ 檢查檔案
    if not file_path:
        return False, "請先選擇 Excel 檔案。"

    # ✅ 2️⃣ 統一編號檢查
    # if not tax_id:
    #     return False, "請輸入統一編號。"
    # if len(tax_id) != 8:
    #     return False, "統一編號必須為 8 碼數字。"
    # if not tax_id.isdigit():
    #     return False, "統一編號僅能包含數字。"

    # ✅ 3️⃣ 填製月份檢查
    if not make_month:
        return False, "請輸入填製月份（例如 114-10）。"

    # ✅ 4️⃣ 最新科餘時間檢查
    if not latest_month:
        return False, "請輸入最新科餘時間（例如 114-09）。"

    # ✅ 5️⃣ 民國年月格式簡單驗證
    # ✅ 僅允許「11410」這種格式（五位數字）
    month_pattern = re.compile(r"^1\d{2}(0[1-9]|1[0-2])$")

    if not month_pattern.match(make_month):
        return False, "填製月份格式錯誤，請使用民國年月（例如 11410，月份01-12，年份1開頭）。"

    if not month_pattern.match(latest_month):
        return False, "最新科餘時間格式錯誤，請使用民國年月（例如 11409，月份01-12，年份1開頭）。"

    # ✅ 若全部檢查通過
    return True, "輸入檢查通過"
