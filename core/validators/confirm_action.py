from tkinter import messagebox

import re

# ✅ validators/precheck.py
def validate_before_action(file_path, tax_id, make_month, latest_month,tasks):
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

    if "update_subjects" in tasks:
        # 格式檢查
        if len(make_month) != 5 or len(latest_month) != 5:
            return False, "月份格式錯誤，應為 5 碼 (YYYMM)"
        try:
            make_year = int(make_month[:3])
            make_mon = int(make_month[3:])
            latest_year = int(latest_month[:3])
            latest_mon = int(latest_month[3:])
        except ValueError:
            return False, "月份必須為數字格式"

        # 規則 1：make_month 必須大於 latest_month
        if (make_year < latest_year) or (make_year == latest_year and make_mon <= latest_mon):
            return False, "製表月份必須大於最新月份"

        # 規則 2：如果年份差 1，latest_month 月份必須是 12 月
        if make_year - latest_year == 1 and latest_mon != 12:
            return False, f"最新月份年份小於製表年份 1，最新月份必須為 12 月"


    # ✅ 若全部檢查通過
    return True, "輸入檢查通過"
