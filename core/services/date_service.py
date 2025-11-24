from datetime import datetime
from tkinter import messagebox

class DateService:
    """民國年月處理與驗證"""

    @staticmethod
    def to_ad(roc_ym: str, label: str = "日期") -> datetime:
        """將民國年月轉西元（含欄位名稱提示）"""
        if not roc_ym:
            raise ValueError(f"請輸入 {label}")
        if len(roc_ym) != 5 or not roc_ym.isdigit():
            raise ValueError(f"{label} 格式錯誤，請輸入 5 碼民國年月，例如 11311")
        year = int(roc_ym[:3]) + 1911
        month = int(roc_ym[3:])
        if not (1 <= month <= 12):
            raise ValueError(f"{label} 的月份需介於 01~12")
        return datetime(year, month, 1)

    @classmethod
    def validate(cls, latest: str, make: str, latest_label="最新科餘時間", make_label="製作科餘時間") -> bool:
        """
        驗證：
        1️⃣ 兩欄位格式正確
        2️⃣ 製作科餘時間 >= 最新科餘時間
        """
        try:
            latest_dt = cls.to_ad(latest, latest_label)
            make_dt = cls.to_ad(make, make_label)

            if make_dt < latest_dt:
                messagebox.showwarning("時間錯誤", f"{make_label} 不可早於 {latest_label}！")
                return False

            return True

        except ValueError as e:
            messagebox.showerror("格式錯誤", str(e))
            return False
