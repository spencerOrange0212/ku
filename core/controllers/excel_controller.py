import os
import time
from tkinter import filedialog, messagebox

from core.services import subject_paste_service
from core.services.SubjectDeleteService import SubjectDeleteService
from core.services.date_service import DateService
from core.services.path_service import PathService
from core.services.excel_service import ExcelService
from core.services.subject_paste_service import SubjectPasteService
from core.services.subject_update_service import SubjectUpdateService


class ExcelController:
    """負責整合 GUI 事件與 Service"""

    def __init__(self, app):
        self.app = app  # app 就是 ExcelToolApp 視窗
        self.excel_service = ExcelService()
        self.date_service = DateService()
        self.path_service = PathService()
        self.file_path = None
        self.output_path = None
        # ⭐ 關鍵修正：在此處初始化 SubjectPasteService
        # 之前就是少了這一行導致 'no attribute subject_paste_service' 錯誤
        self.subject_paste_service = SubjectPasteService()
        # 用於儲存執行當下的環境變數 (廠商ID, 設定, 月份)
        self.context = {
            "vendor_id": None,
            "config": None,
            "make_month": None
        }

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            self.app.load_label.configure(text="上傳失敗")
            return
        self.file_path = path
        self.app.load_label.configure(text=f"📂 已匯入檔案：{path}", width=50, )
        messagebox.showinfo("成功", f"已匯入檔案：{os.path.basename(path)}")

    def choose_output_folder(self):
        from tkinter import filedialog, messagebox
        folder = filedialog.askdirectory(title="選擇輸出資料夾")
        if not folder:
            self.app.output_label.configure(text="未設定輸出路徑")
            return

        self.output_path = folder
        self.app.output_label.configure(text=f"📂 輸出路徑：{folder}")
        messagebox.showinfo("設定成功", f"已設定輸出資料夾：\n{folder}")

    def execute(self, latest: str, make: str) -> str:
        """主要業務邏輯控制流程"""
        if not self.file_path:
            raise ValueError("尚未選擇 Excel 檔案")

        if not self.date_service.validate(latest, make):
            raise ValueError("時間輸入錯誤")

        msg = self.excel_service.process_file(self.file_path, latest, make, self.output_path)
        return msg

        # 模組 1：報表貼入科目
        # =========================================================================

    def run_insert_report(self, file_path: str):  # ⭐️ 必須接受 file_path 參數
        """
        執行「報表貼入科目」
        使用 do_actions_sequential 傳入的 file_path (科餘檔路徑) 作為貼入目標。
        """

        # 1. 取得 GUI 參數
        master_file = file_path
        vendor_id = self.app.tax_id_box.get()
        make_month = self.app.make_var.get().strip()

        # 2. 基礎檢查
        if not master_file:
            raise ValueError("請先匯入 Excel (科餘檔)。")
        if not vendor_id:
            raise ValueError("請輸入廠商代號。")
        if not make_month:
            raise ValueError("請輸入製作年月。")

        # 3. 取得原始資料路徑 (使用唯一正確 Key)
        source_folder = None

        try:
            # 獲取廠商完整設定 (使用正確方法)
            vendor_id_check, config_data = self.app.tax_id_box.get_current_settings()

            if not config_data:
                raise ValueError(f"無法獲取廠商 {vendor_id} 的設定資料。")

            # ⭐️ 修正：只讀取唯一正確的 Key
            source_folder = config_data.get("input_folder")

        except Exception as e:
            # 捕捉任何讀取錯誤，並將其歸類為設定失敗
            raise ValueError(f"讀取廠商設定時發生錯誤：{e}")

        # 4. 路徑最終檢查
        if not source_folder or not os.path.isdir(source_folder):  # 增加檢查路徑是否存在
            # 這裡會捕捉到 "input_folder" 裡面的值為空字串或路徑不存在的情況
            raise ValueError(f"❌ 廠商 [{vendor_id}] 設定中，[原始資料路徑] 無效或不存在，請至廠商資料管理介面設定。")

        self.app.append_log(f"📂 確定使用來源資料夾：{source_folder}")


        service = SubjectPasteService(logger=self.app.append_log, app=self.app)

        service.execute_paste_task(
            input_folder=source_folder,
            make_month=make_month,
            vendor_id=vendor_id,
            master_file_path=master_file
        )

        return f"報表貼入完成！(廠商: {vendor_id})"

    def run_update_subjects(self, file_path):
        """
        執行「科目檢查」：
        1️⃣ 呼叫 SubjectUpdateService 進行餘額比對
        2️⃣ 若有錯誤 → raise 讓外層 thread 捕捉
        3️⃣ 若全部一致 → 進入下一步
        """
        service = SubjectUpdateService(
            file_path=file_path,
            logger=self.app.append_log,  # 寫 log 到 GUI
            app=self.app  # ⭐ 這行很重要，給 _check_cancel 用
        )

        latest_month = self.app.latest_var.get().strip()
        make_month = self.app.make_var.get().strip()  # 製作科餘年月

        # 先跑檢查
        result = service.run_check(latest_month)

        if not isinstance(result, dict):
            raise ValueError("回傳結果格式異常，預期為 dict")

        if result["status"] == "error":
            # 錯誤訊息已經寫進 log 了，這裡丟出去給 do_actions_sequential 處理
            raise Exception(result["message"])

        # ✅ 若成功 → 顯示提示並執行下一步
        if self.app:
            messagebox.showinfo("完成", "✅ 所有項目均一致，開始進入下一步。")

        # ✅ 呼叫下一步（更新科目分頁）
        service.run_copy_data(make_month, latest_month)

        return result["message"]

    def run_delete_details(self, file_path):
        """
        第四模組：科目明細刪除
        - 依「更新清單_XXXX」工作表中的科目列表
        - 到各科目分頁進行摘要分組，若 F/G 加總相等則刪除
        """
        make_month = self.app.make_var.get().strip()
        latest_month = self.app.latest_var.get().strip()

        service = SubjectDeleteService(
            file_path,
            logger=self.app.append_log,  # log 丟到 GUI 右下角紀錄區
            app=self.app  # 讓 service 可以讀取 cancel_requested
        )

        msg = service.run_delete(make_month, latest_month)
        return msg

    def clear_excel(self):
        """清除目前載入的 Excel 檔案與顯示文字"""
        self.file_path = None
        self.app.output_label.configure(text="未設定輸出路徑")
        self.app.load_label.configure(text="請重新上傳檔案", width=50, )
        self.app.status_label.configure(text="狀態：已清除載入的檔案")
