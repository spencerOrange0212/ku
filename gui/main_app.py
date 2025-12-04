import os
from PIL import Image, ImageTk, ImageSequence
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os, sys
from customtkinter import CTkImage

from config.settings import VERSION, APP_NAME
from core.actions.confirm_action import do_actions_sequential
from core.controllers.excel_controller import ExcelController
from core.tool import resource_path
from core.validators.confirm_action import validate_before_action


class ExcelToolApp(ctk.CTk):
    def __init__(self):



        super().__init__()

        self.spinner_frames = None
        self.spinner_label = None
        self.spinner_running = False
        self.wm_iconbitmap(resource_path("ai.ico"))
        self.title(f"{APP_NAME} v{VERSION}")
        self.geometry("600x520")
        self.minsize(700, 650)
        # 控制器（邏輯交由 controller）
        self.controller = ExcelController(self)
        self.cancel_requested = False
        self.create_widgets()
        self.wm_attributes('-topmost', 0)
    def create_widgets(self):
        # =====================
        # 📂 頂部：匯入 Excel + 日期輸入
        # =====================
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=20, pady=15)

        # 匯入 Excel
        self.load_label = ctk.CTkLabel(top_frame, text="📂 匯入 Excel", width=50, wraplength=180, anchor="w")
        self.load_label.grid(row=0, column=0, padx=10, pady=10)

        self.output_label = ctk.CTkLabel(top_frame, text="未設定輸出路徑", width=50, wraplength=180, anchor="w")
        self.output_label.grid(row=0, column=1, padx=10, sticky="w")

        ctk.CTkButton(top_frame, text="選擇檔案", command=self.controller.load_excel).grid(row=1, column=0, padx=5)

        ctk.CTkButton(top_frame, text="設定輸出路徑", command=self.controller.choose_output_folder).grid(row=1,
                                                                                                         column=1,
                                                                                                         padx=5)
        ctk.CTkButton(top_frame, text="🧹 清除檔案", command=self.controller.clear_excel).grid(row=1, column=2, padx=5)

        # 日期輸入
        ctk.CTkLabel(top_frame, text="📅 最新科餘時間（民國年月）").grid(row=2, column=0, padx=5, pady=5)
        ctk.CTkLabel(top_frame, text="📅 製作科餘時間（民國年月）").grid(row=2, column=1, padx=5, pady=5)
        ctk.CTkLabel(top_frame, text="🏢 選取廠商").grid(row=2, column=2, padx=5, pady=5)

        # 變數宣告改成 MemoryEntry
        from gui.widgets.MemoryEntry import MemoryEntry  # 假設你把上次的 MemoryEntry 寫在這個檔案

        self.latest_var = MemoryEntry(top_frame, key="latest_month", default="")
        self.latest_var.grid(row=3, column=0, padx=5)

        self.make_var = MemoryEntry(top_frame, key="make_month", default="")
        self.make_var.grid(row=3, column=1, padx=5)


        # 🏢 統一編號記憶式下拉選單
        from gui.widgets.memory_combobox import VendorConfigManager  # ← 確認有這行

        self.tax_id_box = VendorConfigManager(top_frame, file_path="tax_id_memory.json")
        self.tax_id_box.grid(row=3, column=2, padx=5, sticky="w")

        # =====================
        # ⚙️ 工具列 + 狀態列（合併區塊）
        # =====================
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)


        # 工具選取 (Checkbox)
        tool_frame = ctk.CTkFrame(main_frame)
        tool_frame.pack(pady=(0, 15), fill="x")

        ctk.CTkLabel(tool_frame, text="請選擇要執行的工具：").pack(anchor="w", padx=10, pady=5)

        # 三個工具的選項變數
        self.option_insert = ctk.BooleanVar()
        self.option_update = ctk.BooleanVar()
        self.option_delete = ctk.BooleanVar()

        checkbox_row = ctk.CTkFrame(tool_frame)
        checkbox_row.pack(anchor="w", padx=20, pady=5)  # ⭐ 讓 frame 填滿水平
        checkbox_row.configure(fg_color="transparent")  # ⭐ 不要背景色（最乾淨）
        # 三個水平排列的 Checkbox
        ctk.CTkCheckBox(
            checkbox_row,
            text="📊 報表貼入科目",
            variable=self.option_insert,
        ).pack(side="left", padx=10)

        ctk.CTkCheckBox(
            checkbox_row,
            text="🧩 科目更新",
            variable=self.option_update,
        ).pack(side="left", padx=10)

        ctk.CTkCheckBox(
            checkbox_row,
            text="🗑️ 科目明細刪除",
            variable=self.option_delete,
        ).pack(side="left", padx=10)
        # 建立水平容器（放執行與停止按鈕）
        button_row = ctk.CTkFrame(tool_frame)
        button_row.pack(pady=(10, 5), fill="x")  # ⭐ 讓 frame 填滿水平
        button_row.configure(fg_color="transparent")  # ⭐ 不要背景色（最乾淨）

        # 🚀 執行按鈕
        ctk.CTkButton(
            button_row,
            text="🚀 執行選取的工具",
            height=45,
            width=160,
            command=self.do_exe
        ).pack(side="left", padx=10)

        # ⛔ 停止按鈕
        self.stop_button = ctk.CTkButton(
            button_row,
            text="⛔ 立即停止執行",
            height=45,
            width=140,
            fg_color="red",
            hover_color="#b30000",
            state="disabled",
            command=self.request_cancel
        )
        self.stop_button.pack(side="left", padx=10)
        # ⭐ 停止鍵隱藏
        self.stop_button.pack_forget()

        # ⬇️ 底部統一容器：狀態 + 執行紀錄
        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        bottom_frame.pack(fill="both", expand=True, pady=(5, 10), padx=5)

        # -------------------------------
        # ⭐ 狀態列（不包外框，看起來更自然）
        # -------------------------------
        status_row = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        status_row.pack(fill="x", pady=(0, 5))

        self.spinner_label = ctk.CTkLabel(status_row, text="", anchor="center")
        self.spinner_label.pack(side="left", padx=(0, 5))

        self.status_label = ctk.CTkLabel(status_row, text="狀態：等待操作", anchor="w")
        self.status_label.pack(side="left")



        # -------------------------------
        # ⭐ 執行紀錄（只有 textbox，無多餘框）
        # -------------------------------
        ctk.CTkLabel(bottom_frame, text="📜 執行紀錄：").pack(anchor="w", pady=(0, 0))

        self.log_text = ctk.CTkTextbox(
            bottom_frame,
            height=200,
            wrap="word",
            fg_color="#3f3a3a",
            text_color="#ffffff",
            border_width=1,
            border_color="#CCCCCC"
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_text.configure(state="disabled")

        # ==========================================================
        # 🟢 新增：底部版權字樣
        # ==========================================================
        # 假設您的版權資訊是 "© 2024 Your Company Name. All Rights Reserved."
        copyright_text = f"© 2025 直誠管顧. Designed by spencer. All Rights Reserved. | {APP_NAME} v{VERSION}"
        self.copyright_label = ctk.CTkLabel(
            self,
            text=copyright_text,
            text_color="#888888",  # 柔和的灰色
            font=ctk.CTkFont(size=11)
        )
        # pack 在主視窗底部，給予微小的邊距
        self.copyright_label.pack(side="bottom", pady=(0, 5))

    def run_process(self):
        """GUI 觸發 → 呼叫控制器進行處理"""
        latest = self.latest_var.get()
        make = self.make_var.get()

        try:
            msg = self.controller.execute(latest, make)
            self.status_label.configure(text=f"狀態：{msg}")
            messagebox.showinfo("成功", str(msg))  # ← 確保為字串
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
            self.status_label.configure(text="狀態：發生錯誤")

    def _load_spinner_frames(self):
        """載入 GIF 轉圈圈動畫"""
        spinner_path = "gui/assets/spinner.gif"
        img = Image.open(spinner_path)
        target_size = (32, 32)

        self.spinner_frames = [
            CTkImage(frame.copy().resize(target_size))
            for frame in ImageSequence.Iterator(img)
        ]

    def do_exe(self):

        # 1️⃣ 檢查是否有任何工具被勾選
        if not any([
            self.option_insert.get(),
            self.option_update.get(),
            self.option_delete.get()
        ]):
            messagebox.showwarning("提示", "至少要選擇一個功能")
            return

        # 2️⃣ 根據 checkbox 狀態建立要執行的任務列表
        tasks = []

        if self.option_insert.get():
            tasks.append(("insert_report", "📊 報表貼入科目"))

        if self.option_update.get():
            tasks.append(("update_subjects", "🧩 科目更新"))

        if self.option_delete.get():
            tasks.append(("delete_details", "🗑️ 科目明細刪除"))

        # 3️⃣ 先做一次整體驗證


        ok, msg = validate_before_action(
            file_path=getattr(self.controller, "file_path", None),
            tax_id=self.tax_id_box.get(),
            make_month=self.make_var.get(),
            latest_month=self.latest_var.get(),
            tasks = [task[0] for task in tasks]  # 只傳任務代號列表
        )
        if not ok:
            messagebox.showwarning("錯誤", msg)
            return

        # 4️⃣ 一次確認所有要執行的工具
        tools_text = "\n".join(f"• {name}" for _, name in tasks)
        confirm_msg = f"你選擇要執行以下工具：\n\n{tools_text}\n\n是否確定要執行？"

        if not messagebox.askyesno("確認執行", confirm_msg):
            self.status_label.configure(text="狀態：已取消執行")
            return

        # 5️⃣ 通過確認後，依序執行所有勾選的工具
        do_actions_sequential(self, tasks)
    def append_log(self, msg: str):
        """安全地在 GUI 中追加一行 log（支援背景 thread 呼叫）"""
        print(msg)  # 也印在 console

        def _append():
            if not hasattr(self, "log_text"):
                return

            # 先解鎖才能寫入
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")  # 自動滾到最下面
            # 再鎖回唯讀，使用者就不能輸入/修改
            self.log_text.configure(state="disabled")

        # 丟回主執行緒執行 UI 更新
        self.after(0, _append)


    def request_cancel(self):
        self.cancel_requested = True
        self.append_log("⛔ 使用者要求停止執行。")
        messagebox.showinfo("停止", "正在嘗試停止目前執行的工具…")
