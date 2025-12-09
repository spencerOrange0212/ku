# 檔案: gui/settings_window.py

import customtkinter as ctk
from tkinter import messagebox
from config.ConfigManager import CONFIG
from typing import TYPE_CHECKING, Dict, Any


# 避免循環引用，只在類型註釋時使用 ExcelToolApp
if TYPE_CHECKING:
    from gui.main_app import ExcelToolApp

class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, master_app: 'ExcelToolApp'):
        super().__init__(master_app)
        self.master_app = master_app
        self.title("程式設定")
        self.geometry("450x380")
        self.resizable(False, False)

        self.wm_attributes('-topmost', 1)
        self.grab_set()

        self.module_vars: Dict[str, ctk.BooleanVar] = {}
        # 讀取當前的 overwrite 狀態
        self.overwrite_var = ctk.BooleanVar(value=CONFIG.get("file_handling.overwrite", default=False))

        self.create_settings_widgets()

    def create_settings_widgets(self):
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # -------------------------------
        # 🚀 模組功能設定 (module_management)
        # -------------------------------
        module_frame = ctk.CTkFrame(main_frame)
        module_frame.pack(pady=(5, 15), fill="x", padx=10)

        ctk.CTkLabel(module_frame, text="🚀 執行模組開關 (Check/Uncheck to enable/disable)",
                     font=ctk.CTkFont(weight="bold")).pack(anchor='w', padx=10, pady=(10, 5))

        module_settings = CONFIG.get('module_management', {})
        for key, default_val in module_settings.items():
            var = ctk.BooleanVar(value=default_val)
            self.module_vars[key] = var

            # 建立 Checkbox
            text_name = key.split('_', 1)[-1].capitalize()
            ctk.CTkCheckBox(
                module_frame,
                text=f"{key}. {text_name}",
                variable=var
            ).pack(anchor='w', padx=20, pady=2)

        # -------------------------------
        # 💾 檔案儲存方式 (file_handling.overwrite)
        # -------------------------------
        save_frame = ctk.CTkFrame(main_frame)
        save_frame.pack(pady=(15, 10), fill="x", padx=10)

        ctk.CTkLabel(save_frame, text="💾 檔案儲存方式", font=ctk.CTkFont(weight="bold")).pack(anchor='w', padx=10,
                                                                                              pady=(10, 5))

        # Radiobutton 選項
        ctk.CTkRadioButton(save_frame,
                           text="覆蓋原始檔案 (overwrite: True)",
                           variable=self.overwrite_var,
                           value=True).pack(anchor='w', padx=20, pady=2)
        ctk.CTkRadioButton(save_frame,
                           text="另存新檔 (overwrite: False)",
                           variable=self.overwrite_var,
                           value=False).pack(anchor='w', padx=20, pady=2)

        # --- 儲存按鈕 ---
        ctk.CTkButton(
            self,
            text="✅ 儲存設定並關閉",
            command=self.save_and_close,
            fg_color="#008000",
            hover_color="#006400"
        ).pack(pady=(0, 20))

    def save_and_close(self):
        """將變更寫回 CONFIG 並關閉視窗"""

        # 1. 更新 module_management
        for key, var in self.module_vars.items():
            CONFIG.set(f"module_management.{key}", var.get())

            # 2. 更新 file_handling.overwrite
            CONFIG.set("file_handling.overwrite", self.overwrite_var.get())

        # 3. 儲存到 JSON 檔案
        CONFIG.save()

        messagebox.showinfo("設定成功", "程式設定已儲存並生效！")
        self.destroy()