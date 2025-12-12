import customtkinter as ctk
import json
import os
from tkinter import messagebox

class MemoryEntry(ctk.CTkFrame):
    """可記憶的輸入框，會自動讀取/存檔同一 JSON 檔案的指定欄位"""

    shared_file = "config/memory.json"  # 所有 MemoryEntry 共用檔案

    def __init__(self, master, key, width=120, height=30, default="", **kwargs):
        """
        key: JSON 中的欄位名稱，例如 'latest_month', 'make_month'
        """
        super().__init__(master, width=width, height=height, **kwargs)
        self.key = key
        self.default = default
        self.value = self._load_value()

        self.var = ctk.StringVar(value=self.value)
        self.entry = ctk.CTkEntry(self, textvariable=self.var, width=width)
        self.entry.pack(fill="both", expand=True)

        # 當失焦時自動儲存
        self.entry.bind("<FocusOut>", lambda e: self._save_value())

    def get(self):
        return self.var.get().strip()

    def set(self, value):
        self.var.set(value)
        self._save_value()

    def _load_value(self):
        if os.path.exists(self.shared_file):
            try:
                with open(self.shared_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return data.get(self.key, self.default)
            except:
                return self.default
        return self.default

    def _save_value(self):
        # 先讀取整個檔案
        data = {}
        if os.path.exists(self.shared_file):
            try:
                with open(self.shared_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except:
                data = {}

        # 更新對應欄位
        data[self.key] = self.get()

        try:
            with open(self.shared_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning("儲存失敗", f"無法儲存輸入值：{str(e)}")
