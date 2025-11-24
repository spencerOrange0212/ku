import customtkinter as ctk
from customtkinter import CTkInputDialog
import json
import os
from tkinter import messagebox, simpledialog

class MemoryComboBox(ctk.CTkFrame):
    """可記憶、可新增/刪除的下拉式選單"""

    def __init__(self, master, file_path="tax_id_memory.json", width=220, height=40, **kwargs):
        super().__init__(master, width=width, height=height, **kwargs)

        self.file_path = file_path
        self.options = self._load_memory()

        self.var = ctk.StringVar(value=self.options[0] if self.options else "")
        self.combo = ctk.CTkOptionMenu(self, values=self.options or ["(尚無記錄)"], variable=self.var)
        self.combo.grid(row=0, column=0, padx=5)

        ctk.CTkButton(self, text="➕", width=30, command=self._add_option).grid(row=0, column=1, padx=2)
        ctk.CTkButton(self, text="🗑", width=30, command=self._delete_option).grid(row=0, column=2, padx=2)

    def get(self):
        return self.var.get().strip()

    def _load_memory(self):
        if os.path.exists(self.file_path):
            with open(self.file_path, "r", encoding="utf-8") as f:
                try:
                    return json.load(f)
                except:
                    return []
        return []

    def _save_memory(self):
        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(self.options, f, ensure_ascii=False, indent=2)

    def _add_option(self):
        dialog = CTkInputDialog(
            text="請輸入新的統一編號 (8碼數字)：",
            title="新增統一編號"
        )
        # 等待視窗生成後改字
        dialog.after(20, lambda: (
            dialog._ok_button.configure(text="確定"),
            dialog._cancel_button.configure(text="取消")
        ))
        new_value = dialog.get_input()

        if not new_value:
            return
        if not new_value.isdigit() or len(new_value) != 8:
            messagebox.showwarning("格式錯誤", "統一編號必須是 8 位數字")
            return
        if new_value in self.options:
            messagebox.showinfo("提示", "此統一編號已存在")
            return

        self.options.append(new_value)
        self.options.sort()
        self.combo.configure(values=self.options)
        self.var.set(new_value)
        self._save_memory()

    def _delete_option(self):
        selected = self.var.get()
        if selected not in self.options:
            messagebox.showinfo("提示", "沒有可刪除的項目")
            return
        self.options.remove(selected)
        self.combo.configure(values=self.options or ["(尚無記錄)"])
        self.var.set(self.options[0] if self.options else "")
        self._save_memory()
