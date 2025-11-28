import customtkinter as ctk
from customtkinter import CTkInputDialog
import json
import os
from tkinter import messagebox, filedialog
from copy import deepcopy

# 嘗試引入路徑工具
try:
    from core.utils.path_utils import resource_path
except ImportError:
    def resource_path(p):
        return p

# ====================================================================
# 1. 廠商設定資料結構
# ====================================================================
DEFAULT_CONFIG_MODULES = [
    "分類帳",
    "財產目錄",
    "資產負債表",
    "綜合損益表",
    "綜合損益表期別"
]

DEFAULT_CONFIG_TEMPLATE = {
    "module_options": {name: False for name in DEFAULT_CONFIG_MODULES},
    "input_folder": "",
    # "output_folder": "",  <-- 已完全移除
    "vendor_name": "",
    "note": ""  # 備註欄位
}


# ====================================================================
# 2. 整合型管理視窗
# ====================================================================

class VendorManagerWindow(ctk.CTkToplevel):
    def __init__(self, master_app):
        super().__init__(master_app)
        self.master_app = master_app
        self.title("廠商資料管理")
        self.geometry("900x700")

        # 延遲載入 Icon
        self.after(250, lambda: self._set_icon())

        self.attributes("-topmost", True)

        # 佈局：左 (1/3) 右 (2/3)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1)

        self.current_mode = "view"
        self.editing_id = None

        # --- 左側面板 (清單) ---
        self.left_panel = ctk.CTkFrame(self, corner_radius=0)
        self.left_panel.grid(row=0, column=0, sticky="nsew")
        self.left_panel.grid_rowconfigure(1, weight=1)
        self.left_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.left_panel, text="廠商列表", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0,
                                                                                                      padx=10, pady=10,
                                                                                                      sticky="w")

        self.scroll_list = ctk.CTkScrollableFrame(self.left_panel)
        self.scroll_list.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        self.btn_add_new = ctk.CTkButton(
            self.left_panel,
            text="➕ 新增廠商",
            fg_color="#27AE60", hover_color="#2ECC71",
            height=40,
            command=self._mode_add_new
        )
        self.btn_add_new.grid(row=2, column=0, padx=10, pady=15, sticky="ew")

        # --- 右側面板 (編輯表單) ---
        self.right_panel = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.right_panel.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.right_panel.grid_columnconfigure(1, weight=1)

        self._create_form_widgets()

        self.radio_var = ctk.StringVar(value="")
        self._refresh_list()

        # 預設選取邏輯
        if self.master_app.configs:
            first_id = sorted(self.master_app.configs.keys())[0]
            self.radio_var.set(first_id)
            self._mode_view_edit(first_id)
        else:
            self._mode_add_new()

    def _set_icon(self):
        try:
            self.iconbitmap(resource_path("ai.ico"))
        except:
            pass

    def _create_form_widgets(self):
        """建立右側表單元件"""
        # 標題
        self.lbl_form_title = ctk.CTkLabel(self.right_panel, text="詳細設定", font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_form_title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 20))

        r = 1
        # 代號
        ctk.CTkLabel(self.right_panel, text="廠商代號 (ID)", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0,
                                                                                                   sticky="w", pady=5)
        self.entry_id = ctk.CTkEntry(self.right_panel, placeholder_text="請輸入唯一代號")
        self.entry_id.grid(row=r, column=1, sticky="ew", padx=10, pady=5)
        r += 1

        # 名稱
        ctk.CTkLabel(self.right_panel, text="廠商名稱", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0,
                                                                                              sticky="w", pady=5)
        self.entry_name = ctk.CTkEntry(self.right_panel, placeholder_text="廠商全名")
        self.entry_name.grid(row=r, column=1, sticky="ew", padx=10, pady=5)
        r += 1

        # 模組
        ctk.CTkLabel(self.right_panel, text="執行模組", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0,
                                                                                              sticky="nw", pady=5)
        self.frame_modules = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.frame_modules.grid(row=r, column=1, sticky="ew", padx=10, pady=5)

        self.chk_modules = {}
        for i, mod in enumerate(DEFAULT_CONFIG_MODULES):
            cb = ctk.CTkCheckBox(self.frame_modules, text=mod)
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=10, pady=5)
            self.chk_modules[mod] = cb
        r += 1

        # 路徑 (只保留輸入)
        ctk.CTkLabel(self.right_panel, text="輸入資料夾", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0,
                                                                                                sticky="w", pady=5)
        self.entry_in = ctk.CTkEntry(self.right_panel)
        self.entry_in.grid(row=r, column=1, sticky="ew", padx=(10, 0), pady=5)
        ctk.CTkButton(self.right_panel, text="...", width=30, command=lambda: self._browse(self.entry_in)).grid(row=r,
                                                                                                                column=2,
                                                                                                                padx=5)
        r += 1

        # 按鈕區 (刪除、儲存)
        self.frame_btns = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.frame_btns.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(20, 10))

        self.btn_delete = ctk.CTkButton(self.frame_btns, text="🗑 刪除此廠商", fg_color="#C0392B", hover_color="#E74C3C",
                                        command=self._delete)
        self.btn_delete.pack(side="left")

        self.btn_save = ctk.CTkButton(self.frame_btns, text="💾 儲存設定", fg_color="#2980B9", hover_color="#3498DB",
                                      command=self._save)
        self.btn_save.pack(side="right")
        r += 1

        # 備註區
        ctk.CTkLabel(self.right_panel, text="備註", font=ctk.CTkFont(weight="bold")).grid(row=r, column=0, sticky="w",
                                                                                          pady=(10, 0))
        r += 1

        self.textbox_note = ctk.CTkTextbox(self.right_panel)
        self.textbox_note.grid(row=r, column=0, columnspan=3, sticky="nsew", padx=0, pady=5)
        self.right_panel.grid_rowconfigure(r, weight=1)

    def _refresh_list(self):
        for child in self.scroll_list.winfo_children():
            child.destroy()

        configs = self.master_app.configs
        if not configs:
            ctk.CTkLabel(self.scroll_list, text="(無資料)").pack(pady=10)
            return

        for vid in sorted(configs.keys()):
            name = configs[vid].get("vendor_name", "")
            display = f"[{vid}] {name}"
            btn = ctk.CTkRadioButton(
                self.scroll_list,
                text=display,
                variable=self.radio_var,
                value=vid,
                command=lambda v=vid: self._mode_view_edit(v)
            )
            btn.pack(anchor="w", padx=5, pady=5)

    def _browse(self, entry):
        p = filedialog.askdirectory()
        if p:
            entry.delete(0, "end")
            entry.insert(0, p)

    # --- 模式切換邏輯 ---

    def _mode_add_new(self):
        self.current_mode = "new"
        self.editing_id = None
        self.lbl_form_title.configure(text="新增廠商 (請填寫資料)")

        # 清空
        self.entry_id.configure(state="normal")
        self.entry_id.delete(0, "end")
        self.entry_name.delete(0, "end")
        self.entry_in.delete(0, "end")
        self.textbox_note.delete("0.0", "end")

        for cb in self.chk_modules.values():
            cb.deselect()

        self.btn_delete.pack_forget()
        self.radio_var.set("")

    def _mode_view_edit(self, vid):
        if vid not in self.master_app.configs: return

        self.current_mode = "edit"
        self.editing_id = vid
        self.radio_var.set(vid)
        self.lbl_form_title.configure(text=f"編輯設定：{vid}")

        cfg = self.master_app.configs[vid]

        # 載入
        self.entry_id.delete(0, "end")
        self.entry_id.insert(0, vid)
        self.entry_id.configure(state="disabled")

        self.entry_name.delete(0, "end")
        self.entry_name.insert(0, cfg.get("vendor_name", ""))
        self.entry_in.delete(0, "end")
        self.entry_in.insert(0, cfg.get("input_folder", ""))

        self.textbox_note.delete("0.0", "end")
        self.textbox_note.insert("0.0", cfg.get("note", ""))

        # 載入 Checkbox
        opts = cfg.get("module_options", {})
        for name, cb in self.chk_modules.items():
            val = opts.get(name, False)
            # 處理 BooleanVar
            if hasattr(val, 'get'): val = val.get()
            if val:
                cb.select()
            else:
                cb.deselect()

        self.btn_delete.pack(side="left")

    # --- 儲存與刪除 ---

    def _save(self):
        vid = self.entry_id.get().strip()
        if not vid:
            messagebox.showwarning("錯誤", "廠商代號不能為空")
            return

        new_data = deepcopy(DEFAULT_CONFIG_TEMPLATE)
        new_data["vendor_name"] = self.entry_name.get()
        new_data["input_folder"] = self.entry_in.get()
        new_data["note"] = self.textbox_note.get("0.0", "end").strip()

        mod_opts = {}
        for name, cb in self.chk_modules.items():
            mod_opts[name] = ctk.BooleanVar(value=bool(cb.get()))
        new_data["module_options"] = mod_opts

        configs = self.master_app.configs

        if self.current_mode == "new":
            if vid in configs:
                messagebox.showerror("錯誤", "代號已存在")
                return
            configs[vid] = new_data
            self.master_app._save_configs()  # 存檔
            self._refresh_list()
            self._mode_view_edit(vid)
            self.master_app._refresh_combo()
            self.master_app.current_vendor_id.set(vid)
            messagebox.showinfo("成功", "新增成功！")

        else:  # Edit mode
            configs[self.editing_id] = new_data
            self.master_app._save_configs()  # 存檔
            self._refresh_list()
            self.radio_var.set(self.editing_id)
            messagebox.showinfo("成功", "儲存成功！")

    def _delete(self):
        if not self.editing_id: return
        if messagebox.askyesno("確認", f"確定刪除 {self.editing_id} 嗎？"):
            del self.master_app.configs[self.editing_id]
            self.master_app._save_configs()  # 存檔
            self.master_app._refresh_combo()
            self._refresh_list()
            self._mode_add_new()


# ====================================================================
# 1. 第一層：主畫面組件
# ====================================================================

class VendorConfigManager(ctk.CTkFrame):
    def __init__(self, master, file_path="vendor_configs.json", width=250, height=40, **kwargs):
        super().__init__(master, width=width, height=height, **kwargs)

        self.file_path = file_path
        self.configs = self._load_configs()
        self.current_vendor_id = ctk.StringVar(value="")
        self.win_manager = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)

        self.combo = ctk.CTkOptionMenu(
            self,
            variable=self.current_vendor_id,
            values=[],
            command=self._on_combo_change
        )
        self.combo.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="ew")

        self.btn_settings = ctk.CTkButton(
            self, text="⚙️", width=40,
            fg_color="#555555", hover_color="#777777",
            command=self._open_manager_window
        )
        self.btn_settings.grid(row=0, column=1, padx=(0, 5), pady=5)

        self._refresh_combo()

    def _open_manager_window(self):
        if self.win_manager and self.win_manager.winfo_exists():
            self.win_manager.lift()
        else:
            self.win_manager = VendorManagerWindow(self)

    def _refresh_combo(self):
        ids = sorted(self.configs.keys())
        if not ids:
            self.combo.configure(values=["(尚無資料)"])
            self.current_vendor_id.set("(尚無資料)")
        else:
            self.combo.configure(values=ids)
            if self.current_vendor_id.get() not in ids:
                self.current_vendor_id.set(ids[0])

    def _on_combo_change(self, value):
        self.current_vendor_id.set(value)

    def get_current_id(self):
        val = self.current_vendor_id.get()
        return val if val != "(尚無資料)" else None

    # --- 資料存取 ---

    def _load_configs(self):
        if not os.path.exists(self.file_path):
            return {}
        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                new_data = {}
                for vid in data:
                    new_data[str(vid)] = deepcopy(DEFAULT_CONFIG_TEMPLATE)
                data = new_data
                self._save_configs_direct(data)
            for cfg in data.values():
                if "module_options" not in cfg:
                    cfg["module_options"] = {k: False for k in DEFAULT_CONFIG_MODULES}
                for mod in DEFAULT_CONFIG_MODULES:
                    if mod not in cfg["module_options"]:
                        cfg["module_options"][mod] = False
                cfg["module_options"] = {
                    k: ctk.BooleanVar(master=self, value=v)
                    for k, v in cfg["module_options"].items()
                }
            return data
        except Exception as e:
            print(f"Config load error: {e}")
            return {}

    # ⭐ 關鍵修正：手動建構字典，解決 deepcopy Tkinter 錯誤
    def _save_configs(self):
        serializable = {}
        for vid, cfg in self.configs.items():
            sc = {
                "vendor_name": cfg.get("vendor_name", ""),
                "input_folder": cfg.get("input_folder", ""),
                "note": cfg.get("note", ""),
                "module_options": {
                    k: v.get() if hasattr(v, 'get') else v
                    for k, v in cfg.get("module_options", {}).items()
                }
            }
            serializable[vid] = sc
        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(serializable, f, ensure_ascii=False, indent=2)

    def _save_configs_direct(self, data):
        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def get_current_settings(self):
        vid = self.get_current_id()
        if not vid or vid not in self.configs:
            return None, None

        # 存檔一次確保一致
        self._save_configs()

        src_cfg = self.configs[vid]
        final_cfg = {
            "vendor_name": src_cfg.get("vendor_name", ""),
            "input_folder": src_cfg.get("input_folder", ""),
            "note": src_cfg.get("note", "")
        }
        opts = src_cfg.get("module_options", {})
        enabled = [k for k, v in opts.items() if (v.get() if hasattr(v, 'get') else v)]
        final_cfg["enabled_modules"] = enabled
        return vid, final_cfg