import threading
from tkinter import messagebox


def do_actions_sequential(app, tasks):
    """
    多個工具依序執行（包含只勾一個時的情況）：
    - tasks: List[(action_name, display_name)]
    - 依序執行多個工具
    - 每完成一個工具更新 ProgressBar
    - 任何錯誤或取消會停止後續工具
    """
    if not tasks:
        return

    # 重置取消旗標
    app.cancel_requested = False
    app.status_label.configure(text="狀態：準備開始執行選取的工具...")

    # 啟用停止按鈕
    if getattr(app, "stop_button", None):
        app.stop_button.configure(state="normal")

    total = len(tasks)

    def worker():
        cancelled = False

        try:
            for index, (action_name, display_name) in enumerate(tasks, start=1):

                # 🔴 若使用者按了取消：立即停止
                if getattr(app, "cancel_requested", False):
                    cancelled = True
                    break

                # 更新目前執行中的工具
                app.after(0, lambda name=display_name:
                          app.status_label.configure(text=f"狀態：正在執行「{name}」中..."))

                # --- 執行工具 ---
                if action_name == "insert_report":
                    msg = app.controller.run_insert_report(app.controller.file_path)
                elif action_name == "update_subjects":
                    msg = app.controller.run_update_subjects(app.controller.file_path)
                elif action_name == "delete_details":
                    msg = app.controller.run_delete_details(app.controller.file_path)
                else:
                    msg = f"未知的動作：{action_name}"

                # --- 單一工具完成 ---
                app.after(0, lambda m=msg, name=display_name: [
                    getattr(app, "append_log")(f"✅ 「{name}」執行訊息: {m}"),
                    getattr(app, "append_log")(f"------------- {name} 模組完成 -------------\n"),
                    app.status_label.configure(text=f"狀態：已完成「{name}」")
                ])



            # ---- 收尾 ----
            if cancelled:
                # 使用者中止
                app.after(0, lambda: [

                    app.status_label.configure(text="狀態：已中止執行"),
                    getattr(app, "append_log")("⛔ 任務已被使用者中止，後續工具未執行。")
                ])

            else:
                # 全部工具成功完成
                app.after(0, lambda: [

                    app.status_label.configure(text="狀態：所有選取的工具已依序執行完成"),
                    messagebox.showinfo("完成", "所有選取的工具已全部執行完畢 🟢")
                ])

        except Exception as e:
            # 任一工具發生錯誤
            app.after(0, lambda err=e: [
                app.status_label.configure(text="狀態：發生錯誤，後續工具已停止"),
                messagebox.showerror("錯誤", f"執行過程發生錯誤，已停止後續工具。\n\n{err}")
            ])

        finally:
            # 關閉停止按鈕
            if getattr(app, "stop_button", None):
                app.after(0, lambda: app.stop_button.configure(state="disabled"))

    threading.Thread(target=worker, daemon=True).start()
