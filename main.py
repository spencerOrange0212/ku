from gui.main_app import ExcelToolApp
import customtkinter as ctk
import os, sys

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")


    def resource_path(relative_path):
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)


    ctk.set_default_color_theme(resource_path("theme/orange.json"))
    app = ExcelToolApp()
    app.mainloop()
