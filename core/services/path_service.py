import os, sys, subprocess

class PathService:
    """跨平台開啟資料夾邏輯"""

    @staticmethod
    def open(path: str, select: str | None = None):
        if not os.path.exists(path):
            raise FileNotFoundError(f"路徑不存在：{path}")

        if sys.platform.startswith("win"):
            if select and os.path.exists(select):
                subprocess.run(["explorer", "/select,", os.path.normpath(select)])
            else:
                os.startfile(os.path.normpath(path))
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
