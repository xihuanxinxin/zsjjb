import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("gui.py", base=base, target_name="资产管理系统.exe")]

setup(
    name="资产管理系统",
    version="1.0",
    description="数据库资产管理系统",
    executables=executables,
    options={
        "build_exe": {
            "packages": ["tkinter", "pandas", "datetime"],
            "include_files": []
        }
    }
)
