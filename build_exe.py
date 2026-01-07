"""
Script để build file .exe
Chạy: python build_exe.py
"""
import subprocess
import sys

# Cài PyInstaller nếu chưa có
subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])

# Build exe
subprocess.run([
    "pyinstaller",
    "--onefile",
    "--add-data", "templates;templates",
    "--add-data", "nrl.xlsx;.",
    "--name", "NRL_Lookup",
    "--icon", "NONE",
    "app.py"
])

print("\n✅ Đã tạo file exe tại: dist/NRL_Lookup.exe")
