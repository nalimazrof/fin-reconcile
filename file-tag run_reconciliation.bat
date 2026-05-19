bat_content = """@echo off
:: 切换代码页为 UTF-8，防止中文乱码
chcp 65001 > nul
title 个人财务自动化对账系统

:: 切换到当前批处理文件所在的绝对路径目录
cd /d "%~dp0"

echo ============================================================
echo 🏦 正在启动个人财务自动化对账系统...
echo ============================================================
echo.

:: 使用 uv 免激活运行你的 main.py
uv run main.py

echo.
echo ============================================================
echo 任务执行完毕，按任意键关闭本窗口...
echo ============================================================
pause > nul
"""

file_path = "run_reconciliation.bat"
with open(file_path, "w", encoding="utf-8") as f:
    f.write(bat_content)

import urllib.parse
print(f"[file-tag: {file_path}]")