@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 表格合并清洗器

echo 正在检查依赖（第一次运行需要一两分钟）…
python -m pip install -r requirements.txt -q

echo 启动界面…
python main.py
if errorlevel 1 pause
