@echo off
chcp 65001 >nul
echo ======================================================================
echo   CellPPT Studio Pro v2.0 — 科研 PPT 自动生成与 PowerPoint 实时逐元素输出
echo ======================================================================
"%~dp0.venv\Scripts\python.exe" "%~dp0cellppt_studio.py" demo --record yes --mode sequential --step-ms 25 --keep-open
pause
