@echo off
chcp 65001 >nul
echo ============================================
echo   每日文件摘要 - 定时任务安装
echo ============================================
echo.

:: 获取脚本所在目录
set SCRIPT_DIR=%~dp0

:: 先安装依赖
echo [1/2] 安装 Python 依赖...
pip install -r "%SCRIPT_DIR%requirements.txt"
if %errorlevel% neq 0 (
    echo 依赖安装失败，请确保已安装 Python 和 pip。
    pause
    exit /b 1
)

:: 创建定时任务（每天晚上9点执行）
echo.
echo [2/2] 创建 Windows 定时任务...
schtasks /create /tn "DailyFileSummary" /tr "pythonw \"%SCRIPT_DIR%summarize.py\"" /sc daily /st 21:00 /f
if %errorlevel% neq 0 (
    echo 定时任务创建失败，请以管理员身份运行此脚本。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   安装完成!
echo   定时任务: 每天 21:00 自动运行
echo   任务名称: DailyFileSummary
echo   配置文件: %SCRIPT_DIR%config.json
echo   输出目录: D:\wechat-summary
echo.
echo   如需修改扫描路径，请编辑 config.json
echo   如需卸载定时任务:
echo     schtasks /delete /tn "DailyFileSummary" /f
echo ============================================
pause
