@echo off
chcp 65001 >nul
echo ============================================
echo  疫苗日报 - 一键安装脚本
echo ============================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+
    echo   下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [✓] Python 已就绪

:: 安装依赖
echo.
echo [...] 正在安装依赖库...
python -m pip install --upgrade pip -q
python -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络连接
    pause
    exit /b 1
)
echo [✓] 依赖安装完成

:: 创建输出目录
if not exist "C:\Users\WALVAX\Documents\VaccineDaily" (
    mkdir "C:\Users\WALVAX\Documents\VaccineDaily"
    echo [✓] 已创建输出目录
)

:: 运行一次测试
echo.
echo [...] 正在执行测试运行...
python "%~dp0vaccine_daily.py"
if errorlevel 1 (
    echo [警告] 测试运行遇到问题，请查看上方错误信息
) else (
    echo [✓] 测试运行成功！
    echo     PDF 文件已保存至: C:\Users\WALVAX\Documents\VaccineDaily
)

:: 设置计划任务
echo.
echo [...] 正在创建每日 08:00 计划任务...
powershell -ExecutionPolicy Bypass -File "%~dp0setup_scheduler.ps1"
if errorlevel 1 (
    echo [警告] 计划任务创建失败（可能需要管理员权限）
    echo   请右键点击 setup_scheduler.ps1，选择"以管理员身份运行 PowerShell"
) else (
    echo [✓] 计划任务创建成功！
)

echo.
echo ============================================
echo  安装完成！系统将在每天 08:00 自动生成日报
echo  保存路径: C:\Users\WALVAX\Documents\VaccineDaily
echo ============================================
pause
