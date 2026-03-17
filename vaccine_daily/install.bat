@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo.
echo ╔══════════════════════════════════════════════╗
echo ║         疫苗日报  一键安装程序               ║
echo ║   Vaccine Daily PDF — Auto Installer         ║
echo ╚══════════════════════════════════════════════╝
echo.

set SCRIPT_DIR=%~dp0
set OUT_DIR=C:\Users\WALVAX\Documents\VaccineDaily
set LOG_FILE=%OUT_DIR%\install.log

:: ── 创建输出目录
if not exist "%OUT_DIR%" (
    mkdir "%OUT_DIR%"
    echo [✓] 已创建目录: %OUT_DIR%
) else (
    echo [✓] 输出目录已存在: %OUT_DIR%
)

:: ── 检测 Python
echo.
echo [1/4] 检测 Python...
set PYTHON_CMD=
for %%p in (python python3 py) do (
    if "!PYTHON_CMD!"=="" (
        %%p --version >nul 2>&1
        if not errorlevel 1 set PYTHON_CMD=%%p
    )
)

if "!PYTHON_CMD!"=="" (
    echo [!] 未检测到 Python，尝试自动安装...
    :: 尝试通过 winget 安装（Windows 10/11 自带 winget）
    winget --version >nul 2>&1
    if not errorlevel 1 (
        echo [...] 正在通过 winget 安装 Python 3.12...
        winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
        if not errorlevel 1 (
            set PYTHON_CMD=python
            echo [✓] Python 安装成功
            :: 刷新环境变量
            call refreshenv >nul 2>&1
        ) else (
            echo [✗] 自动安装失败
            goto :manual_python
        )
    ) else (
        :manual_python
        echo.
        echo ╔─────────────────────────────────────────────╗
        echo ║  请手动安装 Python 3.9+ 后重新运行此脚本   ║
        echo ║  下载地址: https://www.python.org/downloads ║
        echo ║  安装时请勾选 "Add Python to PATH"          ║
        echo ╚─────────────────────────────────────────────╝
        pause
        exit /b 1
    )
) else (
    for /f "tokens=*" %%v in ('!PYTHON_CMD! --version 2^>^&1') do echo [✓] 已找到 !PYTHON_CMD!: %%v
)

:: ── 升级 pip
echo.
echo [2/4] 升级 pip...
!PYTHON_CMD! -m pip install --upgrade pip -q
if errorlevel 1 (
    echo [!] pip 升级失败，继续安装依赖...
) else (
    echo [✓] pip 已是最新版本
)

:: ── 安装依赖
echo.
echo [3/4] 安装依赖库 (feedparser / requests / reportlab / deep-translator)...
!PYTHON_CMD! -m pip install -r "%SCRIPT_DIR%requirements.txt" --no-warn-script-location
if errorlevel 1 (
    echo.
    echo [✗] 依赖安装失败！
    echo     请检查网络连接，或尝试使用国内镜像源：
    echo     !PYTHON_CMD! -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    pause
    exit /b 1
)
echo [✓] 所有依赖安装完成

:: ── 测试运行
echo.
echo [4/4] 测试运行（正在抓取新闻并翻译，约需 1-3 分钟）...
echo       请耐心等待，翻译需要联网...
echo.
!PYTHON_CMD! "%SCRIPT_DIR%vaccine_daily.py"
if errorlevel 1 (
    echo.
    echo [!] 测试运行遇到问题，请查看上方错误信息
    echo     常见原因: 网络连接问题 / 翻译服务暂时不可用
    echo     可稍后手动运行: run_now.bat
) else (
    echo.
    echo [✓] 测试运行成功！
    echo     PDF 已保存至: %OUT_DIR%
    :: 自动打开输出目录
    explorer "%OUT_DIR%"
)

:: ── 创建计划任务
echo.
echo [...] 创建每日 08:00 自动运行计划任务...
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%setup_scheduler.ps1"
if errorlevel 1 (
    echo [!] 计划任务创建失败（可能需要管理员权限）
    echo     请右键本 bat 文件 → 以管理员身份运行
) else (
    echo [✓] 计划任务已创建，将在每天 08:00 自动运行
)

echo.
echo ╔══════════════════════════════════════════════╗
echo ║              安装完成！                      ║
echo ║  每天早上 08:00 将自动生成疫苗日报 PDF       ║
echo ║  保存路径: C:\Users\WALVAX\Documents\        ║
echo ║            VaccineDaily\                     ║
echo ║                                              ║
echo ║  手动生成: 双击 run_now.bat                  ║
echo ╚══════════════════════════════════════════════╝
echo.
pause
