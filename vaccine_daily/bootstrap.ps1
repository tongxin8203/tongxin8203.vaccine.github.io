# bootstrap.ps1
# 一键从 GitHub 下载所有文件并完成安装
#
# 使用方法：在 PowerShell 中粘贴运行以下命令（不需要提前下载任何文件）：
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   irm https://raw.githubusercontent.com/tongxin8203/tongxin8203.vaccine.github.io/claude/vaccine-daily-pdf-C1b52/vaccine_daily/bootstrap.ps1 | iex

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       疫苗日报  Bootstrap 安装程序           ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── 配置
$RepoBase  = "https://raw.githubusercontent.com/tongxin8203/tongxin8203.vaccine.github.io/claude/vaccine-daily-pdf-C1b52/vaccine_daily"
$InstallDir = "C:\Users\$env:USERNAME\Documents\VaccineDaily\app"
$OutputDir  = "C:\Users\$env:USERNAME\Documents\VaccineDaily"

$Files = @(
    "vaccine_daily.py",
    "requirements.txt",
    "setup_scheduler.ps1",
    "run_now.bat"
)

# ── 1. 创建安装目录
Write-Host "[1/5] 创建安装目录..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
New-Item -ItemType Directory -Path $OutputDir  -Force | Out-Null
Write-Host "      $InstallDir" -ForegroundColor Gray

# ── 2. 下载文件
Write-Host "[2/5] 从 GitHub 下载程序文件..." -ForegroundColor Yellow
foreach ($file in $Files) {
    $url  = "$RepoBase/$file"
    $dest = Join-Path $InstallDir $file
    Write-Host "      下载: $file" -ForegroundColor Gray
    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
    } catch {
        Write-Host "      [✗] 下载失败: $file - $_" -ForegroundColor Red
        Write-Host "      请检查网络或手动从 GitHub 下载文件" -ForegroundColor Yellow
        exit 1
    }
}
Write-Host "      [✓] 文件下载完成" -ForegroundColor Green

# ── 3. 检测 / 安装 Python
Write-Host "[3/5] 检测 Python 环境..." -ForegroundColor Yellow
$PythonCmd = $null
foreach ($cmd in @("python", "python3", "py")) {
    try {
        $ver = & $cmd --version 2>&1
        if ($LASTEXITCODE -eq 0) { $PythonCmd = $cmd; break }
    } catch {}
}

if (-not $PythonCmd) {
    Write-Host "      未找到 Python，尝试通过 winget 安装..." -ForegroundColor Yellow
    try {
        winget install -e --id Python.Python.3.12 --silent `
            --accept-package-agreements --accept-source-agreements
        $PythonCmd = "python"
        Write-Host "      [✓] Python 安装成功" -ForegroundColor Green
    } catch {
        Write-Host ""
        Write-Host "  请手动安装 Python 3.9+: https://www.python.org/downloads/" -ForegroundColor Red
        Write-Host "  安装时勾选 'Add Python to PATH'，然后重新运行此脚本" -ForegroundColor Red
        exit 1
    }
} else {
    $ver = & $PythonCmd --version 2>&1
    Write-Host "      [✓] $ver" -ForegroundColor Green
}

# ── 4. 安装 Python 依赖
Write-Host "[4/5] 安装 Python 依赖库..." -ForegroundColor Yellow
$reqFile = Join-Path $InstallDir "requirements.txt"
& $PythonCmd -m pip install --upgrade pip -q
& $PythonCmd -m pip install -r $reqFile --no-warn-script-location
if ($LASTEXITCODE -ne 0) {
    Write-Host "      依赖安装失败，尝试使用清华镜像源..." -ForegroundColor Yellow
    & $PythonCmd -m pip install -r $reqFile -i https://pypi.tuna.tsinghua.edu.cn/simple
    if ($LASTEXITCODE -ne 0) {
        Write-Host "      [✗] 依赖安装失败，请检查网络连接" -ForegroundColor Red
        exit 1
    }
}
Write-Host "      [✓] 依赖安装完成" -ForegroundColor Green

# ── 5. 创建计划任务（每天 08:00）
Write-Host "[5/5] 创建 Windows 计划任务（每天 08:00）..." -ForegroundColor Yellow
$scriptPath = Join-Path $InstallDir "vaccine_daily.py"
$logFile    = Join-Path $OutputDir  "vaccine_daily.log"
$taskName   = "VaccineDailyReport"

$action   = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$PythonCmd`" `"$scriptPath`" >> `"$logFile`" 2>&1" `
    -WorkingDirectory $InstallDir

$trigger  = New-ScheduledTaskTrigger -Daily -At "08:00"
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable -RunOnlyIfNetworkAvailable `
    -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME -LogonType InteractiveToken -RunLevel Highest

$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false }

Register-ScheduledTask `
    -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null

Write-Host "      [✓] 计划任务已创建：每天 08:00 自动运行" -ForegroundColor Green

# ── 完成
Write-Host ""
Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║              安装完成！                      ║" -ForegroundColor Green
Write-Host "║                                              ║" -ForegroundColor Green
Write-Host "║  程序目录: $InstallDir" -ForegroundColor Green
Write-Host "║  PDF保存:  $OutputDir" -ForegroundColor Green
Write-Host "║  每天 08:00 自动运行并生成日报               ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""

# 询问是否立即测试运行
$ans = Read-Host "是否立即测试运行？（将抓取新闻并翻译，约需1-3分钟）[Y/N]"
if ($ans -match "^[Yy]") {
    Write-Host "正在运行，请稍候..." -ForegroundColor Yellow
    & $PythonCmd $scriptPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[✓] 运行成功！" -ForegroundColor Green
        Start-Process explorer $OutputDir
    } else {
        Write-Host "[!] 运行遇到问题，请查看日志: $logFile" -ForegroundColor Yellow
    }
}
