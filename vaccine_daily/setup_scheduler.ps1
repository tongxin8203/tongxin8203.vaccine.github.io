# setup_scheduler.ps1
# 一键创建 Windows 计划任务：每天早上 8:00 自动生成疫苗日报 PDF
# 使用方法（以管理员身份运行 PowerShell）:
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   .\setup_scheduler.ps1

# ──────────────────────────────────────────────
# 配置区 - 根据实际情况修改
# ──────────────────────────────────────────────

# Python 解释器路径（示例：Conda 环境 / 系统 Python）
$PythonExe = "C:\Users\WALVAX\AppData\Local\Programs\Python\Python311\python.exe"
# 如果使用 Conda，改为：
# $PythonExe = "C:\Users\WALVAX\anaconda3\envs\vaccine\python.exe"

# 脚本所在目录（请修改为实际路径）
$ScriptDir  = "$PSScriptRoot"
$ScriptPath = Join-Path $ScriptDir "vaccine_daily.py"

# 日志文件
$LogFile    = "C:\Users\WALVAX\Documents\VaccineDaily\vaccine_daily.log"

# 任务名称
$TaskName   = "VaccineDailyReport"
$TaskDesc   = "每天早上8:00自动收集疫苗资讯并生成PDF日报"

# ──────────────────────────────────────────────
# 创建计划任务
# ──────────────────────────────────────────────

Write-Host "=== 疫苗日报计划任务安装程序 ===" -ForegroundColor Cyan
Write-Host ""

# 检查 Python 是否存在
if (-Not (Test-Path $PythonExe)) {
    Write-Host "❌ 未找到 Python: $PythonExe" -ForegroundColor Red
    Write-Host "   请修改脚本中的 `$PythonExe 变量为正确的 Python 路径。" -ForegroundColor Yellow
    Write-Host "   当前系统 Python 路径可通过以下命令查找:" -ForegroundColor Yellow
    Write-Host "     where python" -ForegroundColor Green
    exit 1
}

Write-Host "✓ Python 路径: $PythonExe" -ForegroundColor Green
Write-Host "✓ 脚本路径:   $ScriptPath" -ForegroundColor Green

# 确保输出目录存在
$OutputDir = "C:\Users\WALVAX\Documents\VaccineDaily"
if (-Not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "✓ 已创建输出目录: $OutputDir" -ForegroundColor Green
}

# 删除已有同名任务（如果存在）
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "✓ 已删除旧任务: $TaskName" -ForegroundColor Yellow
}

# 构建任务动作
# 使用 cmd /c 可在后台运行并将日志重定向
$Arguments = "/c `"$PythonExe`" `"$ScriptPath`" >> `"$LogFile`" 2>&1"
$Action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument $Arguments `
    -WorkingDirectory $ScriptDir

# 触发器：每天 08:00
$Trigger = New-ScheduledTaskTrigger -Daily -At "08:00"

# 设置：如果计算机在计划时间处于休眠状态，则在唤醒时立即运行
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -WakeToRun `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew

# 以当前登录用户身份运行
$Principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType InteractiveToken `
    -RunLevel Highest

# 注册任务
Register-ScheduledTask `
    -TaskName $TaskName `
    -Description $TaskDesc `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Principal $Principal `
    -Force | Out-Null

Write-Host ""
Write-Host "============================" -ForegroundColor Cyan
Write-Host "✅ 计划任务创建成功！" -ForegroundColor Green
Write-Host "   任务名称: $TaskName" -ForegroundColor White
Write-Host "   执行时间: 每天 08:00" -ForegroundColor White
Write-Host "   PDF保存至: $OutputDir" -ForegroundColor White
Write-Host "   运行日志: $LogFile" -ForegroundColor White
Write-Host "============================" -ForegroundColor Cyan
Write-Host ""
Write-Host "💡 提示：" -ForegroundColor Yellow
Write-Host "   • 立即测试运行：Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
Write-Host "   • 查看任务状态：Get-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
Write-Host "   • 删除任务：    Unregister-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
Write-Host "   • 在任务计划程序中查看：taskschd.msc" -ForegroundColor Gray
