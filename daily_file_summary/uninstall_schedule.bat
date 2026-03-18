@echo off
chcp 65001 >nul
echo 正在删除定时任务 DailyFileSummary ...
schtasks /delete /tn "DailyFileSummary" /f
echo 已删除。
pause
