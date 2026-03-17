@echo off
chcp 65001 >nul
echo 正在生成今日疫苗日报...
python "%~dp0vaccine_daily.py"
echo.
echo 完成！文件保存至: C:\Users\WALVAX\Documents\VaccineDaily
pause
