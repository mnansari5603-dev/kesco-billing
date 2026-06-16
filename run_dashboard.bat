@echo off
:: 1. Aapki file jahan hai wahan se khud automatic copy karega
copy "C:\Users\Admin\Desktop\BILLING REPORTS\16-06-2026\FINAL SUMMARY.xlsx" "C:\Users\Admin\Downloads\PAYTHON AUTOMATION\FINAL SUMMARY.xlsx" /Y

:: 2. Aapke folder me jaakar data ko GitHub par automatic upload (Push) karega
cd /d "C:\Users\Admin\Downloads\PAYTHON AUTOMATION"
git add .
git commit -m "Auto data update from Desktop"
git push origin main --force

echo ===================================================
echo  ⚡ DATA IS LIVE! DASHBOARD AUTOMATIC UPDATE HO GAYA ⚡
echo ===================================================
pause