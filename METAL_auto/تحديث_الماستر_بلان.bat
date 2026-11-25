@echo off
chcp 65001 > nul
title تحديث الماستر بلان - Metals Dashboard
cd /d "%~dp0"
echo ====================================================================
echo   ربط شيت الماستر بلان بالويب (LV / MV لكل أمر شغل)
echo   الشيت لازم يكون جوه مجلد Master_Plan
echo ====================================================================
python update_wo_voltage.py %*
echo.
echo افتح index.html واعمل Refresh. (ولو الويب على GitHub ارفع wo_voltage.js)
pause > nul
