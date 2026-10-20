@echo off
chcp 65001 >nul
cd /d "%~dp0"
title MigrantInfo Assist

where node >nul 2>nul
if not errorlevel 1 goto hasnode
echo.
echo Node.js topilmadi. Avval uni o'rnating: https://nodejs.org  (LTS versiya)
echo O'rnatgach, shu faylni qayta ishga tushiring.
start https://nodejs.org
pause
exit /b 1

:hasnode
if exist node_modules goto hasmods
echo.
echo Birinchi marta ishga tushirish: kutubxonalar o'rnatilmoqda (1-3 daqiqa)...
call npm install --no-audit --no-fund
if errorlevel 1 goto fail

:hasmods
if exist .env goto hasenv
echo.
set /p KEY=Gemini API kalitini kiriting va Enter bosing: 
> .env echo GEMINI_API_KEY=%KEY%

:hasenv
if exist dist\server.cjs goto hasbuild
echo.
echo Dastur yig'ilmoqda (bir marta)...
call npm run build
if errorlevel 1 goto fail

:hasbuild
set NODE_ENV=production
echo.
echo Dastur ishga tushmoqda. Brauzer ochiladi: http://localhost:3000
echo To'xtatish uchun shu oynani yoping.
start "" cmd /c "ping -n 4 127.0.0.1 >nul & start http://localhost:3000"
node dist\server.cjs
pause
exit /b 0

:fail
echo.
echo Xatolik yuz berdi. Yuqoridagi matnni tekshiring.
pause
exit /b 1
