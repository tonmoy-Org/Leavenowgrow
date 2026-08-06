@echo off
echo ===================================================
echo   Leavenowgrow PM2 & Auto-Restart Setup Script
echo ===================================================
echo.

echo Installing PM2 and PM2 Windows Startup globally...
call npm install -g pm2 pm2-windows-startup

echo.
echo Configuring PM2 Windows Startup Service...
call pm2-startup install

echo.
echo Starting application with PM2 ecosystem config...
cd /d %~dp0
call pm2 start ecosystem.config.js

echo.
echo Saving PM2 process list for auto-boot on server restart...
call pm2 save

echo.
echo ===================================================
echo   SUCCESS! PM2 setup complete.
echo   Your project will now auto-restart on crash
echo   and auto-boot when Windows Server restarts.
echo ===================================================
pause
