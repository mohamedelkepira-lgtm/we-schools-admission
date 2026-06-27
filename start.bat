@echo off
cd /d "%~dp0"
echo.
echo  Starting WE Schools Admission Portal...
echo.
start "WE API Server" cmd /c "node server.js"
timeout /t 2 /nobreak >nul
echo  Opening browser...
start http://localhost:5173
cmd /c "npm run dev"
