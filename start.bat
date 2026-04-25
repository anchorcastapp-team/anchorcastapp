@echo off
echo.
echo   ✝  SermonCast — AI Sermon Display
echo   ---------------------------------
echo.

where node >nul 2>&1
if %errorlevel% neq 0 (
  echo   Node.js not found.
  echo   Please install Node.js 18+ from https://nodejs.org
  echo   Then run this file again.
  pause
  exit /b 1
)

echo   Node.js detected: 
node -v

echo.

if not exist node_modules (
  echo   Installing dependencies ^(first run only^)...
  npm install --silent
  echo   Done.
)

echo.
echo   Launching SermonCast...
echo.
npm start
