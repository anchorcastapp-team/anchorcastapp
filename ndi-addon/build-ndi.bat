@echo off
echo.
echo   ✝  SermonCast NDI Addon Builder
echo   ---------------------------------
echo.
echo   This builds the native NDI sender using the official NDI 6 SDK.
echo   Requirements:
echo     - NDI 6 SDK installed at: C:\Program Files\NDI\NDI 6 SDK\
echo     - Visual Studio Build Tools with C++ workload
echo     - node-gyp installed globally: npm install -g node-gyp
echo.

cd /d "%~dp0"

echo   Checking for NDI SDK...
if not exist "C:\Program Files\NDI\NDI 6 SDK\Include\Processing.NDI.Lib.h" (
    echo.
    echo   ERROR: NDI 6 SDK not found at C:\Program Files\NDI\NDI 6 SDK\
    echo   Download it free from: https://ndi.video/for-developers/ndi-sdk/
    echo.
    if not defined ANCHORCAST_NONINTERACTIVE pause
    exit /b 1
)
echo   NDI SDK found.

echo   Installing node-addon-api...
npm install --ignore-scripts
if %errorlevel% neq 0 (
    echo   ERROR: npm install failed
    if not defined ANCHORCAST_NONINTERACTIVE pause
    exit /b 1
)

echo   Building native addon...
node-gyp rebuild
if %errorlevel% neq 0 (
    echo.
    echo   Build failed. Common fixes:
    echo     1. Make sure Visual Studio Build Tools are installed with C++ workload
    echo     2. Run: npm install -g node-gyp
    echo     3. Check NDI SDK path in binding.gyp
    if not defined ANCHORCAST_NONINTERACTIVE pause
    exit /b 1
)

echo.
echo   ✓ NDI addon built successfully!
echo   File: build\Release\ndi_sender.node
echo.
echo   Restart SermonCast — NDI output will now use the official NDI SDK.
echo   In OBS/vMix: look for "SermonCast" in NDI sources.
echo.
if not defined ANCHORCAST_NONINTERACTIVE pause
