@echo off
setlocal enabledelayedexpansion

:: This bat runs from the app's "Set Up Now" button (not from the installer).
:: VC++ is handled by the NSIS installer itself (already elevated, no UAC).
:: This bat only handles Whisper Python setup.

set "RESOURCES_DIR=%~dp0"
if "%RESOURCES_DIR:~-1%"=="\" set "RESOURCES_DIR=%RESOURCES_DIR:~0,-1%"
set "PYTHON_EXE=%RESOURCES_DIR%\python\python.exe"
set "SETUP_BAT=%RESOURCES_DIR%\setup_whisper.bat"

echo.
echo  ============================================================
echo   AnchorCast - Setting Up Local AI Transcription
echo  ============================================================
echo.

:: Check if Whisper already works
if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" -c "from faster_whisper import WhisperModel" >nul 2>&1
  if !errorlevel! equ 0 (
    echo  [OK] Whisper already installed and working.
    goto :done
  )
)

:: Run whisper setup
if exist "%SETUP_BAT%" (
  set "ANCHORCAST_NONINTERACTIVE=1"
  call "%SETUP_BAT%"
) else (
  echo  [!!] setup_whisper.bat not found at: %SETUP_BAT%
)

:done
echo.
