@echo off
setlocal enabledelayedexpansion
echo.
echo  ================================================
echo   AnchorCast -- Whisper Local Transcription Setup
echo  ================================================
echo.

set "ANCHORCAST_DIR=%~dp0"
set "PYTHON_DIR=%ANCHORCAST_DIR%python"
set "PYTHON_EXE=%PYTHON_DIR%\python.exe"
set "PYTHON_ZIP=%ANCHORCAST_DIR%python312.zip"
set "GETPIP=%ANCHORCAST_DIR%get-pip.py"
set "VCREDIST_EXE=%TEMP%\vc_redist_anchorcast.x64.exe"
set "VCREDIST_URL=https://aka.ms/vs/17/release/vc_redist.x64.exe"
set "PYTHON_URL=https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip"
set "GETPIP_URL=https://bootstrap.pypa.io/get-pip.py"
set "NEEDS_RESTART=0"

:: ── Step 1: System Python version check ─────────────────────────────────────
echo  Checking system Python...
for %%P in (python python3 py) do (
  %%P --version >nul 2>&1
  if !errorlevel! equ 0 (
    for /f "tokens=2" %%V in ('%%P --version 2^>^&1') do (
      echo  Found system Python %%V
      for /f "tokens=1,2 delims=." %%A in ("%%V") do (
        if "%%A"=="3" (
          if %%B GEQ 8 if %%B LEQ 12 (
            echo  [OK] Compatible ^(3.8-3.12^)
          ) else (
            if %%B GTR 12 (
              echo  [!!] Python %%V is too new. faster-whisper requires 3.8-3.12.
              echo       Python 3.13+ is not yet supported by the AI engine.
            ) else (
              echo  [!!] Python %%V is too old. Requires 3.8+.
            )
            echo       Portable Python 3.12 will be used instead.
          )
        )
      )
    )
    goto :system_check_done
  )
)
echo  No system Python found. Portable Python 3.12 will be installed.
:system_check_done
echo.

:: ── Step 2: Visual C++ Redistributable ──────────────────────────────────────
echo  Checking Visual C++ Redistributable...
set "VCREDIST_OK=0"
reg query "HKLM\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" /v Version >nul 2>&1
if !errorlevel! equ 0 set "VCREDIST_OK=1"
if "%VCREDIST_OK%"=="0" (
  reg query "HKLM\SOFTWARE\Classes\Installer\Dependencies\VC,redist.x64,amd64,14.32,bundle" >nul 2>&1
  if !errorlevel! equ 0 set "VCREDIST_OK=1"
)

if "%VCREDIST_OK%"=="1" (
  echo  [OK] Visual C++ Redistributable already installed.
) else (
  echo  Visual C++ Redistributable not found.
  echo  Downloading from Microsoft ^(~25 MB^)...
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Invoke-WebRequest -Uri '%VCREDIST_URL%' -OutFile '%VCREDIST_EXE%' -UseBasicParsing" >nul 2>&1
  if exist "%VCREDIST_EXE%" (
    echo  Installing Visual C++ Redistributable...
    "%VCREDIST_EXE%" /install /quiet /norestart
    del "%VCREDIST_EXE%" >nul 2>&1
    echo  [OK] Visual C++ Redistributable installed.
    echo.
    echo  IMPORTANT: A system restart is recommended for VC++ to fully
    echo  take effect. Setup will continue, but if faster-whisper fails
    echo  to load, please restart Windows and run this setup again.
    echo.
    set "NEEDS_RESTART=1"
  ) else (
    echo  [!!] Could not download VC++ Redistributable.
    echo       Please install manually: https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo       Then run this setup again.
  )
)
echo.

:: ── Step 3: Check if already fully working ──────────────────────────────────
if exist "%PYTHON_EXE%" (
  echo  Portable Python found. Checking faster-whisper...
  "%PYTHON_EXE%" -c "from faster_whisper import WhisperModel; print('[OK] faster-whisper working')" 2>nul
  if !errorlevel! equ 0 goto :download_model
  echo  faster-whisper not working -- reinstalling...
  goto :install_pip
)

:: ── Step 4: Download portable Python 3.12 ───────────────────────────────────
echo  [1/3] Downloading portable Python 3.12 (~11 MB)...
where powershell >nul 2>&1
if %errorlevel% neq 0 (
  echo  ERROR: PowerShell not found.
  goto :setup_failed
)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_ZIP%' -UseBasicParsing" >nul 2>&1
if not exist "%PYTHON_ZIP%" (
  echo  ERROR: Failed to download Python. Check internet connection.
  goto :setup_failed
)

:: ── Step 5: Extract Python ───────────────────────────────────────────────────
echo  [2/3] Extracting Python...
if not exist "%PYTHON_DIR%" mkdir "%PYTHON_DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Expand-Archive -Path '%PYTHON_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force" >nul 2>&1
del "%PYTHON_ZIP%" >nul 2>&1

:: Enable site-packages
for %%F in ("%PYTHON_DIR%\python3*._pth") do (
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "(Get-Content '%%F') -replace '#import site', 'import site' | Set-Content '%%F'" >nul 2>&1
)

:install_pip
:: ── Step 6: Install pip + faster-whisper ────────────────────────────────────
echo  [3/3] Installing pip + faster-whisper AI engine...
echo  Please wait -- this downloads ~500 MB of AI components.
echo  This is a one-time install and will take 3-5 minutes.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Invoke-WebRequest -Uri '%GETPIP_URL%' -OutFile '%GETPIP%' -UseBasicParsing" >nul 2>&1
if exist "%GETPIP%" (
  "%PYTHON_EXE%" "%GETPIP%" --quiet --no-warn-script-location
  del "%GETPIP%" >nul 2>&1
)
"%PYTHON_EXE%" -m pip install faster-whisper --quiet --no-warn-script-location

:: ── Step 7: Verify ───────────────────────────────────────────────────────────
echo.
echo  Verifying installation...
"%PYTHON_EXE%" -c "from faster_whisper import WhisperModel; print('[OK] faster-whisper working')" 2>nul
if !errorlevel! neq 0 (
  echo.
  if "%NEEDS_RESTART%"=="1" (
    echo  ====================================================
    echo   ACTION REQUIRED: Please restart Windows
    echo  ====================================================
    echo.
    echo  The Visual C++ Redistributable was just installed.
    echo  Windows needs a restart for it to take full effect.
    echo.
    echo  After restarting Windows, AnchorCast will
    echo  automatically complete the Whisper setup.
    echo.
    :: Write a flag file so AnchorCast knows to re-run setup on next launch
    echo restart_pending > "%ANCHORCAST_DIR%whisper_restart_pending.flag"
    echo  [Flag written] AnchorCast will retry on next launch.
  ) else (
    echo  ERROR: faster-whisper failed to load.
    echo  Try installing VC++ manually:
    echo  https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo  Then run this setup again.
  )
  goto :done
)

:download_model
:: ── Step 8: Download Whisper model ──────────────────────────────────────────
echo.
echo  Downloading Whisper small.en model (~244 MB)...
echo  Please wait -- this may take a few minutes...
echo.
"%PYTHON_EXE%" -c "from faster_whisper import WhisperModel; m=WhisperModel('small.en',device='cpu',compute_type='int8'); print('[OK] Model ready!')"
if !errorlevel! neq 0 (
  echo  ERROR: Model download failed. Check internet and try again.
  goto :done
)

:: Write success flag for AnchorCast to detect and auto-restart
echo success > "%ANCHORCAST_DIR%whisper_setup_complete.flag"

echo.
echo  ====================================================
echo   Setup complete! AnchorCast will now restart.
echo  ====================================================
echo.
echo  Local Whisper is ready. The microphone icon will
echo  turn green when transcription is active.
echo.

:: Close this window after 3 seconds and signal AnchorCast to restart
timeout /t 3 >nul
goto :done

:setup_failed
echo.
echo  Setup could not complete. Check your internet connection and try again.

:done
if not defined ANCHORCAST_NONINTERACTIVE (
  timeout /t 5 >nul
)
