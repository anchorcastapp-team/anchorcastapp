@echo off
echo SermonCast — Bible Data Downloader
echo ====================================
echo.
echo Downloading complete KJV Bible (31,102 verses)...
echo.

REM Try curl (available on Windows 10+)
curl -L -o kjv_raw.json "https://raw.githubusercontent.com/scrollmapper/bible_databases/master/json/t_kjv.json" 2>nul

IF NOT EXIST kjv_raw.json (
    echo curl failed, trying PowerShell...
    powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/scrollmapper/bible_databases/master/json/t_kjv.json' -OutFile 'kjv_raw.json'"
)

IF NOT EXIST kjv_raw.json (
    echo ERROR: Could not download Bible data.
    echo Please manually download from:
    echo https://github.com/scrollmapper/bible_databases/blob/master/json/t_kjv.json
    echo Save as: kjv.json in this folder
    pause
    exit /b 1
)

REM Convert format: the file uses {b,c,v,t} which is what we expect
REM Just rename it
move kjv_raw.json kjv.json

echo.
echo SUCCESS! kjv.json downloaded with 31102 verses.
echo Restart SermonCast to load the complete Bible.
echo.
pause
