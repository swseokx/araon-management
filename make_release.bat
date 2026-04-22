@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ================================================
echo  ARAON Release Build
echo ================================================

powershell -NoProfile -ExecutionPolicy Bypass -File get_version.ps1 > _ver.tmp
set /p VERSION=<_ver.tmp
del _ver.tmp 2>nul
if "%VERSION%"=="" (
    echo [ERR] version.json read failed
    pause
    exit /b 1
)
echo VERSION: v%VERSION%
echo.

set /p NOTES=Release notes (blank=auto):

echo.
echo [1/5] Building main.exe...
pyinstaller main.spec --clean -y
if errorlevel 1 goto err1

echo.
echo [2/5] Building ARAON.exe...
pyinstaller launcher.spec --clean -y
if errorlevel 1 goto err2

echo.
echo [3/5] Creating release folder...
if not exist release mkdir release
if not exist release\bin mkdir release\bin
if not exist release\img mkdir release\img

copy /y "dist\main.exe"         "release\bin\main.exe"
copy /y "version.json"          "release\bin\version.json"
copy /y "timetable_data.json"   "release\bin\timetable_data.json"
copy /y "credentials.json"      "release\bin\credentials.json"
copy /y "settings.ini.template" "release\bin\settings.ini"
copy /y "favicon.ico"           "release\img\favicon.ico"
for %%f in (*.png) do copy /y "%%f" "release\img\%%f"
copy /y "dist\ARAON.exe"        "release\ARAON.exe"

echo.
echo [4/5] Creating ZIP...
set ZIPNAME=ARAON_v%VERSION%.zip
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'release\*' -DestinationPath '%ZIPNAME%' -Force"
if errorlevel 1 goto err3

echo.
echo [5/5] GitHub Release...
where gh >nul 2>&1
if errorlevel 1 goto nogh
gh auth status >nul 2>&1
if errorlevel 1 goto noauth
if "%NOTES%"=="" (
    gh release create "v%VERSION%" "%ZIPNAME%" --title "ARAON Enterprise v%VERSION%" --generate-notes
) else (
    gh release create "v%VERSION%" "%ZIPNAME%" --title "ARAON Enterprise v%VERSION%" --notes "%NOTES%"
)
if errorlevel 1 goto err4

:done
echo.
echo ================================================
echo  v%VERSION% Release Done!
echo ================================================
pause
exit /b 0

:err1
echo [ERR] main.exe build failed
pause
exit /b 1
:err2
echo [ERR] ARAON.exe build failed
pause
exit /b 1
:err3
echo [ERR] ZIP creation failed
pause
exit /b 1
:nogh
echo [WARN] gh not found - upload ZIP manually: %ZIPNAME%
goto done
:noauth
echo [WARN] gh auth login required - upload ZIP manually: %ZIPNAME%
goto done
:err4
echo [ERR] GitHub Release failed - upload manually: %ZIPNAME%
pause
exit /b 1
