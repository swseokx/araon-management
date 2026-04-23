@echo off
chcp 65001 >/dev/null
setlocal
cd /d "%~dp0"

echo ================================================
echo  ARAON Release Build
echo ================================================

powershell -NoProfile -ExecutionPolicy Bypass -Command "$v=(Get-Content version.json -Raw|ConvertFrom-Json).version; Write-Host $v" > _ver.tmp
set /p VERSION=<_ver.tmp
del _ver.tmp >/dev/null 2>/dev/null
if "%VERSION%"=="" (
    echo VERSION read failed
    pause
    exit /b 1
)
echo VERSION: v%VERSION%
echo.

set /p NOTES="Release notes: "

echo.
echo [1/5] Building main.exe...
pyinstaller main.spec --clean -y
if errorlevel 1 goto err1

echo.
echo [2/5] Building launcher...
pyinstaller launcher.spec --clean -y
if errorlevel 1 goto err2

echo.
echo [3/5] Creating release folder...
set RELEASE=release
if exist %RELEASE% rmdir /s /q %RELEASE%
mkdir %RELEASE%
mkdir %RELEASE%\bin
mkdir %RELEASE%\img

copy /y "dist\main.exe"           "%RELEASE%\bin\main.exe"
copy /y "version.json"            "%RELEASE%\bin\version.json"
copy /y "timetable_data.json"     "%RELEASE%\bin\timetable_data.json"
copy /y "credentials.json"        "%RELEASE%\bin\credentials.json"
copy /y "settings.ini.template"   "%RELEASE%\bin\settings.ini"
copy /y "favicon.ico"             "%RELEASE%\img\favicon.ico"
for %%f in (*.png) do copy /y "%%f" "%RELEASE%\img\%%f"

copy /y "dist\launcher.exe"  "%RELEASE%\launcher.exe"

echo.
echo [4/5] Creating ZIP...
set ZIPNAME=ARAON_v%VERSION%.zip
if exist "%ZIPNAME%" del /f /q "%ZIPNAME%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path '%RELEASE%\*' -DestinationPath '%ZIPNAME%' -Force"
if errorlevel 1 goto err3

echo.
echo [5/5] GitHub Release...
where gh >/dev/null 2>/dev/null
if errorlevel 1 goto nogh

gh auth status >/dev/null 2>/dev/null
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
echo [ERR] launcher build failed
pause
exit /b 1
:err3
echo [ERR] ZIP creation failed
pause
exit /b 1
:nogh
echo [WARN] gh CLI not found - ZIP created: %ZIPNAME%
goto done
:noauth
echo [WARN] gh auth required - ZIP created: %ZIPNAME%
goto done
:err4
echo [ERR] GitHub Release failed - upload manually: %ZIPNAME%
pause
exit /b 1
