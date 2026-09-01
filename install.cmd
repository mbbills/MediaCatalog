@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"
title MediaCatalog Installer

echo.
echo MediaCatalog integrated installer
echo ================================
echo.

call :find_python
if defined MC_PYTHON goto python_ready

echo Python 3.8 or newer was not found.
echo Python 3.8.10 is the final official Python release with a Windows 7 installer.
echo.

set "PYTHON_VERSION=3.8.10"
if /I "%PROCESSOR_ARCHITECTURE%"=="x86" (
    set "PYTHON_URL=https://www.python.org/ftp/python/3.8.10/python-3.8.10.exe"
    set "PYTHON_INSTALLER=%TEMP%\MediaCatalog-python-3.8.10.exe"
    set "PYTHON_TARGET=%LocalAppData%\Programs\Python\Python38-32"
) else (
    set "PYTHON_URL=https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe"
    set "PYTHON_INSTALLER=%TEMP%\MediaCatalog-python-3.8.10-amd64.exe"
    set "PYTHON_TARGET=%LocalAppData%\Programs\Python\Python38"
)

echo Downloading Python %PYTHON_VERSION% from python.org...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=3072; (New-Object Net.WebClient).DownloadFile('%PYTHON_URL%','%PYTHON_INSTALLER%')" >nul 2>&1
if exist "%PYTHON_INSTALLER%" goto python_downloaded

echo PowerShell download failed; trying certutil...
certutil.exe -urlcache -split -f "%PYTHON_URL%" "%PYTHON_INSTALLER%" >nul 2>&1
if not exist "%PYTHON_INSTALLER%" (
    echo.
    echo ERROR: Python could not be downloaded.
    echo Download it manually from:
    echo %PYTHON_URL%
    goto failed
)

:python_downloaded
for %%Z in ("%PYTHON_INSTALLER%") do set "PYTHON_SIZE=%%~zZ"
if %PYTHON_SIZE% LSS 20000000 (
    echo ERROR: The downloaded Python installer is unexpectedly small.
    del /q "%PYTHON_INSTALLER%" >nul 2>&1
    goto failed
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$s=Get-AuthenticodeSignature -LiteralPath '%PYTHON_INSTALLER%'; if($s.Status -ne 'Valid'){exit 1}" >nul 2>&1
if errorlevel 1 (
    echo WARNING: Windows could not validate the Python installer's signature.
    echo This can happen on Windows 7 when root certificates are old.
    choice /C YN /N /M "Continue with the installer downloaded over HTTPS? [Y/N] "
    if errorlevel 2 goto failed
)

echo Installing Python for the current user...
"%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 TargetDir="%PYTHON_TARGET%" Include_launcher=1 PrependPath=1 Include_test=0
set "PYTHON_INSTALL_RC=%ERRORLEVEL%"
if not "%PYTHON_INSTALL_RC%"=="0" if not "%PYTHON_INSTALL_RC%"=="3010" (
    echo ERROR: Python installer returned code %PYTHON_INSTALL_RC%.
    goto failed
)

call :find_python
if not defined MC_PYTHON (
    if exist "%PYTHON_TARGET%\python.exe" set "MC_PYTHON=%PYTHON_TARGET%\python.exe"
)
if not defined MC_PYTHON (
    echo ERROR: Python installation completed but python.exe was not found.
    goto failed
)

:python_ready
echo Using Python: "%MC_PYTHON%"
"%MC_PYTHON%" -E "%~dp0scripts\install_media_catalog.py"
if errorlevel 1 goto failed

echo.
echo Excel template
echo --------------
echo MediaCatalog can build a macro-enabled Excel template with the Media Catalog menu.
echo Excel must be closed. The builder temporarily enables programmatic VBA access
echo for this operation and restores the previous setting afterward.
echo.
choice /C YN /N /M "Build MediaCatalog_template.xlsm now? [Y/N] "
if errorlevel 2 goto finish
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build_excel_template.ps1" -ProjectRoot "%~dp0"
if errorlevel 1 (
    echo.
    echo The data installation succeeded, but the Excel template was not built.
    echo See README.md for the manual Excel module-import procedure.
)

:finish
echo.
echo Installation complete.
echo.
echo Read README.md, open MediaCatalog_template.xlsm or MediaCatalog_template.ods,
echo enable its macros, enter or paste catalog data, select the data rows, and run:
echo.
echo     Media Catalog ^> Resolve Selected Rows
echo.
choice /C YN /N /M "Open README.md now? [Y/N] "
if errorlevel 2 goto success
start "" "%~dp0README.md"

:success
echo.
pause
exit /b 0

:find_python
set "MC_PYTHON="
for /f "usebackq delims=" %%P in (`py -3 -c "import sys; print(sys.executable)" 2^>nul`) do call :accept_python "%%P"
if defined MC_PYTHON exit /b 0
for /f "usebackq delims=" %%P in (`where python 2^>nul`) do call :accept_python "%%P"
exit /b 0

:accept_python
if defined MC_PYTHON exit /b 0
"%~1" -c "import sys; raise SystemExit(0 if sys.version_info >= (3,8) else 1)" >nul 2>&1
if not errorlevel 1 set "MC_PYTHON=%~1"
exit /b 0

:failed
echo.
echo MediaCatalog installation did not complete.
echo Correct the problem above and run install.cmd again.
echo.
pause
exit /b 1
