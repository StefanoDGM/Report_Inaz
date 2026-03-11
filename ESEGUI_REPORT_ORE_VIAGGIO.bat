@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "SCRIPT_FILE=%ROOT_DIR%Script\Genera_Report_Ore_Viaggio.ps1"
set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%POWERSHELL_EXE%" (
    set "POWERSHELL_EXE=powershell"
)

if not exist "%SCRIPT_FILE%" (
    echo ERRORE: file script non trovato: %SCRIPT_FILE%
    pause
    exit /b 1
)

if "%~1"=="" goto run_noargs
if "%~2"=="" goto run_onearg
goto run_twoargs

:run_noargs
"%POWERSHELL_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_FILE%"
goto after_run

:run_onearg
"%POWERSHELL_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_FILE%" "%~1"
goto after_run

:run_twoargs
"%POWERSHELL_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_FILE%" "%~1" "%~2"

:after_run
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo.
    echo Esecuzione terminata con errore, codice %EXIT_CODE%.
    pause
    exit /b %EXIT_CODE%
)

echo.
echo Operazione completata.
pause
exit /b 0
