@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "EXE_LOCAL=%SCRIPT_DIR%FGH_Modbus_Sollwert_Manager.exe"
set "EXE_DIST=%SCRIPT_DIR%dist\Release_FGH_Modbus_Sollwert_Manager\FGH_Modbus_Sollwert_Manager.exe"

if exist "%EXE_LOCAL%" (
    start "" "%EXE_LOCAL%"
    exit /b 0
)

if exist "%EXE_DIST%" (
    start "" "%EXE_DIST%"
    exit /b 0
)

echo Die Release-EXE wurde nicht gefunden:
echo %EXE_LOCAL%
echo %EXE_DIST%
exit /b 1
