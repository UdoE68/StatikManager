@echo off
set EXE=c:\Projekte\StatikManager\bin\Release\net48\StatikManager.exe
if not exist "%EXE%" (
    echo FEHLER: Release-EXE nicht gefunden: %EXE%
    echo Bitte zuerst einen Release-Build durchfuehren.
    pause
    exit /b 1
)
echo Starte Release-Build...
echo %EXE%
start "" "%EXE%"
