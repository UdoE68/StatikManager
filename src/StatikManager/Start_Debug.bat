@echo off
set EXE=c:\Projekte\StatikManager\bin\Debug\net48\StatikManager.exe
if not exist "%EXE%" (
    echo FEHLER: Debug-EXE nicht gefunden: %EXE%
    echo Bitte zuerst einen Debug-Build durchfuehren.
    pause
    exit /b 1
)
echo Starte Debug-Build...
echo %EXE%
start "" "%EXE%"
