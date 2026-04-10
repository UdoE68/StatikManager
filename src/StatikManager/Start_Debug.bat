@echo off
set EXE=C:\KI\StatikManager_V2\src\StatikManager\bin\x64\Debug\net48\StatikManager.exe
if not exist "%EXE%" (
    echo FEHLER: Debug-EXE nicht gefunden: %EXE%
    echo Bitte zuerst einen Debug-Build durchfuehren.
    pause
    exit /b 1
)
echo Starte Debug-Build...
echo %EXE%
start "" "%EXE%"
