@echo off
REM Start Autogen...
cd ..
set "DLL_PATH=AutogenCommandLine.dll"
set "INPUTINFO=Help\InputInfo.csv"
set "MODE=1"

REM Executing command lines
dotnet %DLL_PATH% --inputinfo "%INPUTINFO%" -m %MODE%

pause
