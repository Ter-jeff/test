REM Start Check Script...
cd ..
set "DLL_PATH=AutogenCommandLine.dll"
set EXECUTE_MODE= BincutCheck
set INPUT_FOLDER= ...\input
set PATTERN_FILE= ...\XXX.igxl
set OUTPUT_FOLDER= ...\output

REM Executing command lines
dotnet "%DLL_PATH%" -e %EXECUTE_MODE% -f "%INPUT_FOLDER%" -p "%PATTERN_FILE%" -o "%OUTPUT_FOLDER%"

pause
