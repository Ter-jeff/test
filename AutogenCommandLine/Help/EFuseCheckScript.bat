REM Start Check Script...
cd ..
set "DLL_PATH=AutogenCommandLine.dll"
set EXECUTE_MODE= EFuseCheck
set INPUT_FOLDER= ...\input
set BINDEF_TABLE= ...\EFUSE_BitDef_Table.txt
set CONFIG_TABLE= ...\Config_table.txt
set OUTPUT_FOLDER= ...\output

REM Executing command lines
dotnet "%DLL_PATH%" -e %EXECUTE_MODE% -f "%INPUT_FOLDER%" -b "%BINDEF_TABLE%" -c "%CONFIG_TABLE%" -o "%OUTPUT_FOLDER%"

pause