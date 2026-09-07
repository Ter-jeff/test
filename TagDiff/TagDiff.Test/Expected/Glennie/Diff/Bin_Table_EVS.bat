@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\Bin_Table_EVS.txt" "..\Production\Bin_Table_EVS.txt"
