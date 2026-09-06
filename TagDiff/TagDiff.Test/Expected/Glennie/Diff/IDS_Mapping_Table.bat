@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\IDS_Mapping_Table.txt" "..\Production\IDS_Mapping_Table.txt"
