@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\Flow_EVS_SOC_Scan_EVS0.txt" "..\Production\Flow_EVS_SOC_Scan_EVS0.txt"
