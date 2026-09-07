@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\PatSets_ClockOut.txt" "..\Production\PatSets_ClockOut.txt"
