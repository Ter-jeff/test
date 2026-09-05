@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\PwrScreen_B798.txt" "..\Production\PwrScreen_B798.txt"
