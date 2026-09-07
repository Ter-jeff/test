@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\PwrBinning_V02_CP1.txt" "..\Production\PwrBinning_V02_CP1.txt"
