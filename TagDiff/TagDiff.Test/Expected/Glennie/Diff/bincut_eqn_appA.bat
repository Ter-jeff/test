@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\bincut_eqn_appA.txt" "..\Production\bincut_eqn_appA.txt"
