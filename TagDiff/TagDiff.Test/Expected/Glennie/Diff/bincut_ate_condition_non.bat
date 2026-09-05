@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\bincut_ate_condition_non.txt" "..\Production\bincut_ate_condition_non.txt"
