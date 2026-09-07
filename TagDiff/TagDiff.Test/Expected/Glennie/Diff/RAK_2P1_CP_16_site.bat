@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\RAK_2P1_CP_16_site.txt" "..\Production\RAK_2P1_CP_16_site.txt"
