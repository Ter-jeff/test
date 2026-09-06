@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\TIMESET_GLNA0_S_SC_SI_EQN_3.txt" "..\Production\TIMESET_GLNA0_S_SC_SI_EQN_3.txt"
