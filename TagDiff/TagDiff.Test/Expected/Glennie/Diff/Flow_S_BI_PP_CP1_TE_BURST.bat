@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\Flow_S_BI_PP_CP1_TE_BURST.txt" "..\Production\Flow_S_BI_PP_CP1_TE_BURST.txt"
