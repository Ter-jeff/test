@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\DevChar_Sheet_9S_CP1.txt" "..\Production\DevChar_Sheet_9S_CP1.txt"
