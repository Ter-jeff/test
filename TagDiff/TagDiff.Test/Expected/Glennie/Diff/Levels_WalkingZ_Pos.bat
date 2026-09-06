@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\Levels_WalkingZ_Pos.txt" "..\Production\Levels_WalkingZ_Pos.txt"
