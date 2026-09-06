@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\ChannelMap_CP_EVS_1_site.txt" "..\Production\ChannelMap_CP_EVS_1_site.txt"
