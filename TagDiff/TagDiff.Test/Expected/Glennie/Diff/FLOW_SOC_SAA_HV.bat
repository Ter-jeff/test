@echo off
cd /d %~dp0
echo Current working directory is: %cd%
"C:\Program Files (x86)\teradyne\oasis\KDiff3\kdiff3.exe" "..\Autogen\FLOW_SOC_SAA_HV.txt" "..\Production\FLOW_SOC_SAA_HV.txt"
