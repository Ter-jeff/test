@echo off

echo set variables
set CommandLineTool="%OasisRoot%IG-BenchmarkCL.exe"
set Options=-e "C:\GitLab\igxl_fakes\UnitTestsTw\tp_igbenchmark.json"

echo Execute Command Line: %CommandLineTool% %Options%
%CommandLineTool% %Options%

echo Press any key to exit
exit