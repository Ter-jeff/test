using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

using AutogenCommandLine;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace TagDiff.Test
{
    [TestClass]
    public class TagDiffTests
    {
        private const string Command = "AutogenCommandLine.exe";
        protected static readonly string Root = Directory.GetCurrentDirectory();

        private static readonly string _inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        private static readonly string _outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        private static readonly string _expectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");
        private readonly bool _isDebug = false;

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void GlennieDiffTest_AutogenCommandLineApp()
        {
            string subName = "Glennie";
            string inputPath = Path.Combine(_inputPath, subName);
            string outputPath = Path.Combine(_outputPath, "Glennie_AutogenCommandLineApp");
            string expectPath = Path.Combine(_expectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            Directory.CreateDirectory(outputPath);

            string autoGenTestProgramPath = Path.Combine(inputPath, "Glennie.igxl");
            string productionTestProgramPath = Path.Combine(inputPath, "UR89_CP1_X16_E02_250822_V02A_PROD_BCV0P1_25C.igxl");
            string job = "CP1";
            string exclude_dirs = Path.Combine(inputPath, "ExcludeDirs");

            var arguments = new List<string>
            {
                "-e", "TagDiff",
                "-a", autoGenTestProgramPath,
                "-p", productionTestProgramPath,
                "-o", outputPath,
                "--job", job,
                "-m", "1",
                "--exclude_dirs", exclude_dirs,
            };
            string argument = string.Join(" ", arguments);
            Console.WriteLine($"{Command} {argument}");

            AutogenCommandLineMain.Main([.. arguments]);

            var excluding = new List<string> { "ElapsedTime.json", "TagDiff.log" };
            bool fail = new FileComparisonReport(subName, excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void GlennieDiffTest_AutogenCommandLineApp_exe()
        {
            string subName = "Glennie";
            string inputPath = Path.Combine(_inputPath, subName);
            string outputPath = Path.Combine(_outputPath, "Glennie_AutogenCommandLineApp_exe");
            string expectPath = Path.Combine(_expectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            Directory.CreateDirectory(outputPath);

            string autoGenTestProgramPath = Path.Combine(inputPath, "Glennie.igxl");
            string productionTestProgramPath = Path.Combine(inputPath, "UR89_CP1_X16_E02_250822_V02A_PROD_BCV0P1_25C.igxl");
            string job = "CP1";
            string exclude_dirs = Path.Combine(inputPath, "ExcludeDirs");

            var arguments = new List<string>
                {
                    "-e", "TagDiff",
                    "-a", autoGenTestProgramPath,
                    "-p", productionTestProgramPath,
                    "-o", outputPath,
                    "--job", job,
                    "-m", "1",
                    "--exclude_dirs", exclude_dirs,
                };
            string argument = string.Join(" ", arguments);
            string dir = Root;

            if (_isDebug)
            {
                AutogenCommandLineMain.Main([.. arguments]);
            }
            else
            {
                string batFilePath = "script.bat";
                string batchCommands = $@"
@echo off
echo Step 1: Changing directory to: {dir}
cd /d {"\"" + dir + "\""}

echo Step 2: Executing command
{Command} {argument}

echo Done!
exit
";
                File.WriteAllText(batFilePath, batchCommands);
                ExecuteCmdByWindow(batFilePath, "");
            }

            var excluding = new List<string> { "ElapsedTime.json", "TagDiff.log" };
            bool fail = new FileComparisonReport(subName, excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        public static void ExecuteCmdByWindow(string cmd, string argument)
        {
            var process = new Process();
            var startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/k {cmd} {argument}",  // "/k" keeps the window open
                UseShellExecute = true,  // Allows opening a new window
                CreateNoWindow = false,  // Ensure it is not hidden
                WindowStyle = ProcessWindowStyle.Normal
            };
            process.StartInfo = startInfo;
            _ = process.Start();
            process.WaitForExit();
        }
    }
}
