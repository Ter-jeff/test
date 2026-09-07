using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using AutogenCommandLine;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

// ReSharper disable IdentifierTypo
// ReSharper disable StringLiteralTypo

namespace Automation.Test.UT.Project
{
    [TestClass]
    public class Komodo : TestBase
    {
        private const string Command = "AutogenCommandLine.exe";
        private readonly bool _isDebug = true;

        [TestMethod]
        public void Full_Komodo()
        {
            AssertOnlyWindowsOS("integration test requires IGXL tools and Windows-generated expected output");
            string subName = "komodo_igxl";
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            Directory.CreateDirectory(outputPath);

            string testPlan = Path.Combine("komodo_documents", "Komodo_B0_TestPlan.xlsx");

            string vbtPath = Path.Combine(KPath, "central_library_vb");
            string repoPath = Path.Combine(inputPath, "komodo_documents", "IGXLFiles", "Repo");
            string custom = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Custom");
            string outputFolder = Path.Combine(outputPath, "komodo_documents", "B0_V07A", "CP1");
            string cslibrarypath = Path.Combine(KPath, "Lib");

            string bincutModeSequence = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_BMS_V07A_X_X_7.xlsx");
            string patternListCsv = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "komodo_B0_pattern_dashboard_20251211_2101_20.csv");
            string postBinCut = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_PBC_V07A_X_X_2.xlsx");
            string scgh = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "komodo_B0_SCGH_X_X_X_36.xlsx");
            string efuseTestPlan = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "KOMODO_b0_TESTPLAN_EFUSE_EXTERNAL_75.xlsx");
            string binCut = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_Voltage_Binning_V0P94_121025.xlsx");
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_CP1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_CP2_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_FQA1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_FQA2_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_FT1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_FT2_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_T0TxFT1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_VolTa_V07A_T0TxFT2_X_4.csv")
            });
            string ttr = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_TTR_V07A_X_HardIP_3.xlsx"),
                Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_TTR_V07A_X_DIGHardIP_1.xlsx")
            });

            string eqn = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_B0_EQN_V07A_X_X_2.xlsx");
            string powerBinning = Path.Combine(inputPath, "komodo_documents", "B0_V07A", "Komodo_PowerBinning_V05_2025_1021.xlsx");
            string dramtype = Path.Combine(inputPath, "KOMODO_b0_DRAM_EXTERNAL_1.xlsx");
            string fusechecktable = Path.Combine(inputPath, "Komodo_B0_FuseCheck_V07A_X_X_3.xlsx");

            string testplanfile = Path.Combine(inputPath, testPlan);

            string patternfolder = Path.Combine(KPath, "komodo", "patx");
            string timesetfolder = Path.Combine(KPath, "komodo", "TimeSet");
            string info = Path.Combine(KPath, "komodo", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_komodo.txt");
            string compilePat = Path.Combine(KPath, "komodo", "CompilePat", "komodo_CompiledPat.csv");
            string mbistInfo = Path.Combine(KPath, "komodo", "CSV_Bist_Info", "komodo_Bist_Info_All.csv");
            string settingconfig = "";

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Autogen",
                "--currentproject", "komodo",
                "--skippatcheck", "TRUE",
                "--baslibraryfolder",$"\"{vbtPath}\"" ,
                "--repo", $"\"{repoPath}\"",
                "--outputFolder", $"\"{outputFolder}\"",
                "--job", "CP1",
                "--channelmap", "ChannelMap_CP_4_site",
                "--patternfolder", $"\"{patternfolder}\"",
                "--timesetfolder", $"\"{timesetfolder}\"",
                "-g", $"\"{info}\"",
                "--compilePat", $"\"{compilePat}\"",
                "--mbistInfo",$"\"{mbistInfo}\"" ,
                "--testplanfile", $"\"{testplanfile}\"",
                "--voltagetablefile",$"\"{voltageTables}\"" ,
                "--scghfile", $"\"{scgh}\"",
                "--bincutmodesequence", $"\"{bincutModeSequence}\"",
                "--postbincut",$"\"{postBinCut}\"" ,
                "--patlistcsvfile",$"\"{patternListCsv}\"" ,
                "--bincut", $"\"{binCut}\"",
                "--efusetestplan",$"\"{efuseTestPlan}\"",
                "--cslibrarypath", $"\"{cslibrarypath}\"",
                "--settingconfig", $"\"{settingconfig}\"",
                "--powerbinning", $"\"{powerBinning}\"",
                "--eqn", $"\"{eqn}\"",
                "--dramtype", $"\"{dramtype}\"",
                "--ttr", $"\"{ttr}\"",
                "--fusechecktable", $"\"{fusechecktable}\"",
                "--custom", $"\"{custom}\"",
                "--skippatcheck", "true",
                "-m", "1"
            };

            string argument = string.Join(" ", arguments);
            Console.WriteLine($"{Command} {argument}");
            if (_isDebug)
            {
                AutogenCommandLineMain.Main([.. arguments.Select(x => x.Trim('"'))]);
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

            var excluding = new List<string> { "Versions.txt", "ExecInfo.txt", "output_hashes.txt" };
            bool fail = new FileComparisonReport("Autogen_" + subName, excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
