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
    public class Borneo : TestBase
    {
        private const string Command = "AutogenCommandLine.exe";
        private readonly bool _isDebug = true;

        [TestMethod]
        public void Full_Borneo()
        {
            AssertOnlyWindowsOS("integration test requires IGXL tools and Windows-generated expected output");
            string subName = Path.Combine("borneo_igxl", "borneo_documents", "B0_V07A");
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testPlan = Path.Combine("borneo_documents", "borneo_B0_TestPlan.xlsx");

            string testplanfile = Path.Combine(inputPath, testPlan);

            string binoutreport = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_BinOut_V07A_CP1CP2_X#3.xlsx");
            string bincutModeSequence = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_BMS_V07A_X_X#1.xlsx");
            string dramtype = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "BORNEO_b0_DRAM_EXTERNAL#2.xlsx");
            string eqn = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_EQN_V07A_X_X#1.xlsx");
            string fusechecktable = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_FuseCheck_V07A_X_X#1.xlsx");
            string patternListCsv = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "borneo_B0_pattern_dashboard_20251215_1702_18.csv");
            string postBinCut = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_PBC_V07A_X_X#1.xlsx");
            string scgh = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "borneo_B0_SCGH_X_X_X#57.xlsx");
            string efuseTestPlan = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "BORNEO_b0_TESTPLAN_EFUSE_EXTERNAL#40.xlsx");
            string ttr = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_TTR_V07A_X_DigHIP#1.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_TTR_V07A_X_HardIP_4.xlsx")
            });
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_CP1_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_CP2_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_FT1_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_FT2_X#1.csv")
            });
            string binCut = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_Voltage_Binning_V1P5_111925.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_Voltage_Binning_V1P6_120125_FT1.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_Voltage_Binning_V1P6_120125_FT2.xlsx")
            });
            string powerBinning = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_PowerScreening_V04_2025_1118.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_PowerScreening_V05_2025_1118.xlsx")
            });

            string repoPath = Path.Combine(inputPath, "borneo_documents", "IGXLFiles", "Repo");
            string custom = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Custom");

            string outputFolder = Path.Combine(outputPath, "CP1");

            string patternfolder = Path.Combine(KPath, "borneo", "B0_V07A");
            string timesetfolder = Path.Combine(KPath, "borneo", "B0_V07A", "TimeSet");
            string info = Path.Combine(KPath, "borneo", "B0_V07A", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string compilePat = Path.Combine(KPath, "borneo", "B0_V07A", "CompilePat", "borneo_CompiledPat.csv");
            string mbistInfo = Path.Combine(KPath, "borneo", "B0_V07A", "CSV_Bist_Info", "Borneo_Bist_Info_All.csv");
            string cslibrarypath = Path.Combine(KPath, "Lib");
            string vbtPath = Path.Combine(KPath, "central_library_vb");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Autogen",
                "--currentproject", "borneo",
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
                "--powerbinning", $"\"{powerBinning}\"",
                "--eqn", $"\"{eqn}\"",
                "--dramtype", $"\"{dramtype}\"",
                "--ttr", $"\"{ttr}\"",
                "--fusechecktable", $"\"{fusechecktable}\"",
                "--custom", $"\"{custom}\"",
                "--binoutreport", $"\"{binoutreport}\"",
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
            bool fail = new FileComparisonReport("Autogen_borneo_full", excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void Full_Borneo_T0Tx()
        {
            AssertOnlyWindowsOS("integration test requires IGXL tools and Windows-generated expected output");
            string subName = Path.Combine("borneo_igxl", "borneo_documents", "B0_V07A_T0Tx");
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testPlan = Path.Combine("borneo_documents", "borneo_B0_TestPlan.xlsx");

            string testplanfile = Path.Combine(inputPath, testPlan);

            string binoutreport = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_BinOut_V07A_CP1CP2_X#3.xlsx");
            string bincutModeSequence = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_BMS_V07A_X_X#1.xlsx");
            string dramtype = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "BORNEO_b0_DRAM_EXTERNAL#2.xlsx");
            string eqn = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_EQN_V07A_X_X#1.xlsx");
            string fusechecktable = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_FuseCheck_V07A_X_X#1.xlsx");
            string patternListCsv = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "borneo_B0_pattern_dashboard_20251215_1702_18.csv");
            string postBinCut = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_PBC_V07A_X_X#1.xlsx");
            string scgh = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "borneo_B0_SCGH_X_X_X#57.xlsx");
            string efuseTestPlan = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "BORNEO_b0_TESTPLAN_EFUSE_EXTERNAL#40.xlsx");
            string ttr = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_TTR_V07A_X_DigHIP#1.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_TTR_V07A_X_HardIP_4.xlsx")
            });
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_CP1_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_CP2_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_FT1_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_VolTa_V07A_FT2_X#1.csv")
            });
            string binCut = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_Voltage_Binning_V1P5_111925.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_Voltage_Binning_V1P6_120125_FT1.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_B0_Voltage_Binning_V1P6_120125_FT2.xlsx")
            });
            string powerBinning = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_PowerScreening_V04_2025_1118.xlsx"),
                Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Borneo_PowerScreening_V05_2025_1118.xlsx")
            });

            string repoPath = Path.Combine(inputPath, "borneo_documents", "IGXLFiles", "Repo");
            string custom = Path.Combine(inputPath, "borneo_documents", "B0_V07A", "Custom");

            string outputFolder = Path.Combine(outputPath, "CP1");

            string patternfolder = Path.Combine(KPath, "borneo", "B0_V07A");
            string timesetfolder = Path.Combine(KPath, "borneo", "B0_V07A", "TimeSet");
            string info = Path.Combine(KPath, "borneo", "B0_V07A", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string compilePat = Path.Combine(KPath, "borneo", "B0_V07A", "CompilePat", "borneo_CompiledPat.csv");
            string mbistInfo = Path.Combine(KPath, "borneo", "B0_V07A", "CSV_Bist_Info", "Borneo_Bist_Info_All.csv");
            string cslibrarypath = Path.Combine(KPath, "Lib");
            string vbtPath = Path.Combine(KPath, "central_library_vb");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Autogen",
                "--currentproject", "borneo",
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
                "--powerbinning", $"\"{powerBinning}\"",
                "--eqn", $"\"{eqn}\"",
                "--dramtype", $"\"{dramtype}\"",
                "--ttr", $"\"{ttr}\"",
                "--fusechecktable", $"\"{fusechecktable}\"",
                "--custom", $"\"{custom}\"",
                "--binoutreport", $"\"{binoutreport}\"",
                "--skippatcheck", "true",
                "--gent0tx", "true",
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
            bool fail = new FileComparisonReport("Autogen_borneo_T0Tx", excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void Lite_Borneo()
        {
            AssertOnlyWindowsOS("integration test requires IGXL tools and Windows-generated expected output");
            string subName = Path.Combine("borneo_igxl", "borneo_documents", "A0_V04A");
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testPlan = Path.Combine("borneo_documents", "borneo_A0_TestPlan.xlsx");

            string testplanfile = Path.Combine(inputPath, testPlan);

            string binoutreport = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_BinOut_V04A_X_HardIP_1.xlsx");
            string bincutModeSequence = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_BMS_V04A_X_X#3.xlsx");
            string dramtype = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "BORNEO_a0_DRAM_EXTERNAL#3.xlsx");
            string eqn = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_EQN_V04A_X_X#1.xlsx");
            string fusechecktable = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_FuseCheck_V04A_X_X#3.xlsx");
            string patternListCsv = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "borneo_A0_pattern_dashboard_20251002_1524_15.csv");
            string postBinCut = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_PBC_V04A_X_X#3.xlsx");
            string scgh = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "borneo_A0_SCGH_X_X_X#208.xlsx");
            string efuseTestPlan = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "BORNEO_a0_TESTPLAN_EFUSE_EXTERNAL#20251001.xlsx");
            string ttr = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_TTR_V04A_X_DigHardIP.xlsx")
            });
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_CP1_X#4.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_CP2_X#4.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_FT1_X#5.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_FT2_X#5.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_T0TxFT1_X#2.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_T0TxFT2_X#2.csv")
            });
            string binCut = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_Voltage_Binning_V0P61_092325.xlsx");
            string powerBinning = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_PowerScreening_V01_2025_0716.xlsx");

            string repoPath = Path.Combine(inputPath, "borneo_documents", "IGXLFiles", "Repo");
            string custom = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Custom");

            string outputFolder = Path.Combine(outputPath, "CP1");

            string patternfolder = Path.Combine(KPath, "borneo", "A0_V04A", "patx");
            string timesetfolder = Path.Combine(KPath, "borneo", "A0_V04A", "TimeSet");
            string info = Path.Combine(KPath, "borneo", "A0_V04A", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string compilePat = Path.Combine(KPath, "borneo", "A0_V04A", "CompilePat", "borneo_CompiledPat.csv");
            string mbistInfo = Path.Combine(KPath, "borneo", "A0_V04A", "CSV_Bist_Info", "Borneo_Bist_Info_All.csv");
            string cslibrarypath = Path.Combine(KPath, "Lib");
            string vbtPath = Path.Combine(KPath, "central_library_vb");

            string charPlanPath = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_DFTL_V04A_FT1_9SITE#1.xlsx");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Autogen",
                "--currentproject", "borneo",
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
                "--powerbinning", $"\"{powerBinning}\"",
                "--eqn", $"\"{eqn}\"",
                "--dramtype", $"\"{dramtype}\"",
                "--ttr", $"\"{ttr}\"",
                "--fusechecktable", $"\"{fusechecktable}\"",
                "--custom", $"\"{custom}\"",
                "--binoutreport", $"\"{binoutreport}\"",
                "--search_valid_pattern_rev", "True",
                "-m", "1",
                "--charplan", $"\"{charPlanPath}\""
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

            var excluding = new List<string> { "Versions.txt", "ExecInfo.txt", "output_hashes.txt", "hardip_info.log", "CharPreProcessor.log", "PostProcessor.log", "ErrorReport.txt" };
            bool fail = new FileComparisonReport("Autogen_borneo_lite", excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void Full_Borneo_Validation()
        {
            AssertOnlyWindowsOS("integration test requires IGXL tools and Windows-generated expected output");
            string subName = "borneo_validation";
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testPlan = Path.Combine("borneo_documents", "borneo_A0_TestPlan.xlsx");

            string testplanfile = Path.Combine(inputPath, testPlan);

            string binoutreport = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_BinOut_V04A_X_HardIP_1.xlsx");
            string bincutModeSequence = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_BMS_V04A_X_X#3.xlsx");
            string dramtype = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "BORNEO_a0_DRAM_EXTERNAL#3.xlsx");
            string eqn = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_EQN_V04A_X_X#1.xlsx");
            string fusechecktable = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_FuseCheck_V04A_X_X#3.xlsx");
            string patternListCsv = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "borneo_A0_pattern_dashboard_20251002_1524_15.csv");
            string postBinCut = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_PBC_V04A_X_X#3.xlsx");
            string scgh = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "borneo_A0_SCGH_X_X_X#208.xlsx");
            string efuseTestPlan = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "BORNEO_a0_TESTPLAN_EFUSE_EXTERNAL#20251001.xlsx");
            string ttr = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_TTR_V04A_X_DigHardIP.xlsx")
            });
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_CP1_X#4.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_CP2_X#4.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_FT1_X#5.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_FT2_X#5.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_T0TxFT1_X#2.csv"),
                Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_VolTa_V04A_T0TxFT2_X#2.csv")
            });
            string binCut = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_A0_Voltage_Binning_V0P61_092325.xlsx");
            string powerBinning = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Borneo_PowerScreening_V01_2025_0716.xlsx");

            string repoPath = Path.Combine(inputPath, "borneo_documents", "IGXLFiles", "Repo");
            string custom = Path.Combine(inputPath, "borneo_documents", "A0_V04A", "Custom");

            string outputFolder = Path.Combine(outputPath, "CP1");

            string patternfolder = Path.Combine(KPath, "borneo", "A0_V04A", "patx");
            string timesetfolder = Path.Combine(KPath, "borneo", "A0_V04A", "TimeSet");
            string info = Path.Combine(KPath, "borneo", "A0_V04A", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string compilePat = Path.Combine(KPath, "borneo", "A0_V04A", "CompilePat", "borneo_CompiledPat.csv");
            string mbistInfo = Path.Combine(KPath, "borneo", "A0_V04A", "CSV_Bist_Info", "Borneo_Bist_Info_All.csv");
            string cslibrarypath = Path.Combine(KPath, "Lib");
            string vbtPath = Path.Combine(KPath, "central_library_vb");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Validation",
                "--currentproject", "borneo",
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
                "--powerbinning", $"\"{powerBinning}\"",
                "--eqn", $"\"{eqn}\"",
                "--dramtype", $"\"{dramtype}\"",
                "--ttr", $"\"{ttr}\"",
                "--fusechecktable", $"\"{fusechecktable}\"",
                "--custom", $"\"{custom}\"",
                "--binoutreport", $"\"{binoutreport}\"",
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

            var excluding = new List<string> { "Validation.log" };
            bool fail = new FileComparisonReport("Autogen_" + subName, excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void Lite_Borneo_Csv()
        {
            AssertOnlyWindowsOS("integration test requires IGXL tools and Windows-generated expected output");
            string subName = Path.Combine("borneo_igxl", "borneo_documents", "A0_V04A");
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, "borneo_csv", "borneo_documents", "A0_V04A");
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string inputInfo = Path.Combine(inputPath, "borneo_documents", "borneo_InputInfo.csv");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--inputinfo", $"\"{inputInfo}\"" ,
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

            var excluding = new List<string> { "Versions.txt", "ExecInfo.txt", "output_hashes.txt", "hardip_info.log", "CharPreProcessor.log", "PostProcessor.log", "ErrorReport.txt" };
            bool fail = new FileComparisonReport("Autogen_borneo_csv", excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
