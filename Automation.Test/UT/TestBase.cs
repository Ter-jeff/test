using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

using AutogenCommandLine;
using AutogenCommandLine.CommandLineOptions;
using AutogenCommandLine.Tools.Autogen;

using CommonLib.ErrorReport;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MyCommandLineLib;

namespace Automation.Test.UT
{
    public class TestBase
    {
        protected static readonly string Root = Directory.GetCurrentDirectory();
        protected static readonly string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        protected static readonly string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        protected static readonly string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");
        protected static readonly string KPath = Path.Combine(Directory.GetCurrentDirectory(), "K");

        protected static void AssertOnlyWindowsOS(string reason)
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Assert.Inconclusive($"This test requires Windows OS: {reason}");
            }
        }

        // ReSharper disable once UnusedMember.Global
        protected static void RunBorneoTest(string subName, Action<List<string>, List<string>> runWorkflow)
        {
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testPlan = "borneo_documents/borneo_B0_TestPlan.xlsx";

            string testplanfile = Path.Combine(inputPath, testPlan);

            string binoutreport = Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_BinOut_V07A_CP1CP2_X#3.xlsx");
            string bincutModeSequence = Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_BMS_V07A_X_X#1.xlsx");
            string dramtype = Path.Combine(inputPath, "borneo_documents/B0_V07A/BORNEO_b0_DRAM_EXTERNAL#2.xlsx");
            string eqn = Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_EQN_V07A_X_X#1.xlsx");
            string fusechecktable = Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_FuseCheck_V07A_X_X#1.xlsx");
            string patternListCsv = Path.Combine(inputPath, "borneo_documents/B0_V07A/borneo_B0_pattern_dashboard_20251215_1702_18.csv");
            string postBinCut = Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_PBC_V07A_X_X#1.xlsx");
            string scgh = Path.Combine(inputPath, "borneo_documents/B0_V07A/borneo_B0_SCGH_X_X_X#57.xlsx");
            string efuseTestPlan = Path.Combine(inputPath, "borneo_documents/B0_V07A/BORNEO_b0_TESTPLAN_EFUSE_EXTERNAL#40.xlsx");
            string ttr = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_TTR_V07A_X_DigHIP#1.xlsx"),
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_TTR_V07A_X_HardIP_4.xlsx")
            });
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_VolTa_V07A_CP1_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_VolTa_V07A_CP2_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_VolTa_V07A_FT1_X#1.csv"),
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_VolTa_V07A_FT2_X#1.csv")
            });
            string binCut = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_Voltage_Binning_V1P5_111925.xlsx"),
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_Voltage_Binning_V1P6_120125_FT1.xlsx"),
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_B0_Voltage_Binning_V1P6_120125_FT2.xlsx")
            });
            string powerBinning = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_PowerScreening_V04_2025_1118.xlsx"),
                Path.Combine(inputPath, "borneo_documents/B0_V07A/Borneo_PowerScreening_V05_2025_1118.xlsx")
            });

            string repoPath = Path.Combine(inputPath, "borneo_documents/IGXLFiles/Repo");
            string vbtPath = Path.Combine(inputPath, "borneo_documents/IGXLFiles/VBT");
            string custom = Path.Combine(inputPath, "borneo_documents/B0_V07A/Custom");

            string outputFile = Path.Combine(outputPath, "borneo_documents/B0_V07A/CP1/");

            string patternfolder = Path.Combine(KPath, "borneo", "B0_V07A");
            string timesetfolder = Path.Combine(KPath, "borneo", "B0_V07A", "TimeSet");
            string info = Path.Combine(KPath, "borneo", "B0_V07A", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_borneo.txt");
            string compilePat = Path.Combine(KPath, "borneo", "B0_V07A", "CompilePat", "borneo_CompiledPat.csv");
            string mbistInfo = Path.Combine(KPath, "borneo", "B0_V07A", "CSV_Bist_Info", "Borneo_Bist_Info_All.csv");
            string cslibrarypath = Path.Combine(KPath, "Lib");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Autogen",
                "--currentproject", "borneo",
                "--skippatcheck", "TRUE",
                "--baslibraryfolder",$"\"{vbtPath}\"" ,
                "--repo", $"\"{repoPath}\"",
                "--outputFolder", $"\"{outputFile}\"",
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

            if (!CommandLineManager.ParseArgList([.. arguments.Select(x => x.Trim('"'))], out ICommandLineOptions options, out _))
            {
                return;
            }

            var args = (AutogenOptions)options;
            (List<string> testPlans, List<string> scghSheets) = AutogenCommandLineApp.PreWorkFlow(args);

            runWorkflow(testPlans, scghSheets);

            string report = Path.Combine(outputPath, "Error.txt");
            if (ErrorReportManager.GetErrorList().Count != 0)
            {
                IEnumerable<string> lines = ErrorReportManager.GetSortedErrors().Select(x => x.Print());
                File.AppendAllLines(report, lines);
            }

            bool fail = new FileComparisonReport("Autogen_" + subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        // ReSharper disable once UnusedMember.Global
        protected static void RunKomodoTest(string subName, Action<List<string>, List<string>> runWorkflow)
        {
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testPlan = "komodo_documents/Komodo_B0_TestPlan.xlsx";

            string vbtPath = Path.Combine(inputPath, "komodo_documents/IGXLFiles/VBT");
            string repoPath = Path.Combine(inputPath, "komodo_documents/IGXLFiles/Repo");
            string custom = Path.Combine(inputPath, "komodo_documents/B0_V07A/Custom");
            string outputFolder = Path.Combine(outputPath, "komodo_documents/B0_V07A/CP1/");
            string cslibrarypath = Path.Combine(KPath, "Lib");

            string bincutModeSequence = Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_BMS_V07A_X_X_7.xlsx");
            string patternListCsv = Path.Combine(inputPath, "komodo_documents/B0_V07A/komodo_B0_pattern_dashboard_20251211_2101_20.csv");
            string postBinCut = Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_PBC_V07A_X_X_2.xlsx");
            string scgh = Path.Combine(inputPath, "komodo_documents/B0_V07A/komodo_B0_SCGH_X_X_X_36.xlsx");
            string efuseTestPlan = Path.Combine(inputPath, "komodo_documents/B0_V07A/KOMODO_b0_TESTPLAN_EFUSE_EXTERNAL_75.xlsx");
            string binCut = Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_Voltage_Binning_V0P94_121025.xlsx");
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_CP1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_CP2_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_FQA1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_FQA2_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_FT1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_FT2_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_T0TxFT1_X_4.csv"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_VolTa_V07A_T0TxFT2_X_4.csv")
            });
            string ttr = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_TTR_V07A_X_HardIP_3.xlsx"),
                Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_TTR_V07A_X_DIGHardIP_1.xlsx")
            });

            string eqn = Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_B0_EQN_V07A_X_X_2.xlsx");
            string powerBinning = Path.Combine(inputPath, "komodo_documents/B0_V07A/Komodo_PowerBinning_V05_2025_1021.xlsx");
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
                "--outputFile", $"\"{outputFolder}\"",
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

            if (!CommandLineManager.ParseArgList([.. arguments.Select(x => x.Trim('"'))], out ICommandLineOptions options, out _))
            {
                return;
            }

            var args = (AutogenOptions)options;
            (List<string> testPlans, List<string> scghSheets) = AutogenCommandLineApp.PreWorkFlow(args);

            runWorkflow(testPlans, scghSheets);

            string report = Path.Combine(outputPath, "Error.txt");
            if (ErrorReportManager.GetErrorList().Count != 0)
            {
                IEnumerable<string> lines = ErrorReportManager.GetSortedErrors().Select(x => x.Print());
                File.AppendAllLines(report, lines);
            }

            bool fail = new FileComparisonReport("Autogen_" + subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        protected static void RunHidraTest(string subName, Action<List<string>, List<string>> runWorkflow)
        {
            string inputPath = Path.Combine(InputPath);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testplanfile = Path.Combine(inputPath, "hidra_documents", "Hidra_Ext_ATE_TestPlan_649.xlsx");

            string eqn = Path.Combine(inputPath, "hidra_documents", "Hidra_B0_EQN_V05A_X_X_3.xlsx");
            string patternListCsv = Path.Combine(inputPath, "hidra_documents", "hidra_B0_pattern_dashboard_20241106_0201_16_TW_V2023.12.28.1_2024-11-21_142829.csv");
            string postBinCut = Path.Combine(inputPath, "hidra_documents", "Hidra_B0_PBC_V05A_X_X_2.xlsx");
            string scgh = Path.Combine(inputPath, "hidra_documents", "hidra_B0_scgh_file_96.xlsx");
            string efuseTestPlan = "";
            string voltageTables = string.Join(",", new List<string>
            {
                Path.Combine(inputPath, "hidra_documents", @"Hidra_B0_VolTa_V05A_CP1_X_1.csv")
            });
            string binCut = Path.Combine(inputPath, "hidra_documents", "Hidra_B0_Voltage_Binning_V05AengPCPUInterp_111824.xlsx");
            string bincutModeSequence = Path.Combine(inputPath, "hidra_documents", "Hidra_C0_BMS_V06A_X_X_5.xlsx");
            string fusechecktable = Path.Combine(inputPath, "hidra_documents", "Hidra_C0_FuseCheck_V06A_X_X_1.xlsx");
            string powerBinning = Path.Combine(inputPath, "hidra_documents", "Hidra_PowerBinning_V05_2024_1010.xlsx");
            string dramtype = "";
            string binoutreport = "";
            string ttr = "";

            string cslibrarypath = "";

            string repoPath = Path.Combine(inputPath, "hidra_documents");
            string custom = "";

            string outputFolder = outputPath;

            string patternfolder = Path.Combine(KPath, "hidra", "patx");
            string timesetfolder = Path.Combine(KPath, "hidra", "TimeSet");
            string info = Path.Combine(KPath, "hidra", "HARD_IP", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All.txt");
            string compilePat = Path.Combine(KPath, "hidra", "CompilePat", "hidra_CompiledPat.csv");
            string mbistInfo = Path.Combine(KPath, "hidra", "CSV_Bist_Info", "hidra_Bist_Info_All.csv");
            string vbtPath = Path.Combine(KPath, "central_library_vb");

            string projectconfig = Path.Combine(inputPath, "hidra_documents", "ProjectConfig_hidra.ini");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Autogen",
                "--currentproject", "hidra",
                "--skippatcheck", "FALSE",
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
                "--projectconfig", $"\"{projectconfig}\"",
                "-m", "1"
            };

            if (!CommandLineManager.ParseArgList([.. arguments.Select(x => x.Trim('"'))], out ICommandLineOptions options, out _))
            {
                return;
            }

            var args = (AutogenOptions)options;
            (List<string> testPlans, List<string> scghSheets) = AutogenCommandLineApp.PreWorkFlow(args);

            runWorkflow(testPlans, scghSheets);

            bool fail = new FileComparisonReport("Autogen_" + subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        public static void ExecuteCmdByWindow(string cmd, string argument)
        {
            AssertOnlyWindowsOS("ExecuteCmdByWindow requires cmd.exe");
            var process = new Process();
            var startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/k {cmd} {argument}",
                UseShellExecute = true,
                CreateNoWindow = false,
                WindowStyle = ProcessWindowStyle.Normal
            };
            process.StartInfo = startInfo;
            _ = process.Start();
            process.WaitForExit();
        }
    }
}
