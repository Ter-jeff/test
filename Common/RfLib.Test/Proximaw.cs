using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using AutogenCommandLine;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

// ReSharper disable IdentifierTypo
// ReSharper disable StringLiteralTypo

namespace RfLibLib.Test
{
    [TestClass]
    public class Proximaw
    {
        private const string Command = "AutogenCommandLine.exe";
        private readonly bool _isDebug = true;

        protected static readonly string Root = Directory.GetCurrentDirectory();
        protected static readonly string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        protected static readonly string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        protected static readonly string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");
        protected static readonly string KPath = Path.Combine(Directory.GetCurrentDirectory(), "K");

        [TestMethod]
        public void Full_Proximaw()
        {
            string subName = "proximaw";
            string inputPath = Path.Combine(InputPath, subName);
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            string testPlan = "ProximaW_A0_Test_Plan_99.xlsx";

            string testplanfile = Path.Combine(inputPath, testPlan);

            string binoutreport = ""; //Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_BinOut_V07A_CP1CP2_X#3.xlsx");
            string bincutModeSequence = ""; //Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_BMS_V07A_X_X#1.xlsx");
            string dramtype = ""; //Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_b0_DRAM_EXTERNAL#2.xlsx");
            string eqn = ""; //Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_EQN_V07A_X_X#1.xlsx");
            string fusechecktable = ""; //Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_FuseCheck_V07A_X_X#1.xlsx");
            string patternListCsv = Path.Combine(inputPath, "proximaw_A0_pattern_dashboard_20251112_1800_51.csv");
            string postBinCut = ""; //Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_PBC_V07A_X_X#1.xlsx");
            string scgh = Path.Combine(inputPath, "proximaw_A0_scgh_file_114.xlsx");
            string efuseTestPlan = ""; // Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_b0_TESTPLAN_EFUSE_EXTERNAL#40.xlsx");
            string ttr = ""; // string.Join(",", new List<string>
            //{
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_TTR_V07A_X_DigHIP#1.xlsx"),
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_TTR_V07A_X_HardIP_4.xlsx")
            //});
            string voltageTables = ""; //string.Join(",", new List<string>
            //{
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_VolTa_V07A_CP1_X#1.csv"),
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_VolTa_V07A_CP2_X#1.csv"),
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_VolTa_V07A_FT1_X#1.csv"),
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_VolTa_V07A_FT2_X#1.csv")
            //});
            string binCut = ""; //string.Join(",", new List<string>
            //{
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_Voltage_Binning_V1P5_111925.xlsx"),
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_Voltage_Binning_V1P6_120125_FT1.xlsx"),
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_B0_Voltage_Binning_V1P6_120125_FT2.xlsx")
            //});
            string powerBinning = ""; // string.Join(",", new List<string>
            //{
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_PowerScreening_V04_2025_1118.xlsx"),
            //    Path.Combine(inputPath, "proximaw\\B0_V07A\\Proximaw_PowerScreening_V05_2025_1118.xlsx")
            //});

            string repoPath = Path.Combine(inputPath, "IGXLFiles", "Repo");
            string custom = ""; // Path.Combine(inputPath, "proximaw\\B0_V07A\\Custom");

            string outputFolder = Path.Combine(outputPath, "CP1");

            string patternfolder = ""; // Path.Combine(KPath, "Proximaw", "B0_V07A");
            string timesetfolder = ""; //Path.Combine(KPath, "Proximaw", "B0_V07A", "TimeSet");
            string info = ""; //Path.Combine(KPath, "Proximaw", "B0_V07A", "CSV_HardIP_Info", "HardIP_AutoGen_Info_All_Proximaw.txt");
            string compilePat = ""; //Path.Combine(KPath, "Proximaw", "B0_V07A", "CompilePat", "Proximaw_CompiledPat.csv");
            string mbistInfo = ""; // Path.Combine(KPath, "Proximaw", "B0_V07A", "CSV_Bist_Info", "Proximaw_Bist_Info_All.csv");
            string cslibrarypath = ""; //Path.Combine(KPath, "Lib");
            string vbtPath = ""; //Path.Combine(KPath, "central_library_vb");

            string dir = Root;
            var arguments = new List<string>
            {
                "/c", $"\"{dir}\"",
                "--entry", "Autogen",
                "--currentproject", "proximaw",
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

            var excluding = new List<string> { "Versions.txt", "ExecInfo.txt", "output_hashes.txt" };
            bool fail = new FileComparisonReport("Autogen_Proximaw_full", excluding).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
