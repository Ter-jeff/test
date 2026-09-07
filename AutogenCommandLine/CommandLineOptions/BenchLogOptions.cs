using System;
using System.Collections.Generic;

using CommandLine;

using MyCommandLineLib;

namespace AutogenCommandLine.CommandLineOptions
{
    public class BenchLogOptions : ICommandLineOptions
    {
        [Option("Mode", Required = true, HelpText = "Check mode : PreCheck/GenerateProgram")]
        public string Mode { get; set; } = null!;
        [Option("PinMap", Required = true, HelpText = "Input sheet : PinMap sheet")]
        public string PinMap { get; set; } = null!;
        [Option("EfuseBitDefition", Required = true, HelpText = "Input sheet : EfuseBitDefition sheet")]
        public string EfuseBitDefition { get; set; } = null!;
        [Option("AddrFor64", Required = true, HelpText = "Input file : 64bit_registers file")]
        public string AddrFor64 { get; set; } = null!;
        [Option("BenchLog", Required = true, HelpText = "Input file : BenchLog folder")]
        public string BenchLog { get; set; } = null!;
        [Option("RegisterMap", Required = false, HelpText = "Input file : Register Map file (Sample : null)")]
        public string RegisterMap { get; set; } = null!;
        [Option("PatternType", Required = false, HelpText = "Config : ARF/FW, HTOL (Sample : ARF/FW)")]
        public string PatternType { get; set; } = null!;
        [Option("JTAGFreq", Required = false, HelpText = "Config : JTAG Freq ___ MHz (Sample : 24)")]
        public string JTAGFreq { get; set; } = null!;
        [Option("PatFolder", Required = false, HelpText = "Config : Pat Folder (Sample : B0_ARF)")]
        public string PatFolder { get; set; } = null!;
        [Option("ThreadNumber", Required = false, HelpText = "Config : Thread Number (Sample : 3)")]
        public string ThreadNumber { get; set; } = null!;
        [Option("ProjectName", Required = false, HelpText = "Config : Project Name (Sample : Faraday)")]
        public string ProjectName { get; set; } = null!;
        [Option("ProjectSilicon", Required = false, HelpText = "Config : Project and Silicon (Sample : FAR,B0)")]
        public string ProjectSilicon { get; set; } = null!;
        [Option("PreSetupPatterns", Required = false, HelpText = "Config : PreSetup Patterns (Sample : pat1,pat2,pat3...)")]
        public string PreSetupPatterns { get; set; } = null!;
        [Option("IsR16", Required = false, HelpText = "Config : Is R16 project (Sample : N)")]
        public string IsR16 { get; set; } = null!;
        [Option("IsFullSweep", Required = false, HelpText = "Config : Is need generate full sweep pattern (Sample : N)")]
        public string IsFullSweep { get; set; } = null!;
        [Option("OverWriteLogs", Required = false, HelpText = "Config : BenchLog name need trim field overwrite (Sample : benchlog1,benchlog2,benchlog3...)")]
        public string OverWriteLogs { get; set; } = null!;
        [Option("TimeStamp", Required = false, HelpText = "Config : Add time stamp to the suffix of pattern (Sample : 202412162346)")]
        public string TimeStamp { get; set; } = null!;
        [Option("IsAddComment", Required = false, HelpText = "Config : Generate comments that are required by autogen/library (Sample : N)", DefaultValue = "N")]
        public string IsAddComment { get; set; } = null!;
        [Option("TestPlanShell", Required = false, HelpText = "Input file : TestPlan file")]
        public string TestPlanShell { get; set; } = null!;

        //[Option("Scgh", Required = false, HelpText = "Input file : Scgh file")]
        public string Scgh { get; set; } = null!;

        //[Option("PatternInfo", Required = false, HelpText = "Input file : PatternInfo file")]
        public string PatternInfo { get; set; } = null!;

        [Option("LibraryPath", Required = false, HelpText = "Input file : Library folder")]
        public string LibraryPath { get; set; } = null!;

        //[Option("Program", Required = false, HelpText = "Input file : Program file")]
        //public string Program { get; set; }

        [Option('o', "Output", Required = true, HelpText = "Output path")]
        public string Output { get; set; } = null!;
        // ProximaW release for next Json syntax 
        [Option("CppDspControl", Required = false, HelpText = "Enable CPP DSP control behaviors in pattern/program generation (switch flag)")]
        public string CppDspControl { get; set; } = null!;
        // Draco Perform a pre-check to verify whether the Python library BenchLog is available.
        [Option("PythonLibrary", Required = false, HelpText = "Perform a pre-check to verify whether the Python library BenchLog is available.")]
        public string PythonLibrary { get; set; } = null!;

        public void GetHelp(string[] args)
        {
            var help = new List<string>
            {
                "",
                "**** BenchLog Tool ****",
                "Command-line arguments:",
                "--Mode                        Mode(required)                              -- Check mode : PreCheck/GenerateProgram",
                "--PinMap                      PinMap(required)                            -- Input sheet : PinMap sheet",
                "--EfuseBitDefition            EfuseBitDefition(required)                  -- Input sheet : EfuseBitDefition sheet",
                "--AddrFor64                   AddrFor64(required)                         -- Input file : 64bit_registers file",
                "--BenchLog                    BenchLog(required)                          -- Input file : BenchLog folder",
                "--RegisterMap                 RegisterMap(optional)                       -- Input file : Register Map file (Sample : null)",
                "--PatternType                 PatternType(optional)                       -- Config : ARF/FW, HTOL (Sample : ARF/FW)",
                "--JTAGFreq                    JTAGFreq(optional)                          -- Config : JTAG Freq ___ MHz (Sample : 24)",
                "--PatFolder                   PatFolder(optional)                         -- Config : Pat Folder (Sample : B0_ARF)",
                "--ThreadNumber                ThreadNumber(optional)                      -- Config : Thread Number (Sample : 3)",
                "--ProjectName                 ProjectName(optional)                       -- Config : Project Name (Sample : Faraday)",
                "--ProjectSilicon              ProjectSilicon(optional)                    -- Config : Project and Silicon (Sample : FAR,B0)",
                "--PreSetupPatterns            PreSetupPatterns(optional)                  -- Config : PreSetup Patterns (Sample : pat1,pat2,pat3...)",
                "--IsR16                       IsR16(optional)                             -- Config : Is R16 project (Sample : N)",
                "--IsFullSweep                 IsFullSweep(optional)                       -- Config : Is need generate full sweep pattern (Sample : N)",
                "--OverWriteLogs               OverWriteLogs(optional)                     -- Config : BenchLog name need trim field overwrite (Sample : benchlog1,benchlog2,benchlog3...)",
                "--TimeStamp                   TimeStamp(optional)                         -- Config : Add time stamp to the suffix of pattern (Sample : 202412162346)",
                "--IsAddComment                IsAddComment(optional)                      -- Config : Generate comments that are required by autogen/library (Sample : N)",
                "--TestPlanShell               TestPlanShell(optional)                     -- Input file : TestPlan file",
                "--LibraryPath                 LibraryPath(optional)                       -- Input file : Library folder",
                "-o, --Output                  Output(required)                            -- Output path",
                "--CppDspControl               CppDspControl(optional)                     -- Enable CPP DSP control behaviors in pattern/program generation (switch flag)",
                "--PythonLibrary               PythonLibrary(optional)                     -- Perform a pre-check to verify whether the Python library BenchLog is available.",
                "",
                "Surround file/folder paths with double quotes if they contain spaces",
                "",
            };

            foreach (string line in help)
            {
                Console.WriteLine(line);
            }
        }

        public string GetOutputFolder()
        {
            return Output;
        }
    }
}
