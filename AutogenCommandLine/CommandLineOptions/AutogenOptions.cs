using System;
using System.Collections.Generic;
using System.IO;

using CommandLine;

using MyCommandLineLib;

namespace AutogenCommandLine.CommandLineOptions
{
    public class AutogenOptions : ICommandLineOptions
    {
        [Option("inputinfo", Required = false)]
        public string InputInfo { get; set; } = null!;
        [Option('p', "currentproject", Required = true)]
        public string CurrentProjectName { get; set; } = null!;

        [Option('t', "testplanfile", Required = true)]
        public string TestPlanFile { get; set; } = null!;

        [Option("repo", Required = false)]
        public string RepoFolderPath { get; set; } = null!;

        [Option("basciconfigfile", Required = false)]
        public string BasicConfigFile { get; set; } = null!;

        [Option("hardipinfo", Required = false)]
        public string HardiInfo { get; set; } = null!;

        [Option("binCutinstanceNamingrule", Required = false)]
        public string BinCutInstanceNamingRule { get; set; } = null!;

        [Option("rtoscategoryconfig", Required = false)]
        public string RtosCategoryConfig { get; set; } = null!;

        [Option("scanconfig", Required = false)]
        public string ScanConfig { get; set; } = null!;

        [Option("mbistconfig", Required = false)]
        public string MbistConfig { get; set; } = null!;

        [Option("spiconfig", Required = false)]
        public string SpiConfig { get; set; } = null!;

        [Option("hardipconfig", Required = false)]
        public string HardipConfig { get; set; } = null!;

        [Option("tnassignment", Required = false)]
        public string TnAssignment { get; set; } = null!;

        [Option("binnumberconfig", Required = false)]
        public string BinNumberConfig { get; set; } = null!;

        [Option("inputdir", Required = false)]
        public string InputDirectory { get; set; } = null!;

        [Option('h', "projectconfig", Required = false)]
        public string ProjectConfig { get; set; } = null!;

        [Option('s', "scghfile", Required = false)]
        public string ScghFile { get; set; } = null!;

        [Option('c', "patlistcsvfile", Required = false)]
        public string PatternListCsvFile { get; set; } = null!;

        [Option('b', "bincut", Required = false)]
        public string BinCutFile { get; set; } = null!;

        [Option('n', "postbincut", Required = false)]
        public string PostBinCutFile { get; set; } = null!;

        [Option('k', "bincutmodesequence", Required = false)]
        public string BincutModeSequence { get; set; } = null!;

        [Option('v', "voltagetablefile", Required = false)]
        public string VoltageTable { get; set; } = null!;

        [Option('f', "otpfiles", Required = false)]
        public string OtpFiles { get; set; } = null!;

        [Option('y', "yamlfile", Required = false)]
        public string YamlFile { get; set; } = null!;

        [Option('g', "patterninfofile", Required = false)]
        public string PatternInfoFile { get; set; } = null!;

        [Option('a', "patternfolder", Required = false)]
        public string PatternFolder { get; set; } = null!;

        [Option('i', "timesetfolder", Required = false)]
        public string TimesetFolder { get; set; } = null!;

        [Option('d', "baslibraryfolder", Required = true)]
        public string BasLibraryFolder { get; set; } = null!;

        [Option('m', "mock", Required = false)]
        public int Mock { get; set; } = 0;

        [Option("needsheet", Required = true)]
        public string NeedSheet { get; set; } = null!;
        [Option("compilePat", Required = true)]
        public string CompilePat { get; set; } = null!;

        [Option("eqn", Required = true)]
        public string EqnVoltage { get; set; } = null!;

        [Option("cslibrarypath", Required = false)]
        public string CsLibraryPath { get; set; } = null!;

        [Option("mbistInfo", Required = false)]
        public string MbistInfo { get; set; } = null!;

        [Option("skippatcheck", Required = false)]
        public string SkipPatCheck { get; set; } = null!;

        [Option("skip-iglink", Required = false)]
        public string SkipIgLink { get; set; } = null!;

        [Option("job", Required = false)]
        public string DefaultJob { get; set; } = null!;

        [Option("channelmap", Required = false)]
        public string DefaultChannelMap { get; set; } = null!;

        [Option("powerbinning", Required = false)]
        public string PowerBinning { get; set; } = null!;

        [Option("fusechecktable", Required = false)]
        public string FuseCheckTable { get; set; } = null!;

        [Option("dramtype", Required = false)]
        public string DramType { get; set; } = null!;

        [Option("efusetestplan", Required = false)]
        public string EfuseTestPlan { get; set; } = null!;

        [Option("ttr", Required = false)]
        public string TtrTable { get; set; } = null!;

        [Option("custom", Required = false)]
        public string CustomPath { get; set; } = null!;

        [Option("isufp", Required = false)]
        public string IsUfp { get; set; } = null!;

        [Option("search_valid_pattern_rev", Required = false)]
        public string SearchValidPatternRev { get; set; } = null!;

        [Option("binoutreport", Required = false)]
        public string BinoutReport { get; set; } = null!;

        [Option("gent0tx", Required = false)]
        public string GenT0TX { get; set; } = null!;

        [Option("charplan", Required = false)]
        public string CharPlan { get; set; } = null!;
        [Option("vre_enable", Required = false)]
        public string VreEnable { get; set; } = null!;

        [Option('o', "outputfolder", Required = true)]
        public string OutputDirectory { get; set; } = null!;

        [Option("testprogramname", Required = false)]
        public string TestProgramName { get; set; } = null!;

        void IOptions.GetHelp(string[] args)
        {
            var help = new List<string>
            {
                "",
                "**** Autogen Tool ****",
                "Command-line arguments:",
                "-p                            Current project name(required)              -- Current project name(string)",
                "-repo                         Current project repository path             -- Current project repository path(string)",
                "-t                            Test plan file(required)                    -- Test plan file(.xlsx/.xlsm)",
                "-o                            Output Folder(required)                     -- Output Folder(folder)",
                "--inputdir                    Input directory(optional)                   -- Directory contains all input files(string)",
                "--basciconfigfile             Basci Config file(optional)                 -- Basic config file(.xlsx)",
                "--nwirefile                   Nwrie file(optional)                        -- Nwire file(.xlsx)",
                "--hardipinfo                  Hardi Info file(optional)                   -- Hardi Info file(.log)",
                "--binCutinstanceNamingrule    BinCut instance naming rule file(optional)  -- BinCut instance naming rule file(.xlsx)",
                "--rtoscategoryconfig          Rtos category config file(optional)         -- Rtos category config file(.xlsx)",
                "--tnassignment                TN assignment file(optional)                -- TN assignment file(.xlsx)",
                "--scanconfig                  Scan config file(optional)                  -- Scan config file(.xml)",
                "--mbistonfig                  Mbist config file(optional)                 -- Mbist config file(.xml)",
                "--spiconfig                   Spi config file(optional)                   -- Spi config file(.xml)",
                "--hardipconfig                Hardip config file(optional)                -- Hardip config  file(.xml)",
                "--binnumberconfig             Bin number config file(optional)            -- Bin number config file(.xlsx)",
                "--settingconfig               Setting config file(optional)               -- Setting config file(.ini)",
                "--needsheet                   Need sheet config file(optional)            -- Need sheet config file(.xml)",
                "--gent0tx                     Generate T0TX test program(optional)        -- Generate T0TX test program(true/false)",
                "--vre_enable                  Generate VRE test case scenario(optional)   -- Generate VRE test case scenario(true/false)",
                "-h                            Project config file(optional)               -- Project config file(.ini)",
                "-s                            Scgh file(optional)                         -- Scgh file(.xlsx/.xlsm)",
                "-c                            Pattern list CSV file(optional)             -- Pattern list CSV file(.csv)",
                "-b                            Bincut file(optional)                       -- Bincut file(.xlsx/.xlsm)",
                "-n                            Post bincut file(optional)                  -- Post bincut file(.xlsx/.xlsm)",
                "-v                            Voltage table file(optional)                -- Voltage table file(.csv)",
                "-k                            Bincut mode sequence table file(optional)   -- Bincut mode sequence file(.xlsx/.xlsm)",
                "-f                            OTP files(optional)                         -- OTP files(.otp)",
                "-y                            Yaml file(optional)                         -- Yaml file(.yaml)",
                "-g                            Pattern info file(optional)                 -- Pattern info file(.txt)",
                "-a                            Pattern folder(optional)                    -- Pattern folder(folder)",
                "-i                            Timeset folder(optional)                    -- Timeset folder(folder)",
                "-d                            Bas library folder(optional)                -- Bas library folder(folder)",
                "-isufp                        Is Uxtraflex or UxtraflexPlus(optional)     -- Instrument type",
                "-l                            Log output folder(optional)                 -- Log output folder(folder)",
                "-m                            Mock(optional)                              -- Mock(0/1)",
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
            return OutputDirectory;
        }

        public void Initialize()
        {
            if (string.IsNullOrEmpty(CurrentProjectName))
            {
                BasicConfigFile = GetStringOrDefault(BasicConfigFile, Path.Combine("Settings", "Basic", "Basic_Configure_Default.xlsx"));
                BinCutInstanceNamingRule = GetStringOrDefault(BinCutInstanceNamingRule, Path.Combine("Settings", "SCGH", "BinCutInstanceNamingRule_Default.xlsx"));
                RtosCategoryConfig = GetStringOrDefault(RtosCategoryConfig, Path.Combine("Settings", "Basic", "RtosCategory_Default.xlsx"));
                ScanConfig = GetStringOrDefault(ScanConfig, Path.Combine("Settings", "SCGH", "Scan_Config_Default.xml"));
                MbistConfig = GetStringOrDefault(MbistConfig, Path.Combine("Settings", "SCGH", "Mbist_Config_Default.xml"));
                SpiConfig = GetStringOrDefault(SpiConfig, Path.Combine("Settings", "SCGH", "Spi_Config_Default.xml"));
                HardipConfig = GetStringOrDefault(HardipConfig, Path.Combine("Settings", "SCGH", "HardIP_Config_Default.xml"));
                TnAssignment = GetStringOrDefault(TnAssignment, Path.Combine("Settings", "Basic", "TN_Assignment_Default.xlsx"));
                BinNumberConfig = GetStringOrDefault(BinNumberConfig, Path.Combine("Config", "BinNumberConfig.xlsx"));
                ProjectConfig = GetStringOrDefault(ProjectConfig, Path.Combine("Settings", "ProjectConfig_OTC.ini"));
                NeedSheet = GetStringOrDefault(NeedSheet, Path.Combine("Settings", "Basic", "NeedSheetConfig_Default.xml"));
            }
            else
            {
                BasicConfigFile = GetStringOrDefault(BasicConfigFile, Path.Combine("Settings", "Basic", $"Basic_Configure_{CurrentProjectName}.xlsx"), RepoFolderPath);
                BinCutInstanceNamingRule = GetStringOrDefault(BinCutInstanceNamingRule, Path.Combine("Settings", "SCGH", $"BinCutInstanceNamingRule_{CurrentProjectName}.xlsx"), RepoFolderPath);
                RtosCategoryConfig = GetStringOrDefault(RtosCategoryConfig, Path.Combine("Settings", "Basic", $"RtosCategory_{CurrentProjectName}.xlsx"), RepoFolderPath);
                ScanConfig = GetStringOrDefault(SpiConfig, Path.Combine("Settings", "SCGH", $"Scan_Config_{CurrentProjectName}.xml"), RepoFolderPath);
                MbistConfig = GetStringOrDefault(MbistConfig, Path.Combine("Settings", "SCGH", $"Mbist_Config_{CurrentProjectName}.xml"), RepoFolderPath);
                SpiConfig = GetStringOrDefault(SpiConfig, Path.Combine("Settings", "SCGH", $"Spi_Config_{CurrentProjectName}.xml"), RepoFolderPath);
                HardipConfig = GetStringOrDefault(HardipConfig, Path.Combine("Settings", "SCGH", $"HardIP_Config_{CurrentProjectName}.xml"), RepoFolderPath);
                TnAssignment = GetStringOrDefault(TnAssignment, Path.Combine("Settings", "Basic", $"TN_Assignment_{CurrentProjectName}.xlsx"), RepoFolderPath);
                BinNumberConfig = GetStringOrDefault(BinNumberConfig, Path.Combine("Config", $"BinNumberConfig_{CurrentProjectName}.xlsx"), RepoFolderPath);
                ProjectConfig = GetStringOrDefault(ProjectConfig, Path.Combine("Settings", $"ProjectConfig_{CurrentProjectName}.ini"), RepoFolderPath);
                NeedSheet = GetStringOrDefault(NeedSheet, Path.Combine("Settings", "Basic", $"NeedSheetConfig_{CurrentProjectName}.xml"), RepoFolderPath);
            }
        }

        private static string GetStringOrDefault(string str, string filePath, string workingDir = null)
        {
            string finalWorkingDir = workingDir is null
                ? Directory.GetCurrentDirectory()
                : workingDir;
            return string.IsNullOrEmpty(str)
                ? Path.Combine(finalWorkingDir, filePath)
                : str;
        }
    }
}
