using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

using Automation.Reader;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility;

using CommonLib.Extension;
using CommonLib.Static;
using CommonLib.Utility;

using CommonReaderLib;

using OfficeOpenXml;

using TestPlanLib.PatternListCsvFile;
using TestPlanLib.Static;

// ReSharper disable StringLiteralTypo

namespace Automation
{
    public class UserInfo
    {
        public UserInfo(InputInfoSheet sheet)
        {
            if (sheet.Rows == null || sheet.Rows.Count == 0)
            {
                return;
            }

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (InputInfoRow row in sheet.Rows)
            {
                if (!string.IsNullOrEmpty(row.Item))
                {
                    dict[row.Item.Trim()] = row.Value?.Trim() ?? string.Empty;
                }
            }

            string Get(string key) => dict.TryGetValue(key, out string value) ? value : string.Empty;

            Entry = Get("entry");
            Job = Get("job");
            CurrentProject = Get("currentproject");
            BasLibraryFolder = Get("baslibraryfolder");
            RepoFolderPath = Get("repo");
            OutputFolder = Get("outputfolder");
            ChannelMap = Get("channelmap");
            PatternFolder = Get("patternfolder");
            TimeSetFolder = Get("timesetfolder");
            HardIpInfoFileName = Get("patterninfofile");
            CompilePat = Get("compilepat");
            MbistInfo = Get("mbistinfo");
            TestPlanFile = Get("testplanfile");
            VoltageTableFiles = Get("voltagetablefiles").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            ScghFile = Get("scghfile");
            BinCutModeSeqFileName = Get("bincutmodesequence");
            BinCutPostFileName = Get("postbincut");
            PatternListCsvFile = Get("patlistcsvfile");
            BinCutFileNames = Get("bincuts");
            EfuseTestPlan = Get("efusetestplan");
            CsLibraryPath = Get("cslibrarypath");
            PowerBinningFileNames = Get("powerbinnings");
            EquationVoltagesFileName = Get("eqn");
            DramTypeFileName = Get("dramtype");
            TtrSummaryFileNames = Get("ttrs").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            FuseCheckTable = Get("fusechecktable");
            CustomPath = Get("custom");
            BinOutReportFileName = Get("binoutreport");
            SearchValidPatternRev = Get("search_valid_pattern_rev").Equals("True", StringComparison.CurrentCultureIgnoreCase);
            CharFileName = Get("charplan");

            Options = new Options(sheet);
            LocalSpecs.Options = Options;
        }

        public string CurrentProject { get; set; }
        public string Entry { get; set; }
        public string Job { get; set; }
        public string BasLibraryFolder { get; set; }
        public string RepoFolderPath { get; set; }
        public string OutputFolder { get; set; }

        public string ChannelMap { get; set; }
        public string PatternFolder { get; set; }
        public string TimeSetFolder { get; set; }
        public string HardIpInfoFileName { get; set; }
        public string CompilePat { get; set; }
        public string MbistInfo { get; set; }
        public string TestPlanFile { get; set; }
        public List<string> VoltageTableFiles { get; set; } = new List<string>();
        public string ScghFile { get; set; }
        public string BinCutModeSeqFileName { get; set; }
        public string BinCutPostFileName { get; set; }
        public string PatternListCsvFile { get; set; }
        public string BinCutFileNames { get; set; }
        public string EfuseTestPlan { get; set; }
        public string CsLibraryPath { get; set; }
        public string PowerBinningFileNames { get; set; }
        public string EquationVoltagesFileName { get; set; }
        public string DramTypeFileName { get; set; }
        public List<string> TtrSummaryFileNames { get; set; }
        public string FuseCheckTable { get; set; }
        public string CustomPath { get; set; }
        public string BinOutReportFileName { get; set; }
        public string CharFileName { get; set; }
        public string TestProgramName { get; set; }
        public ConfigFiles ConfigFiles { get; set; } = new ConfigFiles();
        public SettingFiles SettingFiles { get; set; } = new SettingFiles();
        public Options Options { get; set; } = new Options();

        private void ResolveConfigPaths()
        {
            LocalSpecs.Options = Options;
            ErrorReportStatic.StopByCritical = LocalSpecs.Options.StopByCritical;
            string needSheet = "";

            string repoProjectRepoPath = RepoFolderPath;
            string repoProjectName = CurrentProject;

            SettingFiles.BasicConfigFile = ResolveConfigPath(SettingFiles.BasicConfigFile, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "Basic", "Basic_Configure_{0}.xlsx"), Path.Combine("Settings", "Basic", "Basic_Configure_Default.xlsx"));
            SettingFiles.BinCutInstanceNamingRule = ResolveConfigPath(SettingFiles.BinCutInstanceNamingRule, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "SCGH", "BinCutInstanceNamingRule_{0}.xlsx"), Path.Combine("Settings", "SCGH", "BinCutInstanceNamingRule_Default.xlsx"));
            SettingFiles.RtosCategoryConfig = ResolveConfigPath(SettingFiles.RtosCategoryConfig, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "Basic", "RtosCategory_{0}.xlsx"), Path.Combine("Settings", "Basic", "RtosCategory_Default.xlsx"));
            SettingFiles.ScanConfig = ResolveConfigPath(SettingFiles.ScanConfig, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "SCGH", "Scan_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "Scan_Config_Default.xml"));
            SettingFiles.MbistConfig = ResolveConfigPath(SettingFiles.MbistConfig, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "SCGH", "Mbist_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "Mbist_Config_Default.xml"));
            SettingFiles.SpiConfig = ResolveConfigPath(SettingFiles.SpiConfig, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "SCGH", "Spi_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "Spi_Config_Default.xml"));
            SettingFiles.HardipConfig = ResolveConfigPath(SettingFiles.HardipConfig, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "SCGH", "HardIP_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "HardIP_Config_Default.xml"));
            SettingFiles.TnAssignment = ResolveConfigPath(SettingFiles.TnAssignment, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "Basic", "TN_Assignment_{0}.xlsx"), Path.Combine("Settings", "Basic", "TN_Assignment_Default.xlsx"));
            ConfigFiles.BinNumberConfig = ResolveConfigPath(ConfigFiles.BinNumberConfig, repoProjectRepoPath, repoProjectName, Path.Combine("Config", "BinNumberConfig_{0}.xlsx"), Path.Combine("Config", "BinNumberConfig.xlsx"));
            needSheet = ResolveConfigPath(needSheet, repoProjectRepoPath, repoProjectName, Path.Combine("Settings", "Basic", "NeedSheetConfig_{0}.xml"), Path.Combine("Settings", "Basic", "NeedSheetConfig_Default.xml"));

            NeededSheets.InitSheetName(LocalSpecs.Options.Device, RepoFolderPath, CurrentProject, needSheet);
        }

        private void PrepareProjectSettings()
        {
            LocalSpecs.SearchValidPatternRev = SearchValidPatternRev;
            if (LocalSpecs.SearchValidPatternRev)
            {
                LocalSpecs.Options.IsIgnorePatternCheck = false;
            }
            LocalSpecs.CurrentProject = CurrentProject;
            LocalSpecs.TarFolder = string.IsNullOrEmpty(OutputFolder) ? Directory.GetCurrentDirectory() : Path.GetDirectoryName(OutputFolder);
            LocalSpecs.TestProgramName = string.IsNullOrEmpty(TestProgramName) ? CurrentProject : TestProgramName;
        }

        private void PrepareTestPlanFile(List<string> selectedFiles)
        {
            if (File.Exists(TestPlanFile))
            {
                LocalSpecs.TestPlanFileName = string.IsNullOrEmpty(TestPlanFile) ? LocalSpecs.TestPlanFileName : TestPlanFile;
            }
            else if (Directory.Exists(TestPlanFile))
            {
                string dir = Path.GetDirectoryName(TestPlanFile);
                string name = Path.GetFileNameWithoutExtension(TestPlanFile);
                string testPlan = Path.Combine(dir, $"{CurrentProject}_{name}.xlsx");
                var files = Directory.GetFiles(TestPlanFile, "*.*", SearchOption.AllDirectories).Where(f => f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)).ToList();
                FileManager.MergeTestPlanCsv(files, testPlan);
                LocalSpecs.TestPlanFileName = testPlan;
            }
            if (!string.IsNullOrEmpty(LocalSpecs.TestPlanFileName) && !LocalSpecs.TestPlanFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.TestPlanFileName);
            }
            LocalSpecs.EfuseTestPlanFileName = string.IsNullOrEmpty(EfuseTestPlan) ? LocalSpecs.EfuseTestPlanFileName : EfuseTestPlan;
            if (!string.IsNullOrEmpty(LocalSpecs.TestPlanFileName))
            {
                var package = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName));
                EpWorkbook.TestPlanWorkbook = package.Workbook;
                if (!string.IsNullOrEmpty(LocalSpecs.EfuseTestPlanFileName) && LocalSpecs.EfuseTestPlanFileName != "N/A")
                {
                    ExcelWorkbook efuseTestPlan = new ExcelPackage(new FileInfo(LocalSpecs.EfuseTestPlanFileName)).Workbook;
                    foreach (ExcelWorksheet sheet in efuseTestPlan.Worksheets)
                    {
                        if (EpWorkbook.TestPlanWorkbook.Worksheets.Count(x => x.Name.Equals(sheet.Name)) > 0)
                        {
                            EpWorkbook.TestPlanWorkbook.Worksheets.Delete(sheet.Name);
                        }
                        EpWorkbook.TestPlanWorkbook.Worksheets.Add(sheet.Name, sheet);
                    }
                }
            }
        }

        private void CompilePatternListIfPresent(List<string> selectedFiles)
        {
            LocalSpecs.TimeSetFolder = string.IsNullOrEmpty(TimeSetFolder) ? "" : TimeSetFolder;
            LocalSpecs.PatternFolder = string.IsNullOrEmpty(PatternFolder) ? "" : PatternFolder;
            LocalSpecs.CompilePatFileName = string.IsNullOrEmpty(CompilePat) ? "" : CompilePat;

            LocalSpecs.PatternListCsvFileName = string.IsNullOrEmpty(PatternListCsvFile) ? LocalSpecs.PatternListCsvFileName : PatternListCsvFile;
            if (!string.IsNullOrEmpty(LocalSpecs.PatternListCsvFileName) && !LocalSpecs.PatternListCsvFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.PatternListCsvFileName);
            }

            AllSingleton.Initialize();

            bool isUseTheLatestPatVersion = !LocalSpecs.Options.IsIgnorePatternCheck && LocalSpecs.SearchValidPatternRev;
            if (LocalSpecs.PatternListCsvFileName != "N/A")
            {
                var fileInfo = new FileInfo(LocalSpecs.PatternListCsvFileName);
                var patListCsv = new InputPatternListCsv(fileInfo);
                patListCsv.Compile(LocalSpecs.CompilePatFileName, LocalSpecs.PatternFolder, LocalSpecs.TimeSetFolder, isUseTheLatestPatVersion, out Dictionary<string, CompileItem> dicCompileItem, null);
                if (dicCompileItem != null)
                {
                    LocalSpecs.CompileItem = dicCompileItem;
                }
                LocalSpecs.PatternListCsvFileName = patListCsv.FullName;
            }

        }

        private void PrepareReportsAndScghInputs(List<string> selectedFiles)
        {
            LocalSpecs.BinOutReportFileName = BinOutReportFileName;
            if (!string.IsNullOrEmpty(LocalSpecs.BinOutReportFileName) && !LocalSpecs.BinOutReportFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinOutReportFileName);
            }
            LocalSpecs.CharPlanFileName = CharFileName;
            if (!string.IsNullOrEmpty(LocalSpecs.CharPlanFileName) && !LocalSpecs.CharPlanFileName.Equals(@"N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.CharPlanFileName);
            }
            LocalSpecs.ScghFileName = string.IsNullOrEmpty(ScghFile) ? LocalSpecs.ScghFileName : ScghFile;
            if (!string.IsNullOrEmpty(LocalSpecs.ScghFileName) && !LocalSpecs.ScghFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.ScghFileName);
            }
        }

        private void PrepareBinCutInputs(List<string> selectedFiles)
        {
            if (!string.IsNullOrEmpty(BinCutFileNames))
            {
                LocalSpecs.AllBinCutFileNames = new List<string>();
                LocalSpecs.BinCutFileName = "";
                foreach (string filePath in BinCutFileNames.Split(',').ToList())
                {
                    LocalSpecs.AllBinCutFileNames.Add(filePath);
                }
            }
            foreach (string binCutFileName in LocalSpecs.AllBinCutFileNames)
            {
                string shadowStage = Path.GetFileNameWithoutExtension(binCutFileName).Split('_').Last();
                if (LocalSpecs.AllJobs.Contains(shadowStage))
                {
                    LocalSpecs.BinCutShadowFileNames.Add(binCutFileName);
                }
                else
                {
                    if (string.IsNullOrEmpty(LocalSpecs.BinCutFileName) || LocalSpecs.BinCutFileName.Equals("N/A"))
                    {
                        LocalSpecs.BinCutFileName = binCutFileName;
                    }
                    else
                    {
                        throw new Exception("More than one normal BinCut file has been selected!");
                    }
                }
            }
            LocalSpecs.BinCutPostFileName = string.IsNullOrEmpty(BinCutPostFileName) ? LocalSpecs.BinCutPostFileName : BinCutPostFileName;
            if (!string.IsNullOrEmpty(LocalSpecs.BinCutPostFileName) && !LocalSpecs.BinCutPostFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinCutPostFileName);
            }
            LocalSpecs.BinCutModeSeqFileName = string.IsNullOrEmpty(BinCutModeSeqFileName) ? LocalSpecs.BinCutModeSeqFileName : BinCutModeSeqFileName;
            if (!string.IsNullOrEmpty(LocalSpecs.BinCutModeSeqFileName) && !LocalSpecs.BinCutModeSeqFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinCutModeSeqFileName);
            }
        }

        private void PrepareVoltageAndTtrInputs(List<string> selectedFiles)
        {
            LocalSpecs.EquationVoltagesFileName = string.IsNullOrEmpty(EquationVoltagesFileName) ? LocalSpecs.EquationVoltagesFileName : EquationVoltagesFileName;
            if (!string.IsNullOrEmpty(LocalSpecs.EquationVoltagesFileName) && !LocalSpecs.EquationVoltagesFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.EquationVoltagesFileName);
            }

            List<string> ttrSumFiles = TtrSummaryFileNames;
            if (ttrSumFiles.Count == 1)
            {
                LocalSpecs.TtrSummaryFileName = ttrSumFiles[0];
            }
            else if (ttrSumFiles.Count > 1)
            {
                using (var package = new ExcelPackage(new FileInfo(ttrSumFiles[0])))
                {
                    for (int i = 1; i < ttrSumFiles.Count; i++)
                    {
                        ExcelWorkbook ttrFile = new ExcelPackage(new FileInfo(ttrSumFiles[i])).Workbook;
                        foreach (ExcelWorksheet sheet in ttrFile.Worksheets)
                        {
                            if (package.Workbook.Worksheets.Count(x => x.Name.Equals(sheet.Name)) == 0)
                            {
                                package.Workbook.Worksheets.Add(sheet.Name, sheet);
                            }
                        }
                    }
                    string mergeTtr = ttrSumFiles[0].Replace(".xlsx", "_" + TimeContext.Now.ToString("yyMMddHHmm") + "_Merge.xlsx");
                    package.SaveAs(new FileInfo(mergeTtr));
                    LocalSpecs.TtrSummaryFileName = mergeTtr;
                }
            }

            if (!string.IsNullOrEmpty(LocalSpecs.TtrSummaryFileName) && !LocalSpecs.TtrSummaryFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.TtrSummaryFileName);
            }
        }

        private void PreparePowerBinningInputs(List<string> selectedFiles)
        {
            if (string.IsNullOrEmpty(PowerBinningFileNames))
            {
                return;
            }
            LocalSpecs.AllPowerBinningFileName = new List<string>();
            foreach (string filePath in PowerBinningFileNames.Split(',').ToList())
            {
                LocalSpecs.AllPowerBinningFileName.Add(filePath);
                selectedFiles.Add(filePath);
            }
        }

        private void PrepareEfuseTestPlanAndBinCutFileName(List<string> selectedFiles)
        {
            if (!string.IsNullOrEmpty(LocalSpecs.EfuseTestPlanFileName) && !LocalSpecs.EfuseTestPlanFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.EfuseTestPlanFileName);
            }
            if (!string.IsNullOrEmpty(LocalSpecs.BinCutFileName) && !LocalSpecs.BinCutFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinCutFileName);
            }

            if (LocalSpecs.BinCutShadowFileNames.Any())
            {
                selectedFiles.AddRange(LocalSpecs.BinCutShadowFileNames);
            }
        }

        private void PrepareLibraryFolder()
        {
            LocalSpecs.CsLibraryFolder = string.IsNullOrEmpty(CsLibraryPath) ? "" : CsLibraryPath;
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                if (string.IsNullOrEmpty(BasLibraryFolder))
                {
                    LocalSpecs.BasLibraryFolder = Path.Combine(Directory.GetCurrentDirectory(), "Settings", "DefaultVbtForCs");
                }
                else
                {
                    LocalSpecs.BasLibraryFolder = BasLibraryFolder;
                }
            }
            else
            {
                LocalSpecs.BasLibraryFolder = string.IsNullOrEmpty(BasLibraryFolder) ? "" : BasLibraryFolder;
            }
        }

        private void ApplyMiscRuntimeOptions()
        {
            LocalSpecs.HardIpInfoFileName = string.IsNullOrEmpty(HardIpInfoFileName) ? "" : HardIpInfoFileName;
            LocalSpecs.MbistInfoFileName = string.IsNullOrEmpty(MbistInfo) ? "" : MbistInfo;
            LocalSpecs.IsUfp = LocalSpecs.Options.IsUfp;

            LocalSpecs.DefaultChannelMap = string.IsNullOrEmpty(ChannelMap) ? "" : ChannelMap;
            if (!string.IsNullOrEmpty(CustomPath))
            {
                LocalSpecs.CustomPath = [.. CustomPath.Split(',').ToList()];
            }
        }

        private void PrepareSettingsAndConfigFiles()
        {
            LocalSpecs.SettingFiles = new SettingFiles
            {
                BasicConfigFile = string.IsNullOrEmpty(SettingFiles.BasicConfigFile) ? "" : SettingFiles.BasicConfigFile,
                BinCutInstanceNamingRule = string.IsNullOrEmpty(SettingFiles.BinCutInstanceNamingRule) ? "" : SettingFiles.BinCutInstanceNamingRule,
                RtosCategoryConfig = string.IsNullOrEmpty(SettingFiles.RtosCategoryConfig) ? "" : SettingFiles.RtosCategoryConfig,
                ScanConfig = string.IsNullOrEmpty(SettingFiles.ScanConfig) ? "" : SettingFiles.ScanConfig,
                MbistConfig = string.IsNullOrEmpty(SettingFiles.MbistConfig) ? "" : SettingFiles.MbistConfig,
                SpiConfig = string.IsNullOrEmpty(SettingFiles.SpiConfig) ? "" : SettingFiles.SpiConfig,
                HardipConfig = string.IsNullOrEmpty(SettingFiles.HardipConfig) ? "" : SettingFiles.HardipConfig,
                TnAssignment = string.IsNullOrEmpty(SettingFiles.TnAssignment) ? "" : SettingFiles.TnAssignment,
            };
            if (!string.IsNullOrEmpty(RepoFolderPath))
            {
                LocalSpecs.SettingFolder = RepoFolderPath;
            }

            LocalSpecs.ConfigFiles = new ConfigFiles
            {
                BinNumberConfig = string.IsNullOrEmpty(ConfigFiles.BinNumberConfig) ? "" : ConfigFiles.BinNumberConfig,
            };
        }

        private void StartSetAutomation(List<string> selectedFiles)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string version = AssemblyProvider.Current.GetFileVersion(FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);
            LocalSpecs.AutogenVer = version;
            LocalSpecs.SelectedFiles.AddRange(selectedFiles);
            List<string> tpSheetsList = EpWorkbook.TestPlanWorkbook.GetPlanSheets(".*");
            List<string> scghSheetsList = EpWorkbook.ScghWorkbook == null ? new List<string>() : EpWorkbook.ScghWorkbook.GetPlanSheets(".*");
            List<string> binCuts = EpWorkbook.BinCutWorkbook == null ? new List<string>() : EpWorkbook.BinCutWorkbook.GetPlanSheets(".*");
            List<string> postBinCuts = EpWorkbook.BinCutPostWorkbook == null ? new List<string>() : EpWorkbook.BinCutPostWorkbook.GetPlanSheets(".*");
            BlockStatus.SetAutomation(tpSheetsList, scghSheetsList, LocalSpecs.PatternListCsvFileName, binCuts, postBinCuts, null, null);

            FolderStructure.CreateFolder();
        }

        public void Initialize()
        {
            ResolveConfigPaths();
            PrepareProjectSettings();

            var selectedFiles = new List<string>();
            PrepareTestPlanFile(selectedFiles);
            CompilePatternListIfPresent(selectedFiles);
            PrepareReportsAndScghInputs(selectedFiles);
            PrepareBinCutInputs(selectedFiles);
            PrepareVoltageAndTtrInputs(selectedFiles);
            PreparePowerBinningInputs(selectedFiles);
            PrepareEfuseTestPlanAndBinCutFileName(selectedFiles);
            PrepareLibraryFolder();
            ApplyMiscRuntimeOptions();
            PrepareSettingsAndConfigFiles();

            LocalSpecs.FuseCheckFileName = string.IsNullOrEmpty(FuseCheckTable) ? LocalSpecs.FuseCheckFileName : FuseCheckTable;
            LocalSpecs.DramTypeFileName = string.IsNullOrEmpty(DramTypeFileName) ? LocalSpecs.DramTypeFileName : DramTypeFileName;

            if (VoltageTableFiles != null)
            {
                LocalSpecs.VoltageTbFileName = new List<string>();
                foreach (string sheet in VoltageTableFiles)
                {
                    LocalSpecs.VoltageTbFileName.Add(sheet);
                }
            }
            selectedFiles.AddRange(LocalSpecs.VoltageTbFileName);
            if (LocalSpecs.VoltageTbFileName.Any())
            {
                MergeSheet.ParseTestSettingSheetToTestPlan(LocalSpecs.VoltageTbFileName, EpWorkbook.TestPlanWorkbook);
            }

            StartSetAutomation(selectedFiles);
        }

        public bool SearchValidPatternRev { get; set; }

        public static string ResolveConfigPath(string currentValue, string repoFolder, string projectName, string projectPattern, string defaultRelative)
        {
            if (!string.IsNullOrEmpty(currentValue))
            {
                return currentValue;
            }

            string projectFile = Path.Combine(repoFolder, string.Format(projectPattern, projectName));
            if (File.Exists(projectFile))
            {
                return projectFile;
            }

            string projectFile1 = Path.Combine(Directory.GetCurrentDirectory(), string.Format(projectPattern, projectName));
            if (File.Exists(projectFile1))
            {
                return projectFile1;
            }

            string defaultFile = Path.Combine(Directory.GetCurrentDirectory(), defaultRelative);
            return defaultFile;
        }
    }
}
