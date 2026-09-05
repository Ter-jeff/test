using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.GenerateIgxl.PostAction.Relay;
using Automation.GenerateIgxl.PostAction.TempMon.Data;
using Automation.Static.Result;

using TestPlanLib.Basic;
using TestPlanLib.PatternListCsvFile;

namespace Automation.Static
{
    public static class LocalSpecs
    {
        private static string _pwrRes;
        private static string _tarFolder;
        private static string _testPlanFileName;
        private static string _patListCsvFileName;
        private static string _efuseTestPlanFileName;
        private static string _compilePatFileName;
        private static string _mbistInfoFileName;
        private static string _scghFileName;
        private static string _ttrSummaryName;
        private static List<string> _customPath;
        private static string _binCutFileName;
        private static List<string> _binCutShadowFileNames;
        private static string _binCutPostFileName;
        private static string _dramTypeFileName;
        private static string _fuseCheckFileName;
        private static string _binCutModeSeqFileName;
        private static string _eqnVoltagesFileName;
        private static List<string> _voltageTableFileNames;
        private static string _oriTestPlanFileName;
        private static string _oriScghFileName;
        private static string _hardipPatInfo;
        private static string _timeSetFolder;
        private static string _patternFolder = string.Empty;
        private static string _settingFolder;
        private static string _projectIniFileName;
        private static string _binOutReportFileName;
        private static string _charPlanFileName;

        private static List<string> _allBinCutFileNames;
        private static List<string> _allPowerBinningFileName;
        private static string _autogenVer;
        private static Dictionary<string, List<string>> _jobMap;
        private static List<JobTemperatureMap> _jobTemperatureMaps;
        private static List<string> _selectedFiles;

        #region user input file
        public static string CppDspControl { get; set; }
        public static string PinMap { get; set; }
        public static string TestPlanFileName
        {
            get => _testPlanFileName ?? (_testPlanFileName = "N/A");
            set => _testPlanFileName = value;
        }
        public static string ScghFileName
        {
            get => _scghFileName ?? (_scghFileName = "N/A");
            set => _scghFileName = value;
        }
        public static string BinCutFileName
        {
            get => _binCutFileName ?? (_binCutFileName = "N/A");
            set => _binCutFileName = value;
        }
        public static List<string> BinCutShadowFileNames
        {
            get => _binCutShadowFileNames ?? (_binCutShadowFileNames = new List<string>());
            set => _binCutShadowFileNames = value;
        }
        public static string BinCutModeSeqFileName
        {
            get => _binCutModeSeqFileName ?? (_binCutModeSeqFileName = "N/A");
            set => _binCutModeSeqFileName = value;
        }
        public static string BinCutPostFileName
        {
            get => _binCutPostFileName ?? (_binCutPostFileName = "N/A");
            set => _binCutPostFileName = value;
        }
        public static List<string> AllBinCutFileNames
        {
            get
            {
                return _allBinCutFileNames ?? (_allBinCutFileNames = new List<string>());
            }
            set { _allBinCutFileNames = value; }
        }
        public static List<string> AllPowerBinningFileName
        {
            get
            {
                return _allPowerBinningFileName ?? (_allPowerBinningFileName = new List<string>());
            }
            set { _allPowerBinningFileName = value; }
        }
        public static string EquationVoltagesFileName
        {
            get => _eqnVoltagesFileName ?? (_eqnVoltagesFileName = "N/A");
            set => _eqnVoltagesFileName = value;
        }
        public static string EfuseTestPlanFileName
        {
            get => _efuseTestPlanFileName ?? (_efuseTestPlanFileName = "N/A");
            set => _efuseTestPlanFileName = value;
        }
        public static List<string> VoltageTbFileName
        {
            get => _voltageTableFileNames ?? (_voltageTableFileNames = new List<string>());
            set => _voltageTableFileNames = value;
        }
        public static string DramTypeFileName
        {
            get => _dramTypeFileName ?? (_dramTypeFileName = "N/A");
            set => _dramTypeFileName = value;
        }
        public static string FuseCheckFileName
        {
            get => _fuseCheckFileName ?? (_fuseCheckFileName = "N/A");
            set => _fuseCheckFileName = value;
        }
        public static string TtrSummaryFileName
        {
            get => _ttrSummaryName ?? (_ttrSummaryName = "N/A");
            set => _ttrSummaryName = value;
        }
        public static List<string> CustomPath
        {
            get => _customPath ?? (_customPath = new List<string>());
            set => _customPath = value;
        }
        public static string PatternListCsvFileName
        {
            get => _patListCsvFileName ?? (_patListCsvFileName = "N/A");
            set => _patListCsvFileName = value;
        }
        public static string CompilePatFileName
        {
            get
            {
                return _compilePatFileName ?? (_compilePatFileName = "");
            }
            set { _compilePatFileName = value; }
        }
        public static string HardIpInfoFileName
        {
            get
            {
                if (_hardipPatInfo == null)
                {
                    return "";
                }
                return _hardipPatInfo;
            }
            set => _hardipPatInfo = value;
        }
        public static string MbistInfoFileName
        {
            get => _mbistInfoFileName ?? (_mbistInfoFileName = "");
            set => _mbistInfoFileName = value;
        }
        public static string BinOutReportFileName
        {
            get => string.IsNullOrEmpty(_binOutReportFileName) ? _binOutReportFileName = "N/A" : _binOutReportFileName;
            set => _binOutReportFileName = value;
        }
        public static string CharPlanFileName
        {
            get => _charPlanFileName ?? (_charPlanFileName = "N/A");
            set => _charPlanFileName = value;
        }
        public static string OriTestPlanFileName
        {
            get => _oriTestPlanFileName ?? (_oriTestPlanFileName = "N/A");
            set => _oriTestPlanFileName = value;
        }
        public static string OriScghFileName
        {
            get => _oriScghFileName ?? (_oriScghFileName = "N/A");
            set => _oriScghFileName = value;
        }
        #endregion

        #region user input folder
        public static string TarFolder
        {
            get
            {
                if (_tarFolder == null)
                {
                    throw new Exception("Source directory is empty!");
                }
                return _tarFolder;
            }
            set
            {
                _tarFolder = value;
                if (_tarFolder != null && _tarFolder.Length > 2)
                {
                    if (_tarFolder.Substring(_tarFolder.Length - 2, 2) == new string(Path.DirectorySeparatorChar, 2))
                    {
                        _tarFolder = _tarFolder.Substring(0, _tarFolder.Length - 1);
                    }
                }
            }
        }
        public static string SettingFolder
        {
            get => _settingFolder ?? (_settingFolder = Directory.GetCurrentDirectory());
            set => _settingFolder = value;
        }
        public static string TimeSetFolder
        {
            get
            {
                _timeSetFolder = _timeSetFolder.EndsWith(":") ? _timeSetFolder + Path.DirectorySeparatorChar : _timeSetFolder;
                if (_timeSetFolder == null)
                {
                    throw new Exception("TimeSet Path is empty!");
                }
                return _timeSetFolder;
            }
            set
            {
                _timeSetFolder = value;
            }
        }
        public static string PatternFolder
        {
            get
            {
                _patternFolder = _patternFolder.EndsWith(":") ? _patternFolder + Path.DirectorySeparatorChar : _patternFolder;
                if (_patternFolder == null)
                {
                    throw new Exception("Pattern path is empty");
                }

                return _patternFolder;
            }
            set => _patternFolder = value;
        }
        public static string BasLibraryFolder { get; set; }
        public static string CsLibraryFolder { get; set; }
        public static bool IsBenchLog { get; set; }
        #endregion

        #region optional
        public static string CurrentProject { get; set; } = "Default";
        public static string CurrentJob { get; set; }
        public static string BaseTestProgram { get; set; }
        public static string TestProgramName { get; set; }
        public static string DefaultChannelMap { get; set; }
        public static bool IsUnitTest { get; set; }
        public static Options Options { get; set; } = new Options();
        public static bool CheckOnly { get; set; }

        private static readonly List<string> _defaultJobs = new List<string>
        {
            "CP1",
            "CP2",
            "FT1",
            "FT2",
            "FT3"
        };
        public static List<string> AllJobsHardIp
        {
            get
            {
                List<string> allJobs;
                if (TestPlanStatic.MainFlowSheet != null)
                {
                    if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any())
                    {
                        allJobs = TestPlanStatic.TestProgramDefSheet.AllJobs;
                    }
                    else
                    {
                        IEnumerable<MainFlowBase> jobs =
                            TestPlanStatic.MainFlowSheet.Rows
                                .Where(x => !x.MainFlowName.Equals("Main_Flow_Sub", StringComparison.CurrentCultureIgnoreCase)
                                    && !x.MainFlowName.Equals("Main_Flow_T0TX_Room", StringComparison.CurrentCultureIgnoreCase)
                                        && !x.MainFlowName.Equals("Main_Flow_T0TX_Hot", StringComparison.CurrentCultureIgnoreCase)
                                            && x.SequencesNew.Any(c => c.Enable && c.Module.Equals("HardIP", StringComparison.OrdinalIgnoreCase)));
                        allJobs = jobs.Select(x => x.JobName.ToUpper()).ToList();
                    }
                }
                else
                {
                    allJobs = SettingStatic.JobMapSheet.JobMapDictionary.SelectMany(x => x.Value).ToList();
                }

                return allJobs.Count == 0 ? _defaultJobs : allJobs;
            }
        }

        public static List<string> AllJobsRtos
        {
            get
            {
                List<string> allJobs;
                if (TestPlanStatic.MainFlowSheet != null)
                {
                    if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any())
                    {
                        allJobs = TestPlanStatic.TestProgramDefSheet.AllJobs;
                    }
                    else
                    {
                        IEnumerable<MainFlowBase> jobs =
                            TestPlanStatic.MainFlowSheet.Rows
                                .Where(x => !x.MainFlowName.Equals("Main_Flow_Sub", StringComparison.CurrentCultureIgnoreCase)
                                    && !x.MainFlowName.Equals("Main_Flow_T0TX_Room", StringComparison.CurrentCultureIgnoreCase)
                                        && !x.MainFlowName.Equals("Main_Flow_T0TX_Hot", StringComparison.CurrentCultureIgnoreCase)
                                            && x.SequencesNew.Any(c => c.Enable && c.Module.Equals("RTOS", StringComparison.OrdinalIgnoreCase)));
                        allJobs = jobs.Select(x => x.JobName.ToUpper()).ToList();
                    }
                }
                else
                {
                    allJobs = SettingStatic.JobMapSheet.JobMapDictionary.SelectMany(x => x.Value).ToList();
                }

                return allJobs.Count == 0 ? _defaultJobs : allJobs;
            }
        }

        public static List<string> AllJobsIds
        {
            get
            {
                List<string> allJobs;
                if (TestPlanStatic.MainFlowSheet != null)
                {
                    if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any())
                    {
                        allJobs = TestPlanStatic.TestProgramDefSheet.AllJobs;
                    }
                    else
                    {
                        IEnumerable<MainFlowBase> jobs =
                            TestPlanStatic.MainFlowSheet.Rows
                                .Where(x => !x.MainFlowName.Equals("Main_Flow_Sub", StringComparison.CurrentCultureIgnoreCase)
                                    && !x.MainFlowName.Equals("Main_Flow_T0TX_Room", StringComparison.CurrentCultureIgnoreCase)
                                        && !x.MainFlowName.Equals("Main_Flow_T0TX_Hot", StringComparison.CurrentCultureIgnoreCase)
                                            && x.SequencesNew.Any(c => c.Enable && c.Module.Equals("IDS", StringComparison.OrdinalIgnoreCase)));
                        allJobs = jobs.Select(x => x.JobName.ToUpper()).ToList();
                    }
                }
                else
                {
                    allJobs = SettingStatic.JobMapSheet.JobMapDictionary.SelectMany(x => x.Value).ToList();
                }

                return allJobs.Count == 0 ? _defaultJobs : allJobs;
            }
        }
        public static ConfigFiles ConfigFiles = new ConfigFiles();
        public static SettingFiles SettingFiles = new SettingFiles();
        #endregion

        public static bool ExistPowerPinListDcvs { get; set; }
        public static bool ExistPowerPinListDcvi { get; set; }

        public static bool ExistIoSeqHiLo { get; set; }
        public static HardIpInfos HardIpInfos { set; get; } = new HardIpInfos();
        public static ElapsedTimeResults ElapsedTimeResults { get; set; } = new ElapsedTimeResults();
        public static Dictionary<string, CompileItem> CompileItem { get; set; }
        public static string AutogenVer
        {
            get
            {
                return _autogenVer ?? (_autogenVer = "N/A");
            }
            set { _autogenVer = value; }
        }

        public static bool SearchValidPatternRev { get; set; }

        public static bool IsUfp { get; set; }
        public static string ProjectIniFileName
        {
            get
            {
                if (!string.IsNullOrEmpty(_projectIniFileName))
                {
                    return _projectIniFileName;
                }

                return Path.Combine(SettingFolder, "Settings", "ProjectConfig_" + CurrentProject + ".ini");
            }
            set { _projectIniFileName = value; }
        }
        public static string GetProjectNameMapping
        {
            get
            {
                if (string.IsNullOrEmpty(CurrentProject))
                {
                    return "";
                }

                string[] items = CurrentProject.Split('-');
                string projectName = items[0];
                if (items.Length > 1)
                {
                    string domain = items[1];
                    string mappingDomain = "";
                    if (Regex.IsMatch(domain, "CPU*", RegexOptions.IgnoreCase))
                    {
                        mappingDomain = "cpu";
                    }
                    else if (Regex.IsMatch(domain, "GFX*", RegexOptions.IgnoreCase))
                    {
                        mappingDomain = "gpu";
                    }
                    return projectName + "_" + mappingDomain;
                }

                return projectName;
            }
        }

        public static string PwrSupplyRes
        {
            get
            {
                return _pwrRes ?? (_pwrRes = "0.001");

            }
            set { _pwrRes = value; }
        }
        public static Dictionary<string, List<string>> JobMap
        {
            get
            {
                return _jobMap ?? (_jobMap = new Dictionary<string, List<string>>());
            }
            set
            {
                if (UseGenerationQa)
                {
                    _jobMap.Add("FT", new List<string> { "FT1", "FT2" });
                    _jobMap.Add("QA", new List<string> { "QA" });
                }
                else
                {
                    _jobMap = value;
                }

            }
        }
        public static List<string> AllJobs
        {
            get
            {
                List<string> allJobs;
                if (TestPlanStatic.MainFlowSheet != null)
                {
                    if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any())
                    {
                        allJobs = new List<string>(TestPlanStatic.TestProgramDefSheet.AllJobs);
                    }
                    else
                    {
                        allJobs = new List<string>(TestPlanStatic.MainFlowSheet.Jobs);
                    }
                }
                else
                {
                    allJobs = JobMap.SelectMany(x => x.Value).ToList();
                }

                return allJobs;
            }
        }

        public static List<JobTemperatureMap> JobTemperatureMap
        {
            get
            {
                return _jobTemperatureMaps ?? (_jobTemperatureMaps = new List<JobTemperatureMap>());
            }
            set { _jobTemperatureMaps = value; }
        }
        public static List<string> SelectedFiles
        {
            get
            {
                return _selectedFiles ?? (_selectedFiles = new List<string>());
            }
            set
            {
                _selectedFiles.Add(value.ToString());
            }
        }

        #region output
        public static List<RelayItemNew> RelayItems { get; set; }
        public static List<string> FingerPrintList { set; get; }
        #endregion

        public static bool IsPatternValidate { set; get; }
        public static bool UseGenerationQa { get; set; }
        public static bool HasBkmProcess { get; set; }
        public static List<string> EfuseFlowUsedInteger { get; set; } = new List<string>();

        private static List<string> _enableModules;
        public static void SetEnableModules(List<string> enableModules)
        {
            _enableModules = enableModules;
        }

        public static bool IsModuleIncluded(string module)
        {
            return _enableModules == null || (_enableModules != null && _enableModules.Exists(x => x.Equals(module, StringComparison.OrdinalIgnoreCase)));
        }

        public static string GetPatternPath()
        {
            string path = Path.Combine(_patternFolder, CurrentProject);
            if (Directory.Exists(path))
            {
                return path;
            }

            return _patternFolder;
        }

        public static HashSet<TempMonData> TempMonDatas { get; private set; }

        public static void Clear()
        {
            // ==== Basic Flags ====
            UseGenerationQa = false;
            HasBkmProcess = false;
            IsPatternValidate = false;
            ExistPowerPinListDcvs = false;
            ExistPowerPinListDcvi = false;
            ExistIoSeqHiLo = false;
            IsUfp = false;
            SearchValidPatternRev = false;

            // ==== Configuration ====
            ConfigFiles = new ConfigFiles();
            SettingFiles = new SettingFiles();
            _enableModules = null;

            // ==== Core References ====
            CompileItem = null;
            RelayItems = null;
            HardIpInfos = new HardIpInfos();
            FingerPrintList = null;
            ElapsedTimeResults = new ElapsedTimeResults();

            // ==== File Paths ====
            _tarFolder = TarFolder = null;
            _patListCsvFileName = PatternListCsvFileName = null;
            _testPlanFileName = TestPlanFileName = null;
            _efuseTestPlanFileName = null;
            _compilePatFileName = CompilePatFileName = null;
            _mbistInfoFileName = MbistInfoFileName = null;
            _scghFileName = ScghFileName = null;
            _ttrSummaryName = TtrSummaryFileName = null;
            _binCutFileName = BinCutFileName = null;
            _binCutShadowFileNames = null;
            _binCutPostFileName = BinCutPostFileName = null;
            _dramTypeFileName = null;
            _fuseCheckFileName = null;
            _binCutModeSeqFileName = BinCutModeSeqFileName = null;
            _eqnVoltagesFileName = EquationVoltagesFileName = null;
            _oriTestPlanFileName = OriTestPlanFileName = null;
            _oriScghFileName = OriScghFileName = null;
            _timeSetFolder = TimeSetFolder = null;
            _patternFolder = PatternFolder = null;
            _settingFolder = SettingFolder = null;
            _projectIniFileName = ProjectIniFileName = null;
            _autogenVer = AutogenVer = null;
            HardIpInfoFileName = null;
            _binOutReportFileName = null;
            _customPath = null;

            // ==== Collections ====
            _allBinCutFileNames = null;
            _allPowerBinningFileName = null;
            _voltageTableFileNames = null;
            _selectedFiles = null;
            _jobMap = JobMap = null;
            _jobTemperatureMaps = JobTemperatureMap = null;
            TempMonDatas = new HashSet<TempMonData>();

            // ==== Library Paths ====
            BasLibraryFolder = null;
            CsLibraryFolder = null;

            // ==== Job Info ====
            CurrentProject = string.Empty;

            // ==== Misc ====
            _pwrRes = PwrSupplyRes = null;

            Options = new Options();

            IsUnitTest = false;

            EfuseFlowUsedInteger = new List<string>();
        }
    }
}
