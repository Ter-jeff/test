using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Enums;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.DataStruct;
using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.HardIpDc.Business;
using TestPlanLib.PowerMerge;
using TestPlanLib.Static;

namespace Automation.Singleton
{
    public class MultiTestSettingSheetsSingleton
    {
        public const string ValtRowPinNameFlag = "_Valt";
        private static readonly Regex _regTest = new Regex("tdchain|sachain|td|sa|mbist|hardip", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static volatile MultiTestSettingSheetsSingleton _instance;
        private static readonly object _lock = new object();
        private static List<string> _hardIpDcLevelNames;
        private static List<string> _lastReadJoblst;

        private static ExcelWorkbook _testplanWorkbook;
        private static PowerInfoSheet _powerInfo;
        private static string _currentProject;
        private static List<string> _jobs;
        private static EnumDevice _device;

        private List<string> _performanceModeList;

        public List<string> PowerPinList { get; }
        public List<string> VrsPowerPinList { get; }
        public List<string> PerformanceModeList
        {
            get
            {
                if (_performanceModeList == null)
                {
                    IEnumerable<DcCategoryInfo> temp = DcCategoryInfos.Where(x => _regTest.IsMatch(x.Test) && Regex.IsMatch(x.PmodePatternVdip, "^(?!Mbist)M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2}$", RegexOptions.IgnoreCase));
                    _performanceModeList = temp.Select(x => x.PmodePatternVdip).ToList();
                }
                return _performanceModeList;
            }
        }
        public List<TestSettingData> TestSettingSheetsList { get; set; }
        public DcCategoryInfos DcCategoryInfos { get; set; } = new DcCategoryInfos();

        #region Constructor
        //mock
        public MultiTestSettingSheetsSingleton()
        {
        }

        public static MultiTestSettingSheetsSingleton Instance()
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = CreateInstance(_jobs, TestPlanStatic.IoInfoSheet, TestPlanStatic.HardIpDcSheet);
                    }
                    else
                    {
                        UpdateInstance(_jobs);
                    }
                }
            }
            else
            {
                UpdateInstance(_jobs);
            }

            return _instance;
        }

        private MultiTestSettingSheetsSingleton(List<string> jobs, List<Tuple<string, string>> dcCategories = null)
        {
            TestSettingSheetsList = null;
            _performanceModeList = null;

            _hardIpDcLevelNames = new List<string>();

            if (jobs == null || jobs.Count == 0)
            {
                TestSettingSheetsList = new TestSettingReader(TestProgram.IgxlWorkBk.PinMapPair.Value).ReadFlow(_currentProject, _testplanWorkbook, new List<string>());
            }
            else
            {
                TestSettingSheetsList = new TestSettingReader(TestProgram.IgxlWorkBk.PinMapPair.Value).ReadFlowByAllJobs(_currentProject, _testplanWorkbook, jobs, dcCategories);
            }

            _lastReadJoblst = jobs;
            if (_lastReadJoblst.Count == 0)
            {
                _lastReadJoblst.Add("CP1");
            }

            if (TestSettingSheetsList == null)
            {
                DcCategoryInfos = new DcCategoryInfos();
            }
            else
            {
                var categoryInfoList = new DcCategoryInfos();
                foreach (TestSettingData testSetting in TestSettingSheetsList)
                {
                    foreach (DcCategoryName category in testSetting.DcCategorys)
                    {
                        if (!categoryInfoList.Exists(s => s.CategoryName.Equals(category.CategoryName, StringComparison.OrdinalIgnoreCase)))
                        {
                            categoryInfoList.Add(category.DcCategoryInfo);
                        }
                    }
                }

                DcCategoryInfos = categoryInfoList;

                if (PowerPinList == null)
                {
                    PowerPinList = new List<string>();
                    foreach (TestSettingData data in TestSettingSheetsList)
                    {
                        PowerPinList.AddRange(data.GetPowerPins());
                    }
                    PowerPinList = PowerPinList.Distinct().ToList();
                }

                if (VrsPowerPinList == null)
                {
                    VrsPowerPinList = new List<string>();
                    foreach (TestSettingData data in TestSettingSheetsList)
                    {
                        VrsPowerPinList.AddRange(data.GetValtPowerPins());
                    }
                    VrsPowerPinList = VrsPowerPinList.Distinct().ToList();
                }
            }
        }

        private static MultiTestSettingSheetsSingleton CreateInstance(List<string> jobs, IoInfoSheet ioInfo, HardIpDcSheet hardIpDc)
        {
            _instance = new MultiTestSettingSheetsSingleton(jobs, null);

            if (hardIpDc != null)
            {
                var wraper = new HardIpDcWrapper(hardIpDc, _instance.PowerPinList, ioInfo);
                hardIpDc = wraper.WrapHardIpDc();

                foreach (HardIpCategoryDef hipDcCategory in hardIpDc.Rows)
                {
                    _hardIpDcLevelNames.Add(hipDcCategory.GetLevelName());
                }
            }

            IoLevelsSheet ioLevelsSheet = _device == EnumDevice.LCD
                ? new IoLevelsSheetReader().ReadSheet(_testplanWorkbook.Worksheets[NeededSheets.IoLevels])
                : null;
            PowerMergeSheet powerMergeTbl = PowerMergeReader.ReadSheet(_testplanWorkbook.Worksheets[NeededSheets.PowerMerge]);
            var dataParser = new TestSettingDataParser(_testplanWorkbook, _instance, powerMergeTbl, hardIpDc, ioLevelsSheet);
            _instance = dataParser.EditPrintData();

            return _instance;
        }

        private static void UpdateInstance(List<string> jobs)
        {
            List<string> currentJoblist = jobs;
            bool allInvolved = true;
            foreach (string s in currentJoblist)
            {
                if (!_lastReadJoblst.Exists(x => x.Equals(s, StringComparison.OrdinalIgnoreCase)))
                {
                    allInvolved = false;
                }
            }

            if (!allInvolved)
            {
                _instance = CreateInstance(currentJoblist, TestPlanStatic.IoInfoSheet, TestPlanStatic.HardIpDcSheet);
            }
        }

        public static void Initialize(ExcelWorkbook testPlanWorkbook, string currentProject, List<string> jobs, EnumDevice device)
        {
            _instance = null;
            _lastReadJoblst = null;
            _testplanWorkbook = testPlanWorkbook;
            _currentProject = currentProject;
            _jobs = jobs;
            _device = device;
            _powerInfo = TestPlanStatic.PowerInfoSheet;
        }
        #endregion

        public string FindRtosCategory(string performanceMode, out EnumMessageLevel msgLevel, out string errorMsg)
        {
            return DcCategoryInfos.FindRtosCategory(_powerInfo, performanceMode, out msgLevel, out errorMsg);
        }

        //virtual for mock
        public virtual string FindMbistCatgeoryName(string sheetDomain, string type, string performanceMode, List<string> patterns, out EnumMessageLevel msgLevel, out string errorMsg, string chiplet = "", string pmodeDomain = "")
        {
            return DcCategoryInfos.FindMbistCatgeoryName(_powerInfo, sheetDomain, type, performanceMode, patterns, out msgLevel, out errorMsg, chiplet, pmodeDomain);
        }

        public string FindScanCategoryName(string test, string sheetDomain, string performanceMode, List<string> patterns, out EnumMessageLevel msgLevel, out string errorMsg, string pmodeDomain = "", string chiplet = "")
        {
            return DcCategoryInfos.FindScanCategoryName(_powerInfo, test, sheetDomain, performanceMode, patterns, out msgLevel, out errorMsg, pmodeDomain, chiplet);
        }

        //virtual for mock
        public virtual List<string> GetDcCategoryVoltages(string dcCategoryName)
        {
            var voltagelst = new List<string>();
            if (!DcCategoryInfos.Exists(s => s.CategoryName.Equals(dcCategoryName, StringComparison.OrdinalIgnoreCase)))
            {
                return voltagelst;
            }

            DcCategoryValue category = TestSettingSheetsList[0].DataRows[0].DcCategoryValues.Find(s => s.CategoryName.Equals(dcCategoryName, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(category.Nv.OriginValue))
            {
                voltagelst.Add("NV");
            }

            if (!string.IsNullOrEmpty(category.Hv.OriginValue))
            {
                voltagelst.Add("HV");
            }

            if (!string.IsNullOrEmpty(category.Lv.OriginValue))
            {
                voltagelst.Add("LV");
            }

            return voltagelst;
        }
    }
}
