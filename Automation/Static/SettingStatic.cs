using System.IO;

using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Reader.ConfigFile.NamingRule.Business;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Settings;

namespace Automation.Static
{
    public static class SettingStatic
    {
        private static ScanConfig _scanConfig;
        public static ScanConfig ScanConfig
        {
            get
            {
                if (_scanConfig == null)
                {
                    var configReader = new ScanConfigFileReader();
                    _scanConfig = configReader.ReadConfig();
                }

                return _scanConfig;
            }
        }

        private static BinCutInstanceNamingSheet _binCutInstanceNamingSheet;
        public static BinCutInstanceNamingSheet BinCutInstanceNamingSheet
        {
            get
            {
                if (_binCutInstanceNamingSheet == null)
                {
                    string config = LocalSpecs.SettingFiles.BinCutInstanceNamingRule;
                    if (string.IsNullOrEmpty(config) || !File.Exists(config))
                    {
                        config = Path.Combine(LocalSpecs.SettingFolder, "Settings", "SCGH", $"BinCutInstanceNamingRule_{LocalSpecs.CurrentProject}.xlsx");
                        if (!File.Exists(config))
                        {
                            config = Path.Combine(LocalSpecs.SettingFolder, "Settings", "SCGH", "BinCutInstanceNamingRule_Default.xlsx");
                        }
                    }
                    if (File.Exists(config))
                    {
                        var excelPackage = new ExcelPackage(new FileInfo(config));
                        ExcelWorksheet sheet = excelPackage.Workbook.Worksheets["BinCutInstanceNamingRule"];
                        if (sheet != null)
                        {
                            var binCutInstanceNamingRuleReader = new BinCutInstanceNamingRuleReader();
                            _binCutInstanceNamingSheet = binCutInstanceNamingRuleReader.ReadSheet(sheet);
                        }
                    }
                }

                return _binCutInstanceNamingSheet;
            }
        }

        private static ExcelWorkbook _basicConfigWorkbook;
        public static ExcelWorkbook BasicConfigWorkbook
        {
            get
            {
                if (_basicConfigWorkbook == null)
                {
                    if (!string.IsNullOrEmpty(LocalSpecs.SettingFiles.BasicConfigFile) && File.Exists(LocalSpecs.SettingFiles.BasicConfigFile))
                    {
                        var package = new ExcelPackage(new FileInfo(LocalSpecs.SettingFiles.BasicConfigFile));
                        _basicConfigWorkbook = package.Workbook;
                    }
                    else
                    {
                        string basicConfig = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic", $"Basic_Configure_{LocalSpecs.CurrentProject}.xlsx");
                        if (!File.Exists(basicConfig))
                        {
                            basicConfig = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic", "Basic_Configure_Default.xlsx");
                        }
                        if (File.Exists(basicConfig))
                        {
                            var inputExcel = new ExcelPackage(new FileInfo(basicConfig));
                            _basicConfigWorkbook = inputExcel.Workbook;
                        }
                    }
                }

                return _basicConfigWorkbook;
            }
            set => _basicConfigWorkbook = value;
        }

        private static ExcelWorkbook _rtosCategoryConfigWorkbook;
        public static ExcelWorkbook RtosCategoryConfigWorkbook
        {
            get
            {
                if (_rtosCategoryConfigWorkbook == null)
                {
                    string rtosConfig = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic", $"RtosCategory_{LocalSpecs.CurrentProject}.xlsx");
                    if (!string.IsNullOrEmpty(LocalSpecs.SettingFiles.RtosCategoryConfig))
                    {
                        rtosConfig = LocalSpecs.SettingFiles.RtosCategoryConfig;
                    }

                    if (!File.Exists(rtosConfig))
                    {
                        rtosConfig = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic", "RtosCategory_Default.xlsx");
                    }
                    if (File.Exists(rtosConfig))
                    {
                        var inputExcel = new ExcelPackage(new FileInfo(rtosConfig));
                        _rtosCategoryConfigWorkbook = inputExcel.Workbook;
                    }
                }

                return _rtosCategoryConfigWorkbook;
            }
            set => _rtosCategoryConfigWorkbook = value;
        }

        private static JobMapSheet _jobMapSheet;
        public static JobMapSheet JobMapSheet
        {
            get
            {
                if (_jobMapSheet == null && BasicConfigWorkbook != null)
                {
                    foreach (ExcelWorksheet sheet in BasicConfigWorkbook.Worksheets)
                    {
                        string sheetName = sheet.Name;
                        if (sheetName.ContainsIgnoreCase("JobMapping"))
                        {
                            var jobMappingReader = new JobMapReader();
                            _jobMapSheet = jobMappingReader.ReadSheet(sheet);
                        }
                    }
                }

                return _jobMapSheet;
            }
        }

        public static object InstrumentSheet { get; internal set; }

        public static void Clear()
        {
            _scanConfig = null;
            _binCutInstanceNamingSheet = null;
            _basicConfigWorkbook = null;
            _rtosCategoryConfigWorkbook = null;
            _jobMapSheet = null;
        }
    }
}
