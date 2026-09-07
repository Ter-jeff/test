using System.Collections.Generic;
using System.IO;
using System.Linq;

using OfficeOpenXml;

namespace Automation.Static
{
    public static class EpWorkbook
    {
        private static ExcelWorkbook _testPlanWorkbook;
        public static ExcelWorkbook TestPlanWorkbook
        {
            get
            {
                if (_testPlanWorkbook == null)
                {
                    if (!string.IsNullOrEmpty(LocalSpecs.TestPlanFileName) && LocalSpecs.TestPlanFileName != "N/A")
                    {
                        var package = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName));
                        _testPlanWorkbook = package.Workbook;
                    }
                }
                return _testPlanWorkbook;
            }
            set => _testPlanWorkbook = value;
        }
        private static ExcelWorkbook _scghWorkbook;
        public static ExcelWorkbook ScghWorkbook
        {
            get
            {
                if (_scghWorkbook == null)
                {
                    if (!string.IsNullOrEmpty(LocalSpecs.ScghFileName) && LocalSpecs.ScghFileName != "N/A")
                    {
                        var package = new ExcelPackage(new FileInfo(LocalSpecs.ScghFileName));
                        _scghWorkbook = package.Workbook;
                    }
                }
                return _scghWorkbook;
            }
            set => _scghWorkbook = value;
        }
        private static ExcelWorkbook _binCutWorkbook;
        public static ExcelWorkbook BinCutWorkbook
        {
            get
            {
                if (_binCutWorkbook == null)
                {
                    if (!string.IsNullOrEmpty(LocalSpecs.BinCutFileName) && LocalSpecs.BinCutFileName != "N/A")
                    {
                        var package = new ExcelPackage(new FileInfo(LocalSpecs.BinCutFileName));
                        _binCutWorkbook = package.Workbook;
                    }
                }
                return _binCutWorkbook;
            }
            set => _binCutWorkbook = value;
        }
        private static ExcelWorkbook _binCutModeSeqWorkbook;
        public static ExcelWorkbook BinCutModeSeqWorkbook
        {
            get
            {
                if (_binCutModeSeqWorkbook == null)
                {
                    if (!string.IsNullOrEmpty(LocalSpecs.BinCutModeSeqFileName) && LocalSpecs.BinCutModeSeqFileName != "N/A")
                    {
                        var package = new ExcelPackage(new FileInfo(LocalSpecs.BinCutModeSeqFileName));
                        _binCutModeSeqWorkbook = package.Workbook;
                    }
                }
                return _binCutModeSeqWorkbook;
            }
            set => _binCutModeSeqWorkbook = value;
        }
        private static ExcelWorkbook _binCutPostWorkbook;
        public static ExcelWorkbook BinCutPostWorkbook
        {
            get
            {
                if (_binCutPostWorkbook == null)
                {
                    if (!string.IsNullOrEmpty(LocalSpecs.BinCutPostFileName) && LocalSpecs.BinCutPostFileName != "N/A")
                    {
                        var package = new ExcelPackage(new FileInfo(LocalSpecs.BinCutPostFileName));
                        _binCutPostWorkbook = package.Workbook;
                    }
                }
                return _binCutPostWorkbook;
            }
            set => _binCutPostWorkbook = value;
        }
        private static ExcelWorkbook _equationVoltages;
        public static ExcelWorkbook EquationVoltages
        {
            get
            {
                if (_equationVoltages == null)
                {
                    if (!string.IsNullOrEmpty(LocalSpecs.EquationVoltagesFileName) && LocalSpecs.EquationVoltagesFileName != "N/A")
                    {
                        var package = new ExcelPackage(new FileInfo(LocalSpecs.EquationVoltagesFileName));
                        _equationVoltages = package.Workbook;
                    }
                }
                return _equationVoltages;
            }
            set => _equationVoltages = value;
        }

        private static List<ExcelWorkbook> _powerBinningWorkbooks;

        public static List<ExcelWorkbook> PowerBinningWorkbooks
        {
            get
            {
                if (_powerBinningWorkbooks == null)
                {
                    if (LocalSpecs.AllPowerBinningFileName != null && LocalSpecs.AllPowerBinningFileName.Any())
                    {
                        _powerBinningWorkbooks = new List<ExcelWorkbook>();
                        foreach (string fileName in LocalSpecs.AllPowerBinningFileName)
                        {
                            var package = new ExcelPackage(new FileInfo(fileName));
                            _powerBinningWorkbooks.Add(package.Workbook);
                        }
                    }

                }
                return _powerBinningWorkbooks;
            }
            set => _powerBinningWorkbooks = value;
        }

        #region Config Folder
        public static ExcelWorkbook SheetFormatWorkbook { get; set; }     // string.Format("\\Config\\") + "SheetsFormat.xlsx";
        #endregion

        public static ExcelWorkbook EfuseTemplateWorkbook { get; set; }                    //      ..\Settings\eFuse
        public static ExcelWorkbook FreeRunningClockAndRelayFormatWorkbook { get; set; }   //      ..\Settings\Basic\nWire_and_relay_[Project].xlsx
        private static ExcelWorkbook _spiRomWorkBook;

        public static ExcelWorkbook SpiRomWorkBook
        {
            get
            {
                if (_spiRomWorkBook == null)
                {
                    string file = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Spi", "Golden_Spirom.xlsx");
                    if (File.Exists(file))
                    {
                        var package = new ExcelPackage(new FileInfo(file));
                        _spiRomWorkBook = package.Workbook;
                    }
                }

                return _spiRomWorkBook;
            }
            set => _spiRomWorkBook = value;
        }



        public static void Clear()
        {
            // Input Files
            TestPlanWorkbook = null;
            ScghWorkbook = null;
            BinCutWorkbook = null;
            BinCutModeSeqWorkbook = null;
            BinCutPostWorkbook = null;
            EquationVoltages = null;

            // Config Folder
            SheetFormatWorkbook = null;

            EfuseTemplateWorkbook = null;
            SpiRomWorkBook = null;
            PowerBinningWorkbooks = null;
        }
    }
}
