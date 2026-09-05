using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using OfficeOpenXml;

using TestPlanLib.DataStruct;
using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.PowerMerge;
using TestPlanLib.Static;
using TestPlanLib.Utility;

namespace Automation.Singleton
{
    public class TestSettingDataParser
    {
        private readonly MultiTestSettingSheetsSingleton _printData;
        private readonly HardIpDcSheet _hardipDcSheet;
        private readonly PowerMergeSheet _powerMerge;
        private readonly IoLevelsSheet _ioLevelsSheet;
        private delegate void CategoryValueParser(DcCategoryValue category);
        private readonly Dictionary<string, CategoryValueParser> _categoryValueFillDic;

        private bool _hasSpecialUnit;
        private List<Regex> _unitReg;
        private Regex _valueReg;
        private delegate string UnitTreatFunc(string value);
        private Dictionary<string, UnitTreatFunc> _unitTreatDic;
        private EnumTestSettingBasicUnit _currentBasicUnit;

        public const string Base = "Base";
        public const string Zero = "Zero";
        private readonly string _currentFillMethod;
        private readonly ExcelWorkbook _testplanWorkbook;

        public TestSettingDataParser(ExcelWorkbook testplanWorkbook, MultiTestSettingSheetsSingleton originData, PowerMergeSheet powerMerge, HardIpDcSheet hardipDcSheet = null, IoLevelsSheet ioLevelSheet = null)
        {
            _printData = originData;
            _hardipDcSheet = hardipDcSheet;
            _ioLevelsSheet = ioLevelSheet;
            _powerMerge = powerMerge;
            _testplanWorkbook = testplanWorkbook;

            _categoryValueFillDic = new Dictionary<string, CategoryValueParser>
            {
                { Base, TestSettingDataParserHelpers.FillDcCategoryByBase },
                { Zero, TestSettingDataParserHelpers.FillDcCategoryByZero }
            };
            _currentFillMethod = Zero;

            EnableUnitTreatMode();
        }

        public void EnableUnitTreatMode()
        {
            _hasSpecialUnit = true;

            _valueReg = new Regex(@"(?<Value>-?\d+(\.\d+)?)");
            _unitReg = new List<Regex>
            {
                new Regex("mv", RegexOptions.IgnoreCase),
                new Regex("(?<=\\b)V|(?<=\\d)V", RegexOptions.IgnoreCase)
            };

            _unitTreatDic = new Dictionary<string, UnitTreatFunc>
            {
                { "mv", TestSettingDataParserHelpers.ConvertMiliValt },
                { "v", TestSettingDataParserHelpers.ConvertValt }
            };
        }

        public MultiTestSettingSheetsSingleton EditPrintData()
        {
            foreach (TestSettingData entry in _printData.TestSettingSheetsList)
            {
                _currentBasicUnit = entry.BasicUnit;
                if (_currentBasicUnit == EnumTestSettingBasicUnit.mV)
                {
                    EnableUnitTreatMode();
                }

                EditCategoryData(entry.DataRows);
            }

            if (_powerMerge != null)
            {
                ConvertPinsToNetPin(_powerMerge);
            }

            if (_hardipDcSheet != null)
            {
                AddHardipDcCategory(_hardipDcSheet);
            }

            if (_ioLevelsSheet != null)
            {
                AddIoLevelsDcCategory(_ioLevelsSheet);
            }

            return _printData;
        }

        private MultiTestSettingSheetsSingleton ConvertPinsToNetPin(PowerMergeSheet powerMerge)
        {
            foreach (TestSettingData entry in _printData.TestSettingSheetsList)
            {
                ConvertPinsToNetPin(powerMerge, entry.DataRows, entry.Job);
            }
            return _printData;
        }

        public void ConvertPinsToNetPin(PowerMergeSheet powerMerge, List<TestSettingRow> rows, string job)
        {
            Regex vrsMathcer = new Regex(MultiTestSettingSheetsSingleton.ValtRowPinNameFlag, RegexOptions.IgnoreCase);
            string tableName = TestSettingDataParserHelpers.GetPowerMergeTableName(job);  //由job得到powerMerge中对应的table
            List<string> currentAllPinnames = rows.ConvertAll(x => x.PowerPinName);  //获取当前TestSetting的所有Pin Name
            List<string> ballPinNames = powerMerge.GetBallPinsByTable(tableName);  //获得该table中所有的bump pin
            List<string> netPinNames = powerMerge.GetNetPinsByTable(tableName);  //获得该table中所有的net pin
            ExcelWorksheet powerMergeSheet = _testplanWorkbook.Worksheets[NeededSheets.PowerMerge];
            int powerMergeSheetJobHeaderIndex = powerMergeSheet.Cells.FirstOrDefault
                                    (x => x.Text.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) >= 0 &&
                                            (x.Text.IndexOf("BUMP", StringComparison.OrdinalIgnoreCase) >= 0 || x.Text.IndexOf("BALL", StringComparison.OrdinalIgnoreCase) >= 0)
                                    ).Start.Column;
            for (int i = rows.Count - 1; i >= 0; i--)  //反向遍历testsetting所有的row
            {
                string pinName = rows[i].PowerPinName;  //获得当前行的pin name
                if (vrsMathcer.IsMatch(pinName))
                {
                    pinName = vrsMathcer.Replace(pinName, "");  //如果当前行是vrs pin，则取得对应的power pin name
                }
                //if (TestProgram.IgxlWorkBk.PinMapPair.Key == null)
                //{
                //    var pinMapCreater = new PinMapMain();
                //    pinMapCreater.WorkFlow(_testplanWorkbook.Worksheets[NeededSheets.PinMap],
                //        _testplanWorkbook.Worksheets[NeededSheets.IoGroup], _testplanWorkbook.Worksheets[NeededSheets.ContiIo]);
                //}
                //if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pinName))
                //{
                //    extraTestSettingRows.Add(_rows[i]);
                //    foreach (var pin in TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(pinName).PinList)
                //    {
                //        var row = new TestSettingRow();
                //        row.PowerPinName = pin.PinName;
                //        row.DcCategorys = _rows[i].DcCategorys;
                //        extraTestSettingRows.Add(row);

                //        if (isValt)
                //        {
                //            var rowValt = new TestSettingRow();
                //            rowValt.PowerPinName = pin.PinName + MultiTestSettingSheetsSingleton.ValtRowPinNameFlag;
                //            rowValt.DcCategorys = _rows[i].DcCategorys;
                //            extraTestSettingRows.Add(rowValt);
                //        }
                //    }
                //}
                if (netPinNames.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (ballPinNames.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))  //先查看当前pin是否为bump pin
                {
                    string netPinName = powerMerge.CovertNetNameFromBallName(pinName, tableName);  //得到bump pin在该table中的net pin name
                    if (netPinName != string.Empty)
                    {
                        if (currentAllPinnames.Exists(x => x.Equals(netPinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            int pinRowIndex = powerMergeSheet.Cells.FirstOrDefault(x => x.Start.Column == powerMergeSheetJobHeaderIndex && x.Text.Equals(pinName, StringComparison.OrdinalIgnoreCase)).Start.Row;
                            string errorMessage =
                                $"The {tableName} net pin {netPinName} of {pinName} has existed in TestSetting !!!";
                            ErrorReportManager.AddError(BasicErrorType.E_MissingPin_10, _testplanWorkbook.Worksheets[NeededSheets.PowerMerge].Name, pinRowIndex, powerMergeSheetJobHeaderIndex, $"The {tableName} net pin {netPinName} of {pinName} has existed in TestSetting !!!", new string[] { tableName, netPinName, pinName });
                        }
                        else
                        {
                            rows[i].PowerPinName = rows[i].PowerPinName.Replace(pinName, netPinName);  //用net pin name替换bump pin name
                        }
                    }
                }
                else  //当前pin name不是bump pin，则默认其为net pin
                {
                    //string message= "Can not find pin: " + pinName + " in PowerMerge sheet!";
                    //AddError(message, _sheetName, i);
                    //_rows.RemoveAt(i);
                }
            }
        }

        private MultiTestSettingSheetsSingleton AddHardipDcCategory(HardIpDcSheet hardipDcSheet)
        {
            foreach (TestSettingData entry in _printData.TestSettingSheetsList)
            {
                TestSettingDataParserHelpers.AddHardipDcCategory(hardipDcSheet, entry);
            }
            return _printData;
        }

        private MultiTestSettingSheetsSingleton AddIoLevelsDcCategory(IoLevelsSheet iolevelsSheet)
        {
            foreach (TestSettingData entry in _printData.TestSettingSheetsList)
            {
                TestSettingDataParserHelpers.AddIoLevelDcCategory(iolevelsSheet, entry);
            }
            return _printData;
        }

        private void EditCategoryData(List<TestSettingRow> categoryRows)
        {
            foreach (TestSettingRow row in categoryRows)
            {
                foreach (DcCategoryValue categoryValue in row.DcCategoryValues)
                {
                    _categoryValueFillDic[_currentFillMethod](categoryValue);

                    if (_hasSpecialUnit)
                    {
                        categoryValue.Nv.Value = TreatUnit(categoryValue.Nv.Value);
                        categoryValue.Lv.Value = TreatUnit(categoryValue.Lv.Value);
                        categoryValue.Hv.Value = TreatUnit(categoryValue.Hv.Value);
                    }
                }
            }
        }

        private string TreatUnit(string categoryValue)
        {
            foreach (Regex reg in _unitReg)
            {
                if (reg.IsMatch(categoryValue))
                {
                    string unitType = reg.Match(categoryValue).Value.ToLower();
                    Match matcher = _valueReg.Match(categoryValue);
                    string value = matcher.Groups["Value"].Value;
                    return _unitTreatDic[unitType](value);
                }
            }

            return _currentBasicUnit == EnumTestSettingBasicUnit.mV ? TestSettingDataParserHelpers.ConvertMiliValt(categoryValue) : categoryValue;
        }
    }
}
