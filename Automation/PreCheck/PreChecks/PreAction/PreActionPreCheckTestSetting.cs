using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Singleton;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace Automation.PreCheck.PreChecks.PreAction
{
    public class PreActionPreCheckTestSetting : PreActionCheckBase
    {
        private static readonly Regex _valueFormatRegex = new Regex("[*/=]");
        private static readonly Regex _testRegex = new Regex("^(evs|ids|mbist|hardip|conti|nwire|efuse|rtos|rto|sa|sachain|td|tdchain|scan|bincut|htol)$", RegexOptions.IgnoreCase);
        private readonly Dictionary<int, CategoryValueType> _categoryTypeIndex = new Dictionary<int, CategoryValueType>();
        private readonly Dictionary<string, List<int>> _categoryValueDic = new Dictionary<string, List<int>>();
        internal readonly List<int> _fatalErrorColumns = new List<int>();
        private static readonly Regex _unitReg = new Regex("mv|\\dV|\\bV", RegexOptions.IgnoreCase);
        private bool _hasSpecialUnit;
        private int _maxColumnIndex;
        private int _maxRowIndex;

        public PreActionPreCheckTestSetting(ExcelWorkbook excelWorkbook, string sheetName) : base(excelWorkbook, sheetName)
        {
            FirstHeader = "VoltageTable_.*";
        }

        protected internal override bool CheckHeaders()
        {
            return true;
        }

        protected internal override bool CheckFormat()
        {
            _maxColumnIndex = GetMaxColumnIndex();
            _maxRowIndex = GetMaxRowIndex();

            //check category value type
            bool result = CheckCategoryValueType();

            //check category name
            if (!CheckCategoryName())
            {
                result = false;
            }

            //check duplicate and same category must stay together
            if (!CheckCategoryDuplicateAndAdjacency())
            {
                result = false;
            }

            //check category value, must be number or percent, check if the unit of value is valt
            if (!CheckCategoryValue())
            {
                result = false;
            }

            //check pin name, can not have blank in pin name except vrs pin
            if (!CheckPinName())
            {
                result = false;
            }

            return result;
        }

        protected internal override bool CheckBusiness()
        {
            //check category name test part, must be existed in known test naming rule
            bool result = CheckCategoryNamingRule();

            //check if nv value is the same and compare type value for each category
            if (!CheckCategoryValueLogic())
            {
                result = false;
            }

            //check if Mbist category value of power pin is the same as valt pin
            if (!CheckMbistPowerPinWithValt())
            {
                result = false;
            }

            //check if category value is ordinary, HV>=NV>=LV
            if (!CheckCategoryValueOrdinary())
            {
                result = false;
            }

            return result;
        }

        public bool HasSpecialUnit()
        {
            return _hasSpecialUnit;
        }

        internal CategoryValueType GetValueType(string value)
        {
            if (value.ContainsIgnoreCase("LV"))
            {
                return CategoryValueType.LV;
            }

            if (value.ContainsIgnoreCase("HV"))
            {
                return CategoryValueType.HV;
            }

            return CategoryValueType.NV;
        }

        private void CheckHv(string pinName, string standardNv, int row, int col, string categoryName)
        {
            if (standardNv == string.Empty)
            {
                return;
            }
            string hv = EpplusExtensions.GetCellText(ExcelWorksheet, row, col);

            if (hv.Contains("%"))
            {
                if (hv.Contains("-"))
                {
                    ErrorReportManager.AddError(PreActionErrorType.W_RuleViolationVoltage_01, ExcelWorksheet.Name, row, col,
                        [categoryName, pinName, "HV", hv, "<"]);
                }
            }
            else
            {
                bool nvResult = double.TryParse(standardNv, out double nvValue);
                bool hvResult = double.TryParse(hv, out double hvValue);
                if (nvResult && hvResult)
                {
                    if (hvValue < nvValue)
                    {
                        ErrorReportManager.AddError(PreActionErrorType.W_RuleViolationVoltage_01, ExcelWorksheet.Name, row, col,
                            [categoryName, pinName, "HV", hv, "<"]);
                    }
                }
                else
                {
                    ErrorReportManager.AddError(PreActionErrorType.W_InvalidVoltage_01, ExcelWorksheet.Name, row, col,
                        [categoryName, pinName, "HV", standardNv, hv]);
                }
            }
        }

        private void CheckLv(string pinName, string standardNv, int row, int col, string categoryName)
        {
            if (standardNv == string.Empty)
            {
                return;
            }
            string lv = EpplusExtensions.GetCellText(ExcelWorksheet, row, col);

            if (lv.Contains("%"))
            {
                if (!lv.Contains("-"))
                {
                    ErrorReportManager.AddError(PreActionErrorType.W_RuleViolationVoltage_01, ExcelWorksheet.Name, row, col,
                        [categoryName, pinName, "LV", lv, ">"]);
                }
            }
            else
            {
                bool nvResult = double.TryParse(standardNv, out double nvValue);
                bool lvResult = double.TryParse(lv, out double lvValue);
                if (nvResult && lvResult)
                {
                    if (lvValue > nvValue)
                    {
                        ErrorReportManager.AddError(PreActionErrorType.W_RuleViolationVoltage_01, ExcelWorksheet.Name, row, col,
                            [categoryName, pinName, "LV", lv, ">"]);
                    }
                }
                else
                {
                    ErrorReportManager.AddError(PreActionErrorType.W_InvalidVoltage_01, ExcelWorksheet.Name, row, col,
                        [categoryName, pinName, "LV", standardNv, lv]);
                }
            }
        }

        private bool CheckCategoryValueLogic()
        {
            bool result = true;
            for (int i = StartRow + 2; i <= _maxRowIndex; i++)
            {
                string pinName = ExcelWorksheet.GetCellValue(i, StartColumn);
                //find current standard NV
                foreach (KeyValuePair<string, List<int>> entry in _categoryValueDic)
                {
                    if (entry.Value.Count > 1)
                    {
                        foreach (int index in entry.Value)
                        {
                            if (IsFatalErrorColumn(index))
                            {
                                continue;
                            }

                            string currentNv;
                            if (_categoryTypeIndex[index] == CategoryValueType.LV)
                            {
                                currentNv = GetCurrentNvValue(entry.Value, i);
                                CheckLv(pinName, currentNv, i, index, entry.Key);
                            }
                            else if (_categoryTypeIndex[index] == CategoryValueType.HV)
                            {
                                currentNv = GetCurrentNvValue(entry.Value, i);
                                CheckHv(pinName, currentNv, i, index, entry.Key);
                            }
                        }
                    }
                }
            }

            return result;
        }

        private bool CheckCategoryNamingRule()
        {
            bool result = true;
            for (int i = StartColumn + 1; i <= _maxColumnIndex; i++)
            {
                if (IsFatalErrorColumn(i))
                {
                    continue;
                }

                string categoryName = ExcelWorksheet.GetCellValue(StartRow, i).ToUpper();
                string categoryType = ExcelWorksheet.GetCellValue(StartRow + 1, i).ToUpper();
                var currentInfo = new DcCategoryInfo(categoryName);
                CategoryValueType type = GetValueType(categoryType);
                _categoryTypeIndex.Add(i, type);
                if (!_testRegex.IsMatch(currentInfo.Test))
                {
                    ErrorReportManager.AddError(PreActionErrorType.W_RuleViolationDcCategory_01, ExcelWorksheet.Name, StartRow, i,
                        [categoryName, currentInfo.Test]);
                    result = false;
                }

                if (_categoryValueDic.TryGetValue(categoryName, out List<int> value))
                {
                    value.Add(i);
                }
                else
                {
                    var valueList = new List<int> { i };
                    _categoryValueDic.Add(categoryName, valueList);
                }
            }

            return result;
        }

        private bool CheckPinName()
        {
            bool result = true;
            string content = "";
            for (int i = StartRow + 1; i <= _maxRowIndex; i++)
            {
                if (!content.EndsWith(MultiTestSettingSheetsSingleton.ValtRowPinNameFlag, StringComparison.OrdinalIgnoreCase))
                {
                    if (content.Trim().Contains(" "))
                    {
                        ErrorReportManager.AddError(PreActionErrorType.W_InvalidFormat_01, ExcelWorksheet.Name, i, StartColumn, [content]);
                        result = false;
                    }
                }
            }

            return result;
        }

        private bool CheckCategoryValue()
        {
            bool result = true;
            for (int i = StartRow + 1; i <= _maxRowIndex; i++)
            {
                for (int j = StartColumn + 1; j <= _maxColumnIndex; j++)
                {
                    string content = ExcelWorksheet.GetCellValue(i, j);
                    if (_valueFormatRegex.IsMatch(content))
                    {
                        ErrorReportManager.AddError(PreActionErrorType.W_InvalidFormat_02, ExcelWorksheet.Name, i, j, [content]);
                        result = false;
                    }

                    if (!_hasSpecialUnit)
                    {
                        if (_unitReg.IsMatch(content))
                        {
                            _hasSpecialUnit = true;
                        }
                    }

                    if (double.TryParse(content, out double value))
                    {
                        if (value < 0)
                        {
                            Response.Report($"Voltage value cannot be less than 0, Sheet: {ExcelWorksheet.Name} Row: {i}, Colum: {j}", EnumMessageLevel.Error, 73);
                            ErrorReportManager.AddError(PreActionErrorType.W_InvalidVoltage_02, ExcelWorksheet.Name, i, j);
                            result = false;
                        }
                    }
                }
            }
            return result;
        }

        private bool CheckCategoryValueType()
        {
            bool result = true;
            for (int i = StartColumn + 1; i <= _maxColumnIndex; i++)
            {
                string content = ExcelWorksheet.GetCellValue(StartRow, i);
                if (string.IsNullOrEmpty(content))
                {
                    ErrorReportManager.AddError(PreActionErrorType.E_MissingDcCategory_01, ExcelWorksheet.Name, StartRow, i, []);
                    AddFatalErrorColumn(i);
                    result = false;
                }
            }

            return result;
        }

        private bool CheckCategoryName()
        {
            bool result = true;
            for (int i = StartColumn + 1; i <= _maxColumnIndex; i++)
            {
                string content = ExcelWorksheet.GetCellValue(StartRow, i);
                if (string.IsNullOrEmpty(content))
                {
                    ErrorReportManager.AddError(PreActionErrorType.E_MissingDcCategory_01, ExcelWorksheet.Name, StartRow, i, []);
                    AddFatalErrorColumn(i);
                    result = false;
                }
            }

            return result;
        }

        private bool CheckCategoryDuplicateAndAdjacency()
        {
            bool result = true;
            string priorCategory = "";
            var categoryPair = new List<string>();
            var names = new List<string>();
            for (int i = StartColumn + 1; i <= _maxColumnIndex; i++)
            {
                string content = ExcelWorksheet.GetCellValue(StartRow, i).ToUpper();
                string categoryTypeContent = ExcelWorksheet.GetCellValue(StartRow + 1, i).ToUpper();
                string valuePair = $"{content}@{categoryTypeContent}";
                if (categoryPair.Contains(valuePair))
                {
                    ErrorReportManager.AddError(PreActionErrorType.E_DuplicateDcCategory_01, ExcelWorksheet.Name, StartRow, i, [valuePair]);
                    AddFatalErrorColumn(i);
                    result = false;
                }
                else
                {
                    categoryPair.Add(valuePair);
                }

                if (content != priorCategory)
                {
                    if (names.Contains(content))
                    {
                        ErrorReportManager.AddError(PreActionErrorType.E_RuleViolationColumn_01, ExcelWorksheet.Name, StartRow, i, [content]);
                        AddFatalErrorColumn(i);
                        result = false;
                    }
                    names.Add(content);
                }

                priorCategory = content;
            }

            return result;
        }

        private bool CheckMbistPowerPinWithValt()
        {
            bool result = true;
            var valtPins = new Dictionary<int, string>();
            for (int i = StartRow + 1; i <= _maxRowIndex; i++)
            {
                string lStrContent = ExcelWorksheet.GetCellValue(i, StartColumn);
                if (lStrContent.EndsWith(MultiTestSettingSheetsSingleton.ValtRowPinNameFlag, StringComparison.OrdinalIgnoreCase))
                {
                    valtPins.Add(i, lStrContent);
                }
            }

            for (int i = StartColumn + 1; i <= _maxColumnIndex; i++)
            {
                string currentCategory = ExcelWorksheet.GetCellValue(StartRow, i);
                var currentCategoryInfo = new DcCategoryInfo(currentCategory);
                if (!currentCategoryInfo.Test.Equals("Mbist", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                for (int j = StartRow + 1; j <= _maxRowIndex; j++)
                {
                    string powerPin = ExcelWorksheet.GetCellValue(j, StartColumn);
                    if (!powerPin.EndsWith(MultiTestSettingSheetsSingleton.ValtRowPinNameFlag, StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (KeyValuePair<int, string> entry in valtPins)
                        {
                            if (entry.Value.Equals(powerPin + MultiTestSettingSheetsSingleton.ValtRowPinNameFlag, StringComparison.OrdinalIgnoreCase))
                            {
                                string powerPinValue = ExcelWorksheet.GetCellValue(j, i);
                                string valtPinValue = ExcelWorksheet.GetCellValue(entry.Key, i);
                                if (!powerPinValue.Equals(valtPinValue))
                                {
                                    string currentSelector = ExcelWorksheet.GetCellValue(StartRow + 1, i);
                                    ErrorReportManager.AddError(PreActionErrorType.W_RuleViolationVoltage_02, ExcelWorksheet.Name, entry.Key, i,
                                        [powerPin, currentSelector, powerPinValue, valtPinValue, currentCategory]);
                                    result = false;
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        private bool CheckCategoryValueOrdinary()
        {
            bool result = true;
            var errors = new List<string>();
            for (int i = StartRow + 1; i <= _maxRowIndex; i++)
            {
                string pinName = ExcelWorksheet.GetCellValue(i, StartColumn);
                foreach (KeyValuePair<string, List<int>> entry in _categoryValueDic)
                {
                    List<int> columns = entry.Value;
                    if (columns.Count > 3)
                    {
                        Response.Report("Duplicate category : " + entry.Key + " in " + ExcelWorksheet.Name, EnumMessageLevel.Error);
                    }
                    else if (columns.Count > 1)
                    {
                        for (int j = 0; j < columns.Count; j++)
                        {
                            for (int k = j + 1; k < columns.Count; k++)
                            {
                                string value1 = ExcelWorksheet.GetCellValue(i, columns[j]);
                                string value2 = ExcelWorksheet.GetCellValue(i, columns[k]);
                                CategoryValueType valueType1 = _categoryTypeIndex[columns[j]];
                                CategoryValueType valueType2 = _categoryTypeIndex[columns[k]];
                                bool flag = CompareTwoCategoryValue(value1, value2, valueType1, valueType2, out string connect);
                                if (!flag)
                                {
                                    string error = "Pin: " + pinName + ", " + "Category: " + entry.Key + ", " + valueType1 + ":" + value1 + " " + valueType2 + ":" + value2 + " " + valueType1 + connect + valueType2;
                                    ErrorReportManager.AddError(PreActionErrorType.E_RuleViolationVoltage_01, ExcelWorksheet.Name, i, columns[k],
                                        new string[] { pinName, entry.Key, valueType1.ToString(), value1, valueType2.ToString(), value2, connect });
                                    errors.Add(error);
                                    result = false;
                                }
                            }
                        }
                    }
                }
            }

            if (errors.Count > 0)
            {
                Response.Report("==== Compare Sheet " + ExcelWorksheet.Name + " Logic Error ====", EnumMessageLevel.Error);
                foreach (string message in errors)
                {
                    Response.Report(message, EnumMessageLevel.Error);
                }
            }

            return result;
        }

        internal bool CompareTwoCategoryValue(string value1, string value2, CategoryValueType valueType1, CategoryValueType valueType2, out string connect)
        {
            bool bValue1 = double.TryParse(value1, out double dValue1);
            bool bValue2 = double.TryParse(value2, out double dValue2);
            if (bValue1 && bValue2)
            {
                if (valueType1 == CategoryValueType.HV)
                {
                    if (dValue1 < dValue2)
                    {
                        connect = "<";
                        return false;
                    }
                }
                else if (valueType1 == CategoryValueType.NV)
                {
                    if (valueType2 == CategoryValueType.HV)
                    {
                        if (dValue1 > dValue2)
                        {
                            connect = ">";
                            return false;
                        }
                    }
                    else if (valueType2 == CategoryValueType.LV)
                    {
                        if (dValue1 < dValue2)
                        {
                            connect = "<";
                            return false;
                        }
                    }
                }
                else if (valueType1 == CategoryValueType.LV)
                {
                    if (dValue1 > dValue2)
                    {
                        connect = ">";
                        return false;
                    }
                }
            }

            connect = "X";
            return true;
        }

        private string GetCurrentNvValue(List<int> categoryColumns, int currentRow)
        {
            string result = "";
            foreach (int col in categoryColumns)
            {
                CategoryValueType type = GetValueType(ExcelWorksheet.GetCellValue(StartRow + 1, col));
                if (type == CategoryValueType.NV)
                {
                    result = EpplusExtensions.GetCellText(ExcelWorksheet, currentRow, col);
                }
            }
            return result;
        }

        internal bool IsBlankColumn(int columnNumber)
        {
            for (int i = StartRow + 1; i <= ExcelWorksheet.Dimension.End.Row; i++)
            {
                string content = ExcelWorksheet.GetCellValue(i, columnNumber);
                if (content != string.Empty)
                {
                    return false;
                }
            }
            return true;
        }

        internal int GetMaxColumnIndex()
        {
            for (int i = ExcelWorksheet.Dimension.End.Column; i >= 1; i--)
            {
                if (!IsBlankColumn(i))
                {
                    return i;
                }
            }
            return StartColumn;
        }

        internal int GetMaxRowIndex()
        {
            for (int i = StartRow + 1; i <= ExcelWorksheet.Dimension.End.Row; i++)
            {
                string content = ExcelWorksheet.GetCellValue(i, StartColumn);
                if (content == string.Empty)
                {
                    return i - 1;
                }
            }
            return ExcelWorksheet.Dimension.End.Row;
        }

        internal bool IsFatalErrorColumn(int col)
        {
            return _fatalErrorColumns.Contains(col);
        }

        internal void AddFatalErrorColumn(int col)
        {
            if (!IsFatalErrorColumn(col))
            {
                _fatalErrorColumns.Add(col);
            }
        }
    }
}
