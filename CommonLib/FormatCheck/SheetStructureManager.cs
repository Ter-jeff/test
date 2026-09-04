using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

using CommonLib.Utility;

using Microsoft.Office.Interop.Excel;

using Range = Microsoft.Office.Interop.Excel.Range;

using OfficeOpenXml;

namespace CommonLib.FormatCheck
{
    public class SheetStructureManager
    {
        public static List<SheetConfig> SheetConfigs { get; set; }
        public static ExcelWorksheet ExcelWorksheet;
        private static readonly List<string> _pinMapType = new List<string> { "I/O", "Input", "Output", "Analog", "Power", "Gnd", "Utility", "Voltage", "Current", "Unknown" };
        private static readonly List<string> _channelMapType = new List<string> { "DCDiffMeter", "DCTime", "DCTimeHP", "DCTimeTrig", "DCVI", "DCVIMerfed", "DCVS", "DCVSMerged2", "DCVSMerged4", "DCVSMerged6", "DCVSMerged8", "Gnd", "I/O", "N/C", "Utility" };
        private static readonly List<string> _dctestContinuityCategoryList = new List<string> { "Continuity" };

        public static List<string> UnitList = new List<string>();
        private static readonly List<string> _unitPrefixs = new List<string> { "f", "p", "n", "u", "m", "K", "M", "G", "T" };
        private static readonly List<string> _units = new List<string> { "V", "A", "Hz" };
        private static readonly Regex _regex = new Regex("[a-zA-Z]", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex("[a-zA-Z]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static void Initialize(Assembly assembly)
        {
            var excelPackage = new ExcelPackage(new FileInfo("Test"));
            ExcelWorksheet = excelPackage.Workbook.Worksheets.Add("Test");
            if (SheetConfigs == null)
            {
                SheetConfigs = GetSheetConfig(assembly);
            }

            foreach (string unitPrefix in _unitPrefixs)
            {
                foreach (string unit in _units)
                {
                    UnitList.Add(unitPrefix + unit);
                }
            }
        }

        public static List<SheetConfig> GetSheetConfig(Assembly assembly)
        {
            string[] resourceNames = assembly.GetManifestResourceNames();
            foreach (string resourceName in resourceNames)
            {
                if (resourceName.EndsWith(".Config.xlsx", StringComparison.CurrentCultureIgnoreCase))
                {
                    var inputExcel = new ExcelPackage(assembly.GetManifestResourceStream(resourceName));
                    ExcelWorksheet workSheet = inputExcel.Workbook.Worksheets["SheetConfig"];
                    if (workSheet != null)
                    {
                        var projectConfigSettingReader = new SheetConfigReader();
                        SheetConfigSheet sheetConfigSheet = projectConfigSettingReader.ReadSheet(workSheet);
                        var sheetConfigs = new List<SheetConfig>();
                        foreach (SheetConfigRow row in sheetConfigSheet.Rows)
                        {
                            var sheetConfig = new SheetConfig();
                            sheetConfig.SheetName = row.SheetName;
                            sheetConfig.FirstHeaderName = row.FirstHeaderName;
                            sheetConfig.HeaderName = row.HeaderName;
                            sheetConfig.Optional = row.Optional.Equals("True", StringComparison.CurrentCultureIgnoreCase);
                            if (row.Type.Equals("Pin", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.Pin;
                            }
                            else if (row.Type.Equals("PinMapType", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.PinMapType;
                            }
                            else if (row.Type.Equals("ChannelMapType", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.ChannelMapType;
                            }
                            else if (row.Type.Equals("Decimal", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.Decimal;
                            }
                            else if (row.Type.Equals("Integer", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.Integer;
                            }
                            else if (row.Type.Equals("Binary", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.Binary;
                            }
                            else if (row.Type.Equals("DctestContinuityCategory", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.DctestContinuityCategory;
                            }
                            else if (row.Type.Equals("Formula", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.Formula;
                            }
                            else if (row.Type.Equals("Unit", StringComparison.CurrentCultureIgnoreCase))
                            {
                                sheetConfig.Type = EnumColumn.Unit;
                            }

                            sheetConfigs.Add(sheetConfig);

                        }
                        return sheetConfigs;
                    }
                }
            }
            return null;
        }

        public static void SetInitialColumns(Workbook workbook, Worksheet worksheet, string firstHeader, string targetHeader, EnumColumn columnType)
        {
            int headerRowNumber;
            int columnIndex = worksheet.GetColumnIndexByHeader(firstHeader, targetHeader, out headerRowNumber);
            if (columnIndex != -1)
            {
                worksheet.Unprotect();

                Validation validation = worksheet.Columns[columnIndex].Validation;

                if (columnType == EnumColumn.Pin)
                {
                    validation.Delete();
                    Worksheet sheet = workbook.GetSheet("IoPinMap");
                    if (sheet != null)
                    {
                        validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                            XlFormatConditionOperator.xlBetween, "=" + sheet.Name + "!" + GetRangeAddress(sheet));
                    }
                }
                else if (columnType == EnumColumn.PinMapType)
                {
                    validation.Delete();
                    validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                           XlFormatConditionOperator.xlBetween, string.Join(",", _pinMapType));
                }
                else if (columnType == EnumColumn.ChannelMapType)
                {
                    validation.Delete();
                    validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                         XlFormatConditionOperator.xlBetween, string.Join(",", _channelMapType));
                }
                else if (columnType == EnumColumn.Decimal)
                {
                    validation.Delete();
                    Range range = worksheet.Columns[columnIndex];
                    string letter = range.GetColumnLetter();
                    validation.Add(XlDVType.xlValidateCustom, XlDVAlertStyle.xlValidAlertStop,
                         XlFormatConditionOperator.xlBetween, "=IsNumber(" + letter + "1)");
                    validation.ErrorMessage = "It should be number !!!";
                }
                else if (columnType == EnumColumn.DctestContinuityCategory)
                {
                    validation.Delete();
                    validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                         XlFormatConditionOperator.xlBetween, string.Join(",", _dctestContinuityCategoryList));
                }
                else if (columnType == EnumColumn.Formula)
                {
                }
                else if (columnType == EnumColumn.Unit)
                {
                }


                worksheet.GetRange(1, columnIndex, headerRowNumber, columnIndex).Validation.Delete();
                // worksheet.Protect(Type.Missing, false, true, false, false, true, true, true, true, true, true, true,true, true, true, true);

            }
        }

        private static string GetRangeAddress(Worksheet worksheet)
        {
            int headerRowNumber;
            int columnIndex = worksheet.GetColumnIndexByHeader("Group Name", "Pin Name", out headerRowNumber);
            if (columnIndex != -1)
            {
                dynamic finalRow = worksheet.Cells[headerRowNumber, columnIndex].End(XlDirection.xlDown).Row();
                return worksheet.Range[worksheet.Cells[headerRowNumber + 1, columnIndex], worksheet.Cells[finalRow, columnIndex]]
                        .Address();
            }
            return "";
        }

        public static bool JudgeCell(EnumColumn columnType, string value, out string errorMessage)
        {
            errorMessage = "";
            if (string.IsNullOrEmpty(value))
            {
                return true;
            }

            if (columnType == EnumColumn.PinMapType)
            {
                if (_pinMapType.Exists(x => x.Equals(value, StringComparison.CurrentCultureIgnoreCase)))
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! => {1}", value, string.Join(",", _pinMapType));
                return false;
            }
            else if (columnType == EnumColumn.ChannelMapType)
            {
                if (_channelMapType.Exists(x => x.Equals(value, StringComparison.CurrentCultureIgnoreCase)))
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, string.Join(",", _channelMapType));
                return false;
            }
            else if (columnType == EnumColumn.Decimal)
            {
                double number;
                if (double.TryParse(value, out number))
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not number !!!");
                return false;
            }
            else if (columnType == EnumColumn.Integer)
            {
                int number;
                if (int.TryParse(value, out number))
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not integer !!!");
                return false;
            }
            else if (columnType == EnumColumn.Binary)
            {
                if (value == "0" || value == "1")
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not binary !!!");
                return false;
            }
            else if (columnType == EnumColumn.DctestContinuityCategory)
            {
                if (_dctestContinuityCategoryList.Exists(x => x.Equals(value, StringComparison.CurrentCultureIgnoreCase)))
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, string.Join(",", _channelMapType));
                return false;
            }
            else if (columnType == EnumColumn.Formula)
            {
                if (IsFormula(value, ""))
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not formula !!!");
                return false;
            }
            else if (columnType == EnumColumn.Unit)
            {
                if (IsUnit(value))
                {
                    return true;
                }

                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not value with unit !!!");
                return false;
            }
            return true;
        }

        private static bool IsUnit(string value)
        {
            try
            {
                value = value.Replace(" ", "");
                foreach (string unit in UnitList)
                {
                    if (value.EndsWith(unit))
                    {
                        value = Regex.Replace(value, unit, "", RegexOptions.IgnoreCase);
                        break;
                    }
                }

                double outValue;
                if (double.TryParse(value, out outValue))
                {
                    return true;
                }

                if (_regex.IsMatch(value))
                {
                    return false;
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsFormula(string formula, string variable)
        {
            try
            {
                if (!string.IsNullOrEmpty(variable))
                {
                    formula = formula.Replace(" ", "").TrimStart('=').Replace(variable, "1");
                }
                else
                {
                    formula = formula.Replace(" ", "").TrimStart('=');
                }

                if (_regex2.IsMatch(formula))
                {
                    return false;
                }

                //ExcelWorksheet.Cells["A1"].Formula = formula;
                //ExcelWorksheet.Cells["A1"].Calculate();

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
