using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.PreCheck.PreChecks.Basic
{
    public class BasicPreCheckIoInfo : BasicPreCheckBase
    {
        private readonly Dictionary<int, string> _titleIndexDic = new Dictionary<int, string>();
        private readonly List<string> _commonIoPins = new List<string>();

        private const string ConBlockCommon = "Common";
        private const string ConBlockHardip = "HardIP";
        private const string ConPinGrpName = "PinGrpName";

        private static readonly Regex _pinReg = new Regex("[a-zA-Z]", RegexOptions.Compiled);

        public BasicPreCheckIoInfo(ExcelWorkbook excelWorkbook, string sheetName) : base(excelWorkbook, sheetName)
        {
            FirstHeader = "PinGrpName";
        }

        protected internal override bool CheckHeaders()
        {
            //must have "common" above first data location
            bool result = CheckPinBlock();

            //check table column titles
            if (!CheckColumnTitle())
            {
                result = false;
            }

            return result;
        }

        protected internal override bool CheckFormat()
        {
            int currentRow = 1;

            //check common io pin name and table data
            bool result = CheckIoPinName(ref currentRow);

            //check other block io pin name exist in common io pins
            if (!CheckBlock(ConBlockHardip, ref currentRow))
            {
                result = false;
            }

            return result;
        }

        protected internal override bool CheckBusiness()
        {
            return true;
        }

        private bool ContainsCurrentTitle(string title)
        {
            var titleReg = new Regex(title, RegexOptions.IgnoreCase);
            for (int i = StartColumn + 1; i <= ExcelWorksheet.Dimension.Columns; i++)
            {
                string content = ExcelWorksheet.GetCellValue(StartRow, i);
                if (titleReg.IsMatch(content))
                {
                    _titleIndexDic.Add(i, title);
                    return true;
                }
            }
            return false;
        }

        private bool CheckPinBlock()
        {
            bool result = true;
            if (StartRow == -1)
            {
                ErrorReportManager.AddError(BasicErrorType.E_MissingBlock_01, ExcelWorksheet.Name, StartRow, 1, "Can not find common block.");
                result = false;
            }

            bool pass = false;
            for (int row = 1; row <= StartRow; row++)
            {
                for (int col = 1; col <= ExcelWorksheet.Dimension.Columns; col++)
                {
                    string commonBlock = ExcelWorksheet.GetCellValue(row, col);
                    if (commonBlock.Equals(ConBlockCommon, StringComparison.OrdinalIgnoreCase))
                    {
                        pass = true;
                    }
                }
            }

            if (!pass)
            {
                ErrorReportManager.AddError(BasicErrorType.E_MissingBlock_02, ExcelWorksheet.Name, StartRow, 1, "Can not find common block.");
                result = false;
            }

            return result;
        }

        private bool CheckColumnTitle()
        {
            bool result = true;
            var titles = new HashSet<string> { "NV", "HV", "LV", "Vil", "Vih", "Vol", "Voh" };
            foreach (string title in titles)
            {
                if (!ContainsCurrentTitle(title))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingBlock_03, ExcelWorksheet.Name, StartRow, 1, "Omit column \"" + title + "\"", new string[] { title });
                    result = false;
                }
            }
            return result;
        }

        private bool CheckIoPinName(ref int currentRow)
        {
            bool result = true;
            for (int i = StartRow + 1; i <= ExcelWorksheet.Dimension.Rows; i++)
            {
                string cell = ExcelWorksheet.GetCellValue(i, StartColumn);
                if (cell == string.Empty)
                {
                    currentRow = i;
                    break;
                }
                _commonIoPins.Add(cell);
                bool nvIsPin = NvIsPin(i, StartColumn + 1, ExcelWorksheet.Dimension.End.Column, out string pinName);
                for (int j = StartColumn + 1; j <= ExcelWorksheet.Dimension.Columns; j++)
                {
                    if (ExcelWorksheet.GetCellValue(StartRow, j) == string.Empty ||
                        ExcelWorksheet.GetCellValue(StartRow, j).Equals("Comment", StringComparison.CurrentCultureIgnoreCase))
                    {
                        continue;
                    }

                    string content = ExcelWorksheet.GetCellValue(i, j);
                    if (nvIsPin)
                    {
                        if (content == string.Empty)
                        {
                            continue;
                        }

                        if (content.Contains("*"))
                        {
                            string[] items = content.Split('*');
                            if (!items[0].Trim().Equals(pinName))
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_11, ExcelWorksheet.Name, i, j, "Different pin name from NV: " + pinName, new string[] { pinName });
                                result = false;
                            }

                            if (!double.TryParse(items[1].Trim(), out _))
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_12, ExcelWorksheet.Name, i, j, "Wrong format ratio: " + items[1], new string[] { items[1] });
                                result = false;
                            }
                        }
                        else
                        {
                            if (double.TryParse(content.Trim(), out _))
                            {
                                continue;
                            }
                            if (!content.Trim().Equals(pinName))
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_13, ExcelWorksheet.Name, i, j, "Different pin name from NV: " + pinName, new string[] { pinName });
                                result = false;
                            }
                        }
                    }
                    else
                    {
                        if (content == string.Empty)
                        {
                            if (_titleIndexDic.TryGetValue(j, out string value1))
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_14, ExcelWorksheet.Name, i, 1, "Omit \"" + value1 + "\" value for pin: " + cell, new string[] { value1, cell });
                                result = false;
                            }
                            else
                            {
                                string header = ExcelWorksheet.GetCellValue(StartRow, j);
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_15, ExcelWorksheet.Name, i, 1, $"Omit value for pin: {cell} in column {header}", new string[] { cell, header });
                                result = false;
                            }
                            continue;
                        }

                        if (!double.TryParse(content, out _))
                        {
                            if (_titleIndexDic.TryGetValue(j, out string value1))
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_16, ExcelWorksheet.Name, i, 1, "Non numerical \"" + value1 + "\" value for pin: " + cell, new string[] { value1, cell });
                                result = false;
                            }
                        }
                    }
                }
            }

            return result;
        }

        private bool CheckBlock(string blockName, ref int currentRow)
        {
            bool result = true;
            if (FindBlock(blockName, ref currentRow))
            {
                int startRow = currentRow;
                int startCol = 0;
                if (FindCell(ConPinGrpName, ref startRow, ref startCol))
                {
                    for (int row = startRow + 1; row <= ExcelWorksheet.Dimension.Rows; row++)
                    {
                        string content = ExcelWorksheet.GetCellValue(row, startCol);
                        if (content == string.Empty)
                        {
                            currentRow = row;
                            break;
                        }

                        if (!_commonIoPins.Exists(x => x.Equals(content, StringComparison.OrdinalIgnoreCase)))
                        {
                            ErrorReportManager.AddError(BasicErrorType.E_MissingBlock_04, ExcelWorksheet.Name, row, startCol, content + " in " + blockName + " block does not exist in common io pins.", new string[] { content, blockName });
                            result = false;
                        }
                    }
                }
            }
            return result;
        }

        private bool FindBlock(string blockName, ref int currentRow)
        {
            var blockReg = new Regex(blockName, RegexOptions.IgnoreCase);
            for (int i = currentRow + 1; i <= ExcelWorksheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= ExcelWorksheet.Dimension.Columns; j++)
                {
                    string content = ExcelWorksheet.GetCellValue(i, j);
                    if (blockReg.IsMatch(content))
                    {
                        currentRow = i;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool IsPin(string content)
        {
            return _pinReg.IsMatch(content);
        }

        private bool NvIsPin(int row, int startCol, int endCol, out string pinName)
        {
            for (int i = startCol; i <= endCol; i++)
            {
                string content = ExcelWorksheet.GetCellValue(row, i);
                if (IsPin(content))
                {
                    if (_titleIndexDic.TryGetValue(i, out string value))
                    {
                        if (value.Equals("NV", StringComparison.OrdinalIgnoreCase))
                        {
                            pinName = content.Trim();
                            return true;
                        }
                    }
                }
            }

            pinName = string.Empty;
            return false;
        }

        private bool FindCell(string cellContent, ref int row, ref int col)
        {
            var dataLocationReg = new Regex(cellContent, RegexOptions.IgnoreCase);
            for (int i = row + 1; i <= ExcelWorksheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= ExcelWorksheet.Dimension.Columns; j++)
                {
                    string value = ExcelWorksheet.GetCellValue(i, j);
                    if (dataLocationReg.IsMatch(value))
                    {
                        row = i;
                        col = j;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
