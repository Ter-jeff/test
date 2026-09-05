using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.Utility;

using OfficeOpenXml;

namespace Cautogen.common.ReaderWriter.Reader.InputReader
{
    public class AdaptiveCoolingReader : ExcelReader
    {
        private readonly Dictionary<string, int> _headerOrder = new Dictionary<string, int>();
        private readonly AdaptiveCooling _daptiveCooling = new AdaptiveCooling();

        public AdaptiveCoolingReader(string filePath = "", Action callbackFunc = null) : base(filePath, callbackFunc)
        {
        }

        public AdaptiveCooling Read(ExcelWorksheet sh)
        {
            _Read(sh);
            return _daptiveCooling;
        }

        protected override void _Read(ExcelWorksheet sh)
        {
            int startRow = _SearchStartRow(sh);
            int endRow = sh.Dimension.End.Row;
            if (startRow == endRow)
            {
                return;
            }

            _GetHeader(sh, startRow);

            for (int row = startRow + 1; row <= endRow; row++)
            {
                string label = ExcelOperation.GetCellValue(sh, row, 1).ToUpper();
                if (string.IsNullOrEmpty(label))
                {
                    continue;
                }

                _daptiveCooling[label] = new AdaptiveCoolingData();

                foreach (string regName in _headerOrder.Keys)
                {
                    int col = _headerOrder[regName];
                    string value = sh.Cells[row, col].Value != null ? sh.Cells[row, col].Value.ToString().Trim() : "";

                    if (Regex.IsMatch(regName, "TemperatureC", RegexOptions.IgnoreCase))
                    {
                        _daptiveCooling[label].TemperatureC = value;
                    }

                    if (Regex.IsMatch(regName, "Enable", RegexOptions.IgnoreCase))
                    {
                        _daptiveCooling[label].Enable = value;
                    }

                    if (Regex.IsMatch(regName, "minDeltaC", RegexOptions.IgnoreCase))
                    {
                        _daptiveCooling[label].MinDeltaC = value;
                    }

                    if (Regex.IsMatch(regName, "maxDeltaC", RegexOptions.IgnoreCase))
                    {
                        _daptiveCooling[label].MaxDeltaC = value;
                    }

                    if (Regex.IsMatch(regName, "TimeoutSec", RegexOptions.IgnoreCase))
                    {
                        _daptiveCooling[label].TimeoutSec = value;
                    }

                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    //var reg = new AdaptiveCoolingData() { RegName = regName, RegValue = value };
                    //_daptiveCooling[label].Add(reg);
                }
            }

        }
        private void _GetHeader(ExcelWorksheet sheet, int startRow)
        {
            for (int col = 2; col <= sheet.Dimension.End.Column; col++)
            {

                string regName = sheet.Cells[startRow, col].Value != null ? sheet.Cells[startRow, col].Value.ToString().Trim() : "";
                if (!string.IsNullOrEmpty(regName))
                {
                    _headerOrder[regName] = col;
                }
            }
        }
        private int _SearchStartRow(ExcelWorksheet sheet)
        {
            for (int i = 1; i <= sheet.Dimension.End.Row; i++)
            {
                if (sheet.Cells[i, 1].Value != null && sheet.Cells[i, 1].Value.ToString().Trim() == "Insertion")
                {
                    return i;
                }
            }
            return sheet.Dimension.End.Row;
        }
    }
}
