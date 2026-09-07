using System;
using System.Collections.Generic;

using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.Utility;

using OfficeOpenXml;

namespace Cautogen.common.ReaderWriter.Reader.InputReader
{
    public class DigSrcRegReader : ExcelReader
    {
        private readonly DigSrcReg _digSrcReg = new DigSrcReg();
        private readonly Dictionary<string, int> _headerOrder = new Dictionary<string, int>();
        public DigSrcRegReader(string filePath = "", Action callbackFunc = null) : base(filePath, callbackFunc)
        {

        }
        public DigSrcReg Read(ExcelWorksheet sh)
        {
            _Read(sh);
            return _digSrcReg;
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

                _digSrcReg[label] = new List<DigSrcRegData>();

                foreach (string regName in _headerOrder.Keys)
                {
                    int col = _headerOrder[regName];
                    string value = sh.Cells[row, col].Value != null ? sh.Cells[row, col].Value.ToString().Trim() : "";
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    var reg = new DigSrcRegData() { RegName = regName, RegValue = value };
                    _digSrcReg[label].Add(reg);
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
                if (sheet.Cells[i, 1].Value != null && sheet.Cells[i, 1].Value.ToString().Trim() == "DataInfo")
                {
                    return i;
                }
            }
            return sheet.Dimension.End.Row;
        }
    }
}
