using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public class VariableInitTableReader : MySheetReader<VariableInitTable>
    {
        private const string ConOpCode = "OP Code";
        private const string ConParameter = "Parameter";
        private const string ConComment = "Comment";
        private readonly Dictionary<string, int> _jobIdx = [];
        private int _indexOpCode = -1;
        private int _indexParameter = -1;
        private int _indexComment = -1;
        private readonly VariableInitTable _variableInitTable = new();

        public override VariableInitTable? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadSheet();

            return _variableInitTable;
        }

        private void ReadSheet()
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new VariableRow();
                if (_indexOpCode != -1)
                {
                    row.Opcode = ExcelWorksheet.GetCellValue(i, _indexOpCode).Trim();
                }
                if (_indexParameter != -1)
                {
                    row.Parameter = ExcelWorksheet.GetCellValue(i, _indexParameter).Trim();
                }
                if (_indexComment != -1)
                {
                    row.Comment = ExcelWorksheet.GetCellValue(i, _indexComment).Trim();
                }
                foreach (KeyValuePair<string, int> job in _jobIdx)
                {
                    int jobCol = job.Value;
                    string jobName = job.Key;
                    string value = ExcelWorksheet.GetCellValue(i, jobCol).Trim();
                    if (!string.IsNullOrEmpty(value))
                    {
                        row.JobValues.Add(jobName, value);
                    }
                }
                _variableInitTable.Rows.Add(row);
            }
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConOpCode))
                {
                    _indexOpCode = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConParameter))
                {
                    _indexParameter = i;
                    continue;
                }

                if (header.EqualsIgnoreCase(ConComment))
                {
                    _indexComment = i;
                    break;
                }
                if (_jobIdx.TryAdd(header, i))
                {
                    continue;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow;
            int colNum = EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConOpCode))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
