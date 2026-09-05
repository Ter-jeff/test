using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;


namespace TestPlanLib
{
    public class VreTestCaseTableReader : MySheetReader<VreTestCaseTable>
    {
        private const string ConCaseId = "Case ID";
        private const string ConSubProgram = "SubProgram";
        private const string ConProcesureName = "ProcesureName";
        private const string ConInstanceName = "InstanceName";
        private const string ConPattern = @"^Pattern\d+";
        private const string ConUserDef = "User_Def";
        private const string ConLevelCheck = "Level check";
        private const string ConHardBin = "Expected Hard Bin";
        private const string ConSorfdBin = "Expected Soft Bin";
        private const string ConComment = "Comment";

        private int _indexCaseId = -1;
        private int _indexSubProgram = -1;
        private int _indexProcesureName = -1;
        private int _indexInstanceName = -1;
        private List<int> _indexPattern = new List<int>();
        private int _indexUserDef = -1;
        private int _indexLevelCheck = -1;
        private int _indexHardBin = -1;
        private int _indexSorfdBin = -1;
        private int _indexComment = -1;
        private Dictionary<string, int> _headerIndex = new Dictionary<string, int>();
        private readonly VreTestCaseTable _vreTestCaseTableTable = new VreTestCaseTable();
        public override VreTestCaseTable? ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            ExcelWorksheet = worksheet;

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

            _vreTestCaseTableTable.SheetName = ExcelWorksheet.Name;
            return _vreTestCaseTableTable;
        }

        private void ReadSheet()
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new VreTestCaseRow();
                if (_indexCaseId != -1 && int.TryParse(ExcelWorksheet.GetCellValue(i, _indexCaseId).Trim(), out int value))
                {
                    row.CaseId = value;
                }
                if (_indexSubProgram != -1)
                {
                    row.SubProgram = ExcelWorksheet.GetCellValue(i, _indexSubProgram).Trim();
                }
                if (_indexProcesureName != -1)
                {
                    row.ProcesureName = ExcelWorksheet.GetCellValue(i, _indexProcesureName).Trim();
                }
                if (_indexInstanceName != -1)
                {
                    row.InstanceName = ExcelWorksheet.GetCellValue(i, _indexInstanceName).Trim();
                }
                if (_indexUserDef != -1)
                {
                    row.UserDef = ExcelWorksheet.GetCellValue(i, _indexUserDef).Trim();
                }
                if (_indexHardBin != -1)
                {
                    row.HardBin = ExcelWorksheet.GetCellValue(i, _indexHardBin).Trim();
                }
                if (_indexSorfdBin != -1)
                {
                    row.SorfdBin = ExcelWorksheet.GetCellValue(i, _indexSorfdBin).Trim();
                }
                if (_indexComment != -1)
                {
                    row.Comment = ExcelWorksheet.GetCellValue(i, _indexComment).Trim();
                }
                if (_indexLevelCheck != -1)
                {
                    row.LevelCheck = ExcelWorksheet.GetCellValue(i, _indexLevelCheck).Trim();
                }
                if (_indexPattern.Any())
                {
                    foreach (int index in _indexPattern)
                    {
                        row.Pattern.Add(ExcelWorksheet.GetCellValue(i, index).Trim());
                    }
                }
                _vreTestCaseTableTable.Rows.Add(row);
            }
            _vreTestCaseTableTable.HeaderIndex = _headerIndex;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                bool foundIndex = true;
                if (header.EqualsIgnoreCase(ConCaseId))
                {
                    _indexCaseId = i;
                }
                else if (header.EqualsIgnoreCase(ConSubProgram))
                {
                    _indexSubProgram = i;
                }
                else if (header.EqualsIgnoreCase(ConProcesureName))
                {
                    _indexProcesureName = i;
                }
                else if (header.EqualsIgnoreCase(ConInstanceName))
                {
                    _indexInstanceName = i;
                }
                else if (header.EqualsIgnoreCase(ConUserDef))
                {
                    _indexUserDef = i;
                }
                else if (header.EqualsIgnoreCase(ConHardBin))
                {
                    _indexHardBin = i;
                }
                else if (header.EqualsIgnoreCase(ConSorfdBin))
                {
                    _indexSorfdBin = i;
                }
                else if (header.EqualsIgnoreCase(ConComment))
                {
                    _indexComment = i;
                }
                else if (Regex.IsMatch(header, ConPattern, RegexOptions.IgnoreCase))
                {
                    _indexPattern.Add(i);
                }
                else if (header.EqualsIgnoreCase(ConLevelCheck))
                {
                    _indexLevelCheck = i;
                }
                else
                {
                    foundIndex = false;
                }
                if (foundIndex)
                {
                    _headerIndex.Add(header, i);
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConCaseId))
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
