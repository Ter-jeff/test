using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Concurrent
{
    public class ConcurrentFlowSheetReader : MySheetReader<ConcurrentFlowSheet>
    {
        private const string ConHeaderSequenceName = "SEQUENCE NAME";
        private const string ConHeaderSubflow = "SUBFLOW";

        private int _sequenceNameCol = -1;
        private int _subFlowStartCol = -1;
        private int _subFlowEndCol = -1;

        public override ConcurrentFlowSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition();

            GetHeaders();

            ConcurrentFlowSheet sheet = ReadSheet(sheetName);

            return sheet;
        }

        protected void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderSequenceName))
                    {
                        StartRow = i;
                        break;
                    }
                }
            }
        }

        protected void GetHeaders()
        {
            bool flowStepStart = false;
            for (int i = 1; i <= EndCol; i++)
            {
                string lStrHeader = GetCellValue(StartRow, i).ToUpper().Trim();

                if (lStrHeader.ToUpper().EqualsIgnoreCase(ConHeaderSequenceName))
                {
                    _sequenceNameCol = i;
                    continue;
                }
                if (lStrHeader.Contains(ConHeaderSubflow))
                {
                    if (!flowStepStart)
                    {
                        _subFlowStartCol = i;
                        flowStepStart = true;
                    }
                    else
                    {
                        _subFlowEndCol = i;
                    }
                }
            }
        }

        protected ConcurrentFlowSheet ReadSheet(string sheetName)
        {
            var sheet = new ConcurrentFlowSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new ConcurrentFlowSheetRow
                {
                    RowNum = i,
                    SheetName = sheetName,
                    SequenceName = _sequenceNameCol == -1 ? "" : ExcelWorksheet.GetCellValue(i, _sequenceNameCol).Trim()
                };
                if (_subFlowStartCol != -1 && _subFlowEndCol != -1)
                {
                    for (int j = _subFlowStartCol; j <= _subFlowEndCol; j++)
                    {
                        if (string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(i, j).Trim()))
                        {
                            continue;
                        }

                        row.Subflows.Add(ExcelWorksheet.GetCellValue(i, j).Trim());
                    }
                }
                sheet.Rows.Add(row);
            }
            sheet.SequenceNameColNumber = _sequenceNameCol;
            sheet.SubflowColStart = _subFlowStartCol;
            sheet.SubflowColEnd = _subFlowEndCol;

            return sheet;
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = ExcelWorksheet.Cells[rowNumber, columnNumber].Value;
            return value?.ToString() ?? "";
        }
    }
}
