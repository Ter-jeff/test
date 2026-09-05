using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib
{
    public class TestNameWidthReader : MySheetReader<TestNameWidthTable>
    {
        private const string ConModule = "Module";
        private const string ConptrTestNameWidth = "ptrTestNameWidth";
        private const string ConptrPinWidth = "ptrPinWidth";
        private const string ConftrTestNameWidth = "ftrTestNameWidth";
        private const string ConftrPatternWidth = "ftrPatternWidth";
        private int _indexModule = -1;
        private int _indexptrTestNameWidt = -1;
        private int _indexptrPinWidth = -1;
        private int _indexftrTestNameWidt = -1;
        private int _indexftrPatternWidth = -1;
        private readonly TestNameWidthTable _testNameWidthTable = new();

        public override TestNameWidthTable? ReadSheet(ExcelWorksheet excelWorksheet)
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

            return _testNameWidthTable;
        }

        private void ReadSheet()
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new TestNameWidthRow();
                if (_indexModule != -1)
                {
                    row.Module = ExcelWorksheet.GetCellValue(i, _indexModule).Trim();
                }
                if (_indexptrTestNameWidt != -1)
                {
                    row.PtrTestNameWidt = ExcelWorksheet.GetCellValue(i, _indexptrTestNameWidt).Trim();
                }
                if (_indexptrPinWidth != -1)
                {
                    row.PtrPinWidth = ExcelWorksheet.GetCellValue(i, _indexptrPinWidth).Trim();
                }
                if (_indexftrTestNameWidt != -1)
                {
                    row.FtrTestNameWidt = ExcelWorksheet.GetCellValue(i, _indexftrTestNameWidt).Trim();
                }
                if (_indexftrPatternWidth != -1)
                {
                    row.FtrPatternWidth = ExcelWorksheet.GetCellValue(i, _indexftrPatternWidth).Trim();
                }
                _testNameWidthTable.Rows.Add(row);
            }
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConModule))
                {
                    _indexModule = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConptrTestNameWidth))
                {
                    _indexptrTestNameWidt = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConptrPinWidth))
                {
                    _indexptrPinWidth = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConftrTestNameWidth))
                {
                    _indexftrTestNameWidt = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConftrPatternWidth))
                {
                    _indexftrPatternWidth = i;
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConModule))
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
