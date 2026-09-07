using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public class HarvPmodeTableOffsetReader : MySheetReader<HarvPmodeTableOffsetSheet>
    {
        private const string ConHeaderByMode = "By_Mode";
        private const string ConHeaderPmode = "Pmode";
        private const string ConHeaderOffset = "Offset";

        private int _indexByMode = -1;
        private int _indexPmode = -1;
        private int _indexOffset = -1;
        public override HarvPmodeTableOffsetSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            HarvPmodeTableOffsetSheet harvPmodeTableOffsetSheet = ReadSheet(sheetName);

            return harvPmodeTableOffsetSheet;
        }

        private HarvPmodeTableOffsetSheet ReadSheet(string sheetName)
        {
            var harvPmodeTableOffsetSheet = new HarvPmodeTableOffsetSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvOffsetRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexPmode != -1)
                {
                    row.Pmode = ExcelWorksheet.GetCellValue(i, _indexPmode).Trim();
                }

                if (_indexByMode != -1)
                {
                    row.ByMode = ExcelWorksheet.GetCellValue(i, _indexByMode).Trim();
                }

                if (_indexOffset != -1)
                {
                    row.Offset = ExcelWorksheet.GetCellValue(i, _indexOffset).Trim();
                }

                if (row.Pmode?.Length == 0)
                {
                    break;
                }

                if (row.Offset?.Length != 0)
                {
                    harvPmodeTableOffsetSheet.Rows.Add(row);
                }
            }

            harvPmodeTableOffsetSheet.IndexByMode = _indexByMode;
            harvPmodeTableOffsetSheet.IndexPmode = _indexPmode;
            harvPmodeTableOffsetSheet.IndexOffset = _indexOffset;
            return harvPmodeTableOffsetSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderByMode))
                {
                    _indexByMode = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPmode))
                {
                    _indexPmode = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderOffset))
                {
                    _indexOffset = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            _ = EndRow > 10 ? 10 : EndRow;

            _ = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= EndRow; i++)
            //for (int j = 1; j <= colNum; j++)
            {
                if (ExcelWorksheet.GetCellValue(i, 1).Trim().EqualsIgnoreCase(ConHeaderPmode))
                //if (ExcelWorksheet.GetCellValue( i, j).Trim().Equals(ConHeaderByMode, StringComparison.OrdinalIgnoreCase))
                {
                    StartRow = i;
                    return true;
                }
            }
            return false;
        }
    }

    public class HarvPmodeTableOffsetSheet : MySheet
    {
        public List<HarvOffsetRow> Rows { set; get; }

        public int IndexByMode = -1;
        public int IndexPmode = -1;
        public int IndexOffset = -1;

        #region Constructor
        public HarvPmodeTableOffsetSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

    }

    public class HarvOffsetRow : MyRow
    {
        public string Pmode { set; get; } = "";
        public string ByMode { set; get; } = "";
        public string Offset { set; get; } = "";
        #region Constructor
        public HarvOffsetRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
