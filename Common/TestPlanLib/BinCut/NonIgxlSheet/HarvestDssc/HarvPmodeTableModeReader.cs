using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public class HarvPmodeTableModeReader : MySheetReader<HarvPmodeTableModeSheet>
    {
        private const string ConHeaderPmode = "Pmode";
        private const string ConHeaderByMode = "By_Mode";

        private int _indexPmode = -1;
        private int _indexByMode = -1;

        public override HarvPmodeTableModeSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            HarvPmodeTableModeSheet harvPmodeTableModeSheet = ReadSheet(sheetName);

            return harvPmodeTableModeSheet;
        }

        private HarvPmodeTableModeSheet ReadSheet(string sheetName)
        {
            var harvPmodeTableModeSheet = new HarvPmodeTableModeSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvPmodeTableModeRow(sheetName)
                {
                    Bymode = [],
                    RowNum = i
                };
                if (_indexPmode != -1)
                {
                    row.Pmode = ExcelWorksheet.GetCellValue(i, _indexPmode).Trim();
                }

                if (_indexByMode != -1)
                {
                    string bymode = ExcelWorksheet.GetCellValue(i, _indexByMode).Trim();
                    string[] split = bymode.Split(',');
                    for (int j = 0; j < split.Length; j++)
                    {
                        row.Bymode.Add(split[j]);
                    }
                }
                //row.ByMode = ExcelWorksheet.GetCellValue( i, _indexByMode).Trim();

                if (string.IsNullOrEmpty(row.Pmode))
                {
                    break;
                }

                harvPmodeTableModeSheet.Rows.Add(row);
            }

            harvPmodeTableModeSheet.IndexPmode = _indexPmode;
            harvPmodeTableModeSheet.IndexByMode = _indexByMode;

            return harvPmodeTableModeSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderPmode))
                {
                    _indexPmode = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderByMode))
                {
                    _indexByMode = i;
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
                {
                    StartRow = i;
                    return true;
                }
            }
            return false;
        }
    }

    public class HarvPmodeTableModeSheet : MySheet
    {
        public List<HarvPmodeTableModeRow> Rows { set; get; }

        public int IndexPmode = -1;
        public int IndexByMode = -1;

        #region Constructor
        public HarvPmodeTableModeSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion
    }

    public class HarvPmodeTableModeRow : MyRow
    {
        public string Pmode { set; get; } = "";
        public List<string> Bymode { set; get; } = [];

        #region Constructor
        public HarvPmodeTableModeRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
