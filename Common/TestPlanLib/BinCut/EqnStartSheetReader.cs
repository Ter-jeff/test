using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class EqnStartSheetReader : MySheetReader<EqnStartSheet>
    {
        private const string ConHeaderMode = "Mode";
        private const string ConHeaderEqnstart = "EqnStart";

        private int _indexMode = -1;
        private int _indexEqnstart = -1;

        public override EqnStartSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            EqnStartSheet eqnStartSheet = ReadSheet(sheetName);

            return eqnStartSheet;
        }

        private EqnStartSheet ReadSheet(string sheetName)
        {
            var eqnStartSheet = new EqnStartSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new EqnStartRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexMode != -1)
                {
                    row.Mode = ExcelWorksheet.GetCellValue(i, _indexMode).Trim();
                }

                if (_indexEqnstart != -1)
                {
                    row.Eqnstart = ExcelWorksheet.GetCellValue(i, _indexEqnstart).Trim();
                }

                eqnStartSheet.Rows.Add(row);
            }

            eqnStartSheet.IndexMode = _indexMode = -1;
            eqnStartSheet.IndexEqnstart = _indexEqnstart = -1;

            return eqnStartSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string lStrHeader = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (lStrHeader.EqualsIgnoreCase(ConHeaderMode))
                {
                    _indexMode = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderEqnstart))
                {
                    _indexEqnstart = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i < rowNum; i++)
            {
                for (int j = 1; j < colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderMode))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class EqnStartSheet : MySheet
    {
        public List<EqnStartRow> Rows { get; }

        public int IndexMode = -1;
        public int IndexEqnstart = -1;

        #region Constructor
        public EqnStartSheet(string sheetname)
        {
            SheetName = sheetname;
            Rows = [];
        }
        #endregion
    }

    public class EqnStartRow : MyRow
    {
        public string Mode { get; set; }
        public string Eqnstart { get; set; }

        #region Constructor
        public EqnStartRow(string sheetName = "")
        {
            SheetName = sheetName;
            Mode = "";
            Eqnstart = "";
        }
        #endregion
    }
}
