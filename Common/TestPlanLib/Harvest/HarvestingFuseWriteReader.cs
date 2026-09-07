using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Harvest
{
    public class HarvestingFuseWriteReader : MySheetReader<HarvestingFuseWriteSheet>
    {
        private const string ConHeaderFuseName = "FuseName";
        private const string ConHeaderValue = "Value to write";

        private int _indexName = -1;
        private int _indexValue = -1;

        public override HarvestingFuseWriteSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            HarvestingFuseWriteSheet harvestingFuseWriteSheet = ReadSheet(sheetName);

            return harvestingFuseWriteSheet;
        }

        private HarvestingFuseWriteSheet ReadSheet(string sheetName)
        {
            var harvestingFuseWriteSheet = new HarvestingFuseWriteSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvestingFuseWriteRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexName != -1)
                {
                    row.FuseName = ExcelWorksheet.GetCellValue(i, _indexName).Trim();
                }

                if (_indexValue != -1)
                {
                    row.Value = ExcelWorksheet.GetCellValue(i, _indexValue).Trim();
                }

                harvestingFuseWriteSheet.Rows.Add(row);
            }
            harvestingFuseWriteSheet.IndexFuseName = _indexName;
            harvestingFuseWriteSheet.IndexValue = _indexValue;

            return harvestingFuseWriteSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderFuseName))
                {
                    _indexName = i;
                    continue;
                }
                if (header.StartsWithIgnoreCase(ConHeaderValue))
                {
                    _indexValue = i;
                    continue;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderFuseName))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class HarvestingFuseWriteSheet : MySheet
    {
        public List<HarvestingFuseWriteRow> Rows { set; get; }
        public int IndexFuseName = -1;
        public int IndexValue = -1;
        #region Constructor

        public HarvestingFuseWriteSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }

        #endregion
    }
    public class HarvestingFuseWriteRow : MyRow
    {
        public string FuseName = "";
        public string Value = "";
        public string Job = "";
        public string BlockName = "";

        #region Constructor
        public HarvestingFuseWriteRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
