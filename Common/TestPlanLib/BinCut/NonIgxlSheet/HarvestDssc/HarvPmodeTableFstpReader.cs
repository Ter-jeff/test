using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public class HarvPmodeTableFstpReader : MySheetReader<HarvPmodeTableFstpSheet>
    {
        private const string ConHeaderEnableCore = "Enable_Core";
        private const string ConHeaderFstp = "FSTP";

        private int _indexEnableCore = -1;
        private int _indexFstp = -1;

        public override HarvPmodeTableFstpSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            HarvPmodeTableFstpSheet harvPmodeTableFstpTableSheet = ReadSheet(sheetName);

            return harvPmodeTableFstpTableSheet;
        }

        private HarvPmodeTableFstpSheet ReadSheet(string sheetName)
        {
            var harvPmodeTableFstpTableSheet = new HarvPmodeTableFstpSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvPmodeTableFstpRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexEnableCore != -1)
                {
                    row.EnableCore = ExcelWorksheet.GetCellValue(i, _indexEnableCore).Trim();
                }

                if (_indexFstp != -1)
                {
                    row.Fstp = ExcelWorksheet.GetCellValue(i, _indexFstp).Trim();
                }

                if (string.IsNullOrEmpty(row.EnableCore))
                {
                    break;
                }

                harvPmodeTableFstpTableSheet.Rows.Add(row);
            }

            harvPmodeTableFstpTableSheet.IndexEnableCore = _indexEnableCore;
            harvPmodeTableFstpTableSheet.IndexFstp = _indexFstp;

            return harvPmodeTableFstpTableSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderEnableCore))
                {
                    _indexEnableCore = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderFstp))
                {
                    _indexFstp = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            //int rowNum = EndRowNumber > 10 ? 10 : EndRowNumber;
            //int colNum = EndColNumber > 10 ? 10 : EndColNumber;
            for (int i = 1; i <= EndRow; i++)
            //for (int j = 1; j <= colNum; j++)
            {
                if (ExcelWorksheet.GetCellValue(i, 1).Trim().EqualsIgnoreCase(ConHeaderEnableCore))
                //if (ExcelWorksheet.GetCellValue( i, j).Trim().Equals(ConHeaderEnableCore, StringComparison.OrdinalIgnoreCase))
                {
                    StartRow = i;
                    return true;
                }
            }
            return false;
        }
    }

    public class HarvPmodeTableFstpSheet : MySheet
    {
        public List<HarvPmodeTableFstpRow> Rows { set; get; }

        public int IndexEnableCore = -1;
        public int IndexFstp = -1;

        #region Constructor
        public HarvPmodeTableFstpSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

        public List<string> GetFstp(string groupCondition)
        {
            var fstp = new List<string>();
            string[] arr = groupCondition.Split(',');
            foreach (string item in arr)
            {
                if (Rows.Exists(x => x.EnableCore.EqualsIgnoreCase(item)))
                {
                    fstp.Add(Rows.Find(x => x.EnableCore.EqualsIgnoreCase(item))!.Fstp);
                }
            }
            return fstp;
        }
    }

    public class HarvPmodeTableFstpRow : MyRow
    {
        public string EnableCore { set; get; } = "";
        public string Fstp { set; get; } = "";

        #region Constructor
        public HarvPmodeTableFstpRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
