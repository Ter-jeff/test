using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public class HarvMappingTableReader : MySheetReader<HarvMappingTableSheet>
    {
        private const string ConHeaderPatternname = "PatternName";
        private const string ConHeaderSequence = "Sequence";
        private const string ConHeaderHarvestResult0 = "Harvest Result0";

        private int _indexPatternname = -1;
        private int _indexSequence = -1;
        private int _indexHarvestResult0 = -1;

        public override HarvMappingTableSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            HarvMappingTableSheet harvMappingTableSheet = ReadSheet(sheetName);

            return harvMappingTableSheet;
        }

        private HarvMappingTableSheet ReadSheet(string sheetName)
        {
            var harvMappingTableSheet = new HarvMappingTableSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvMappingTableRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexPatternname != -1)
                {
                    row.Patternname = ExcelWorksheet.GetCellValue(i, _indexPatternname).Trim();
                }

                if (_indexSequence != -1)
                {
                    _ = int.TryParse(ExcelWorksheet.GetCellValue(i, _indexSequence).Trim(), out int value);
                    row.Sequence = value;
                }
                if (_indexHarvestResult0 != -1)
                {
                    row.HarvestResult0 = ExcelWorksheet.GetCellValue(i, _indexHarvestResult0).Trim();
                }

                var dic = new Dictionary<string, bool>();
                for (int j = _indexHarvestResult0 + 1; j < EndCol; j++)
                {
                    string header = ExcelWorksheet.GetCellValue(1, j).Trim();
                    string context = ExcelWorksheet.GetCellValue(i, j).Trim();
                    dic.Add(header, context == "1");
                }
                row.DsscDictionary = dic;
                harvMappingTableSheet.Rows.Add(row);
            }

            harvMappingTableSheet.BincutPatternCount =
                harvMappingTableSheet.Rows.Count(x => x.Patternname?.IndexOf("_TDF_", StringComparison.OrdinalIgnoreCase) >= 0);
            harvMappingTableSheet.IndexPatternname = _indexPatternname;
            harvMappingTableSheet.IndexSequence = _indexSequence;
            harvMappingTableSheet.IndexHarvestResult0 = _indexHarvestResult0;
            harvMappingTableSheet.Rows = [.. harvMappingTableSheet.Rows.OrderBy(x => x.Sequence)];

            return harvMappingTableSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderPatternname))
                {
                    _indexPatternname = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderSequence))
                {
                    _indexSequence = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderHarvestResult0))
                {
                    _indexHarvestResult0 = i;
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderPatternname))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class HarvMappingTableSheet : MySheet
    {
        public List<HarvMappingTableRow> Rows { set; get; }

        public int IndexPatternname = -1;
        public int IndexSequence = -1;
        public int IndexHarvestResult0 = -1;
        public int BincutPatternCount;

        #region Constructor
        public HarvMappingTableSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

        public List<bool> GetBits(string item)
        {
            var bits = new List<bool>();
            foreach (HarvMappingTableRow row in Rows)
            {
                foreach (KeyValuePair<string, bool> dic in row.DsscDictionary)
                {
                    if (dic.Key.EqualsIgnoreCase(item))
                    {
                        bits.Add(dic.Value);
                    }
                }
            }
            return bits;
        }
    }

    public class HarvMappingTableRow : MyRow
    {
        public string Patternname { set; get; } = "";
        public int Sequence { set; get; }
        public string HarvestResult0 { set; get; } = "";
        public Dictionary<string, bool> DsscDictionary = [];

        #region Constructor
        public HarvMappingTableRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
