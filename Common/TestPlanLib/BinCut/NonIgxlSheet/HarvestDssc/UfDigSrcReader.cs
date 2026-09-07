using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public class UfDigSrcReader : MySheetReader<List<UfDigSrcSheet>>
    {
        private const string ConHeaderPatternname = "PatternName";
        private const string ConHeaderSequence = "Sequence";
        private const string ConHeaderBitNum = "Dig SRC bit number";
        private const string ConHeaderBitDescription = "Bit description";
        private const string ConHeaderFunction = "Function";
        private const string ConHeaderDirective = "Directive";

        private int _indexPatternname = -1;
        private int _indexSequence = -1;
        private int _indexBitDescription = -1;
        private int _indexFunction = -1;
        private int _indexDirective = -1;

        public override List<UfDigSrcSheet>? ReadSheet(ExcelWorksheet excelWorksheet)
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

            List<UfDigSrcSheet> ufDigSrcSheet = ReadSheets(sheetName);

            return ufDigSrcSheet;
        }

        private List<UfDigSrcSheet> ReadSheets(string sheetName)
        {
            var ufDigSrcSheets = new List<UfDigSrcSheet>();
            var ufDigSrcSheet = new UfDigSrcSheet(sheetName);
            const int headerRow = 1;
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new UfDigSrcSheetRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexPatternname != -1)
                {
                    row.PatternName = ExcelWorksheet.GetCellValue(i, _indexPatternname).Trim();
                }

                if (_indexSequence != -1)
                {
                    _ = int.TryParse(ExcelWorksheet.GetCellValue(i, _indexSequence).Trim(), out int value);
                    row.Sequence = value;
                }
                if (_indexBitDescription != -1)
                {
                    row.BitDescription = ExcelWorksheet.GetCellValue(i, _indexBitDescription).Trim();
                }

                if (_indexFunction != -1)
                {
                    row.Function = ExcelWorksheet.GetCellValue(i, _indexFunction).Trim();
                }

                if (_indexDirective != -1)
                {
                    row.Directive = ExcelWorksheet.GetCellValue(i, _indexDirective).Trim();
                }

                if (row.BitDescription.EqualsIgnoreCase("skip"))
                {
                    ufDigSrcSheet.BincutPatternCount = ufDigSrcSheet.Rows.Count(x => x.PatternName != null);
                    ufDigSrcSheet.IndexPatternname = _indexPatternname;
                    ufDigSrcSheet.IndexSequence = _indexSequence;
                    ufDigSrcSheet.IndexDirective = _indexDirective;
                    ufDigSrcSheet.IndexBitDescription = _indexBitDescription;
                    ufDigSrcSheet.IndexFunction = _indexFunction;
                    ufDigSrcSheet.Rows = [.. ufDigSrcSheet.Rows.OrderBy(x => x.Sequence)];
                    ufDigSrcSheets.Add(ufDigSrcSheet);
                    ufDigSrcSheet = new UfDigSrcSheet(sheetName);
                    continue;
                }
                var dic = new Dictionary<string, string>();
                for (int j = _indexDirective + 1; j < EndCol; j++)
                {
                    string header = ExcelWorksheet.GetCellValue(headerRow, j).Trim();
                    string context = ExcelWorksheet.GetCellValue(i, j).Trim();
                    if (header.EqualsIgnoreCase("end"))
                    {
                        break;
                    }

                    dic.Add(header, context);
                }
                row.DsscDictionary = dic;
                ufDigSrcSheet.Rows.Add(row);
            }

            return ufDigSrcSheets;
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
                if (header.EqualsIgnoreCase(ConHeaderSequence) || header.EqualsIgnoreCase(ConHeaderBitNum))
                {
                    _indexSequence = i;
                    continue;
                }
                //if (header.Equals(ConHeaderHarvestResult0, StringComparison.OrdinalIgnoreCase))
                //{
                //    _indexHarvestResult0 = i;
                //}
                if (header.EqualsIgnoreCase(ConHeaderBitDescription))
                {
                    _indexBitDescription = i;
                }
                if (header.EqualsIgnoreCase(ConHeaderFunction))
                {
                    _indexFunction = i;
                }
                if (header.EqualsIgnoreCase(ConHeaderDirective))
                {
                    _indexDirective = i;
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

    public class UfDigSrcSheet : MySheet
    {
        public List<UfDigSrcSheetRow> Rows { set; get; }

        public int IndexPatternname = -1;
        public int IndexSequence = -1;
        public int IndexDirective = -1;
        public int IndexBitDescription = -1;
        public int IndexFunction = -1;
        public int BincutPatternCount;

        #region Constructor
        public UfDigSrcSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

        public List<string> GetBits(string item)
        {
            var bits = new List<string>();
            foreach (UfDigSrcSheetRow row in Rows)
            {
                foreach (KeyValuePair<string, string> dic in row.DsscDictionary)
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

    public class UfDigSrcSheetRow : MyRow
    {
        public string PatternName { set; get; } = string.Empty;
        public int Sequence { set; get; }
        public string HarvestResult0 { set; get; } = string.Empty;
        public string BitDescription { set; get; } = string.Empty;
        public string Function { set; get; } = string.Empty;
        public string Directive { set; get; } = string.Empty;
        public Dictionary<string, string> DsscDictionary = [];

        public UfDigSrcSheetRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
    }
}
