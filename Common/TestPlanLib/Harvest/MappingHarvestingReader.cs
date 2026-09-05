using System.Collections.Generic;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Harvest
{
    public class MappingHarvestingReader : MySheetReader<MappingHarvestingSheet>
    {
        private const string ConHeaderPattern = "PLPATTERN";
        private const string ConHeaderCoreName = "CoreName/PinGroup";
        private const string ConHeaderHarvestFlag = "HarvestFlag";
        private const string ConHeaderInitPattern = "INPATTERN";
        private const string ConHeaderPowerSupply = "PowerSupply (multiple rails bincut search)";
        private const string ConHeaderComment = "Comment";
        private int _indexPattern = -1;
        private int _indexCoreName = -1;
        private int _indexHarvestFlag = -1;
        private int _indexInitPattern = -1;
        private int _indexPowerSupply = -1;
        private int _indexComment = -1;
        private readonly Dictionary<string, int> _headerIndex = [];
        private string _sheetName = "";
        private MappingHarvestingSheet? _sheet;

        public override MappingHarvestingSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;
            _sheet = new MappingHarvestingSheet(_sheetName);
            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadSheet(_sheetName);
            // _error.Add(ErrorReportManager.ErrorInstance);

            return _sheet!;
        }

        private void ReadSheet(string sheetName)
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new MappingHarvestingRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexPattern != -1)
                {
                    row.Pattern = ExcelWorksheet.GetCellValue(i, _indexPattern).Trim();
                }

                if (row.Pattern?.Length == 0)
                {
                    continue;
                }

                if (_indexCoreName != -1)
                {
                    row.CoreName = ExcelWorksheet.GetCellValue(i, _indexCoreName).Trim();
                }

                if (_indexHarvestFlag != -1)
                {
                    row.HarvestFlag = ExcelWorksheet.GetCellValue(i, _indexHarvestFlag).Trim();
                }

                if (_indexInitPattern != -1)
                {
                    row.InitPattern = ExcelWorksheet.GetCellValue(i, _indexInitPattern).Trim();
                }

                if (_indexPowerSupply != -1)
                {
                    row.PowerSupply = ExcelWorksheet.GetCellValue(i, _indexPowerSupply).Trim();
                }

                if (_indexComment != -1)
                {
                    row.Comment = ExcelWorksheet.GetCellValue(i, _indexComment).Trim();
                }

                _sheet!.Rows.Add(row);
            }
            _sheet!.IndexPowerSupply = _indexPowerSupply;
            _sheet!.IndexInpattern = _indexInitPattern;
            _sheet!.IndexPattern = _indexPattern;
            _sheet!.IndexCoreName = _indexCoreName;
            _sheet!.IndexHarvestFlag = _indexHarvestFlag;
            _sheet!.IndexComment = _indexComment;
            _sheet!.HeaderIndex = _headerIndex;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                bool foundIndex = true;
                if (header.EqualsIgnoreCase(ConHeaderPattern))
                {
                    _indexPattern = i;
                }
                else if (header.EqualsIgnoreCase(ConHeaderCoreName))
                {
                    _indexCoreName = i;
                }
                else if (header.EqualsIgnoreCase(ConHeaderHarvestFlag))
                {
                    _indexHarvestFlag = i;
                }
                else if (header.EqualsIgnoreCase(ConHeaderInitPattern))
                {
                    _indexInitPattern = i;
                }
                else if (header.EqualsIgnoreCase(ConHeaderPowerSupply))
                {
                    _indexPowerSupply = i;
                }
                else if (header.EqualsIgnoreCase(ConHeaderComment))
                {
                    _indexComment = i;
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
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderPattern))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class MappingHarvestingSheet : MySheet
    {
        public List<MappingHarvestingRow> Rows { set; get; }

        public int IndexPattern = -1;
        public int IndexCoreName = -1;
        public int IndexHarvestFlag = -1;
        public int IndexInpattern = -1;
        public int IndexPowerSupply = -1;
        public int IndexComment = -1;
        public Dictionary<string, int> HeaderIndex = [];
        #region Constructor
        public MappingHarvestingSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion
    }

    public partial class MappingHarvestingRow : MyRow
    {
        [GeneratedRegex("(SSC|SSU)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public string Pattern = "";
        public string CoreName = "";
        public string HarvestFlag = "";
        public string InitPattern = "";
        public string PowerSupply = "";
        public string Comment = "";

        public bool IsSsn { get { return _regex.IsMatch(Pattern); } }

        #region Constructor
        public MappingHarvestingRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion

    }
}
