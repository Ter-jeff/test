using System;
using System.Collections.Generic;

using CommonLib.Utility;

using OfficeOpenXml;

namespace CommonLib.FormatCheck
{
    public class SheetConfigRow
    {
        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string SheetName { get; set; }
        public string FirstHeaderName { get; set; }
        public string HeaderName { get; set; }
        public string Optional { get; set; }
        public string Type { get; set; }

        #endregion

        #region Constructor
        public SheetConfigRow()
        {
            SheetName = "";
            FirstHeaderName = "";
            HeaderName = "";
            Optional = "";
            Type = "";
        }

        public SheetConfigRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            SheetName = "";
            FirstHeaderName = "";
            HeaderName = "";
            Optional = "";
            Type = "";
        }
        #endregion
    }

    public class SheetConfigSheet
    {
        #region Field
        #endregion

        #region Property
        public string SheetName { get; set; }
        public List<SheetConfigRow> Rows { get; }
        public Dictionary<string, int> HeaderIndex { get; } = new Dictionary<string, int>();
        #endregion

        #region Constructor
        public SheetConfigSheet(string sheetname)
        {
            SheetName = sheetname;
            Rows = new List<SheetConfigRow>();
        }
        #endregion
    }

    public class SheetConfigReader
    {
        private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
        private SheetConfigSheet _sheetConfigSheet;

        private const string ConHeaderSheetname = "SheetName";
        private const string ConHeaderFirstheadername = "FirstHeaderName";
        private const string ConHeaderHeadername = "HeaderName";
        private const string ConHeaderOptional = "Optional";
        private const string ConHeaderType = "Type";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _sheetnameIndex = -1;
        private int _firstheadernameIndex = -1;
        private int _headernameIndex = -1;
        private int _optionalIndex = -1;
        private int _typeIndex = -1;
        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            { "SheetName", true }, { "FirstHeaderName", true }, { "HeaderName", true }, { "Optional", true }, { "Type", true }
        };

        public SheetConfigSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _sheetConfigSheet = new SheetConfigSheet(_sheetName);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _sheetConfigSheet = ReadSheetData();

            return _sheetConfigSheet;
        }

        private SheetConfigSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new SheetConfigRow(_sheetName);
                row.RowNum = i;
                if (_sheetnameIndex != -1)
                {
                    row.SheetName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _sheetnameIndex).Trim();
                }

                if (_firstheadernameIndex != -1)
                {
                    row.FirstHeaderName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _firstheadernameIndex).Trim();
                }

                if (_headernameIndex != -1)
                {
                    row.HeaderName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _headernameIndex).Trim();
                }

                if (_optionalIndex != -1)
                {
                    row.Optional = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionalIndex).Trim();
                }

                if (_typeIndex != -1)
                {
                    row.Type = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _typeIndex).Trim();
                }

                _sheetConfigSheet.Rows.Add(row);
            }
            return _sheetConfigSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderSheetname, StringComparison.OrdinalIgnoreCase))
                {
                    _sheetnameIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(ConHeaderSheetname, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderFirstheadername, StringComparison.OrdinalIgnoreCase))
                {
                    _firstheadernameIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(ConHeaderFirstheadername, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderHeadername, StringComparison.OrdinalIgnoreCase))
                {
                    _headernameIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(ConHeaderHeadername, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderOptional, StringComparison.OrdinalIgnoreCase))
                {
                    _optionalIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(ConHeaderOptional, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(ConHeaderType, i);
                    continue;
                }
            }

            foreach (KeyValuePair<string, int> header in _sheetConfigSheet.HeaderIndex)
            {
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                {
                    return false;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderSheetname, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }
            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _sheetnameIndex = -1;
            _firstheadernameIndex = -1;
            _headernameIndex = -1;
            _optionalIndex = -1;
            _typeIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            var dics = new List<Dictionary<string, string>>();
            foreach (SheetConfigRow row in _sheetConfigSheet.Rows)
            {
                var dic = new Dictionary<string, string>();
                dic.Add("SheetName", row.SheetName);
                dic.Add("FirstHeaderName", row.FirstHeaderName);
                dic.Add("HeaderName", row.HeaderName);
                dic.Add("Optional", row.Optional);
                dic.Add("Type", row.Type);
                dics.Add(dic);
            }
            return dics;
        }
    }
}
