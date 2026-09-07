using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.DataStruct
{
    public class IfoldPowerTableReader
    {
        private const string ConHeaderPinname = "PinName";
        private const string ConHeaderCurrentA = "Current (A)";

        private ExcelWorksheet? _excelWorksheet;
        private string _sheetName = "";
        private IfoldPowerTableSheet? _powerTableSheet;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _pinnameIndex = -1;
        private int _currentAIndex = -1;

        public IfoldPowerTableSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            _excelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _powerTableSheet = new IfoldPowerTableSheet(_sheetName);

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

            _powerTableSheet = ReadSheetData();

            return _powerTableSheet!;
        }

        private IfoldPowerTableSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                IfoldPowerTableRow row = new IfoldPowerTableRow(_sheetName)
                {
                    RowNum = i
                };
                if (_pinnameIndex != -1)
                {
                    row.PinName = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _pinnameIndex).Trim();
                }

                if (_currentAIndex != -1)
                {
                    row.Current = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _currentAIndex).Trim();
                }

                _powerTableSheet!.Rows.Add(row);
            }
            return _powerTableSheet!;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string header = EpplusExtensions.GetCellValue(_excelWorksheet!, _startRowNumber, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderPinname))
                {
                    _pinnameIndex = i;
                    _powerTableSheet!.HeaderIndex.Add(ConHeaderPinname, i);
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderCurrentA))
                {
                    _currentAIndex = i;
                    _powerTableSheet!.HeaderIndex.Add(ConHeaderCurrentA, i);
                    continue;
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
                    if (EpplusExtensions.GetCellValue(_excelWorksheet!, i, j).Trim().EqualsIgnoreCase(ConHeaderPinname))
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
            if (_excelWorksheet!.Dimension != null)
            {
                _startColNumber = _excelWorksheet!.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet!.Dimension.Start.Row;
                _endColNumber = _excelWorksheet!.Dimension.End.Column;
                _endRowNumber = _excelWorksheet!.Dimension.End.Row;
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
            _pinnameIndex = -1;
            _currentAIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            List<Dictionary<string, string>> dics = [];
            foreach (IfoldPowerTableRow row in _powerTableSheet!.Rows)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>
                {
                    { "PinName", row.PinName },
                    { "Current (A)", row.Current }
                };
                dics.Add(dic);
            }
            return dics;
        }
    }

    public class IfoldPowerTableRow
    {
        #region Properity
        public string SourceSheetName { get; set; } = string.Empty;
        public int RowNum { get; set; }
        public string PinName { get; set; }
        public string Current { get; set; }
        #endregion

        #region Constructor
        public IfoldPowerTableRow()
        {
            PinName = "";
            Current = "";
        }

        public IfoldPowerTableRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            PinName = "";
            Current = "";
        }
        #endregion
    }

    public class IfoldPowerTableSheet(string sheetname)
    {
        #region Properity

        public string SheetName { get; set; } = sheetname;
        public List<IfoldPowerTableRow> Rows { get; } = [];

        public Dictionary<string, int> HeaderIndex { get; } = [];

        #endregion
        #region Constructor

        #endregion
    }
}
