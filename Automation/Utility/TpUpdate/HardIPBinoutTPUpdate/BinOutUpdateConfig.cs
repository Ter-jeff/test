using System;
using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace Automation.Utility.TpUpdate.HardIPBinoutTPUpdate
{
    public class BinOutUpdateConfigReader : MySheetReader<BinOutUpdateConfigSheet>
    {
        private string _sheetName;
        private BinOutUpdateConfigSheet _updateConfigsheet;

        private const string HeaderOption = "Option";
        private const string HeaderValue = "Value";

        private int _indexOption = -1;
        private int _indexValue = -1;

        public override BinOutUpdateConfigSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _updateConfigsheet = new BinOutUpdateConfigSheet(_sheetName);

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

            _updateConfigsheet = ReadSheet();

            return _updateConfigsheet;
        }

        private BinOutUpdateConfigSheet ReadSheet()
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                BinOutUpdateConfigRow row = new BinOutUpdateConfigRow(_sheetName)
                {
                    RowNum = i
                };

                if (_indexOption != -1)
                {
                    row.Option = ExcelWorksheet.GetCellValue(i, _indexOption).Trim();
                }

                if (_indexValue != -1)
                {
                    row.Value = ExcelWorksheet.GetCellValue(i, _indexValue).Trim();
                }

                if (!_updateConfigsheet.Rows.ContainsKey(row.Option))
                {
                    _updateConfigsheet.Rows[row.Option] = row.Value;
                }
            }

            _updateConfigsheet.IndexOption = _indexOption;
            _updateConfigsheet.IndexValue = _indexValue;

            return _updateConfigsheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string lStrHeader = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (lStrHeader.Equals(HeaderOption, StringComparison.OrdinalIgnoreCase))
                {
                    _indexOption = i;
                    continue;
                }
                if (lStrHeader.Equals(HeaderValue, StringComparison.OrdinalIgnoreCase))
                {
                    _indexValue = i;
                    continue;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().Equals(HeaderOption, StringComparison.OrdinalIgnoreCase))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }

        private void Reset()
        {
            StartCol = -1;
            StartRow = -1;
            EndCol = -1;
            EndRow = -1;
            _indexOption = -1;
            _indexValue = -1;
        }
    }

    public class BinOutUpdateConfigSheet : MySheet
    {
        public Dictionary<string, string> Rows { get; set; }

        public int IndexOption = -1;
        public int IndexValue = -1;

        public const string OptionOverWriteTestPlanLimits = "OverWriteTestPlanLimits";

        public bool EnableOverWriteTestLimits
        {
            get { return IsSettingEnabled(OptionOverWriteTestPlanLimits); }
        }

        public BinOutUpdateConfigSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
        }

        public bool IsSettingEnabled(string option)
        {
            // Default set the option to be enabled
            if (Rows == null || !Rows.ContainsKey(option) || string.IsNullOrEmpty(Rows[option]))
            {
                return true;
            }

            // TRUE => enable the setting
            // FALSE => disable the setting
            bool enabled = Rows[option].Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            return enabled;
        }
    }

    public class BinOutUpdateConfigRow : MyRow
    {
        public string Option { get; set; }
        public string Value { get; set; }

        public BinOutUpdateConfigRow(string sheetName)
        {
            SheetName = sheetName;
        }
    }
}
