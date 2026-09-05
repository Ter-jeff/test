using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public class TimeSettingReader : MySheetReader<TimeSettingSheet>
    {
        private const string ConHeaderSymbol = "Symbol";
        private Dictionary<string, int> _symbolIdx = [];

        private int _indexSymbol = -1;

        public override TimeSettingSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            string sheetName = excelWorksheet.Name;

            var timeSettingSheet = new TimeSettingSheet(sheetName);

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                timeSettingSheet.AddDimensionError();
                return timeSettingSheet;
            }

            //if (!GetFirstHeaderPosition()) return (TimeSettingsReaderSheet)timeSettingsReaderSheet.FirstHeaderError(ConHeaderSymbol);

            GetHeaderIndex();

            return ReadSheet(sheetName);
        }

        private TimeSettingSheet ReadSheet(string sheetName)
        {
            var timeSettingSheet = new TimeSettingSheet(sheetName);
            for (int i = StartCol + 1; i <= EndCol; i++)
            {
                var row = new TimeSettingRow(sheetName)
                {
                    RowNum = i
                };
                foreach (KeyValuePair<string, int> symbol in _symbolIdx)
                {
                    if (symbol.Key == ConHeaderSymbol)
                    {
                        row.SheetName = ExcelWorksheet.GetCellValue(symbol.Value, i).Trim();
                    }
                    else if (ExcelWorksheet.GetCellValue(symbol.Value, i).Trim()?.Length != 0)
                    {
                        if (!row.SymbolValues.ContainsKey(symbol.Key))
                        {
                            row.SymbolValues.Add(symbol.Key, ExcelWorksheet.GetCellValue(symbol.Value, i).Trim());
                        }
                    }
                }
                timeSettingSheet.Rows.Add(row);
            }

            timeSettingSheet.IndexSymbol = _indexSymbol;

            return timeSettingSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartRow; i <= EndRow; i++)
            {
                string header = ExcelWorksheet.GetCellValue(i, StartCol).Trim();
                if (string.IsNullOrEmpty(header))
                {
                    continue;
                }

                if (header.EqualsIgnoreCase(ConHeaderSymbol))
                {
                    _indexSymbol = i;
                    _symbolIdx = [];
                }
                if (_symbolIdx?.ContainsKey(header) == false)
                {
                    _symbolIdx.Add(header, i);
                }
            }
        }
    }
}
