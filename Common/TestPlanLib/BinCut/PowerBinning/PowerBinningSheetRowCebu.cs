using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.PowerBinning
{
    public class PowerBinningSheetRowCebu : PowerBinningSheetRow
    {
    }

    public class PowerBinningSheetCebu : PowerBinningSheet
    {
        public PowerBinningSheetCebu(string name)
        {
            SheetName = name;
            BinnedModeList = [];
            OtherModeList = [];
        }

        public PowerBinningSheetCebu()
        {
            BinnedModeList = [];
            OtherModeList = [];
        }

        public PowerBinningSheetCebu(PowerBinningSheetCebu powerBinningSheetCebu) : base(powerBinningSheetCebu)
        {
        }

        public new PowerBinningSheetCebu Copy()
        {
            return new PowerBinningSheetCebu(this);
        }
    }

    public partial class PowerBinningSheetReaderCebu
    {
        private const string ConHeaderBinnedMode = "Binned_Mode";
        private const string ConHeaderBinVoltage = @"Bin Voltage\.*";
        private const string ConHeaderIds = @"IDS\.*";
        private const string ConHeaderOtherMode = "Other_Mode";
        private const string ConHeaderOther = "Other";
        private const string ConHeaderOffset = "Offset";
        private const string ConHeaderSpec = "Spec";
        private const string ConVdd0 = "Vdd0 (mV)";

        [GeneratedRegex(@"\(*\s*mv\s*\)*", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(ConHeaderBinVoltage, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(ConHeaderIds, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();

        private int _binnedModeStartRowNumber = 1;
        private int _otherModeStartRowNumber = 1;
        private int _endColNumber = 1;
        private int _endRowNumber = 1;

        private int _binnedModeIndex = -1;
        private int _binnedBinVoltageIndex = -1;
        private int _binnedIdsIndex = -1;
        private int _otherModeIndex = -1;
        private int _otherBinVoltageIndex = -1;
        private int _otherIdsIndex = -1;
        private int _otherIndex = -1;

        private ExcelWorksheet? _excelWorksheet;
        private string _sheetName = "";

        public PowerBinningSheetCebu? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            _excelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            Initialize();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetBinnedModeFirstHeaderPosition())
            {
                return null;
            }

            //20201214 new PwrScreen no other_mode header
            GetOtherModeFirstHeaderPosition();
            //if (GetOtherModeFirstHeaderPosition() == false) return null;

            if (!GetOtherFirstHeaderPosition())
            {
                return null;
            }

            ReadHeader();

            PowerBinningSheetCebu dataList = ReadAllData();

            return dataList;
        }

        private PowerBinningSheetCebu ReadAllData()
        {
            var sheet = new PowerBinningSheetCebu(_sheetName)
            {
                BinnedModeList = GetModeList(_binnedModeStartRowNumber, _endRowNumber, _binnedModeIndex, _binnedBinVoltageIndex, _binnedIdsIndex),
                OtherModeList = GetModeList(_otherModeStartRowNumber, _endRowNumber, _otherModeIndex, _otherBinVoltageIndex, _otherIdsIndex),
                NewMethod = GetMethod()
            };

            #region get offset and spec
            for (int i = _binnedModeStartRowNumber + 1; i <= _endRowNumber; i++)
            {
                string item = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _otherIndex).Trim();
                if (item.EqualsIgnoreCase(ConHeaderOffset))
                {
                    _ = double.TryParse(EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _otherIndex + 1).Trim(), out double value);
                    sheet.Offset = value;
                }
                if (item.EqualsIgnoreCase(ConHeaderSpec))
                {
                    _ = double.TryParse(EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _otherIndex + 1).Trim(), out double value);
                    sheet.Spec = value;
                }
            }
            #endregion

            return sheet;
        }

        private List<PowerBinningSheetRow> GetModeList(int startRow, int endRow, int index, int voltageIndex, int idsIndex)
        {
            var rows = new List<PowerBinningSheetRow>();
            if (index != -1)
            {
                for (int i = startRow + 1; i <= endRow; i++)
                {
                    var row = new PowerBinningSheetRow(_sheetName) { RowNum = i, BinnedMode = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, index).Trim() };
                    if (voltageIndex != -1)
                    {
                        row.BinVoltage = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, voltageIndex).Trim();
                    }

                    if (idsIndex != -1)
                    {
                        row.Ids = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, idsIndex).Trim();
                    }

                    int startCol = new[] { index, voltageIndex, idsIndex }.Max();
                    Dictionary<string, double> factorDictionary = GetFactorDictionary(i, startRow, startCol);

                    row.FactorDictionary = factorDictionary;
                    if (row.BinnedMode?.Length != 0)
                    {
                        rows.Add(row);
                    }
                }
            }

            return rows;
        }

        private Dictionary<string, double> GetFactorDictionary(int i, int startRowNumber, int startColNumber)
        {

            var factorDictionary = new Dictionary<string, double>();
            int cnt = 1;
            string value;
            do
            {
                string key = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, startRowNumber, startColNumber + cnt).Trim();
                key = MyRegex().Replace(key, "").Trim();
                value = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, startColNumber + cnt).Trim();
                if (value?.Length != 0)
                {
                    if (double.TryParse(value, out double factor))
                    {
                        factorDictionary.TryAdd(key, factor);
                    }
                }
                cnt++;
            } while (value?.Length != 0);
            return factorDictionary;
        }

        private bool GetMethod()
        {
            var headerList = new List<string>();
            for (int col = 1; col <= _endColNumber; col++)
            {
                headerList.Add(EpplusExtensions.GetCellValue(_excelWorksheet!, 1, col).Trim());
            }
            return headerList.Contains(ConVdd0);
        }

        private void ReadHeader()
        {
            for (int i = _binnedModeIndex; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusExtensions.GetCellValue(_excelWorksheet!, _binnedModeStartRowNumber, i).Trim();
                if (string.IsNullOrEmpty(lStrHeader))
                {
                    break;
                }

                if (MyRegex1().IsMatch(lStrHeader))
                {
                    _binnedBinVoltageIndex = i;
                    continue;
                }
                if (MyRegex2().IsMatch(lStrHeader) && lStrHeader != "IDS_REF")
                {
                    _binnedIdsIndex = i;
                    continue;
                }
            }

            if (_otherModeIndex != -1)
            {
                for (int i = _otherModeIndex; i <= _endColNumber; i++)
                {
                    string lStrHeader = EpplusExtensions.GetCellValue(_excelWorksheet!, _binnedModeStartRowNumber, i).Trim();
                    if (string.IsNullOrEmpty(lStrHeader))
                    {
                        break;
                    }

                    if (MyRegex1().IsMatch(lStrHeader))
                    {
                        _otherBinVoltageIndex = i;
                        continue;
                    }
                    if (MyRegex2().IsMatch(lStrHeader) && lStrHeader != "IDS_REF")
                    {
                        _otherIdsIndex = i;
                        continue;
                    }
                }
            }
        }

        private bool GetBinnedModeFirstHeaderPosition()
        {
            for (int i = 1; i < _endRowNumber; i++)
            {
                for (int j = 1; j < _endColNumber; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet!, i, j).Trim().EqualsIgnoreCase(ConHeaderBinnedMode))
                    {
                        _binnedModeStartRowNumber = i;
                        _binnedModeIndex = j;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool GetOtherModeFirstHeaderPosition()
        {
            for (int i = 1; i < _endRowNumber; i++)
            {
                for (int j = 1; j < _endColNumber; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet!, i, j).Trim().EqualsIgnoreCase(ConHeaderOtherMode))
                    {
                        _otherModeStartRowNumber = i;
                        _otherModeIndex = j;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool GetOtherFirstHeaderPosition()
        {
            for (int i = 1; i < _endRowNumber; i++)
            {
                for (int j = 1; j < _endColNumber; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet!, i, j).Trim().EqualsIgnoreCase(ConHeaderOther))
                    {
                        _otherIndex = j;
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
                _endColNumber = _excelWorksheet!.Dimension.End.Column;
                _endRowNumber = _excelWorksheet!.Dimension.End.Row;
                return true;
            }
            return false;
        }

        private void Initialize()
        {
            _binnedModeStartRowNumber = 1;
            _otherModeStartRowNumber = 1;
            _endColNumber = 1;
            _endRowNumber = 1;

        }
    }
}
