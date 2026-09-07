using System;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.DataStruct
{
    public class IoLevelsSheetReader
    {
        private const string ConHeaderType = "Type";
        private const string ConHeaderPinName = "PinName";
        private const string ConHeaderFsdd = "FS/DD";
        private const string ConHeaderDomain = "Domain";
        private const string ConHeaderVdd = "VDD";
        private const string ConHeaderVih = "VIH";
        private const string ConHeaderVil = "VIL";
        private const string ConHeaderVoh = "VOH";
        private const string ConHeaderVol = "VOL";

        private ExcelWorksheet? _excelWorksheet;
        private string _sheetName = "";
        private IoLevelsSheet? _iOLevelsSheet;
        private object[,]? _dataArray;
        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _pinGroupIndex = -1;
        private int _pinIndex = -1;
        private int _fsddIndex = -1;

        public IoLevelsSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            _excelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _iOLevelsSheet = new IoLevelsSheet(_sheetName);

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

            _dataArray = RangeToArray(_excelWorksheet!.Cells[_excelWorksheet!.Dimension.Start.Row, _excelWorksheet!.Dimension.Start.Column, _excelWorksheet!.Dimension.End.Row, _excelWorksheet!.Dimension.End.Column]);

            //ReplaceDomain();

            _iOLevelsSheet = ReadSheetData();

            return _iOLevelsSheet!;
        }

        //private void FillEmptyCell()
        //{
        //    object lastContext = null;
        //    for (int j = _startColNumber + 1; j <= _endColNumber; j++)
        //    {
        //        string headerName = ExcelOperation.GetMergerdCellValue(_excelWorksheet!, _startRowNumber, j).Trim();
        //        for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
        //        {
        //            if (headerName.Equals(ConHeaderVdd, StringComparison.OrdinalIgnoreCase) ||
        //                headerName.Equals(ConHeaderVih, StringComparison.OrdinalIgnoreCase) ||
        //                headerName.Equals(ConHeaderVil, StringComparison.OrdinalIgnoreCase) ||
        //                headerName.Equals(ConHeaderVoh, StringComparison.OrdinalIgnoreCase) ||
        //                headerName.Equals(ConHeaderVol, StringComparison.OrdinalIgnoreCase))
        //            {
        //                if (_dataArray![i - 1, j - 1] == null)
        //                    _dataArray![i - 1, j - 1] = lastContext;
        //                else
        //                    lastContext = _dataArray![i - 1, j - 1];
        //            }
        //        }
        //    }
        //}

        private void FillEmptyCell()
        {
            foreach (IoLevelsRow row in _iOLevelsSheet!.RowList)
            {
                foreach (IoLevelsItem item in row.IoLevelDate)
                {
                    IoLevelsItem? firstRow = _iOLevelsSheet!.RowList.SelectMany(y => y.IoLevelDate)
                        .ToList().Find(x => x.Level == item.Level && x.Domain == item.Domain && !string.IsNullOrEmpty(x.Vdd));
                    if (string.IsNullOrEmpty(item.Vdd) && firstRow != null)
                    {
                        item.Vdd = firstRow.Vdd;
                        item.Vih = firstRow.Vih;
                        item.Vil = firstRow.Vil;
                        item.Voh = firstRow.Voh;
                        item.Vol = firstRow.Vol;
                    }
                }
            }
        }

        private void GetTheSameDomain()
        {
            foreach (IoLevelsRow row in _iOLevelsSheet!.RowList)
            {
                row.IsTheSameRow = true;
                row.IsGroupPin = true;
                foreach (IoLevelsItem data in row.IoLevelDate)
                {
                    if (!data.Domain.EqualsIgnoreCase(row.IoLevelDate[0].Domain))
                    {
                        data.IsSameDomain = false;
                    }

                    IoLevelsItem? ioLevelDate = _iOLevelsSheet!.RowList.SelectMany(x => x.IoLevelDate).ToList()
                        .Find(x => x.Level.EqualsIgnoreCase(data.Level)
                            && x.Domain.EqualsIgnoreCase(data.Domain));
                    if (ioLevelDate != null && !data.Vdd.EqualsIgnoreCase(ioLevelDate.Vdd))
                    {
                        row.IsGroupPin = false;
                    }

                    if (ioLevelDate != null && !data.Vih.EqualsIgnoreCase(ioLevelDate.Vih))
                    {
                        row.IsGroupPin = false;
                    }

                    if (ioLevelDate != null && !data.Vil.EqualsIgnoreCase(ioLevelDate.Vil))
                    {
                        row.IsGroupPin = false;
                    }

                    if (ioLevelDate != null && !data.Voh.EqualsIgnoreCase(ioLevelDate.Voh))
                    {
                        row.IsGroupPin = false;
                    }

                    if (ioLevelDate != null && !data.Vol.EqualsIgnoreCase(ioLevelDate.Vol))
                    {
                        row.IsGroupPin = false;
                    }
                }

                if (row.IoLevelDate.Select(x => x.Domain).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                    row.IsGroupPin = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vdd).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vih).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vil).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Voh).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vol).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }

                row.Domain = row.IoLevelDate.Select(x => x.Domain).Distinct().First();
            }
        }

        private IoLevelsSheet ReadSheetData()
        {
            IoLevelsItem ioLevelsItem = new IoLevelsItem();
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                IoLevelsRow row = new IoLevelsRow(_sheetName)
                {
                    RowNum = i
                };
                if (_pinGroupIndex != -1)
                {
                    row.Type = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _pinGroupIndex).Trim();
                }

                if (_pinIndex != -1)
                {
                    row.PinName = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _pinIndex).Trim();
                }

                if (_fsddIndex != -1)
                {
                    row.Fsdd = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, i, _fsddIndex).Trim();
                }

                PopulateRowColumns(row, i, ref ioLevelsItem);

                _iOLevelsSheet!.RowList.Add(row);
            }

            FillEmptyCell();

            GetTheSameDomain();

            return _iOLevelsSheet!;
        }

        private void PopulateRowColumns(IoLevelsRow ioLevelsRow, int i, ref IoLevelsItem ioLevelsItem)
        {
            int cnt = 0;
            for (int j = _startColNumber + 1; j <= _endColNumber; j++)
            {
                string levelName = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, _startRowNumber - 1, j).Trim();
                string headerName = EpplusExtensions.GetMergedCellValue(_excelWorksheet!, _startRowNumber, j).Trim();

                if (!string.IsNullOrEmpty(levelName) && headerName.EqualsIgnoreCase(ConHeaderDomain))
                {
                    if (cnt != 0)
                    {
                        ioLevelsRow.IoLevelDate.Add(ioLevelsItem);
                    }

                    ioLevelsItem = new IoLevelsItem(levelName);
                    cnt++;
                }

                SetIoLevelsItemValue(ioLevelsItem, headerName, i, j);

                if (j == _endColNumber)
                {
                    if (cnt != 0)
                    {
                        ioLevelsRow.IoLevelDate.Add(ioLevelsItem);
                    }

                    ioLevelsItem = new IoLevelsItem(levelName);
                    cnt++;
                }
            }
        }

        private void SetIoLevelsItemValue(IoLevelsItem ioLevelsItem, string headerName, int i, int j)
        {
            string cellValue = GetTrimmedCellValue(i, j);
            if (headerName.EqualsIgnoreCase(ConHeaderDomain))
            {
                ioLevelsItem.Domain = cellValue;
            }
            else if (headerName.EqualsIgnoreCase(ConHeaderVdd))
            {
                ioLevelsItem.Vdd = cellValue.Replace(" ", "");
            }
            else if (headerName.EqualsIgnoreCase(ConHeaderVih))
            {
                ioLevelsItem.Vih = cellValue.Replace(" ", "");
            }
            else if (headerName.EqualsIgnoreCase(ConHeaderVil))
            {
                ioLevelsItem.Vil = cellValue.Replace(" ", "");
            }
            else if (headerName.EqualsIgnoreCase(ConHeaderVoh))
            {
                ioLevelsItem.Voh = cellValue.Replace(" ", "");
            }
            else if (headerName.EqualsIgnoreCase(ConHeaderVol))
            {
                ioLevelsItem.Vol = cellValue.Replace(" ", "");
            }
        }

        private string GetTrimmedCellValue(int i, int j)
        {
            return _dataArray![i - 1, j - 1]?.ToString()?.Trim() ?? "";
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string header = EpplusExtensions.GetCellValue(_excelWorksheet!, _startRowNumber, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderType))
                {
                    _pinGroupIndex = i;
                    _iOLevelsSheet!.HeaderIndex.Add(ConHeaderType, i);
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPinName))
                {
                    _pinIndex = i;
                    _iOLevelsSheet!.HeaderIndex.Add(ConHeaderPinName, i);
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderFsdd))
                {
                    _fsddIndex = i;
                    _iOLevelsSheet!.HeaderIndex.Add(ConHeaderFsdd, i);
                    continue;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i < rowNum; i++)
            {
                for (int j = 1; j < colNum; j++)
                {
                    if (EpplusExtensions.GetCellValue(_excelWorksheet!, i, j).Trim().EqualsIgnoreCase(ConHeaderPinName))
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
            _pinGroupIndex = -1;
            _pinIndex = -1;
            _fsddIndex = -1;
        }

        private static object[,]? RangeToArray(ExcelRange excelRange)
        {
            if (excelRange == null)
            {
                return null;
            }

            int startRow = excelRange.Start.Row;
            int startCol = excelRange.Start.Column;
            int endRow = excelRange.End.Row;
            int endCol = excelRange.End.Column;

            object[,] strAry = new object[excelRange.Rows, excelRange.Columns];

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    strAry[row - 1, col - 1] = excelRange[row, col].Value;
                }
            }
            return strAry;
        }
    }
}
