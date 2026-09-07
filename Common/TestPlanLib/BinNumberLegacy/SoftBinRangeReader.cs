using System;
using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinNumberLegacy
{
    public class SoftBinRangeReader
    {
        private const string ConDescription = "Description";
        private const string ConRangeStart = "SW BIN Range Start";
        private const string ConRangeEnd = "SW BIN Range End";
        private const string ConRangeState = "good or bad";
        private const string ConRangeBin = "Bin #";
        private const string ConBlock = "Block";
        private const string ConCondition = "Condition";
        private const string ConHv = "HV";
        private const string ConLv = "LV";
        private const string ConNv = "NV";
        private const string ConHlv = "HLV";
        private const string ConHardBin = "Hard Bin";

        private int _columnDescription = 0;
        private int _columnRangeStart = 0;
        private int _columnRangeEnd = 0;
        private int _columnRangeState = 0;
        private int _columnRangeBin = 0;
        private int _columBlock = 0;
        private int _columnCondition = 0;
        private int _columnHardHvBin = 0;
        private int _columnHardLvBin = 0;
        private int _columnHardNvBin = 0;
        private int _columnHardHlvBin = 0;
        private int _columnHardBin = 0;

        public List<SoftBinRangeData> ReadSheet(ExcelWorksheet excelWorksheet)
        {
            List<SoftBinRangeData> rangeDatas = [];
            int startRow = 0;
            for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= excelWorksheet.Dimension.End.Column; j++)
                {
                    if (EpplusExtensions.GetCellValue(excelWorksheet, i, j).EqualsIgnoreCase(ConDescription))
                    {
                        startRow = i;
                        _columnDescription = j;
                        break;
                    }
                }
                if (startRow != 0)
                {
                    break;
                }
            }

            if (startRow == 0 || _columnDescription == 0)
            {
                throw new Exception("Error occur when Initial Soft Bin Range ");
            }

            for (int i = _columnDescription; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                GetIndex(excelWorksheet, startRow, i);
            }

            if (_columnRangeBin == 0 || _columnRangeEnd == 0 || _columnRangeStart == 0 || _columnRangeState == 0 || _columnHardBin == 0 || _columBlock == 0 || _columnCondition == 0)
            {
                throw new Exception("Error occur when Initial Soft Bin Range ");
            }

            ReadData(excelWorksheet, rangeDatas, startRow);

            return rangeDatas;
        }

        private void GetIndex(ExcelWorksheet excelWorksheet, int startRow, int i)
        {
            if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConRangeStart))
            {
                _columnRangeStart = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConRangeEnd))
            {
                _columnRangeEnd = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConRangeState))
            {
                _columnRangeState = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConRangeBin))
            {
                _columnRangeBin = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConBlock))
            {
                _columBlock = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConCondition))
            {
                _columnCondition = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConHardBin))
            {
                _columnHardBin = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConHlv))
            {
                _columnHardHlvBin = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConLv))
            {
                _columnHardLvBin = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConNv))
            {
                _columnHardNvBin = i;
            }
            else if (EpplusExtensions.GetCellValue(excelWorksheet, startRow, i).EqualsIgnoreCase(ConHv))
            {
                _columnHardHvBin = i;
            }
        }

        private void ReadData(ExcelWorksheet excelWorksheet, List<SoftBinRangeData> softBinRangeDatas, int startRow)
        {
            for (int i = startRow + 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                SoftBinRangeData data = new SoftBinRangeData
                {
                    Description = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnDescription),
                    Start = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnRangeStart),
                    End = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnRangeEnd),
                    State = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnRangeState),
                    Block = EpplusExtensions.GetCellValue(excelWorksheet, i, _columBlock).Replace("_", "").Replace(" ", ""),
                    Condition = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnCondition).Replace("_", "").Replace(" ", ""),
                    HardBin = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnHardBin)
                };
                if (_columnHardHlvBin != 0)
                {
                    data.HardHlvBin = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnHardHlvBin);
                }

                if (_columnHardHvBin != 0)
                {
                    data.HardHvBin = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnHardHvBin);
                }

                if (_columnHardLvBin != 0)
                {
                    data.HardLvBin = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnHardLvBin);
                }

                if (_columnHardNvBin != 0)
                {
                    data.HardNvBin = EpplusExtensions.GetCellValue(excelWorksheet, i, _columnHardNvBin);
                }

                if (!string.IsNullOrEmpty(data.Block) && !string.IsNullOrEmpty(data.Condition))
                {
                    softBinRangeDatas.Add(data);
                }
            }
        }
    }
}
