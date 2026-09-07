using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.DataStruct
{
    public class IoContiReader
    {
        #region Field
        private const int MaxSearchRow = 10;
        private const int MaxSearchCol = 10;
        private ExcelWorksheet? _ioContiWorksheet;
        private IoContiSheet? _ioContiSheet;
        private int _startColNum = 1;
        private int _startRowNum = 1;
        private int _endColNum = 1;
        private int _endRowNum = 1;
        private Headers? _headerItems;

        #endregion

        #region Public Function
        public IoContiSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            _ioContiWorksheet = excelWorksheet;
            _ioContiSheet = new IoContiSheet
            {
                Name = excelWorksheet.Name
            };

            GetDimension();

            ReadHeader();

            ReadData();

            return _ioContiSheet!;
        }
        #endregion

        #region Get dimension

        /// <summary>
        /// GetDimension
        /// </summary>
        private void GetDimension()
        {
            for (int i = 1; i < MaxSearchRow; i++)
            {
                for (int j = 1; j < MaxSearchCol; j++)
                {
                    string startHeader = GetCellValue(i, j);
                    if (startHeader.Length != 0 &&
                        startHeader.Trim().EqualsIgnoreCase(IoContiRow.ConHeaderBumpName))
                    {
                        _startColNum = j;
                        _startRowNum = i;
                        break;
                    }
                }
            }
            _endColNum = _ioContiWorksheet!.Dimension.End.Column;
            _endRowNum = _ioContiWorksheet!.Dimension.End.Row;
        }

        #endregion
        #region Read Header

        /// <summary>
        /// ReadHeader
        /// </summary>
        private void ReadHeader()
        {
            _headerItems = new Headers();
            for (int i = _startColNum; i <= _endColNum; i++)
            {
                string header = GetCellValue(_startRowNum, i);
                if (header.Length != 0)
                {
                    _headerItems!.AddHeaderItem(header, i);
                }
            }
        }
        #endregion

        #region Read Data

        private void ReadData()
        {
            int lIBumpNameColNum = _headerItems!.GetHeaderIndex(IoContiRow.ConHeaderBumpName);
            int lIBallColNum = _headerItems!.GetHeaderIndex(IoContiRow.ConHeaderBallName, false);
            int lIIoVolColNum = _headerItems!.GetHeaderIndex(IoContiRow.ConHeaderIoVoltage);
            int lIFsDdColNum = _headerItems!.GetHeaderIndex(IoContiRow.ConHeaderFsDd);
            int lIChipleteColNum = _headerItems!.GetHeaderIndex(IoContiRow.ConHeaderChiplet, false);
            int lICellNameColNum = _headerItems!.GetHeaderIndex(IoContiRow.ConHeaderCellName, false);
            for (int i = _startRowNum + 1; i <= _endRowNum; i++)
            {
                var lRowTemp = new IoContiRow
                {
                    RowNum = i,
                    BumpName = FormatPinName(GetCellValue(i, lIBumpNameColNum))
                };
                if (lIBallColNum > 0)   // Not must to have this column
                {
                    lRowTemp.BallName = FormatPinName(GetCellValue(i, lIBallColNum));
                }

                lRowTemp.IoVoltage = GetCellValue(i, lIIoVolColNum);
                lRowTemp.FsDd = GetCellValue(i, lIFsDdColNum).TrimStart().TrimEnd();
                if (lICellNameColNum > 0)   // Not must to have this column
                {
                    lRowTemp.CellName = GetCellValue(i, lICellNameColNum);
                }

                if (!(lRowTemp.FsDd == "DD" || lRowTemp.FsDd == "FS" || lRowTemp.FsDd == "SCR"))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_15, _ioContiWorksheet!.Name, i, i, $"Incorrect Value : {lRowTemp.FsDd} in column FS/DD", [lRowTemp.FsDd]);
                    continue;
                }
                // 2017 07 11 for Cyprus used if IO pin only exist in FT 
                if (lIBallColNum != 0 && lRowTemp.BumpName.Length == 0 && lRowTemp.BallName.Length != 0)
                {
                    lRowTemp.BumpName = lRowTemp.BallName;
                }
                if (lIChipleteColNum > 0)   // Not must to have this column
                {
                    lRowTemp.Chiplet = GetCellValue(i, lIChipleteColNum);
                }

                _ioContiSheet!.AddRow(lRowTemp);
            }
        }

        /// <summary>
        /// Delete "[" and "]" for pinGroup name
        /// </summary>
        /// <param name="pPinName">pPinName</param>
        /// <returns>pinName</returns>
        private static string FormatPinName(string pPinName)
        {
            string lStrTargetPinName = pPinName;
            if (lStrTargetPinName.Contains('['))
            {
                lStrTargetPinName = lStrTargetPinName.Replace("[", "");
            }
            if (lStrTargetPinName.Contains(']'))
            {
                lStrTargetPinName = lStrTargetPinName.Replace("]", "");
            }

            return lStrTargetPinName;
        }
        #endregion

        #region Comon Function

        /// <summary>
        /// Get cell value
        /// </summary>
        /// <param name="rowNumber">rowNumber</param>
        /// <param name="columnNumber">columnNumber</param>
        /// <returns></returns>
        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = _ioContiWorksheet!.Cells[rowNumber, columnNumber].Value;
            return value?.ToString() ?? "";
        }
        #endregion
    }
}
