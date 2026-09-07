using System;
using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public class HarvestFieldReader
    {
        #region Field

        #region Const Field
        private const string HarvestFieldName = "HarvestFieldName";
        private const int MaxSearchRow = 10;
        private const int MaxSearchCol = 10;
        #endregion

        private ExcelWorksheet? _sheet;
        private int _startCol = 1;
        private int _startRow = 1;
        private int _endColNumber;
        private int _endRowNumber;
        private Headers? _headerItems;
        public static readonly List<string> HarvestFieldList = [];
        #endregion

        public void ReadFlow(ExcelWorksheet excelWorksheet)
        {
            try
            {
                _sheet = excelWorksheet;
                ReadHeader();
                ReadData();
            }
            catch (Exception e)
            {
                throw new Exception("Read HarvestField sheet Failed! " + e.StackTrace);
            }
        }

        private void ReadData()
        {
            int i = _startRow + 1;
            try
            {
                int startCol = _headerItems!.GetHeaderIndex(HarvestFieldName);

                for (; i <= _endRowNumber; i++)
                {
                    HarvestFieldList.Add(_sheet!.Cells[i, startCol].Value.ToString()!);
                }
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Read harvest field at row: {0} " + e.StackTrace, i));
            }
        }

        #region Init

        private void ReadHeader()
        {
            SetStartIndex();
            GetHeaderItems();
        }

        private void SetStartIndex()
        {
            for (int i = 1; i < MaxSearchRow; i++)
            {
                for (int j = 1; j < MaxSearchCol; j++)
                {
                    string startHeader = EpplusExtensions.GetCellValue(_sheet!, i, j);
                    if (!string.IsNullOrEmpty(startHeader) &&
                        startHeader.Trim().EqualsIgnoreCase(HarvestFieldName))
                    {
                        _startCol = j;
                        _startRow = i;
                        break;
                    }
                }
            }
            _endColNumber = _sheet!.Dimension.End.Column;
            _endRowNumber = _sheet.Dimension.End.Row;
        }

        private void GetHeaderItems()
        {
            _headerItems = new Headers();
            for (int i = _startCol; i <= _endColNumber; i++)
            {
                string header = EpplusExtensions.GetCellValue(_sheet!, _startRow, i);
                if (!string.IsNullOrEmpty(header))
                {
                    _headerItems.AddHeaderItem(header, i);
                }
            }
        }
        #endregion
    }
}
