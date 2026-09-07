using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace CommonReaderLib
{
    public abstract class MySheetReader<T>
    {
        protected int EndCol = -1;
        protected int EndRow = -1;
        protected ExcelWorksheet ExcelWorksheet = null!;
        protected int StartCol = -1;
        protected int StartRow = -1;

        public abstract T? ReadSheet(ExcelWorksheet excelWorksheet);

        protected bool GetDimensions()
        {
            if (ExcelWorksheet.Dimension != null)
            {
                StartCol = ExcelWorksheet.Dimension.Start.Column;
                StartRow = ExcelWorksheet.Dimension.Start.Row;
                EndCol = ExcelWorksheet.Dimension.End.Column;
                EndRow = ExcelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        protected bool GetFirstHeaderPosition(string header)
        {
            return GetFirstHeaderPosition([header]);
        }

        protected bool GetFirstHeaderPosition(List<string> headers)
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    string cellValue = ExcelWorksheet.GetCellValue(i, j).Trim();
                    if (string.IsNullOrEmpty(cellValue))
                    {
                        continue;
                    }

                    foreach (string header in headers)
                    {
                        if (cellValue.EqualsIgnoreCase(header))
                        {
                            StartRow = i;
                            return true;
                        }
                    }
                }
            }

            return false;
        }
    }
}
