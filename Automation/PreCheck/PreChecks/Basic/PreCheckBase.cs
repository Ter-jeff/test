using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using OfficeOpenXml;

namespace Automation.PreCheck.PreChecks.Basic
{
    public abstract class PreCheckBase
    {
        protected readonly string SheetName;

        protected ExcelWorkbook ExcelWorkbook { get; set; }
        internal ExcelWorksheet ExcelWorksheet { get; set; }

        protected const int ConMaxStartRowIndex = 10;
        protected const int ConMaxStartColumnIndex = 10;

        public int StartRow = -1;
        public int StopRow = -1;
        public int StartColumn = -1;
        public int StopColumn = -1;
        protected string FirstHeader = "";

        protected PreCheckBase(ExcelWorkbook excelWorkbook, string sheetName)
        {
            ExcelWorkbook = excelWorkbook;
            SheetName = sheetName;
        }

        protected internal abstract bool CheckHeaders();
        protected internal abstract bool CheckFormat();
        protected internal abstract bool CheckBusiness();

        public bool CheckMain()
        {
            bool checkResult = CheckExist();
            if (!checkResult)
            {
                return false;
            }

            checkResult = CheckHeaders();
            if (!checkResult)
            {
                return false;
            }

            _ = CheckFormat();

            checkResult = CheckBusiness();
            if (!checkResult)
            {
                return false;
            }

            return true;
        }

        protected internal virtual bool CheckExist()
        {
            bool checkResult = CheckSheetName();
            if (!checkResult)
            {
                return false;
            }

            StartRow = ExcelWorksheet.Dimension.Start.Row;
            StopRow = ExcelWorksheet.Dimension.End.Row;
            StartColumn = ExcelWorksheet.Dimension.Start.Column;
            StopColumn = ExcelWorksheet.Dimension.End.Column;

            checkResult = CheckFirstHeaderLocation();
            return checkResult;
        }

        protected bool CheckSheetName()
        {
            if (ExcelWorkbook.Worksheets.Any(o => string.Equals(o.Name, SheetName, StringComparison.CurrentCultureIgnoreCase)))
            {
                ExcelWorksheet = ExcelWorkbook.Worksheets[SheetName];
            }
            else
            {
                ErrorReportManager.AddError(BasicErrorType.E_MissingBlock_07, SheetName, 1, 1, "Sheet: " + SheetName + " Not Exist.", new string[] { SheetName });
                return false;
            }

            return true;
        }

        protected bool CheckFirstHeaderLocation()
        {
            StopRow = ExcelWorksheet.Dimension.End.Row;
            StopColumn = ExcelWorksheet.Dimension.End.Column;
            if (FirstHeader != "")
            {
                for (int i = StartRow; i <= ConMaxStartRowIndex; i++)
                {
                    for (int j = StartColumn; j <= ConMaxStartColumnIndex; j++)
                    {
                        string value = ExcelWorksheet.GetCellValue(i, j);
                        if (CellDiff.IsLiked(value, FirstHeader))
                        {
                            StartRow = i;
                            StartColumn = j;
                            return true;
                        }
                    }
                }
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_20, SheetName, 1, 1, $"Key Don't Exist In {SheetName} : " + FirstHeader, new string[] { SheetName, FirstHeader });
                return false;
            }

            return true;
        }

        protected bool CheckColumnHeadByList(int row, List<string> headers)
        {
            bool result = true;
            for (int i = 0; i <= headers.Count - 1; i++)
            {
                string header = headers[i];
                bool flag = false;
                for (int j = StartColumn; j <= ExcelWorksheet.Dimension.End.Column; j++)
                {
                    string value = ExcelWorksheet.GetCellValue(row, j);
                    if (CellDiff.IsLiked(value, header))
                    {
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatError_21, ExcelWorksheet.Name, row, 1, "Must Exist Item :" + header + " Not Exist.", new string[] { header });
                    result = false;
                }
            }

            return result;
        }
    }
}
