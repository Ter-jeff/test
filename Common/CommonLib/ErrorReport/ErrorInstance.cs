using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport.Base;
using CommonLib.Extension;

using OfficeOpenXml;

namespace CommonLib.ErrorReport
{
    public class ErrorInstance
    {
        private ConcurrentBag<Error> _errors;

        public ErrorInstance()
        {
            _errors = [];
        }

        public void GenErrorReport(ExcelPackage excelPackage, List<string> copyFiles, string errorReportName, string summaryReport = "SummaryReport")
        {
            ExcelWorkbook workbook = excelPackage.Workbook;
            workbook.CopyWorkSheets(copyFiles);

            var workbooks = new List<ExcelWorkbook> { excelPackage.Workbook };
            if (GetErrorCount() > 0)
            {
                List<Error> errorList = GetErrorList();
                var errorReport = new ErrorReportEpplus(errorList);
                errorReport.WriteReport(workbooks, errorReportName, summaryReport);
            }
            else
            {
                ExcelWorksheet worksheet = workbook.Worksheets.Add(errorReportName);
                worksheet.Cells[1, 1].Value = "No error occurs";
                worksheet.Cells.TryAutoFitColumns();
                ExcelWorksheet firstSheet = workbook.Worksheets.First();
                workbook.Worksheets.MoveBefore(worksheet.Name, firstSheet.Name);
            }
        }

        public void GenErrorReport(ExcelPackage excelPackage, string errorReportName)
        {
            var workbooks = new List<ExcelWorkbook> { excelPackage.Workbook };
            if (GetErrorCount() > 0)
            {
                List<Error> errorList = GetSortedErrors();
                var errorReport = new ErrorReportEpplus(errorList);
                errorReport.WriteReport(workbooks, errorReportName, "");
            }
        }

        public List<Error> GetErrorList()
        {
            return [.. _errors];
        }

        public List<Error> GetSortedErrors()
        {
            var distinctErrors = _errors.GroupBy(x => new { x.SheetName, x.Message, x.RowNum, x.ColNum, x.ErrorCode.EnumErrorBehavior, x.ErrorCode.EnumErrorTarget })
                .Select(g => g.First()).ToList();

            return [.. distinctErrors.OrderBy(x => x.SheetName)
                .ThenBy(x => x.RowNum)
                .ThenBy(x => x.ColNum)
                .ThenBy(x => x.ErrorCode.EnumErrorBehavior.ToString())
                .ThenBy(x => x.ErrorCode.EnumErrorTarget.ToString())
                .ThenBy(x => x.Message)];
        }

        public void AddError(Error error)
        {
            _errors.Add(error);
        }

        public void AddErrors(List<Error> errors)
        {
            foreach (Error error in errors)
            {
                _errors.Add(error);
            }
        }

        public void ClearErrors()
        {
            _errors = [];
        }

        public int GetErrorCount()
        {
            return _errors.Count;
        }

        public bool Exist(string message)
        {
            return _errors.Any(x => x.Message.EqualsIgnoreCase(message));
        }

        public int GetErrorCountByCategory(EnumErrorCategory enumErrorCategory)
        {
            return _errors.Count(a => a.ErrorCode.EnumErrorCategory == enumErrorCategory);
        }
    }
}
