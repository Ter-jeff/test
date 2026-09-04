using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport.Epplus;
using CommonLib.ErrorReport.Interop;
using CommonLib.Extension;

using Microsoft.Office.Interop.Excel;

using OfficeOpenXml;

namespace CommonLib.ErrorReport
{
    public class ErrorInstance
    {
        private ConcurrentBag<Error> _errors;
        public static ErrorInstance Instance = new ErrorInstance();

        private ErrorInstance()
        {
            _errors = new ConcurrentBag<Error>();
        }

        #region Interop
        public void GenErrorReport(Workbook workbook, string errorReportName, string summaryReport = "SummaryReport")
        {
            List<Error> errorList = GetErrorList();
            var errorReport = new ReportDefault(errorList);
            errorReport.WriteReport(workbook, errorReportName);
        }
        #endregion

        #region Epplus

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
        }

        public void GenErrorReport(ExcelPackage excelPackage, string errorReportName)
        {
            var workbooks = new List<ExcelWorkbook> { excelPackage.Workbook };
            if (GetErrorCount() > 0)
            {
                List<Error> errorList = GetSortedErrorList();
                var errorReport = new ErrorReportEpplus(errorList);
                errorReport.WriteReport(workbooks, errorReportName, "");
            }
        }

        public void GenErrorReportRaw(ExcelPackage excelPackage, string errorReportName)
        {
            var workbooks = new List<ExcelWorkbook> { excelPackage.Workbook };
            if (GetErrorCount() > 0)
            {
                List<Error> errorList = GetSortedErrorList();
                var errorReport = new ErrorReportEpplus(errorList);
                errorReport.WriteReportRaw(workbooks, errorReportName, "");
            }
        }
        #endregion

        public List<Error> GetErrorList()
        {
            return _errors.ToList();
        }

        public List<Error> GetSortedErrorList()
        {
            var distinctErrors = _errors.GroupBy(x => new { x.SheetName, x.Message, x.RowNum, x.ColNum, x.ErrorType })
                .Select(g => g.First()).ToList();

            return distinctErrors.OrderBy(x => x.SheetName).ThenBy(x => x.RowNum).ThenBy(x => x.ColNum).ThenBy(x => x.ErrorType.ToString()).ThenBy(x => x.Message).ToList();
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

        public void Clear()
        {
            _errors = new ConcurrentBag<Error>();
        }

        public int GetErrorCount()
        {
            return _errors.Count;
        }

        public bool Exist(string message)
        {
            return _errors.Any(x => x.Message.Equals(message, StringComparison.CurrentCultureIgnoreCase));
        }

        public int GetErrorCountByType(Type type)
        {
            return GetErrorsByType(type).Count;
        }

        public List<Error> GetErrorsByType(Type type)
        {
            return _errors.Where(a => a.ErrorType.GetType() == type).ToList();
        }

        public List<Error> GetErrorsByType(List<Error> errorItems, Type type)
        {
            return errorItems.Where(a => a.ErrorType.GetType() == type).OrderBy(x => x.SheetName).ThenBy(x => x.RowNum).ToList();
        }

        public List<Type> GetErrorTypeList()
        {
            var typeList = _errors.GroupBy(p => p.ErrorType.GetType()).Select(p => p.Key).ToList();
            return typeList;
        }
    }
}
