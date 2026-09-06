using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Static;

using OfficeOpenXml;

namespace CommonLib.ErrorReport
{
    public static class ErrorReportManager
    {
        private static readonly ErrorInstance _errorInstance = new();

        public static void GenErrorReport(ExcelPackage excelPackage, List<string> copyFiles, string errorReportName, string summaryReport = "SummaryReport")
        {
            _errorInstance.GenErrorReport(excelPackage, copyFiles, errorReportName, summaryReport);
        }

        public static void GenErrorReport(ExcelPackage excelPackage, string errorReportName)
        {
            _errorInstance.GenErrorReport(excelPackage, errorReportName);
        }

        public static void AddError(ErrorCode errorCode, string sheetName, int rowNum, int columnNum, string message, ErrorInfo? errorInfo = null)
        {
            if (errorCode.ErrorLevel == EnumErrorLevel.Critical && ErrorReportStatic.StopByCritical)
            {
                Environment.Exit(-1);
            }
            _errorInstance.AddError(new Error(errorCode, errorCode.ErrorLevel, sheetName, rowNum, columnNum < 1 ? 0 : columnNum, message, errorInfo));
        }

        public static void AddError(ErrorCode errorCode, string sheetName, int rowNum, int columnNum, string message, string[] messageArgs, ErrorInfo? errorInfo = null)
        {
            if (errorCode.ErrorLevel == EnumErrorLevel.Critical && ErrorReportStatic.StopByCritical)
            {
                Environment.Exit(-1);
            }
            _errorInstance.AddError(new Error(errorCode, sheetName, rowNum, columnNum < 1 ? 0 : columnNum, messageArgs, errorInfo));
        }

        public static void AddError(ErrorCode errorCode, string sheetName, int rowNum, int columnNum, string[]? messageArgs = null, ErrorInfo? errorInfo = null)
        {
            if (errorCode.ErrorLevel == EnumErrorLevel.Critical && ErrorReportStatic.StopByCritical)
            {
                Environment.Exit(-1);
            }
            _errorInstance.AddError(new Error(errorCode, sheetName, rowNum, columnNum < 1 ? 0 : columnNum, messageArgs, errorInfo));
        }

        public static bool Contains(string message)
        {
            return _errorInstance.Exist(message);
        }

        public static void ClearErrors()
        {
            _errorInstance.ClearErrors();
        }

        public static int GetErrorCountByCategory(EnumErrorCategory enumErrorCategory)
        {
            return _errorInstance.GetErrorCountByCategory(enumErrorCategory);
        }

        public static List<Error> GetErrorList()
        {
            return _errorInstance.GetErrorList();
        }

        public static int GetErrorCount()
        {
            return _errorInstance.GetErrorCount();
        }

        public static void PrintErrors(string file)
        {
            var lines = new List<string>();
            var headers = new List<string> { "SheetName", "ErrorBehavior", "ErrorTarget", "Link", "ErrorLevel", "RowNum", "ColNum", "Message", "Guidance", "Comments" };
            lines.Add(string.Join("\t", headers));
            foreach (Error error in _errorInstance.GetErrorList())
            {
                var list = new List<string>
                {
                    error.SheetName,
                    error.ErrorCode.EnumErrorBehavior?.ToString() ??  "",
                    error.ErrorCode.EnumErrorTarget?.ToString() ?? "",
                    error.Link,
                    error.ErrorLevel.ToString(),
                    error.RowNum.ToString(),
                    error.ColLetter,
                    error.Message,
                    error.ErrorCode.Guidance.ToString(),
                    string.Join("\t", error.Comments)
                };
                string line = string.Join("\t", list);
                lines.Add(line);
            }
            File.WriteAllLines(file, lines);
        }

        public static void RemoveNonUsedPattern(HashSet<string> usedPatternList)
        {
            if (usedPatternList == null || usedPatternList.Count == 0)
            {
                return;
            }

            List<Error> keep = [];
            foreach (Error error in _errorInstance.GetErrorList())
            {
                bool isMissingPattern = error.ErrorCode?.FullCode == PatternMissing.E_MissingPatternFile_01.FullCode;
                if (isMissingPattern)
                {
                    if (error.Pattern != null && usedPatternList.Contains(error.Pattern))
                    {
                        keep.Add(error);
                    }
                }
                else
                {
                    keep.Add(error);
                }
            }
            _errorInstance.ClearErrors();
            _errorInstance.AddErrors(keep);
        }

        public static List<Error> GetSortedErrors()
        {
            return _errorInstance.GetSortedErrors();
        }
    }
}
