using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using CommonLib.ErrorReport.Base;
using CommonLib.Utility;

using Microsoft.Office.Interop.Excel;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace CommonLib.ErrorReport.Interop
{
    public abstract class ErrorReportInterop
    {
        protected List<Error> ErrorList;

        protected string ReportName
        {
            get
            {
                if (ErrorList.Count == 0)
                {
                    return "";
                }

                return ErrorList[0].ErrorType.GetType().Name + "Report";
            }
        }

        protected ErrorReportInterop(List<Error> errorList)
        {
            ErrorList = errorList;
        }

        public void WriteReport(Workbook workbook, string errorReportName)
        {
            List<Error> errors = ErrorList;
            if (workbook.IsSheetExist(errorReportName))
            {
                Application app = workbook.Parent;
                app.DisplayAlerts = false;
                workbook.Worksheets[errorReportName].Delete();
                app.DisplayAlerts = true;
            }
            if (errors.Any())
            {
                Worksheet worksheet = workbook.AddSheet(errorReportName);
                Range range = worksheet.Cells[1, 1];
                var printErrors = new List<object[]>();
                object[] array = new object[] { "SheetName", "ErrorType", "Link", "ErrorLevel", "RowNum", "ColNum", "Message" };
                printErrors.Add(array);
                foreach (Error error in errors.OrderBy(x => x.ErrorType))
                {
                    printErrors.Add(new object[] {error.SheetName, error.ErrorType.ToString(),
                        error.Link, error.ErrorLevel.ToString(), error.RowNum.ToString(), error.ColNum.ToString(), error.Message });
                }
                range.LoadFromArray(printErrors);
                worksheet.Columns.AutoFit();
                worksheet.Select();
                //worksheet.Cells["1:1"].AutoFilter = true;
            }
        }

        public abstract void Write(Workbook workbook);

        protected static void WriteSummary(Workbook wBook, string reportSheetName, int errorCount, int warningCount = 0)
        {
            bool flag = wBook.Worksheets.Cast<Worksheet>().Any(sheet => sheet.Name.Equals("SummaryReport", StringComparison.OrdinalIgnoreCase));
            if (!flag)
            {
                return;
            }

            Worksheet wSheet = wBook.Worksheets["SummaryReport"];
            int startRow = wSheet.UsedRange.Rows.Count + 1;
            wSheet.Cells[startRow, 1] = reportSheetName;
            wSheet.Cells[startRow, 2] = errorCount;
            wSheet.Cells[startRow, 3] = warningCount;
            Range linkRange = wSheet.Cells[startRow, 1];
            wSheet.Hyperlinks.Add(linkRange, "#" + "'" + reportSheetName + "'" + "!A1", Type.Missing, Type.Missing, Type.Missing);
            wSheet.Columns["A:B"].AutoFit();
            wSheet.Activate();
        }

        protected void WriteReport(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            try
            {
                WriteErrors(workbook, ReportName, ErrorList);
            }
            catch (Exception e)
            {
                throw new Exception("Write General ErrorReport failed for " + ReportName + "  " + e.StackTrace);
            }
        }

        private void WriteErrors(Workbook wBook, string reportSheetName, List<Error> errorList)
        {
            if (errorList.Count == 0)
            {
                return;
            }

            // Delete sheet if alreadey exists
            foreach (Worksheet sheet in wBook.Worksheets)
            {
                if (sheet.Name == reportSheetName)
                {
                    wBook.Application.DisplayAlerts = false;
                    sheet.Delete();
                    break;
                }
            }
            bool flag = wBook.Worksheets.Cast<Worksheet>().Any(sheet => sheet.Name.Equals("SummaryReport", StringComparison.OrdinalIgnoreCase));
            Worksheet wSheet = wBook.Worksheets.Add(flag ? wBook.Worksheets[2] : wBook.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
            wSheet.Name = reportSheetName;

            dynamic condtion = wSheet.Range["B:B"].FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, ErrorLevel.Error.ToString());
            condtion.Interior.Pattern = XlPattern.xlPatternSolid;
            condtion.Interior.Color = Color.Red;

            string[] headers = new string[8] { "ErrorType", "Level", "Link", "SheetName", "Row", "Col", "ErrorMessage", "Count" };
            dynamic headerRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
            headerRange.Value2 = headers;
            headerRange.Font.Bold = true;

            var sheetNames = wBook.Worksheets.Cast<Worksheet>().Select(x => x.Name).ToList();

            int currentRow = 2;
            int count = 0;
            int errorCount = 0;
            int warnCount = 0;
            foreach (
                KeyValuePair<ErrorLevel, List<Error>> spilterrorList in errorList.GroupBy(x => x.ErrorLevel).ToDictionary(x => x.Key, x => x.ToList()).OrderByDescending(x => x.Key))
            {
                count = 0;
                List<object> typeList = GetErrorSubType(spilterrorList.Value);
                foreach (object errorType in typeList)
                {

                    int start = currentRow;

                    List<Error> errors = GetErrorsBySubType(errorType, spilterrorList.Value);
                    string[,] content = new string[errors.Count, 7];
                    for (int i = 0; i < errors.Count; i++)
                    {
                        content[i, 0] = errors[i].ErrorType.ToString();
                        content[i, 1] = errors[i].ErrorLevel.ToString();
                        if (sheetNames.Exists(p => p.Equals(errors[i].SheetName, StringComparison.OrdinalIgnoreCase)) &&
                            errors[i].RowNum > 0)
                        {
                            content[i, 2] = errors[i].Link;
                        }

                        content[i, 3] = errors[i].SheetName;
                        content[i, 4] = errors[i].RowNum.ToString();
                        content[i, 5] = errors[i].ColNum.ToString();
                        content[i, 6] = errors[i].Message;
                    }
                    wSheet.Range[wSheet.Cells[start, 1], wSheet.Cells[errors.Count + start - 1, 7]].Value2 = content;
                    //PaintColor(wSheet, start, errors);

                    #region Add comment

                    foreach (Error error in errors)
                    {
                        int commentCount = 0;
                        foreach (string comment in error.Comments)
                        {
                            wSheet.Cells[currentRow, 9 + commentCount].Value = comment;
                            commentCount++;
                        }
                        //if (sheetNames.Exists(p => p.Equals(error.SheetName, StringComparison.OrdinalIgnoreCase)) &&
                        //    error.RowNum > 0)
                        //{

                        //    var sameCell =
                        //        errors.Where(
                        //            x =>
                        //                x.SheetName == error.SheetName && x.RowNum == error.RowNum &&
                        //                x.ColNum == error.ColNum).Max(y => (int)y.ErrorLevel);
                        //    var errorLevel = sameCell == 2 ? ErrorLevel.Error : ErrorLevel.Warning;

                        //    PaintErrorCellColor(wBook, error, errorLevel);

                        //}
                        currentRow++;
                    }

                    #endregion

                    int end = currentRow;

                    wSheet.Range[wSheet.Cells[start, 1], wSheet.Cells[end - 1, 8]].Rows.Group();
                    wSheet.Cells[currentRow, 1].Value = errorType.ToString();
                    wSheet.Cells[currentRow, 8].Value = end - start;
                    count += end - start;
                    currentRow++;
                }


                if (spilterrorList.Key.Equals(ErrorLevel.Error))
                {
                    wSheet.Cells[currentRow, 1].Value = "Total error";
                    wSheet.Cells[currentRow, 8].Value = count;
                    errorCount += count;
                    wSheet.Range[wSheet.Cells[currentRow, 1], wSheet.Cells[currentRow, 8]].Interior.Pattern =
                        XlPattern.xlPatternSolid;
                    wSheet.Range[wSheet.Cells[currentRow, 1], wSheet.Cells[currentRow, 8]].Interior.Color = Color.Red;
                }
                if (spilterrorList.Key.Equals(ErrorLevel.Warning))
                {
                    wSheet.Cells[currentRow, 1].Value = "Total Warning";
                    wSheet.Cells[currentRow, 8].Value = count;
                    warnCount += count;
                    wSheet.Range[wSheet.Cells[currentRow, 1], wSheet.Cells[currentRow, 8]].Interior.Pattern =
                        XlPattern.xlPatternSolid;
                    wSheet.Range[wSheet.Cells[currentRow, 1], wSheet.Cells[currentRow, 8]].Interior.Color = Color.Red;
                }
                currentRow++;
            }
            wSheet.UsedRange.Value = wSheet.UsedRange.Value;
            wSheet.Rows["1:1"].AutoFilter();
            WriteSummary(wBook, reportSheetName, errorCount, warnCount);
            wSheet.Columns["A:I"].AutoFit();
        }

        //private void PaintErrorCellColor(Workbook wBook, Error error, ErrorLevel errorLevel)
        //{
        //    Worksheet errorSheet = wBook.Worksheets[error.SheetName];
        //    Range errorRange = error.ColNum > 0 ?
        //        errorSheet.Range[errorSheet.Cells[error.RowNum, error.ColNum], errorSheet.Cells[error.RowNum, error.ColNum]] :
        //        errorSheet.Range[errorSheet.Cells[error.RowNum, 1], errorSheet.Cells[error.RowNum, wBook.Worksheets[error.SheetName].UsedRange.Columns.Count]];

        //    if (errorLevel == ErrorLevel.Error)
        //    {
        //        errorRange.Interior.Pattern = XlPattern.xlPatternSolid;
        //        errorRange.Interior.Color = Color.Red;
        //    }
        //    else if (error.ErrorLevel == ErrorLevel.Warning)
        //    {
        //        errorRange.Interior.Pattern = XlPattern.xlPatternSolid;
        //        errorRange.Interior.Color = Color.Yellow;
        //    }
        //}

        private static List<object> GetErrorSubType(List<Error> errorList)
        {
            var subTypeList = errorList.GroupBy(p => p.ErrorType).Select(p => p.Key).ToList();
            return subTypeList;
        }

        private static List<Error> GetErrorsBySubType(object subtype, List<Error> errorList)
        {
            return errorList.Where(p => p.ErrorType.Equals(subtype)).ToList();
        }
    }
}
