using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using CommonLib.ErrorReport.Base;
using CommonLib.Extension;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

namespace CommonLib.ErrorReport
{
    internal class ErrorReportEpplus(List<Error> errors)
    {
        private readonly List<Error> _errors = errors;

        private string ReportName
        {
            get
            {
                if (_errors.Count == 0)
                {
                    return "";
                }

                return _errors[0].ErrorCode.EnumErrorCategory + "Report";
            }
        }

        private List<string> GetErrorCategory()
        {
            return [.. _errors.GroupBy(p => p.ErrorCode.EnumErrorCategory.ToString()).Select(p => p.Key ?? "").Distinct()];
        }

        private List<Error> GetErrorsByCategory(string category)
        {
            var list = _errors.Where(p => p.ErrorCode.EnumErrorCategory.ToString() == category).ToList();
            return [.. list.OrderBy(x => x.SheetName).ThenBy(x => x.RowNum)];
        }

        public void WriteReport(List<ExcelWorkbook> excelWorkbooks, string errorReportName = "", string summaryReport = "SummaryReport")
        {
            if (_errors.Count == 0)
            {
                return;
            }

            string finalErrorReportName = !string.IsNullOrEmpty(errorReportName) ? errorReportName : ReportName;

            try
            {
                var errorsHaveWritten = new List<Error>();
                for (int i = 0; i < excelWorkbooks.Count - 1; i++)
                {
                    var names = excelWorkbooks[i].Worksheets.Select(x => x.Name).ToList();
                    var errorsNeedToWrite = _errors.Where(x => names.Exists(y => y.EqualsIgnoreCase(x.SheetName))).ToList();
                    WriteErrors(excelWorkbooks[i], errorsNeedToWrite, finalErrorReportName, summaryReport);
                    errorsHaveWritten.AddRange(errorsNeedToWrite);
                    excelWorkbooks[i].Worksheets.First().Select();
                }

                WriteErrors(excelWorkbooks.Last(), [.. _errors.Where(p => !errorsHaveWritten.Exists(a => a.SheetName == p.SheetName))], finalErrorReportName, summaryReport);
                excelWorkbooks.Last().Worksheets.First().Select();
            }
            catch (Exception e)
            {
                throw new Exception("Write General ErrorReport failed for " + finalErrorReportName, e);
            }
        }

        private void WriteErrors(ExcelWorkbook excelWorkbook, List<Error> errors, string reportName, string summaryReport = "SummaryReport")
        {
            if (errors.Count == 0)
            {
                return;
            }

            excelWorkbook.DeleteSheet(reportName);

            ExcelWorksheet workSheet = excelWorkbook.AddSheet(reportName);
            IExcelConditionalFormattingEqual condition = workSheet.ConditionalFormatting.AddEqual(new ExcelAddress("$B:$B"));
            condition.Style.Fill.PatternType = ExcelFillStyle.Solid;
            condition.Style.Fill.BackgroundColor.Color = Color.Red;
            condition.Formula = "\"Error\"";

            object[] headers = ["Category", "ErrorCode", "Level", "ErrorBehavior", "ErrorTarget", "Link", "SheetName", "Row", "Col", "ErrorMessage", "Guidance", "Comments"];
            workSheet.Cells[1, 1].LoadFromArrays([headers]).Style.Font.Bold = true;
            var rowsToOutline = new List<int>();
            int currentRow = 2;
            int count = 0;
            List<string> categoryList = GetErrorCategory();
            var content = new List<object[]>();
            Dictionary<string, (int, int)> categoryCountDict = [];
            foreach (string errorCategory in categoryList)
            {
                List<Error> subErrors = GetErrorsByCategory(errorCategory);
                foreach (Error error in subErrors)
                {
                    object[] row = new object[12];
                    row[0] = error.ErrorCode.EnumErrorCategory.ToString() ?? "";
                    row[1] = error.ErrorCode.FullCode ?? "";
                    row[2] = error.ErrorLevel.ToString();
                    row[3] = error.ErrorCode.EnumErrorBehavior?.ToString() ?? "";
                    row[4] = error.ErrorCode.EnumErrorTarget?.ToString() ?? "";
                    if (excelWorkbook.Worksheets.Any(x => x.Name.EqualsIgnoreCase(error.SheetName)) && error.RowNum > 0)
                    {
                        row[5] = error.Link;
                    }
                    else
                    {
                        row[5] = "";
                    }
                    row[6] = error.SheetName;
                    row[7] = error.RowNum;
                    row[8] = error.ColLetter;
                    row[9] = error.Message;
                    row[10] = error.ErrorCode.Guidance;
                    row[11] = error.Comments != null ? string.Join(",", error.Comments) : "";
                    content.Add(row);
                    rowsToOutline.Add(currentRow);
                    currentRow++;
                }

                object[] subSumRow = [errorCategory.ToString(), "", "", "", "", "", "", "", "", "", "", subErrors.Count];
                content.Add(subSumRow);
                count += subErrors.Count;
                categoryCountDict[errorCategory] = (subErrors.Count, currentRow);
                currentRow++;
            }

            object[] totalRow = ["Total error", "", "", "", "", "", "", "", "", "", "", count];
            content.Add(totalRow);
            workSheet.Cells[2, 1].LoadFromArrays(content);

            foreach (int rowNum in rowsToOutline)
            {
                workSheet.Row(rowNum).OutlineLevel = 1;
                workSheet.Row(rowNum).Collapsed = true;
            }

            workSheet.Cells[currentRow, 1, currentRow, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[currentRow, 1, currentRow, 12].Style.Fill.BackgroundColor.SetColor(Color.Red);

            workSheet.SetFormula(6);
            workSheet.Cells["1:1"].AutoFilter = true;
            workSheet.Cells.TryAutoFitColumns();
            workSheet.Column(7).Width = 23;
            workSheet.Column(10).Width = 100;
            workSheet.Column(11).Width = 100;

            if (categoryList.Count > 0 && !string.IsNullOrEmpty(summaryReport))
            {
                AppendSummaryReport(excelWorkbook, reportName, summaryReport, categoryCountDict);
            }
        }

        private static void AppendSummaryReport(ExcelWorkbook excelWorkbook, string reportName, string summaryReport, Dictionary<string, (int, int)> categoryCountDict)
        {
            ExcelWorksheet sheet;
            if (excelWorkbook.Worksheets.Any(x => x.Name.EqualsIgnoreCase(summaryReport)))
            {
                sheet = excelWorkbook.Worksheets[summaryReport];
            }
            else
            {
                sheet = excelWorkbook.AddSheet(summaryReport);
                object[] headers = ["ReportType", "Count", "Link"];
                sheet.Cells[1, 1].LoadFromArrays([headers]).Style.Font.Bold = true;
            }

            int startRow = 2;
            if (sheet.Dimension != null && sheet.Dimension.End != null)
            {
                startRow = sheet.Dimension.End.Row + 1;
            }

            foreach (KeyValuePair<string, (int, int)> category in categoryCountDict)
            {
                sheet.Cells[startRow, 1].Value = category.Key;
                sheet.Cells[startRow, 2].Value = category.Value.Item1;
                sheet.Cells[startRow, 3].Value = "Link";
                sheet.Cells[startRow, 3].Hyperlink = new Uri("#" + "'" + reportName + "'" + "!A" + category.Value.Item2, UriKind.Relative);
                sheet.Cells[startRow, 3].Style.Font.UnderLine = true;
                sheet.Cells[startRow, 3].Style.Font.Color.SetColor(Color.Blue);
                sheet.Cells.TryAutoFitColumns();
                startRow++;
            }
        }
    }
}
