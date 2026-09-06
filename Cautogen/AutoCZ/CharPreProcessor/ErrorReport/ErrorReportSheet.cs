using System;
using System.Drawing;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Writer;

using CommonLib.Extension;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Cautogen.AutoCZ.CharPreProcessor.ErrorReport
{
    internal class ErrorReportSheet : IExcelSheetWriter
    {
        public void Write(ExcelWorkbook wb)
        {
            int count = 0;
            ExcelWorksheet wSheet = wb.Worksheets.Add("ErrorReport");
            wb.Worksheets.MoveToStart("ErrorReport");

            wSheet.Cells[1, 1].Value = "PatternListFile: ";
            wSheet.Cells[1, 2].Value = UtilityMain.UtilityData.InputParam.PatListFile; //AutoGenParse._param.PatListFile;
            wSheet.Cells[2, 1].Value = "Error(Red)/Warning(Yellow)";
            wSheet.Cells[2, 2].Value = "ErrorType";
            wSheet.Cells[2, 3].Value = "SheetName";
            wSheet.Cells[2, 4].Value = "Address";
            wSheet.Cells[2, 5].Value = "ErrorMessage";
            wSheet.Cells[2, 6].Value = "Count";
            wSheet.Cells[2, 1, 2, 6].Style.Font.Bold = true;
            wSheet.Cells[3, 1].Value = "Below shows Warning findings";
            int startRow = ErrorManager.ErrorListDict.Keys
                .Where(errorType => ErrorManager.ErrorListDict[errorType][0].ErrorLevel == ErrorLevel.Warning)
                .Aggregate(4, (current, errorType) => _WriteErrorByType(wSheet, errorType, current, wb, ref count));

            wSheet.Cells[startRow++, 1].Value = "Below shows Error findings";

            startRow = ErrorManager.ErrorListDict.Keys
                .Where(errorType => ErrorManager.ErrorListDict[errorType][0].ErrorLevel == ErrorLevel.Error)
                .Aggregate(startRow,
                    (current, errorType) => _WriteErrorByType(wSheet, errorType, current, wb, ref count));

            wSheet.Cells[startRow, 2].Value = "Total error";
            wSheet.Cells[startRow, 6].Value = count;
            wSheet.Cells[startRow, 1, startRow, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            wSheet.Cells[startRow, 1, startRow, 6].Style.Fill.BackgroundColor.SetColor(Color.Tomato);

            wSheet.Column(1).TryAutoFit();
            wSheet.Column(2).Width = 40;
            wSheet.Column(3).TryAutoFit();
            wSheet.Column(4).TryAutoFit();
            wSheet.Column(5).TryAutoFit();
            wSheet.Column(6).TryAutoFit();
        }

        private static int _WriteErrorByType(ExcelWorksheet wSheet, ErrorType errorType, int startRow, ExcelWorkbook planWorkbook, ref int count)
        {
            Color fillColor = ErrorManager.ErrorListDict[errorType][0].ErrorLevel == ErrorLevel.Error
                ? Color.Tomato
                : Color.Yellow;

            int start = startRow;
            foreach (ErrorMessage error in ErrorManager.ErrorListDict[errorType])
            {
                wSheet.Row(startRow).OutlineLevel = 1;
                wSheet.Row(startRow).Collapsed = true;
                wSheet.Cells[startRow, 1].Value = error.ErrorLevel;
                wSheet.Cells[startRow, 2].Value = errorType.ToString();
                ExcelRange myRange = wSheet.Cells[startRow, 1, startRow, 2];
                myRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                myRange.Style.Fill.BackgroundColor.SetColor(fillColor);

                wSheet.Cells[startRow, 3].Value = error.SheetName;
                if (error.ColList.Count > 0 && error.ColList[0] != 0)
                {
                    ErrorMessage error1 = error;
                    wSheet.Cells[startRow, 4].Value = string.Join(",", error.ColList.Select(p => ExcelCellBase.GetAddress(error1.RowNum, p)));
                }
                else
                {
                    wSheet.Cells[startRow, 4].Value = $"Row:{error.RowNum}";
                }

                wSheet.Cells[startRow, 5].Value = error.Message;

                int commentCount = 0;
                foreach (string comment in error.CommentsList)
                {
                    wSheet.Cells[startRow, 7 + commentCount].Value = comment;
                    commentCount++;
                }
                if (planWorkbook.Worksheets[error.SheetName] == null)
                {
                    continue;
                }

                if (error.SheetName != "" && error.RowNum > 0)
                {
                    try
                    {
                        wSheet.Cells[startRow, 5].Style.Font.Color.SetColor(Color.Blue);
                        wSheet.Cells[startRow, 5].Style.Font.UnderLine = true;
                        int endCol = planWorkbook.Worksheets[error.SheetName].Dimension.End.Column;
                        ExcelWorksheet errorSheet = planWorkbook.Worksheets[error.SheetName];
                        ExcelRange errorRange = null;
                        //todo fill color with independent cell not ready
                        if (error.ColList.Count > 0 && error.ColList[0] != 0)
                        {
                            foreach (int col in error.ColList)
                            {
                                errorRange = errorSheet.Cells[error.RowNum, col];
                                errorRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                errorRange.Style.Fill.BackgroundColor.SetColor(fillColor);
                            }
                        }
                        else
                        {
                            errorRange = errorSheet.Cells[error.RowNum, 1, error.RowNum, endCol];
                            errorRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            errorRange.Style.Fill.BackgroundColor.SetColor(fillColor);
                        }

                        if (errorRange != null)
                        {
                            wSheet.Cells[startRow, 5].Hyperlink =
                                new ExcelHyperLink("'" + error.SheetName + "'" + "!" + errorRange.Address, error.Message);
                        }

                    }
                    catch (Exception)
                    {
                    }
                }
                startRow++;
            }

            int end = startRow;
            wSheet.Cells[startRow, 1].Value = ErrorManager.ErrorListDict[errorType][0].ErrorLevel;
            wSheet.Cells[startRow, 2].Value = errorType.ToString();
            ExcelRange titleRange = wSheet.Cells[startRow, 1, startRow, 2];
            titleRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            titleRange.Style.Fill.BackgroundColor.SetColor(fillColor);
            wSheet.Cells[startRow, 6].Value = end - start;
            count += end - start;
            startRow++;
            return startRow;
        }
    }
}
