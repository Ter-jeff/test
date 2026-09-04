using System;
using System.Collections.Generic;
using System.Drawing;

using Microsoft.Office.Interop.Excel;

namespace CommonLib.ErrorReport.Interop
{
    internal class ReportUnusedDcCategory : ErrorReportInterop
    {
        public ReportUnusedDcCategory(List<Error> errorList)
            : base(errorList)
        {
        }

        public override void Write(Workbook workbook)
        {
            WriteUnusedDcCategoryReport(workbook);
        }

        public void WriteUnusedDcCategoryReport(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            const string reportSheetName = "Unused DcCategorys";
            // Delete sheet if alreadey exists
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == reportSheetName)
                {
                    workbook.Application.DisplayAlerts = false;
                    sheet.Delete();
                    break;
                }
            }

            Worksheet wSheet = workbook.Worksheets.Add(workbook.Worksheets[2], Type.Missing, Type.Missing, Type.Missing);
            wSheet.Name = reportSheetName;
            string[,] headers = new string[1, 1];
            headers[0, 0] = "DcCategory Name";

            wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 1]].Value2 = headers;
            wSheet.Range[(Range)wSheet.Cells[1, 1], (Range)wSheet.Cells[1, 1]].Font.Bold = true;

            string[,] content = new string[ErrorList.Count, 1];
            for (int i = 0; i < ErrorList.Count; i++)
            {
                content[i, 0] = ErrorList[i].Message;
            }

            wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[ErrorList.Count + 1, 1]].Value2 = content;
            int startRow = ErrorList.Count + 2;
            wSheet.Cells[startRow, 1].Value = "Total Unused Categorys";
            wSheet.Cells[startRow, 2].Value = ErrorList.Count;
            wSheet.Range[(Range)wSheet.Cells[startRow, 1], (Range)wSheet.Cells[startRow, 2]].Interior.Pattern = XlPattern.xlPatternSolid;
            wSheet.Range[(Range)wSheet.Cells[startRow, 1], (Range)wSheet.Cells[startRow, 2]].Interior.Color = Color.Red;

            wSheet.Columns["A:B"].AutoFit();
            WriteSummary(workbook, reportSheetName, ErrorList.Count);
        }
    }
}
