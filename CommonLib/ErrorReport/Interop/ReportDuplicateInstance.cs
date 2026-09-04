using System;
using System.Collections.Generic;
using System.Drawing;

using Microsoft.Office.Interop.Excel;

namespace CommonLib.ErrorReport.Interop
{
    internal class ReportDuplicateInstance : ErrorReportInterop
    {
        public ReportDuplicateInstance(List<Error> errorList)
            : base(errorList)
        {
        }

        public override void Write(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            WriteDuplicatInstancesReport(workbook);
        }

        public void WriteDuplicatInstancesReport(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            const string reportSheetName = "Duplicate Instances";
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
            string[,] headers = new string[1, 4];
            headers[0, 0] = "ErrorType";
            headers[0, 1] = "Instance Sheet Name";
            headers[0, 2] = "Instance Name";
            headers[0, 3] = "Count";

            wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 4]].Value2 = headers;
            wSheet.Range[(Range)wSheet.Cells[1, 1], (Range)wSheet.Cells[1, 4]].Font.Bold = true;

            string[,] content = new string[ErrorList.Count, 4];
            for (int i = 0; i < ErrorList.Count; i++)
            {
                content[i, 0] = "Duplicate instance";
                content[i, 1] = ErrorList[i].SheetName;
                content[i, 2] = ErrorList[i].Comments[0];
                content[i, 3] = ErrorList[i].Comments[1];
            }

            wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[ErrorList.Count + 1, 4]].Value2 = content;
            int startRow = ErrorList.Count + 2;
            wSheet.Cells[startRow, 1].Value = "Total error";
            wSheet.Cells[startRow, 4].Value = ErrorList.Count;
            wSheet.Range[(Range)wSheet.Cells[startRow, 1], (Range)wSheet.Cells[startRow, 5]].Interior.Pattern = XlPattern.xlPatternSolid;
            wSheet.Range[(Range)wSheet.Cells[startRow, 1], (Range)wSheet.Cells[startRow, 5]].Interior.Color = Color.Red;

            wSheet.Columns["A:D"].AutoFit();
            WriteSummary(workbook, reportSheetName, ErrorList.Count);
        }
    }
}
