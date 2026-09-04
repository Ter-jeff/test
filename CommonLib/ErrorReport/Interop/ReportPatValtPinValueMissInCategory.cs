using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

namespace CommonLib.ErrorReport.Interop
{
    internal class ReportPatValtPinValueMissInCategory : ErrorReportInterop
    {
        public ReportPatValtPinValueMissInCategory(List<Error> errorList)
            : base(errorList)
        {
        }

        public override void Write(Workbook workbook)
        {
            WritePatValtPinValueMissInCategoryReport(workbook);
        }

        public void WritePatValtPinValueMissInCategoryReport(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            const string reportSheetName = "PatValtPinsValueMissed";
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
            // Header
            string[,] headers = new string[1, 4];
            //DC Category	Select	Source	SelectAllZero	Instance Apply 
            headers[0, 0] = "Pattern";
            headers[0, 1] = "Dc Category Used";
            headers[0, 2] = "Job";
            headers[0, 3] = "Valt Pins value missed";

            wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 4]].Value2 = headers;
            wSheet.Range[(Range)wSheet.Cells[1, 1], (Range)wSheet.Cells[1, 4]].Font.Bold = true;

            // Data
            string[,] content = new string[ErrorList.Count, 4];
            for (int i = 0; i < ErrorList.Count; i++)
            {
                string[] data = ErrorList[i].Message.Split('|');
                content[i, 0] = data[0];
                content[i, 1] = data[1];
                content[i, 2] = data[2];
                content[i, 3] = data[3];
            }
            wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[ErrorList.Count + 1, 4]].Value2 = content;

            wSheet.Columns["A:D"].AutoFit();
            WriteSummary(workbook, reportSheetName, ErrorList.Count);


        }
    }
}
