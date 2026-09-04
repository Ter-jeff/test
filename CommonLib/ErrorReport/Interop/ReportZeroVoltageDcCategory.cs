using System;
using System.Collections.Generic;
using System.Drawing;

using CommonLib.Utility;

using Microsoft.Office.Interop.Excel;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace CommonLib.ErrorReport.Interop
{
    internal class ReportZeroVoltageDcCategory : ErrorReportInterop
    {
        public ReportZeroVoltageDcCategory(List<Error> errorList)
            : base(errorList)
        {
        }

        public override void Write(Workbook workbook)
        {
            WriteZeroVoltageDcCategoryReport(workbook);
        }

        public void WriteZeroVoltageDcCategoryReport(Workbook workbook)
        {
            if (ErrorList.Count == 0)
            {
                return;
            }

            const string reportSheetName = "ZeroVol DcCategorys";
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
            headers[0, 0] = "DcCategory Name";
            headers[0, 1] = "Select";
            headers[0, 2] = "Source";
            //headers[0, 3] = "SelectAllZero";
            headers[0, 3] = "Instance Apply";

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
                content[i, 3] = data[4];
                //content[i, 4] = data[4];
            }
            wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[ErrorList.Count + 1, 4]].Value2 = content;

            //Format
            for (int i = 2; i <= ErrorList.Count + 1; i++)
            {
                if (wSheet.Cells[i, 3].Value == "Plan" && wSheet.Cells[i, 4].Value == "X")
                {
                    wSheet.Range[wSheet.Cells[i, 4], wSheet.Cells[i, 4]].Interior.Pattern = XlPattern.xlPatternSolid;
                    wSheet.Range[wSheet.Cells[i, 4], wSheet.Cells[i, 4]].Interior.Color = Color.Orange;
                }

                if (wSheet.Cells[i, 3].Value == "Autogen rule all zero" && wSheet.Cells[i, 4].Value == "V")
                {
                    wSheet.Range[wSheet.Cells[i, 4], wSheet.Cells[i, 4]].Interior.Pattern = XlPattern.xlPatternSolid;
                    wSheet.Range[wSheet.Cells[i, 4], wSheet.Cells[i, 4]].Interior.Color = Color.Red;
                }
            }

            // Link Data
            var dicInstance = new Dictionary<string, List<string>>();
            int column = 0;
            for (int i = 0; i < ErrorList.Count; i++)
            {
                string[] data = ErrorList[i].Message.Split('|');
                if (data[2] == "Autogen rule all zero" && data[4] == "V")
                {
                    Range linkRange = wSheet.Cells[i + 2, 4];
                    column += 1;
                    wSheet.Hyperlinks.Add(linkRange, "#'DCCateInsTable'!R1C" + column, Type.Missing, Type.Missing, Type.Missing);
                    dicInstance.Add(data[0] + " " + data[1], ErrorList[i].Comments);
                }
            }

            // Footer
            int startRow = ErrorList.Count + 2;
            wSheet.Cells[startRow, 1].Value = "Total Zero Voltage Categorys";
            wSheet.Cells[startRow, 2].Value = ErrorList.Count;
            wSheet.Range[(Range)wSheet.Cells[startRow, 1], (Range)wSheet.Cells[startRow, 2]].Interior.Pattern = XlPattern.xlPatternSolid;
            wSheet.Range[(Range)wSheet.Cells[startRow, 1], (Range)wSheet.Cells[startRow, 2]].Interior.Color = Color.Red;

            wSheet.Columns["A:F"].AutoFit();
            WriteSummary(workbook, reportSheetName, ErrorList.Count);

            // Add Detail Sheet
            AddDetailSheet(workbook, dicInstance);
        }

        private void AddDetailSheet(Workbook wBook, Dictionary<string, List<string>> dicInstance)
        {
            const string sheetName = "DCCateInsTable";
            Worksheet wSheet = wBook.IsSheetExist(sheetName) ?
                wBook.Worksheets[sheetName] :
                wBook.Worksheets.Add(wBook.Worksheets[3], Type.Missing, Type.Missing, Type.Missing);
            wSheet.Name = sheetName;
            wSheet.Cells.Clear();

            int column = 1;
            foreach (KeyValuePair<string, List<string>> dcCatePair in dicInstance)
            {
                wSheet.Cells[1, column].Value = dcCatePair.Key;
                for (int row = 0; row < dcCatePair.Value.Count; row++)
                {
                    wSheet.Cells[row + 2, column].Value = dcCatePair.Value[row];
                }
                column += 1;
            }
        }
    }
}
