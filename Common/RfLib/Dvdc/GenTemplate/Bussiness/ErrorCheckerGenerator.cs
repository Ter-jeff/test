using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    public class ErrorCheckerGenerator(ExcelPackage excelPackage)
    {
        public ExcelPackage XlPackage = excelPackage;
        public ExcelWorkbook XlWorkBook = excelPackage.Workbook;

        public void GenMissingReg(Dictionary<string, List<string>> errorcheck)
        {
            ExcelWorksheet excelWorksheet = XlWorkBook.AddSheet("Missing Register Name");
            excelWorksheet.Cells[1, 1].Value = "Pattern";
            excelWorksheet.Cells[1, 2].Value = "Missing Register Name in DigSrcAssignment";
            int rowindx = 2;
            foreach (KeyValuePair<string, List<string>> sheet in errorcheck)
            {
                excelWorksheet.Cells[rowindx, 1].Value = sheet.Key;

                foreach (string reg in sheet.Value)
                {
                    excelWorksheet.Cells[rowindx, 2].Value = reg;
                    rowindx++;
                }
            }

        }

        public void GendupRegName(Dictionary<string, List<string>> errorcheck)
        {
            ExcelWorksheet excelWorksheet = XlWorkBook.AddSheet("Duplicate Register Name");
            excelWorksheet.Cells[1, 1].Value = "Pattern";
            excelWorksheet.Cells[1, 2].Value = "Duplicate Register Name in DigSrcAssignment";
            int rowindx = 2;
            foreach (KeyValuePair<string, List<string>> sheet in errorcheck)
            {
                excelWorksheet.Cells[rowindx, 1].Value = sheet.Key;

                foreach (string reg in sheet.Value)
                {
                    excelWorksheet.Cells[rowindx, 2].Value = reg;
                    rowindx++;
                }
            }
        }
    }
}
