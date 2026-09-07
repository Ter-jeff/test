using OfficeOpenXml;

namespace TestPlanLib
{
    public class MySheetReader
    {
        protected ExcelWorksheet ExcelWorksheet;

        protected int StartColNumber = -1;
        protected int StartRowNumber = -1;
        protected int EndColNumber = -1;
        protected int EndRowNumber = -1;

        public static string GetMergedCellValue(ExcelWorksheet sheet, int rowNumber, int columnNumber)
        {
            string range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null ?
                GetCellValue(sheet, rowNumber, columnNumber) :
                GetCellValue(sheet, (new ExcelAddress(range).Start.Row), (new ExcelAddress(range).Start.Column));
        }

        public static string GetCellValue(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            if (sheet.Cells[row, column].Value != null)
            {
                return sheet.Cells[row, column].Value.ToString().Trim();
            }

            if (sheet.Cells[row, column].Text != null)
            {
                return sheet.Cells[row, column].Text.Trim();
            }

            return "";
        }

        protected bool GetDimensions()
        {
            if (ExcelWorksheet.Dimension != null)
            {
                StartColNumber = ExcelWorksheet.Dimension.Start.Column;
                StartRowNumber = ExcelWorksheet.Dimension.Start.Row;
                EndColNumber = ExcelWorksheet.Dimension.End.Column;
                EndRowNumber = ExcelWorksheet.Dimension.End.Row;
                return true;
            }
            return false;
        }

        public static ExcelWorksheet ConvertCsvToExcelSheet(string fileName)
        {
            var excelPackage = new ExcelPackage();
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            int index = 0;
            using (var sr = new StreamReader(fileName))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    index++;
                    if (line != null)
                    {
                        string[] arr = line.Split(new[] { ',' }, StringSplitOptions.None);
                        int cnt = 0;
                        foreach (string item in arr)
                        {
                            sheet.Cells[index, 1 + cnt].Value = item;
                            cnt++;
                        }
                    }
                }
            }
            return sheet;
        }
    }
}
