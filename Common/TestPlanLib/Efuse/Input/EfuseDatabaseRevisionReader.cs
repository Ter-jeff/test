using System.Text.RegularExpressions;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Efuse.Input
{
    public partial class EfuseDatabaseRevisionReader : MySheetReader<int>
    {
        private const string ConHeaderDatabaseRevision = "^Database Revision";

        [GeneratedRegex(ConHeaderDatabaseRevision, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        private int _fuseDatabaseColNumber = -1;

        public override int ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition();

            GetHeaders();

            return ReadValue(sheetName);
        }

        protected int ReadValue(string sheetName)
        {
            int revision = -1;
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                if (!int.TryParse(GetCellValue(i, _fuseDatabaseColNumber), out revision))
                {
                    break;
                }
            }
            return revision;
        }

        protected void GetHeaders()
        {
            for (int i = 1; i <= EndCol; i++)
            {
                string lStrHeader = GetCellValue(StartRow, i).ToUpper().Trim();

                if (MyRegex().IsMatch(lStrHeader))
                {
                    _fuseDatabaseColNumber = i;
                }
            }
        }

        protected void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (MyRegex().IsMatch(GetCellValue(i, j).Trim()))
                    {
                        StartRow = i;
                        break;
                    }
                }
            }
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = ExcelWorksheet.Cells[rowNumber, columnNumber].Value;
            return value?.ToString() ?? "";
        }
    }
}
