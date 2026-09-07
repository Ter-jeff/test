using System.Globalization;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace Cautogen.Utility
{
    public class ExcelOperation
    {
        public static string GetCellValue(ExcelWorksheet wSheet, int row, int column)
        {
            return wSheet.Cells[row, column].Value == null
                ? ""
                : wSheet.Cells[row, column].Value.ToString().Trim();
        }

        public static string GetMergerdCellValue(ExcelWorksheet sheet, int rowNumber, int columnNumber)
        {
            string range = sheet.MergedCells[rowNumber, columnNumber];

            string mergeCellValue = range == null
                ? GetCellValue(sheet, rowNumber, columnNumber)
                : GetCellValue(sheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column);

            return mergeCellValue;
        }

        public static string ConvertUnits(string limitStr)
        {
            if (limitStr.Contains("10^"))
            {
                limitStr = limitStr.Replace("*10^", "E");
            }

            if (limitStr == "" || limitStr.Contains("E") || Regex.IsMatch(limitStr, @"^(\d|\.|-)+$")) // Limit value may be 1.2E-5
            {
                return limitStr;
            }

            if (!Regex.IsMatch(limitStr, @"^(\d|\.|-)+(\w)*$"))
            {
                return limitStr;
            }

            string limitNum = Regex.Match(limitStr, @"(?<num>((\d|\.|-)+))[^\*]*").Groups["num"].ToString();

            if (limitNum == "0")
            {
                return limitNum;
            }

            string limitUnit = limitStr.Replace(limitNum, "").Trim();

            double rate = 1;
            if (Regex.IsMatch(limitUnit, "^m.*"))
            {
                rate = 1E-3;
            }
            else if (Regex.IsMatch(limitUnit, "^u.*"))
            {
                rate = 1E-6;
            }
            else if (Regex.IsMatch(limitUnit, "^n.*"))
            {
                rate = 1E-9;
            }
            else if (Regex.IsMatch(limitUnit.ToLower(), "^k.*"))
            {
                rate = 1E3;
            }
            else if (Regex.IsMatch(limitUnit, "^M.*"))
            {
                rate = 1E6;
            }
            else if (Regex.IsMatch(limitUnit, "^G.*"))
            {
                rate = 1E9;
            }

            return double.TryParse(limitNum, out double value)
                ? (value * rate).ToString("G15", CultureInfo.InvariantCulture)
                : limitStr;
        }

    }
}
