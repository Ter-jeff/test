using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.ExcelOp
{
    public class ExcelOperation
    {
        /// <summary>
        /// Return the column index of header, throw exception if missing header
        /// </summary>
        /// <param name="headerOrder">A dictionary recorded all headers and their column index</param>
        /// <param name="header">header name</param>
        /// <returns></returns>
        public static int GetHeaderIndex(Dictionary<string, int> headerOrder, string header)
        {
            header = header.Replace(" ", "").Replace("(", @"\(").Replace(")", @"\)");
            return headerOrder.FirstOrDefault(a => Regex.IsMatch(a.Key, header, RegexOptions.IgnoreCase)).Value;
        }

        /// <summary>
        /// Record all headers' column index in a dictionary
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <returns></returns>
        public static Dictionary<string, int> GetHeaderOrder(ExcelWorksheet sheet, int startRow = 1)
        {
            var headerOrder = new Dictionary<string, int>();
            if (sheet.Dimension == null)
            {
                return headerOrder;
            }

            int endCol = sheet.Dimension.End.Column;
            for (int i = 1; i <= endCol; i++)
            {
                if (sheet.Cells[startRow, i].Value == null)
                {
                    continue;
                }

                string header = sheet.Cells[startRow, i].Value.ToString().Replace(" ", "");
                if (header.ToLower() == "ipuse")
                {
                    header = header + "_" + sheet.Cells[startRow + 1, i].Value.ToString().Replace(" ", "");
                }

                if (!headerOrder.ContainsKey(header))
                {
                    headerOrder.Add(header, i); //Remember column index of each header
                }
            }
            return headerOrder;
        }

        public static string GetCellValue(ExcelWorksheet wSheet, int row, int column)
        {
            if (column <= 0 || wSheet.Cells[row, column].Value == null)
            {
                return "";
            }

            object cell = wSheet.Cells[row, column].Value;
            if (cell is double value)
            {
                return value.ToString("G15", CultureInfo.InvariantCulture);
            }
            else
            {
                string text = cell.ToString().Trim();
                return text.ToLower() == "n/a" ? "" : text;
            }
        }

        public static void ConvertTimeSetVersion(string workbookPath)
        {

        }
    }
}
