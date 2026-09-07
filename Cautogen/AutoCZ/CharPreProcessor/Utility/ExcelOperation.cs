using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.Utility
{
    public class ExcelOperation
    {
        public static string GetCellValue(ExcelWorksheet wSheet, int row, int column)
        {
            string value = "";
            if (wSheet.Cells[row, column].Value != null)
            {
                value = wSheet.Cells[row, column].Value.ToString().Trim();
            }

            if (value.ToLower() == "n/a")
            {
                value = "";
            }

            return value;
        }
    }
}
