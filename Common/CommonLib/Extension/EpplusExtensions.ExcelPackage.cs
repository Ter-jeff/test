using System.IO;

using OfficeOpenXml;

namespace CommonLib.Extension
{
    public static partial class EpplusExtensions
    {
        public static void ExportWorkBook2Txt(this ExcelPackage excelPackage, string outPath)
        {
            if (!Directory.Exists(outPath))
            {
                Directory.CreateDirectory(outPath);
            }

            foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
            {
                ExportWorkBook2Txt(worksheet, outPath);
            }
        }
    }
}
