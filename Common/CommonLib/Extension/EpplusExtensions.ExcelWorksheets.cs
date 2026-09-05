using OfficeOpenXml;

namespace CommonLib.Extension
{
    public static partial class EpplusExtensions
    {
        public static ExcelWorksheet AddSheet(this ExcelWorksheets excelWorksheets, string name, bool insertFlag)
        {
            bool isExist = excelWorksheets[name] != null;

            if (isExist)
            {
                excelWorksheets[name].Cells.Clear();
                return excelWorksheets[name];
            }

            ExcelWorksheet excelWorksheet = excelWorksheets.Add(name);
            if (insertFlag)
            {
                excelWorksheets.MoveBefore(name, excelWorksheets[0].Name);
            }

            return excelWorksheet;
        }

        public static void AddSheet(this ExcelWorksheets excelWorksheets, ExcelWorksheet excelWorksheet)
        {
            bool isExist = excelWorksheets[excelWorksheet.Name] != null;
            ExcelWorksheet target;

            if (isExist)
            {
                target = excelWorksheets[excelWorksheet.Name];
                target.Cells.Clear();
                CopyCellsAcrossPackage(excelWorksheet, target);
            }
            else
            {
                target = excelWorksheets.Add(excelWorksheet.Name);
                CopyCellsAcrossPackage(excelWorksheet, target);
            }

            excelWorksheets.MoveBefore(target.Name, excelWorksheets[0].Name);
        }

        public static ExcelWorksheet InsertSheet(this ExcelWorksheets excelWorksheets, string name)
        {
            bool isExist = excelWorksheets[name] != null;

            if (isExist)
            {
                excelWorksheets[name].Cells.Clear();
            }
            else
            {
                excelWorksheets.Add(name);
            }

            excelWorksheets.MoveBefore(name, excelWorksheets[0].Name);
            return excelWorksheets[0];
        }

        public static ExcelWorksheet AddSheet(this ExcelWorksheets excelWorksheets, string name)
        {
            return excelWorksheets.InsertSheet(name);
        }
    }
}
