using System;
using System.Collections.Generic;
using System.IO;

using Automation.Static;

using OfficeOpenXml;

namespace EfuseCheckCmdLib.EFuse.EFuseApp
{
    public class GenerateConfigTable(ExcelWorkbook excelWorkbook)
    {
        public ExcelWorkbook Workbook = excelWorkbook;
        public List<ExcelWorksheet> ConfigTables = [];

        public static void WorkFlow()
        {
            EpWorkbook.SheetFormatWorkbook = new ExcelPackage(new FileInfo(Path.Combine(AppContext.BaseDirectory, "Config", "SheetsFormat.xlsx"))).Workbook;
        }
    }
}
