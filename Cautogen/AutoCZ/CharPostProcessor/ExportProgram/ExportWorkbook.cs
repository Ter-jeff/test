using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPostProcessor.ExportProgram
{
    public class ExportWorkbook
    {
        public static void Export(string workbook, string outputDir)
        {
            if (!File.Exists(workbook))
            {
                Console.WriteLine("Workbook :" + workbook + "Does not exist!");
                return;
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var fileInfo = new FileInfo(workbook);
            using (var pck = new ExcelPackage(fileInfo))
            {
                _ExportSheets(pck, outputDir);
                _ExportVba(pck, outputDir);
            }
        }

        private static void _ExportVba(ExcelPackage pck, string outputDir)
        {
            OfficeOpenXml.VBA.ExcelVbaModuleCollection modules = pck.Workbook.VbaProject.Modules;

            foreach (OfficeOpenXml.VBA.ExcelVBAModule module in modules)
            {
                // get extension by type
                string extension = "";
                switch (module.Type.ToString())
                {
                    case "Class":
                        extension = ".cls";
                        break;

                    case "Module":
                        extension = ".bas";
                        break;
                }
                if (extension == "")
                {
                    continue;
                }

                string vbaFileName = Path.Combine(outputDir, module.Name + extension);
                string header = $"Attribute VB_Name = \"{module.Name}\"\r\n";
                string content = extension == ".bas" ? header + module.Code : module.Code;
                File.WriteAllText(vbaFileName, content);
            }
        }

        private static void _ExportSheets(ExcelPackage pck, string outputDir)
        {
            ExcelWorkbook wb = pck.Workbook;
            foreach (ExcelWorksheet sh in wb.Worksheets)
            {
                if (sh == null)
                {
                    continue;
                }

                if (sh.Dimension == null)
                {
                    continue;
                }

                int maxColumnNumber = sh.Dimension.End.Column;
                int totalRowCount = sh.Dimension.End.Row;

                using (var writer = new StreamWriter(Path.Combine(outputDir, sh.Name + ".txt")))
                {
                    for (int rowNum = 1; rowNum <= totalRowCount; rowNum++)
                    {
                        writer.WriteLine(_GetRowStr(sh, rowNum, maxColumnNumber));
                    }
                }
            }
        }

        private static string _GetRowStr(ExcelWorksheet worksheet, int currentRowNum, int maxColumnNumber)
        {
            var rowList = new List<string>();
            for (int i = 1; i <= maxColumnNumber; i++)
            {
                ExcelRange cell = worksheet.Cells[currentRowNum, i];
                string val;
                if (cell.Value is double doubleValue)
                {
                    // .NET Framework 4.8's default "G" was effectively G15. .NET 5+
                    // changed it to G17 (round-trippable), producing strings like
                    // 0.10000000000000001 that IGXL rejects with MDLC0008.
                    val = doubleValue.ToString("G15", CultureInfo.InvariantCulture);
                }
                else
                {
                    val = cell.Value != null ? cell.Value.ToString() : string.Empty;
                }
                string formula = cell.Formula ?? string.Empty;
                string rowStr = formula != string.Empty ? "=" + formula : val;
                rowList.Add(rowStr != "#N/A" ? rowStr : "#N/A");
            }
            rowList.Add(string.Empty);
            return string.Join("\t", rowList);
        }
    }
}
