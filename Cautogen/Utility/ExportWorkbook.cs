using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;

using OfficeOpenXml;

namespace Cautogen.Utility
{
    public static class ExportWorkbook
    {
        public static void ExportCmd(string workbookPath, string exportDir)
        {
            FileUtility.CleanDir(exportDir);
            string option = "-w \"" + workbookPath + "\" -d \"" + exportDir + "\"";
            _RunCmd("ExportWorkbook", option);
        }

        private static void _RunCmd(string cmd, string argment = "")
        {
            var nProcess = new Process();
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                FileName = cmd,
                Arguments = argment
            };
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
        }

        public static void ExportEpplus(string workbookPath, string exportDir)
        {
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine("Workbook :" + workbookPath + "Does not exist!");
                return;
            }

            if (!Directory.Exists(exportDir))
            {
                Directory.CreateDirectory(exportDir);
            }

            var fileInfo = new FileInfo(workbookPath);
            using (var pck = new ExcelPackage(fileInfo))
            {
                _ExportSheets(pck, exportDir);
                _ExportVba(pck, exportDir);
            }
        }

        private static void _ExportVba(ExcelPackage pck, string outputDir)
        {
            if (pck.Workbook.VbaProject == null)
            {
                return;
            }

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
                    // Force G15 / InvariantCulture so .NET 8's default G17
                    // doesn't emit 0.10000000000000001 strings that IGXL rejects.
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
