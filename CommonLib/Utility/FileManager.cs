using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

using CommonLib.Extension;

using OfficeOpenXml;

namespace CommonLib.Utility
{

    public static class FileManager
    {
        public static void CopyFolder(string sourceDir, string destDir, bool overwrite = true)
        {
            // Ensure destination exists
            Directory.CreateDirectory(destDir);

            // Copy all files
            foreach (string filePath in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(filePath);
                string destFile = Path.Combine(destDir, fileName);
                File.Copy(filePath, destFile, overwrite);
            }

            // Copy all subdirectories
            foreach (string dirPath in Directory.GetDirectories(sourceDir))
            {
                string dirName = Path.GetFileName(dirPath);
                string destSubDir = Path.Combine(destDir, dirName);
                CopyFolder(dirPath, destSubDir, overwrite);
            }
        }

        public static string CopyFile(string sourcePath, string targetPath)
        {
            if (File.Exists(targetPath))
            {
                File.Delete(targetPath);
            }

            File.Copy(sourcePath, targetPath);
            return targetPath;
        }

        public static void MergeTestPlan(List<string> files, string excel)
        {
            if (File.Exists(excel))
            {
                File.Delete(excel);
            }

            string dir = Path.GetDirectoryName(excel);
            if (!Directory.Exists(dir))
            {
                _ = Directory.CreateDirectory(dir);
            }

            _ = GetExcelFormat();
            using (var excelPackage = new ExcelPackage(new FileInfo(excel)))
            {
                foreach (string file in files)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(file);
                    string ext = Path.GetExtension(file).ToLower();

                    if (ext == ".csv")
                    {
                        ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                        List<List<string>> lists = file.ConvertToLists();
                        _ = sheet.Cells[1, 1].PrintExcelRowByList(lists);

                        //var table = ParseCsvToDataTable(file);
                        //var sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                        //sheet.Cells[1, 1].PrintExcelRowByList(table);
                    }
                    else if (ext == ".xlsx")
                    {
                        using (var sourcePackage = new ExcelPackage(new FileInfo(file)))
                        {
                            foreach (ExcelWorksheet sheet in sourcePackage.Workbook.Worksheets)
                            {
                                string originalName = sheet.Name;
                                string uniqueName = GetUniqueSheetName(excelPackage, originalName);
                                ExcelWorksheet targetSheet = excelPackage.Workbook.Worksheets.Add(uniqueName);
                                ExcelAddressBase dim = sheet.Dimension;
                                if (dim == null)
                                {
                                    continue;
                                }

                                ExcelRange range = sheet.Cells[dim.Start.Row, dim.Start.Column, dim.End.Row, dim.End.Column];
                                targetSheet.Cells[range.Address].Value = range.Value;

                                // Copy merged cells
                                foreach (string mergedCell in sheet.MergedCells)
                                {
                                    targetSheet.Cells[mergedCell].Merge = true;
                                }
                            }
                        }
                    }
                }

                excelPackage.Save();
            }
        }

        private static string GetUniqueSheetName(ExcelPackage excelPackage, string baseName)
        {
            string name = baseName;
            int count = 1;
            while (excelPackage.Workbook.Worksheets.Any(ws => ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                name = $"{baseName}_{count++}";
            }

            return name;
        }

        private static ExcelTextFormat GetExcelFormat()
        {
            var format = new ExcelTextFormat
            {
                Delimiter = ',',
                Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                {
                    DateTimeFormat =
                    {
                        ShortDatePattern = "dd-mm-yyyy"
                    }
                },
                Encoding = Encoding.GetEncoding(950)
            };
            return format;
        }
    }
}
