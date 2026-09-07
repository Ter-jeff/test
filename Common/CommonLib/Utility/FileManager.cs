using System.Collections.Generic;
using System.IO;

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

        public static void MergeTestPlanCsv(List<string> files, string outputExcelPath)
        {
            if (File.Exists(outputExcelPath))
            {
                File.Delete(outputExcelPath);
            }

            using var pkg = new ExcelPackage();
            foreach (string file in files)
            {
                string ext = Path.GetExtension(file).ToLower();
                string baseSheetName = Path.GetFileNameWithoutExtension(file);
                if (ext == ".csv")
                {
                    string fileContent = File.ReadAllText(file);
                    string detectedEol = fileContent.Contains("\r\n") ? "\r\n" : "\n";
                    var format = new ExcelTextFormat
                    {
                        Delimiter = ',',
                        EOL = detectedEol
                    };
                    string sheetName = GetUniqueSheetName(pkg, baseSheetName);
                    ExcelWorksheet sheet = pkg.Workbook.Worksheets.Add(sheetName);
                    sheet.Cells["A1"].LoadFromText(new FileInfo(file), format);
                }
                else if (ext == ".xlsx" || ext == ".xls")
                {
                    using var srcPkg = new ExcelPackage(new FileInfo(file));
                    foreach (ExcelWorksheet srcSheet in srcPkg.Workbook.Worksheets)
                    {
                        string uniqueName = GetUniqueSheetName(pkg, srcSheet.Name);
                        ExcelWorksheet destSheet = pkg.Workbook.Worksheets.Add(uniqueName);

                        // Copy cells row-by-row
                        if (srcSheet.Dimension != null)
                        {
                            int endRow = srcSheet.Dimension.End.Row;
                            int endCol = srcSheet.Dimension.End.Column;
                            for (int r = 1; r <= endRow; r++)
                            {
                                for (int c = 1; c <= endCol; c++)
                                {
                                    destSheet.Cells[r, c].Value = srcSheet.Cells[r, c].Value;
                                }
                            }
                        }
                    }
                }
            }

            // Sort sheets alphabetically
            SortSheetsAlphabetically(pkg);

            pkg.SaveAs(new FileInfo(outputExcelPath));
        }

        private static void SortSheetsAlphabetically(ExcelPackage excelPackage)
        {
            ExcelWorksheets worksheets = excelPackage.Workbook.Worksheets;
            var sortedNames = new List<string>();
            foreach (ExcelWorksheet ws in worksheets)
            {
                sortedNames.Add(ws.Name);
            }
            sortedNames.Sort(StringExtensions.IgnoreCase);

            for (int i = sortedNames.Count - 1; i >= 0; i--)
            {
                worksheets.MoveToStart(sortedNames[i]);
            }
        }

        public static void CopyDirectory(string sourceDir, string destinationDir, bool copySubDirs = true)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(sourceDir);

            if (!directoryInfo.Exists)
            {
                throw new DirectoryNotFoundException($"Source directory does not exist: {sourceDir}");
            }

            DirectoryInfo[] dirs = directoryInfo.GetDirectories();

            // If destination directory doesn't exist, create it
            if (!Directory.Exists(destinationDir))
            {
                _ = Directory.CreateDirectory(destinationDir);
            }

            // Copy all files
            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                string targetFilePath = Path.Combine(destinationDir, file.Name);
                _ = file.CopyTo(targetFilePath, true);
            }

            // Copy subdirectories
            if (copySubDirs)
            {
                foreach (DirectoryInfo dir in dirs)
                {
                    string targetSubDir = Path.Combine(destinationDir, dir.Name);
                    CopyDirectory(dir.FullName, targetSubDir);
                }
            }
        }

        private static string GetUniqueSheetName(ExcelPackage excelPackage, string baseName)
        {
            string name = baseName.Length > 31 ? baseName[..31] : baseName;
            int counter = 1;
            while (SheetExists(excelPackage, name))
            {
                name = baseName + "_" + counter++;
                if (name.Length > 31)
                {
                    name = name[..31];
                }
            }

            return name;
        }

        private static bool SheetExists(ExcelPackage excelPackage, string name)
        {
            foreach (ExcelWorksheet sheet in excelPackage.Workbook.Worksheets)
            {
                if (sheet.Name == name)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
