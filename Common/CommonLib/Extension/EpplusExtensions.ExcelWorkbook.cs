using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

using Microsoft.Office.Interop.Excel;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace CommonLib.Extension
{
    public static partial class EpplusExtensions
    {
        public static List<string> GetMatchPlanSheets(this ExcelWorkbook excelWorkbook, string pattern)
        {
            List<string> sheets = [];
            foreach (ExcelWorksheet worksheet in excelWorkbook.Worksheets)
            {
                string wsName = worksheet.Name;
                if (Regex.IsMatch(wsName, pattern, RegexOptions.IgnoreCase))
                {
                    sheets.Add(wsName);
                }
            }
            return sheets;
        }

        public static bool IsExist(this ExcelWorkbook excelWorkbook, string name)
        {
            foreach (ExcelWorksheet sheet in excelWorkbook.Worksheets)
            {
                if (sheet.Name.EqualsIgnoreCase(name))
                {
                    return true;
                }
            }

            return false;
        }

        public static void ExportToTxt(this ExcelWorkbook excelWorkbook, string folder)
        {
            if (excelWorkbook == null)
            {
                return;
            }

            foreach (ExcelWorksheet excelWorksheet in excelWorkbook.Worksheets)
            {
                excelWorksheet.ExportToTxt(Path.Combine(folder, excelWorksheet.Name + ".txt"));
            }
        }

        public static List<string> GetPlanSheets(this ExcelWorkbook excelWorkbook, string key)
        {
            var sheets = new List<string>();
            foreach (ExcelWorksheet worksheet in excelWorkbook.Worksheets)
            {
                string wsName = worksheet.Name;
                if (Regex.IsMatch(wsName, key, RegexOptions.IgnoreCase))
                {
                    sheets.Add(wsName);
                }
            }

            return sheets;
        }

        public static void CopyWorkSheets(this ExcelWorkbook excelWorkbook, List<string> files)
        {
            if (files == null)
            {
                return;
            }

            foreach (string file in files)
            {
                CopyWorkSheet(excelWorkbook, file);
            }
        }

        public static void CopyWorkSheet(this ExcelWorkbook excelWorkbook, string file)
        {
            if (string.IsNullOrEmpty(file))
            {
                return;
            }

            string fileName = Path.GetFileNameWithoutExtension(file);
            if (Path.GetExtension(file).EqualsIgnoreCase(".csv"))
            {
                string fileContent = File.ReadAllText(file);
                string detectedEol = fileContent.Contains("\r\n") ? "\r\n" : "\n";
                var format = new ExcelTextFormat
                {
                    Delimiter = ',',
                    Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                    {
                        DateTimeFormat = { ShortDatePattern = "dd-mm-yyyy" }
                    },
                    Encoding = new UTF8Encoding(),
                    EOL = detectedEol
                };
                var fileInfo = new FileInfo(file);
                ExcelWorksheet worksheet = excelWorkbook.Worksheets.Add(fileName);
                excelWorkbook.Worksheets.MoveBefore(worksheet.Name, excelWorkbook.Worksheets[0].Name);
                worksheet.Cells["A1"].LoadFromText(fileInfo, format);
            }
            else if (Path.GetExtension(file).EqualsIgnoreCase(".txt"))
            {
                var format = new ExcelTextFormat
                {
                    Delimiter = '\t',
                    Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                    {
                        DateTimeFormat = { ShortDatePattern = "dd-mm-yyyy" }
                    },
                    Encoding = new UTF8Encoding()
                };
                var fileInfo = new FileInfo(file);
                ExcelWorksheet worksheet = excelWorkbook.Worksheets.Add(fileName);
                worksheet.Cells["A1"].LoadFromText(fileInfo, format);
            }
            else
            {
                using var package = new ExcelPackage(new FileInfo(file));
                foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                {
                    excelWorkbook.AddSheet(worksheet);
                }
            }
        }

        public static void AddSheet(this ExcelWorkbook excelWorkbook, ExcelWorksheet excelWorksheet)
        {
            bool isExist = excelWorkbook.Worksheets[excelWorksheet.Name] != null;
            ExcelWorksheet target;

            if (isExist)
            {
                target = excelWorkbook.Worksheets[excelWorksheet.Name];
                target.Cells.Clear();
                CopyCellsAcrossPackage(excelWorksheet, target);
            }
            else
            {
                target = excelWorkbook.Worksheets.Add(excelWorksheet.Name);
                CopyCellsAcrossPackage(excelWorksheet, target);
            }

            excelWorkbook.Worksheets.MoveBefore(target.Name, excelWorkbook.Worksheets[0].Name);
        }

        public static void DeleteSheet(this ExcelWorkbook excelWorkbook, string name)
        {
            foreach (ExcelWorksheet sheet in excelWorkbook.Worksheets)
            {
                if (name.EqualsIgnoreCase(sheet.Name))
                {
                    excelWorkbook.Worksheets.Delete(sheet);
                    break;
                }
            }
        }

        public static ExcelWorksheet AddSheet(this ExcelWorkbook excelWorkbook, string name)
        {
            return excelWorkbook.Worksheets.InsertSheet(name);
        }

        private static void AddHeaderStyle(this ExcelWorkbook excelWorkbook)
        {
            foreach (ExcelNamedStyleXml item in excelWorkbook.Styles.NamedStyles)
            {
                if (item.Name == "Header")
                {
                    return;
                }
            }

            ExcelNamedStyleXml namedStyle = excelWorkbook.Styles.CreateNamedStyle("Header");
            namedStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyle.Style.Fill.BackgroundColor.SetColor(Color.YellowGreen);
        }

        public static void AddTxt(this ExcelWorkbook excelWorkbook, string fileName)
        {
            string[] lines = File.ReadAllLines(fileName);
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            ExcelWorksheet newSheet = excelWorkbook.Worksheets[sheetName] ?? excelWorkbook.AddSheet(sheetName);

            int rowCount = lines.Length + 1;
            int columnCount = lines.Max(x => x.Split('\t').Length);

            object[,] values = new object[rowCount, columnCount];
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                List<string> arr = [.. line.Split('\t')];
                for (int j = 0; j < arr.Count; j++)
                {
                    values[i, j] = arr[j];
                }
            }

            newSheet.Name = sheetName;
            newSheet.Cells[1, 1, rowCount, columnCount].Value = values;
        }

        public static void PrintInputFileInformation(this ExcelWorkbook excelWorkbook, Dictionary<string, List<string>> dicInfo)
        {
            ExcelWorksheet sheet = excelWorkbook.Worksheets.Add("InputFile_Information");
            sheet.Cells[1, 1].Value = "Input File";
            sheet.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#305496"));
            sheet.Cells[1, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            sheet.Cells[1, 2].Value = "Information";
            sheet.Cells[1, 2].Style.Font.Color.SetColor(Color.White);
            sheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#305496"));
            sheet.Cells[1, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;

            sheet.Column(1).Width = 26.29;
            sheet.Column(2).Width = 66.86;

            int index = 2;
            foreach (KeyValuePair<string, List<string>> item in dicInfo)
            {
                sheet.Cells[index, 1].Value = item.Key;
                sheet.Cells[index, 1].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[index, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[index, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00B050"));
                sheet.Cells[index, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[index, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                sheet.Cells[index, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                if (item.Value.Count > 0)
                {
                    foreach (string val in item.Value)
                    {
                        sheet.Cells[index, 2].Value = val;
                        sheet.Cells[index, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        sheet.Cells[index, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        index++;
                    }
                }
                else
                {
                    index++;
                }
            }
        }

        [SupportedOSPlatform("windows")]
        public static void MergeWorkbooks(List<string> sourceFiles, string outputPath)
        {
            var app = new Application { Visible = false, DisplayAlerts = false };
            Workbook? target = null;

            try
            {
                target = app.Workbooks.Add();

                while (target.Sheets.Count > 1)
                {
                    ((Worksheet)target.Sheets[target.Sheets.Count]).Delete();
                }

                Dictionary<string, int> usedNames = new Dictionary<string, int>(StringExtensions.IgnoreCase);
                bool firstSheet = true;

                foreach (string file in sourceFiles)
                {
                    Workbook? source = null;
                    try
                    {
                        source = app.Workbooks.Open(file, ReadOnly: true);

                        foreach (Worksheet srcSheet in source.Worksheets)
                        {
                            string name = GetUniqueName(srcSheet.Name, usedNames);

                            if (firstSheet)
                            {
                                srcSheet.Copy(Before: target.Sheets[1]);
                                ((Worksheet)target.Sheets[1]).Name = name;
                                firstSheet = false;
                            }
                            else
                            {
                                srcSheet.Copy(After: target.Sheets[target.Sheets.Count]);
                                ((Worksheet)target.Sheets[target.Sheets.Count]).Name = name;
                            }
                        }
                    }
                    finally
                    {
                        if (source != null)
                        {
                            source.Close(false);
                            Marshal.ReleaseComObject(source);
                        }
                    }

                    target.SaveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
                }
            }
            finally
            {
                if (target != null)
                {
                    target.Close(false);
                    Marshal.ReleaseComObject(target);
                }
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
        }

        private static string GetUniqueName(string baseName, Dictionary<string, int> usedNames)
        {
            if (baseName.Length > 31)
            {
                baseName = baseName[..31];
            }

            if (!usedNames.ContainsKey(baseName))
            {
                usedNames[baseName] = 0;
                return baseName;
            }

            int count = ++usedNames[baseName];
            string candidate;
            do
            {
                string suffix = $"_{count}";
                int maxBase = 31 - suffix.Length;
                string trimmed = baseName.Length > maxBase ? baseName[..maxBase] : baseName;
                candidate = trimmed + suffix;
                count++;
            } while (usedNames.ContainsKey(candidate));

            usedNames[baseName] = count - 1;
            usedNames[candidate] = 0;
            return candidate;
        }
    }
}
