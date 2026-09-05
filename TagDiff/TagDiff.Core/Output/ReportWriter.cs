using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.Enums;

using LogLib.Static;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

using TagDiff.Core.Const;
using TagDiff.Core.Static;
using TagDiff.Core.Utility;

namespace TagDiff.Core.Output
{
    internal class ReportWriter(bool isUnitTest, string outputPath, Dictionary<string, List<string>> fileInfos)
    {
        public const string Data = "Data";
        public const string Pivot = "Pivot";

        public static readonly IReadOnlyDictionary<string, int> CategoryIndexMap = new Dictionary<string, int>(StringExtensions.IgnoreCase)
        {
            { Category.Basic, 0 },
            { Category.Dc, 1 },
            { Category.Hardip, 2 },
            { Category.Atpg, 3 },
            { Category.Efuse, 4 },
            { Category.Rtos, 5 },
            { Category.BinCut, 6 },
            { Category.Cz, 7 },
            { Category.RegAssign, 8 },
            { Category.Other, 9 },
            { Category.NotTracked, 10 }
        };

        private readonly bool _isUnitTest = isUnitTest;
        private readonly string _outputPath = outputPath;
        private readonly Dictionary<string, List<string>> _fileInfos = fileInfos;

        internal void WorkFlow(string batFolder, ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)> fileMap, string reportPath, List<CompareReport> compareReports, IEnumerable<string>? excludeDirs = null)
        {
            using var excelPackage = new ExcelPackage();
            Response.Report("Writing Detail Sheets ...", EnumMessageLevel.General, 90);
            WriteDetailSheet(fileMap, excelPackage);

            Response.Report("Writing Legend Sheet ...", EnumMessageLevel.General, 90);
            WriteLegend(excelPackage);

            if (!_isUnitTest)
            {
                Response.Report("Writing InputFile_Information Sheet ...", EnumMessageLevel.General, 90);
                excelPackage.Workbook.PrintInputFileInformation(_fileInfos);
            }

            HashSet<string> excludedSheetNames = GetExcludedSheetNames(excludeDirs);

            string relative = batFolder.Replace(_outputPath, "").TrimStart('\\');
            Response.Report("Writing Data Sheet ...", EnumMessageLevel.General, 90);
            WriteDataReport(excelPackage, compareReports, relative, fileMap, excludedSheetNames);

            Response.Report("Writing Pivot Sheet ...", EnumMessageLevel.General, 90);
            List<CompareReport> pivotCompareReports = excludedSheetNames.Count == 0
                ? compareReports
                : [.. compareReports.Where(report => !excludedSheetNames.Contains(report.SheetName.Split('(')[0]))];
            WritePivotTable(excelPackage, pivotCompareReports);

            if (excelPackage.Workbook.IsExist(Pivot))
            {
                Response.Report("Writing OverView.txt ...", EnumMessageLevel.General, 90);
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[Pivot];
                ExcelToTxt(worksheet, reportPath);
            }

            Response.Report("Saving Excel ...", EnumMessageLevel.General, 90);
            excelPackage.Compression = CompressionLevel.BestSpeed;
            excelPackage.SaveAs(new FileInfo(reportPath));
        }

        private static HashSet<string> GetExcludedSheetNames(IEnumerable<string>? excludeDirs)
        {
            var excludedSheetNames = new HashSet<string>(StringExtensions.IgnoreCase);
            if (excludeDirs == null)
            {
                return excludedSheetNames;
            }

            foreach (string excludeDir in excludeDirs)
            {
                if (string.IsNullOrWhiteSpace(excludeDir) || !Directory.Exists(excludeDir))
                {
                    continue;
                }

                foreach (string file in Directory.GetFiles(excludeDir, "*.*", SearchOption.AllDirectories))
                {
                    excludedSheetNames.Add(Path.GetFileNameWithoutExtension(file));
                }
            }

            return excludedSheetNames;
        }

        private static void WriteDetailSheet(ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)> fileMap, ExcelPackage excelPackage)
        {
            // Phase 1: read and preprocess all files in parallel (I/O bound)
            var fileContents = fileMap
                .AsParallel()
                .Select(file =>
                {
                    (EnumSheetType sheetType, string bf, string _) = file.Value;
                    if (string.IsNullOrWhiteSpace(bf) || !File.Exists(bf))
                    {
                        return (key: file.Key, sheetType, text: "");
                    }

                    // due to format error of patSets_all in igxl
                    string text = File.ReadAllText(bf);
                    if (text.Contains('\0'))
                    {
                        string[] lines = text.Split('\n');
                        if (lines.Length >= 2)
                        {
                            lines[^2] = lines[^2].Replace("\0", "");
                            lines[^1] = lines[^1].Replace("\0", "");
                        }
                        else if (lines.Length == 1)
                        {
                            lines[0] = lines[0].Replace("\0", "");
                        }
                        text = string.Join("\n", lines);
                    }
                    text = text.TrimEnd('\r', '\n');
                    return (key: file.Key, sheetType, text);
                })
                .ToList();

            // every column must be forced to String, otherwise EPPlus auto-detects cell types
            // and misreads pin names like "9.z412" as dates (see TagDiffMain history)
            int maxColumnCount = fileContents.Count == 0
                ? 0
                : fileContents.Max(f => f.text.Length == 0 ? 0 : f.text.Split('\n').Max(line => line.Split('\t').Length));
            var format = new ExcelTextFormat
            {
                Delimiter = '\t',
                DataTypes = [.. Enumerable.Repeat(eDataTypes.String, maxColumnCount)]
            };

            // Phase 2: write to EPPlus sequentially (EPPlus is not thread-safe)
            foreach ((string key, EnumSheetType sheetType, string text) in fileContents)
            {
                ExcelWorksheet sheet = excelPackage.Workbook.Worksheets[key] ?? excelPackage.Workbook.AddSheet(key);
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                sheet.Cells["A1"].LoadFromText(text, format);
                SheetFormat.SetSheetFormat(sheet);
                if (sheetType != EnumSheetType.DTUnknown)
                {
                    sheet.Cells[1, 3].Formula = SheetFormat.GetHyperlinkFormula(Data, 1, 0, "Return");
                }
            }
        }

        public static void ExcelToTxt(ExcelWorksheet excelWorksheet, string filePath)
        {
            if (excelWorksheet.Dimension == null)
            {
                return;
            }

            string outputTxt = Path.Combine(Path.GetDirectoryName(filePath) ?? string.Empty, Path.GetFileNameWithoutExtension(filePath) + "_OverView.txt");
            using var writer = new StreamWriter(outputTxt);
            int rowCount = excelWorksheet.Dimension.Rows;
            int colCount = excelWorksheet.Dimension.Columns;
            string[] rowValues = new string[colCount];
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    rowValues[col - 1] = excelWorksheet.Cells[row, col].Text;
                }
                writer.WriteLine(string.Join("\t", rowValues));
            }
        }

        private static (int passCount, int failCount, double rate) ComputeMetrics(int total, int manual, int semiAuto)
        {
            int passCount = total - manual - semiAuto;
            int failCount = manual + semiAuto;
            double rate = total > 0 ? (double)passCount / total : 0;
            return (passCount, failCount, rate);
        }

        private static void SetCellColorFormat(ExcelRange excelRange, Color color)
        {
            if (excelRange == null)
            {
                return;
            }

            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Color.SetColor(Color.White);
            excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelRange.Style.Fill.BackgroundColor.SetColor(color);
        }

        public static void WriteDataReport(ExcelPackage excelPackage, List<CompareReport> compareReports, string relative, ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)> fileMap, HashSet<string>? excludedSheetNames = null)
        {
            excludedSheetNames ??= new HashSet<string>(StringExtensions.IgnoreCase);

            compareReports = [.. compareReports
                .OrderBy(x => CategoryIndexMap.TryGetValue(x.Category ?? "", out int idx) ? idx : int.MaxValue)
                .ThenBy(x => x.Block)
                .ThenBy(x => x.SheetName)];

            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(Data);
            excelPackage.Workbook.Worksheets.MoveBefore(Data, excelPackage.Workbook.Worksheets[0].Name);
            sheet.Cells[1, 1].Value = "Category";
            sheet.Cells[1, 2].Value = "Sub Category";
            sheet.Cells[1, 3].Value = "Block";
            sheet.Cells[1, 4].Value = "Sheet Name";
            sheet.Cells[1, 5].Value = "Link";
            sheet.Cells[1, 6].Value = "Exclude";
            sheet.Cells[1, 7].Value = "T-AutoGen";
            sheet.Cells[1, 8].Value = "Manual";
            sheet.Cells[1, 9].Value = "Amount of item";
            sheet.Cells[1, 10].Value = "ManualCount";
            sheet.Cells[1, 11].Value = "Semi-auto";
            sheet.Cells[1, 12].Value = "FailCount";
            sheet.Cells[1, 13].Value = "Description";

            SetCellColorFormat(sheet.Cells[1, 1, 1, 9], Color.Gray);
            SetCellColorFormat(sheet.Cells[1, 10], Color.Red);
            sheet.Cells[1, 10].Style.Font.Color.SetColor(Color.Black);
            SetCellColorFormat(sheet.Cells[1, 11], Color.Yellow);
            sheet.Cells[1, 11].Style.Font.Color.SetColor(Color.Black);
            SetCellColorFormat(sheet.Cells[1, 12, 1, 14], Color.Gray);

            foreach (CompareReport compareReport in compareReports)
            {
                string sheetName = compareReport.SheetName;
                string friendName = sheetName;
                if (fileMap.TryGetValue(sheetName, out (EnumSheetType sheetType, string baseFile, string compareFile) tuple))
                {
                    if (string.IsNullOrEmpty(tuple.compareFile))
                    {
                        friendName = compareReport.SheetName + "(X)";
                    }
                }
                compareReport.Exclude = excludedSheetNames.Contains(sheetName) || compareReport.Category == Category.NotTracked;
                compareReport.SheetName = friendName;
                compareReport.Link = "=HYPERLINK(\"#\'" + sheetName + "\'!A1\",\"" + sheetName + "\")";
            }

            sheet.Cells[2, 1].LoadFromCollection(compareReports, false);

            IExcelConditionalFormattingEqual excludeRule = sheet.ConditionalFormatting.AddEqual(sheet.Cells[2, 6, sheet.Dimension.End.Row, 6]);
            excludeRule.Formula = "TRUE";
            excludeRule.Style.Fill.PatternType = ExcelFillStyle.Solid;
            excludeRule.Style.Fill.BackgroundColor.Color = Color.Red;

            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT") ?? "";
            if (!string.IsNullOrEmpty(oasisRootFolder))
            {
                sheet.Cells[1, 14].Value = "KDiff";
                SetCellColorFormat(sheet.Cells[1, 13], Color.Gray);
                for (int i = 0; i < compareReports.Count; i++)
                {
                    CompareReport compareReport = compareReports[i];
                    string sheetName = compareReport.SheetName;
                    string bat = Path.Combine(".", relative, sheetName + ".bat");
                    string friendName = sheetName;
                    if (fileMap.TryGetValue(sheetName, out (EnumSheetType sheetType, string baseFile, string compareFile) tuple))
                    {
                        if (string.IsNullOrEmpty(tuple.compareFile))
                        {
                            friendName = sheetName + "(X)";
                        }
                    }

                    sheet.Cells[i + 2, 14].Formula = $@"HYPERLINK(""{bat}"",""{friendName}"")";
                }
            }

            int startRow = 2;
            int startCol = 1;
            int endRow = sheet.Dimension.End.Row;
            int endCol = sheet.Dimension.End.Column;
            for (int currentRow = startRow; currentRow <= endRow; currentRow++)
            {
                sheet.Cells[currentRow, 7].SetPercentFormat($"=1-$H{currentRow}");
                sheet.Cells[currentRow, 8].SetPercentFormat($"=($J{currentRow}+$K{currentRow})/$I{currentRow}");
            }

            sheet.Cells[startRow, startCol, endRow, endCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[startRow, startCol, endRow, endCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells.TryAutoFitColumns();
            sheet.Column(2).Width = 0;
            sheet.Column(4).Width = 36;
            sheet.Column(5).Width = 36;
            sheet.Column(13).Width = 40;
            sheet.Column(13).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Column(14).Width = 40;
            sheet.SetFormula(5);
            sheet.Cells[1, 1, 1, endCol].AutoFilter = true;
            sheet.View.FreezePanes(2, 1);
            sheet.MergeColumn(1);
            sheet.MergeColumn(3);
        }

        public static void WritePivotTable(ExcelPackage excelPackage, List<CompareReport> compareReports)
        {
            ExcelWorksheet dataSheet = excelPackage.Workbook.Worksheets[Data];
            if (dataSheet?.Dimension == null)
            {
                return;
            }

            if (excelPackage.Workbook.IsExist(Pivot))
            {
                excelPackage.Workbook.Worksheets.Delete(Pivot);
            }

            ExcelWorksheet pivotSheet = excelPackage.Workbook.Worksheets.Add(Pivot);
            excelPackage.Workbook.Worksheets.MoveBefore(Pivot, dataSheet.Name);

            var groups = new Dictionary<(string category, string subCategory), (int total, int manual, int semiAuto)>();
            foreach (CompareReport report in compareReports)
            {
                if (TagDiffStatic.IsNotTracked(report.SheetName.Split('(')[0]))
                {
                    continue;
                }

                string category = report.Category ?? string.Empty;
                string subCategory = report.SubCategory ?? string.Empty;

                (string category, string subCategory) key = (category, subCategory);
                if (groups.TryGetValue(key, out (int total, int manual, int semiAuto) existing))
                {
                    groups[key] = (existing.total + report.TotalCount, existing.manual + report.ManualCount, existing.semiAuto + report.SemiAutoCount);
                }
                else
                {
                    groups[key] = (report.TotalCount, report.ManualCount, report.SemiAutoCount);
                }
            }

            pivotSheet.Cells[1, 1].Value = "Sub Category";
            pivotSheet.Cells[1, 2].Value = "Category";
            pivotSheet.Cells[1, 3].Value = "Automation Rate";
            pivotSheet.Cells[1, 4].Value = "Pass Count";
            pivotSheet.Cells[1, 5].Value = "Fail Count";
            pivotSheet.Cells[1, 6].Value = "Total Count";
            SetCellColorFormat(pivotSheet.Cells[1, 1, 1, 6], Color.Gray);

            int currentRow = 2;
            var sortedGroups = groups
                .OrderByDescending(g => g.Key.subCategory)
                .ThenBy(g => CategoryIndexMap.TryGetValue(g.Key.category, out int idx) ? idx : int.MaxValue)
                .ToList();

            int grandTotal = 0;
            int grandManual = 0;
            int grandSemiAuto = 0;

            int subTotalTotal = 0;
            int subTotalManual = 0;
            int subTotalSemiAuto = 0;
            string? currentSubCategory = null;
            foreach (KeyValuePair<(string category, string subCategory), (int total, int manual, int semiAuto)> kvp in sortedGroups)
            {
                if (currentSubCategory != null && currentSubCategory != kvp.Key.subCategory)
                {
                    WriteSubTotalRow(pivotSheet, currentSubCategory, subTotalTotal, subTotalManual, subTotalSemiAuto, ref currentRow);
                    subTotalTotal = 0;
                    subTotalManual = 0;
                    subTotalSemiAuto = 0;
                }

                currentSubCategory = kvp.Key.subCategory;

                int total = kvp.Value.total;
                int manual = kvp.Value.manual;
                int semiAuto = kvp.Value.semiAuto;
                (int passCount, int failCount, double rate) = ComputeMetrics(total, manual, semiAuto);

                pivotSheet.Cells[currentRow, 1].Value = kvp.Key.subCategory;
                pivotSheet.Cells[currentRow, 2].Value = kvp.Key.category;
                pivotSheet.Cells[currentRow, 3].Value = rate;
                pivotSheet.Cells[currentRow, 3].Style.Numberformat.Format = "0.00%";
                pivotSheet.Cells[currentRow, 4].Value = passCount;
                pivotSheet.Cells[currentRow, 5].Value = failCount;
                pivotSheet.Cells[currentRow, 6].Value = total;
                currentRow++;

                subTotalTotal += total;
                subTotalManual += manual;
                subTotalSemiAuto += semiAuto;
                grandTotal += total;
                grandManual += manual;
                grandSemiAuto += semiAuto;
            }

            if (currentSubCategory != null)
            {
                WriteSubTotalRow(pivotSheet, currentSubCategory, subTotalTotal, subTotalManual, subTotalSemiAuto, ref currentRow);
            }

            (int grandPass, int grandFail, double grandRate) = ComputeMetrics(grandTotal, grandManual, grandSemiAuto);

            pivotSheet.Cells[currentRow, 1].Value = "Grand Total";
            pivotSheet.Cells[currentRow, 2].Value = string.Empty;
            pivotSheet.Cells[currentRow, 3].Value = grandRate;
            pivotSheet.Cells[currentRow, 3].Style.Numberformat.Format = "0.00%";
            pivotSheet.Cells[currentRow, 4].Value = grandPass;
            pivotSheet.Cells[currentRow, 5].Value = grandFail;
            pivotSheet.Cells[currentRow, 6].Value = grandTotal;
            SetCellColorFormat(pivotSheet.Cells[currentRow, 1, currentRow, 6], Color.DarkGray);

            pivotSheet.Cells.TryAutoFitColumns();
            pivotSheet.MergeColumn(1);
        }

        public static void WriteLegend(ExcelPackage excelPackage)
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.AddSheet("Legend", true);
            int row = 1;
            // Add headers
            worksheet.Cells[row, 1].Value = "Instruction";
            worksheet.Cells[row, 2].Value = "Example";
            worksheet.Cells[row, 1].FormatHeader();

            // Add data
            worksheet.Cells[++row, 1].Value = "Green = cell values are identical";
            worksheet.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);

            worksheet.Cells[++row, 1].Value = "Yellow = cells are different." + Environment.NewLine + "  • ReferenceProgramValue => AutogenProgramValue";
            worksheet.Cells[row, 1].Style.WrapText = true;
            // Optional: Adjust row height for better visibility
            worksheet.Row(row).Height = 30;
            worksheet.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

            worksheet.Cells[++row, 1].Value = "Red = in Reference program, but not in Autogen program.";
            worksheet.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Red);

            worksheet.Cells[++row, 1].Value = "Grey = in Autogen program, but not in Reference program.";
            worksheet.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.Gray);

            worksheet.Cells[++row, 1].Value = "no color = no tested";

            worksheet.Cells[++row, 1].Value = "Output_hashes, ExecInfo, Versions will be excluded from the summary sheet.";

            worksheet.Column(1).Width = 120;
            worksheet.Column(2).Width = 15;
        }

        private static void WriteSubTotalRow(ExcelWorksheet excelWorksheet, string subCategory, int total, int manual, int semiAuto, ref int row)
        {
            (int passCount, int failCount, double rate) = ComputeMetrics(total, manual, semiAuto);

            excelWorksheet.Cells[row, 1].Value = subCategory + " Total";
            excelWorksheet.Cells[row, 2].Value = string.Empty;
            excelWorksheet.Cells[row, 3].Value = rate;
            excelWorksheet.Cells[row, 3].Style.Numberformat.Format = "0.00%";
            excelWorksheet.Cells[row, 4].Value = passCount;
            excelWorksheet.Cells[row, 5].Value = failCount;
            excelWorksheet.Cells[row, 6].Value = total;
            SetCellColorFormat(excelWorksheet.Cells[row, 1, row, 6], Color.SlateGray);
            row++;
        }
    }
}
