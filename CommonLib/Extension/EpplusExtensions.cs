using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.Utility;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;

namespace CommonLib.Extension
{
    public static class EpplusExtensions
    {
        #region workbook

        public static bool IsExist(this ExcelWorkbook workbook, string name)
        {
            foreach (ExcelWorksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == name)
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

        public static void CopyWorkSheets(this ExcelWorkbook workbook, List<string> files)
        {
            if (files == null)
            {
                return;
            }

            foreach (string file in files)
            {
                CopyWorkSheet(workbook, file);
            }
        }

        public static void CopyWorkSheet(this ExcelWorkbook workbook, string file)
        {
            if (file == null)
            {
                return;
            }

            string fileName = Path.GetFileNameWithoutExtension(file);
            if (Path.GetExtension(file).Equals(".csv", StringComparison.CurrentCultureIgnoreCase))
            {
                var format = new ExcelTextFormat
                {
                    Delimiter = ',',
                    Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                    {
                        DateTimeFormat = { ShortDatePattern = "dd-mm-yyyy" }
                    }
                };
                format.Encoding = new UTF8Encoding();
                var fileInfo = new FileInfo(file);
                ExcelWorksheet worksheet = workbook.Worksheets.Add(fileName);
                workbook.Worksheets.MoveBefore(worksheet.Name, workbook.Worksheets[1].Name);
                worksheet.Cells["A1"].LoadFromText(fileInfo, format);
            }
            else if (Path.GetExtension(file).Equals(".txt", StringComparison.CurrentCultureIgnoreCase))
            {
                var format = new ExcelTextFormat
                {
                    Delimiter = '\t',
                    Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                    {
                        DateTimeFormat = { ShortDatePattern = "dd-mm-yyyy" }
                    }
                };
                format.Encoding = new UTF8Encoding();
                var fileInfo = new FileInfo(file);
                ExcelWorksheet worksheet = workbook.Worksheets.Add(fileName);
                worksheet.Cells["A1"].LoadFromText(fileInfo, format);
            }
            else
            {
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                    {
                        workbook.AddSheet(worksheet);
                    }
                }
            }
        }

        public static void AddSheet(this ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            bool isExist = false;
            foreach (ExcelWorksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == worksheet.Name)
                {
                    isExist = true;
                }
            }

            if (isExist)
            {
                workbook.Worksheets[worksheet.Name].Cells.Clear();
            }
            else
            {
                workbook.Worksheets.Add(worksheet.Name, worksheet);
            }

            workbook.Worksheets.MoveBefore(worksheet.Name, workbook.Worksheets[1].Name);
        }

        public static void DeleteSheet(this ExcelWorkbook workbook, string name)
        {
            foreach (ExcelWorksheet sheet in workbook.Worksheets)
            {
                if (name.Equals(sheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    workbook.Worksheets.Delete(sheet);
                    break;
                }
            }
        }

        public static ExcelWorksheet AddSheet(this ExcelWorkbook workbook, string name)
        {
            bool isExist = false;
            foreach (ExcelWorksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == name)
                {
                    isExist = true;
                }
            }

            if (isExist)
            {
                workbook.Worksheets[name].Cells.Clear();
            }
            else
            {
                workbook.Worksheets.Add(name);
            }

            workbook.Worksheets.MoveBefore(name, workbook.Worksheets[1].Name);
            return workbook.Worksheets[1];
        }

        private static void AddHeaderStyle(this ExcelWorkbook workbook)
        {
            foreach (OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml item in workbook.Styles.NamedStyles)
            {
                if (item.Name == "Header")
                {
                    return;
                }
            }

            OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyle = workbook.Styles.CreateNamedStyle("Header");
            namedStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyle.Style.Fill.BackgroundColor.SetColor(Color.YellowGreen);
        }

        public static void AddTxt(this ExcelWorkbook workbook, string fileName)
        {
            string[] lines = File.ReadAllLines(fileName);
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            ExcelWorksheet newSheet = workbook.Worksheets[sheetName];
            if (newSheet == null)
            {
                newSheet = workbook.AddSheet(sheetName);
            }

            int rowCount = lines.Count() + 1;
            int columnCount = lines.Max(x => x.Split('\t').Count());

            object[,] values = new object[rowCount, columnCount];
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                List<string> arr = line.Split('\t').ToList();
                for (int j = 0; j < arr.Count; j++)
                {
                    values[i, j] = arr[j];
                }
            }

            newSheet.Name = sheetName;
            newSheet.Cells[1, 1, rowCount, columnCount].Value = values;
        }

        #endregion

        #region worksheets

        public static void AddSheet(this ExcelWorksheets worksheets, ExcelWorksheet worksheet)
        {
            bool isExist = false;
            foreach (ExcelWorksheet sheet in worksheets)
            {
                if (sheet.Name == worksheet.Name)
                {
                    isExist = true;
                }
            }

            if (isExist)
            {
                worksheets[worksheet.Name].Cells.Clear();
            }
            else
            {
                worksheets.Add(worksheet.Name, worksheet);
            }

            worksheets.MoveBefore(worksheet.Name, worksheets[1].Name);
        }

        public static ExcelWorksheet InsertSheet(this ExcelWorksheets excelWorksheets, string name)
        {
            bool isExist = false;
            foreach (ExcelWorksheet sheet in excelWorksheets)
            {
                if (sheet.Name == name)
                {
                    isExist = true;
                }
            }

            if (isExist)
            {
                excelWorksheets[name].Cells.Clear();
            }
            else
            {
                excelWorksheets.Add(name);
            }

            excelWorksheets.MoveBefore(name, excelWorksheets[1].Name);
            return excelWorksheets[1];
        }

        public static ExcelWorksheet AddSheet(this ExcelWorksheets excelWorksheets, string name)
        {
            bool isExist = false;
            foreach (ExcelWorksheet sheet in excelWorksheets)
            {
                if (sheet.Name == name)
                {
                    isExist = true;
                }
            }

            if (isExist)
            {
                excelWorksheets[name].Cells.Clear();
            }
            else
            {
                excelWorksheets.Add(name);
            }

            excelWorksheets.MoveBefore(name, excelWorksheets[1].Name);
            return excelWorksheets[1];
        }

        #endregion

        #region worksheet
        public static string GetCellValue(this ExcelWorksheet excelWorksheet, int row, int col)
        {
            if (col <= 0 || row <= 0)
            {
                return "";
            }

            return excelWorksheet.GetMergeCellValue(row, col);
        }

        private static string GetMergeCellValue(this ExcelWorksheet excelWorksheet, int row, int col)
        {
            string mergedCell = excelWorksheet.MergedCells[row, col];
            if (mergedCell == null)
            {
                if (excelWorksheet.Cells[row, col].Value is string)
                {
                    return excelWorksheet.Cells[row, col].Text.Trim();
                }

                if (excelWorksheet.Cells[row, col].Value is double)
                {
                    return excelWorksheet.Cells[row, col].Value.ToString().Trim();
                }

                if (excelWorksheet.Cells[row, col].Value == null)
                {
                    return excelWorksheet.Cells[row, col].Text.Trim();
                }

                if (excelWorksheet.Cells[row, col].Value is bool)
                {
                    return excelWorksheet.Cells[row, col].Value.ToString().Trim();
                }

                return excelWorksheet.Cells[row, col].Text != null ? excelWorksheet.Cells[row, col].Text.Trim() : string.Empty;
            }

            ExcelRange cell = excelWorksheet.Cells[new ExcelAddress(mergedCell).Start.Row, new ExcelAddress(mergedCell).Start.Column];

            if (cell.Value is string)
            {
                return cell.Text.Trim();
            }

            if (cell.Value is double o)
            {
                return o.ToString(CultureInfo.InvariantCulture).Trim();
            }

            if (cell.Value == null)
            {
                return cell.Text.Trim();
            }

            if (cell.Value is bool value)
            {
                return value.ToString();
            }

            return cell.Text != null ? cell.Text.Trim() : string.Empty;
        }

        public static void FindCellByValue(this ExcelWorksheet sheet, ref int row, ref int col, string value)
        {
            for (int i = 1; i <= sheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= sheet.Dimension.Columns; j++)
                {
                    string lCellContent = sheet.GetMergeCellValue(i, j);
                    if (lCellContent.Equals(value, StringComparison.OrdinalIgnoreCase))
                    {
                        row = i;
                        col = j;
                        return;
                    }
                }
            }
        }

        public static Dictionary<string, int> GetHeaderOrder(this ExcelWorksheet sheet, int startRow = 1)
        {
            var headerOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (sheet.Dimension == null)
            {
                return headerOrder;
            }

            int endCol = sheet.Dimension.End.Column;
            for (int i = 1; i <= endCol; i++)
            {
                if (sheet.Cells[startRow, i].Value != null)
                {
                    string header = sheet.Cells[startRow, i].Value.ToString().Trim();
                    if (!headerOrder.ContainsKey(header))
                    {
                        headerOrder.Add(header, i);
                    }
                }
            }

            return headerOrder;
        }

        public static string GetMergedCellValueAndAddress(this ExcelWorksheet sheet, int rowNumber, int columnNumber, out string address)
        {
            string range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null
                ? GetCellValueAndAddress(sheet, rowNumber, columnNumber, out address)
                : GetCellValueAndAddress(sheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column,
                    out address);
        }

        public static string GetCellValueAndAddress(this ExcelWorksheet sheet, int row, int column, out string address)
        {
            address = string.Empty;
            if (sheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            if (sheet.Cells[row, column] == null)
            {
                return "";
            }

            if (sheet.Cells[row, column].Value != null)
            {
                ExcelRange range = sheet.Cells[row, column];
                address = range.Address;
                return sheet.Cells[row, column].Value.ToString();
            }

            if (sheet.Cells[row, column].Text != null)
            {
                ExcelRange range = sheet.Cells[row, column];
                address = range.Address;
                return sheet.Cells[row, column].Text;
            }

            return "";
        }

        public static void ExportToTxt(this ExcelWorksheet worksheet, string filePath, string delimiter = "\t", string[,] pColumnNameFixedValue = null, List<int> skipColumnList = null, Dictionary<int, List<string>> extraRows = null)
        {
            if (worksheet == null)
            {
                return;
            }

            if (worksheet.Dimension == null)
            {
                return;
            }

            int columnCount = worksheet.Dimension.End.Column;
            int rowCount = worksheet.Dimension.End.Row;
            string dir = Path.GetDirectoryName(filePath);
            if (dir != null && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (rowCount > 0)
            {
                using (var sw = new StreamWriter(filePath))
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        var arr = new List<string>();
                        for (int j = 1; j <= columnCount; j++)
                        {
                            //skip specified column, can't export specified column
                            if (skipColumnList != null && skipColumnList.Contains(j))
                            {
                                continue;
                            }

                            object objCellValue = worksheet.Cells[i, j].Value;
                            object cellValue = objCellValue != null ? objCellValue : "";
                            arr.Add(cellValue.ToString());
                        }

                        if (pColumnNameFixedValue != null)
                        {
                            //it's the last column of the first row
                            if (i == 1)
                            {
                                for (int lIntRow = 0;
                                     lIntRow < pColumnNameFixedValue.GetLength(0);
                                     lIntRow++)
                                //getLength() 1->0, by Ze
                                {
                                    int lIntIndex = int.Parse(pColumnNameFixedValue[lIntRow, 2]);
                                    arr.Insert(lIntIndex, pColumnNameFixedValue[lIntRow, 0]);
                                }
                            }
                            else
                            //it's the last column of non-first row
                            {
                                for (int lIntColumn = 0;
                                     lIntColumn < pColumnNameFixedValue.GetLength(0);
                                     lIntColumn++)
                                //getLength() 1->0, by Ze
                                {
                                    int lIntIndex = int.Parse(pColumnNameFixedValue[lIntColumn, 2]);
                                    arr.Insert(lIntIndex, pColumnNameFixedValue[lIntColumn, 1]);
                                }
                            }
                        }

                        sw.WriteLine(string.Join(delimiter, arr));
                    }

                    if (extraRows != null)
                    {
                        int extraRowCount = extraRows.Values.Max(o => o.Count);
                        for (int i = 1; i <= extraRowCount; i++)
                        {
                            List<string> extraArr = new string[extraRows.Keys.Max() + 1].ToList();
                            foreach (int index in extraRows.Keys)
                            {
                                if (extraRows[index].Count >= i)
                                {
                                    extraArr[index] = extraRows[index][i - 1];
                                }
                            }

                            sw.WriteLine(string.Join(delimiter, extraArr));
                        }
                    }
                }
            }
        }

        public static void VddLevelExportToTxt(this ExcelWorksheet worksheet, string filePath, int seqIndex, string delimiter = "\t", string[,] pColumnNameFixedValue = null)
        {
            if (worksheet == null)
            {
                return;
            }

            if (worksheet.Dimension == null)
            {
                return;
            }

            int columnCount = worksheet.Dimension.End.Column;
            int rowCount = worksheet.Dimension.End.Row;
            string dir = Path.GetDirectoryName(filePath);
            if (dir != null && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (rowCount > 0)
            {
                using (var sw = new StreamWriter(filePath))
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        object seqColumnValue = worksheet.Cells[i, seqIndex].Value;
                        object seqColumnStr = seqColumnValue != null ? seqColumnValue : "";
                        if (seqColumnStr.ToString().Equals("x", StringComparison.CurrentCultureIgnoreCase))
                        {
                            continue;
                        }

                        var arr = new List<string>();
                        for (int j = 1; j <= columnCount; j++)
                        {
                            object objCellValue = worksheet.Cells[i, j].Value;
                            object cellValue = objCellValue != null ? objCellValue : "";
                            arr.Add(cellValue.ToString());
                        }

                        sw.WriteLine(string.Join(delimiter, arr));
                    }
                }
            }
        }

        public static List<int> SplitByEmptyRow(this ExcelWorksheet worksheet)
        {
            var worksheets = new List<ExcelWorksheet>();
            int startColNumber = worksheet.Dimension.Start.Column;
            int startRowNumber = worksheet.Dimension.Start.Row;
            int endColNumber = worksheet.Dimension.End.Column;
            int endRowNumber = worksheet.Dimension.End.Row + 1;
            var splitRows = new List<int>();
            for (int i = startRowNumber; i <= endRowNumber; i++)
            {
                IEnumerable<string> range = worksheet.Cells[i, startColNumber, i, endColNumber].ToList().Select(x => x.Text);
                if (range.All(string.IsNullOrEmpty))
                {
                    splitRows.Add(i);
                }
            }

            for (int index = 0; index < splitRows.Count; index++)
            {
                int splitRow = splitRows[index];
                if (index - 1 >= 0 && splitRow == splitRows[index - 1] + 1)
                {
                    splitRows.RemoveAt(index);
                }
            }

            return splitRows;
        }

        public static void InsertColumn(this ExcelWorksheet worksheet, string title, int row, int column)
        {
            worksheet.InsertColumn(column, 1);
            worksheet.Cells[row, column].Value = title;
            worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.DeepPink);
        }

        public static List<Error> CheckFormula(this ExcelWorksheet worksheet)
        {
            int endColumn = worksheet.Dimension.End.Column;
            int endRow = worksheet.Dimension.End.Row;
            var errors = new List<Error>();
            for (int i = 1; i <= endRow; i++)
            {
                for (int j = 1; j <= endColumn; j++)
                {
                    bool formulaError = false;

                    if (!string.IsNullOrEmpty(worksheet.Cells[i, j].Formula))
                    {
                        if (worksheet.Cells[i, j].Formula.Contains("_xlfn"))
                        {
                            formulaError = true;
                            //if (worksheet.Cells[i, j].Formula.Contains("_xlfn.SINGLE"))
                            //{
                            //    var text = worksheet.Cells[i, j].Formula;
                            //    text = Regex.Replace(text, "_xlfn.SINGLE", "", RegexOptions.IgnoreCase);
                            //    worksheet.Cells[i, j].Formula = text;
                            //    continue;
                            //}
                        }
                    }

                    string value = EpplusOperation.GetCellValue(worksheet, i, j);
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    if (Regex.IsMatch(value, @"#NULL!", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(value, @"#DIV/0!", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(value, @"#VALUE!", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(value, @"#REF!", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(value, @"#NAME?", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(value, @"#NUM!", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(value, @"#N/A", RegexOptions.IgnoreCase) ||
                        formulaError)
                    {
                        string errorMessage = string.Format("The formula incorrect!!!");
                        var error = new Error
                        {
                            ErrorType = BinCutErrorType.Formula,
                            SheetName = worksheet.Name,
                            ErrorLevel = ErrorLevel.Error,
                            RowNum = i,
                            ColNum = j,
                            Message = errorMessage,
                        };
                        errors.Add(error);
                    }
                }
            }

            return errors;
        }

        public static void FixFormula(this ExcelWorksheet worksheet)
        {
            int endColumn = worksheet.Dimension.End.Column;
            int endRow = worksheet.Dimension.End.Row;
            for (int i = 1; i <= endRow; i++)
            {
                for (int j = 1; j <= endColumn; j++)
                {
                    if (!string.IsNullOrEmpty(worksheet.Cells[i, j].Formula))
                    {
                        if (worksheet.Cells[i, j].Formula.Contains("_xlfn.SINGLE"))
                        {
                            string text = worksheet.Cells[i, j].Formula;
                            text = Regex.Replace(text, "_xlfn.SINGLE", "", RegexOptions.IgnoreCase);
                            worksheet.Cells[i, j].Formula = text;

                        }
                        else if (worksheet.Cells[i, j].Formula.Contains("_xlfn.CONCAT"))
                        {
                            string text = worksheet.Cells[i, j].Formula;
                            text = Regex.Replace(text, "_xlfn.CONCAT", "CONCATENATE", RegexOptions.IgnoreCase);
                            worksheet.Cells[i, j].Formula = text;
                        }
                    }
                }
            }
        }

        public static DataTable ReadSheetToDataSet(this ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null)
            {
                return null;
            }

            int columnCount = worksheet.Dimension.End.Column;
            int rowCount = worksheet.Dimension.End.Row;
            var dt = new DataTable(worksheet.Name);
            if (rowCount > 0)
            {
                object objCellValue;
                object cellValue;
                for (int j = 0; j < columnCount; j++) //設定DataTable列名
                {
                    objCellValue = worksheet.Cells[1, j + 1].Value;
                    cellValue = objCellValue == null ? "" : objCellValue.ToString();
                    dt.Columns.Add(cellValue.ToString(), typeof(object));
                }

                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        objCellValue = worksheet.Cells[i, j].Value;
                        cellValue = objCellValue ?? "";
                        dr[j - 1] = cellValue;
                    }

                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        public static void OutLineColumnByCell(this ExcelWorksheet worksheet, int colNum, int rowNum = 0)
        {
            const int count = 1;
            int startRow = rowNum != 0 ? rowNum : worksheet.Dimension.Start.Row;
            var list = new Dictionary<int, int>();
            for (int i = startRow; i <= worksheet.Dimension.End.Row; i++)
            {
                object now = worksheet.Cells[i, colNum].Value;
                object next = i + 1 < worksheet.Dimension.End.Row ? worksheet.Cells[i + 1, colNum].Value : null;
                if (now != null && next != null)
                {
                    if (now.Equals(next))
                    {
                        int start = i;
                        do
                        {
                            i++;
                            next = i + 1 <= worksheet.Dimension.End.Row ? worksheet.Cells[i + 1, colNum].Value : null;
                        } while (now.Equals(next));

                        list.Add(start, i);
                    }
                }
            }

            foreach (KeyValuePair<int, int> item in list)
            {
                for (int i = item.Key; i < item.Value; i++)
                {
                    worksheet.Row(i).OutlineLevel = count;
                    worksheet.Row(i).Collapsed = true;
                }
            }

        }

        public static void MergeColumn(this ExcelWorksheet worksheet, int colNum, bool isHeaders = true)
        {
            for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
            {
                object now = worksheet.Cells[i, colNum].Value;
                object next = i + 1 <= worksheet.Dimension.End.Row ? worksheet.Cells[i + 1, colNum].Value : null;
                if (now != null && next != null && now.Equals(next))
                {
                    int start = i;
                    do
                    {
                        i++;
                        next = i + 1 <= worksheet.Dimension.End.Row ? worksheet.Cells[i + 1, colNum].Value : null;
                    } while (now.Equals(next));

                    worksheet.Cells[start, colNum, i, colNum].Merge = true;
                    worksheet.Cells[start, colNum, i, colNum].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
            }
        }

        public static void SetFormula(this ExcelWorksheet worksheet, int colNum, bool isHeaders = true)
        {
            int start = isHeaders ? 2 : 1;
            for (int i = start; i <= worksheet.Dimension.End.Row; i++)
            {
                if (worksheet.Cells[i, colNum].Value != null)
                {
                    SetHyperLinkFormat(worksheet.Cells[i, colNum]);
                }
            }
        }

        public static void SetHyperLinkFormat(this ExcelRange excelRange)
        {
            excelRange.Formula = excelRange.Value.ToString();
            excelRange.Style.Font.UnderLine = true;
            excelRange.Style.Font.Color.SetColor(Color.Blue);
        }

        public static void SetHeaderStyle(this ExcelWorksheet worksheet)
        {
            worksheet.Workbook.AddHeaderStyle();
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.Start.Row, worksheet.Dimension.End.Column].StyleName =
                "Header";
        }

        public static void SetHairBorder(this ExcelWorksheet worksheet)
        {
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Top.Style =
                ExcelBorderStyle.Hair;
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Bottom.Style =
                ExcelBorderStyle.Hair;
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Left.Style =
                ExcelBorderStyle.Hair;
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Right.Style =
                ExcelBorderStyle.Hair;
        }

        public static void ExportSheet(this ExcelWorksheet worksheet, string txtFileName, string delimiter = "\t")
        {
            int maxRow = worksheet.Dimension.Rows;
            int maxCol = worksheet.Dimension.Columns;
            StreamWriter sw = File.CreateText(txtFileName);
            for (int iRow = 1; iRow <= maxRow; iRow++)
            {
                var row = new List<string>();
                for (int iCol = 1; iCol <= maxCol; iCol++)
                {
                    object value = worksheet.Cells[iRow, iCol].Value ?? "";
                    row.Add(value.ToString());
                }

                sw.WriteLine(string.Join(delimiter, row));
            }

            sw.Close();
        }

        public static List<string> ConvertToLines(this ExcelWorksheet worksheet)
        {
            int maxRow = worksheet.Dimension.Rows;
            int maxCol = worksheet.Dimension.Columns;
            var lines = new List<string>();
            for (int iRow = 1; iRow <= maxRow; iRow++)
            {
                var row = new List<string>();
                for (int iCol = 1; iCol <= maxCol; iCol++)
                {
                    object value = worksheet.Cells[iRow, iCol].Value ?? "";
                    row.Add(value.ToString());
                }

                lines.Add(string.Join("\t", row));
            }

            return lines;
        }

        #endregion

        #region Print

        public static int PrintExcelRow<T>(this ExcelRange range, T[] data)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            var arrays = new List<object[]>();
            object[] array = new object[data.Count()];
            for (int index = 0; index < data.Length; index++)
            {
                T item = data[index];
                array[index] = item == null ? "" : item;
            }

            arrays.Add(array);
            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
            return startRow + arrays.Count();
        }

        public static void PrintExcelCol<T>(this ExcelRange range, T[] data)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            var dataArray = new T[data.GetLength(0), 1];
            for (int i = 0; i < dataArray.GetLength(0); i++)
            {
                for (int j = 0; j < dataArray.GetLength(1); j++)
                {
                    dataArray[i, j] = data[i];
                }
            }

            worksheet.Cells[startRow, startCol].LoadFromCollection(data);
        }

        public static int PrintExcelRange<T>(this ExcelRange range, T[,] data)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            var arrays = new List<object[]>();
            for (int i = 0; i < data.GetLength(0); i++)
            {
                object[] array = new object[data.GetLength(1)];
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    array[j] = data[i, j];
                }

                arrays.Add(array);
            }

            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
            return startRow + data.GetLength(0);
        }

        public static void PrintExcelColByList<T>(this ExcelRange range, List<List<T>> list)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            for (int index = 0; index < list.Count; index++)
            {
                List<T> col = list[index];
                worksheet.Cells[startRow, startCol + index].PrintExcelCol(col.ToArray());
            }

            //for (var index = startCol - 1; index < list.Count; index++)
            //{
            //    var item = list[index];
            //    worksheet.Cells[startRow, index, startRow + rowCnt - 1, index].LoadFromCollection(item);
            //}
        }

        public static int PrintExcelRowByList<T>(this ExcelRange range, List<List<T>> list)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            if (list.Count == 0)
            {
                return 0;
            }

            for (int index = 0; index < list.Count; index++)
            {
                List<T> row = list[index];
                worksheet.Cells[startRow + index, startCol].PrintExcelRow(row.ToArray());
            }

            return startRow + list.Count;
        }

        public static void AddHyperLink(this ExcelRange range, string sheetName2, int x2, int y2)
        {
            if (x2 > 0 && y2 > 0)
            {
                string cellPosBase = ExcelCellBase.GetAddress(x2, y2);
                range.Hyperlink = new ExcelHyperLink(sheetName2 + "!" + cellPosBase, range.Text);
                range.StyleName = "HyperLink";
                range.Style.Font.UnderLine = true;
            }
        }

        #endregion

        public static void AddMarcoFromBas(this ExcelPackage excel, string filePath, string moduleName)
        {
            if (File.Exists(filePath))
            {
                OfficeOpenXml.VBA.ExcelVBAModule module = IsExistModule(excel, moduleName)
                    ? excel.Workbook.VbaProject.Modules[moduleName]
                    : excel.Workbook.VbaProject.Modules.AddModule(moduleName);
                var sb = new StringBuilder();
                using (var reader = new StreamReader(filePath))
                {
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine() + "\r\n";
                        if (!(line.StartsWith("Attribute")))
                        {
                            sb.Append(line);
                        }
                    }
                }

                module.Code = sb.ToString();
            }
        }

        public static bool IsExistModule(this ExcelPackage excel, string moduleName)
        {
            bool flag = false;
            foreach (OfficeOpenXml.VBA.ExcelVBAModule item in excel.Workbook.VbaProject.Modules)
            {
                if (item.Name == moduleName)
                {
                    flag = true;
                }
            }

            return flag;
        }

        public static void SetPercentFormat(this ExcelRange excelRange, string value, bool fill = true)
        {
            excelRange.Formula = value;
            excelRange.Style.Numberformat.Format = "0.00%";

            if (!fill)
            {
                return;
            }

            OfficeOpenXml.ConditionalFormatting.Contracts.IExcelConditionalFormattingDataBarGroup cf = excelRange.ConditionalFormatting.AddDatabar(Color.FromArgb(153, 0, 255));
            cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf.HighValue.Value = 1;
            cf.LowValue.Value = 0;

            XmlNodeList nodes = cf.Node.ChildNodes;
            if (nodes != null)
            {
                foreach (XmlElement node in nodes)
                {
                    node.SetAttribute("minLength", "0");
                    node.SetAttribute("maxLength", "100");
                }
            }
        }

        public static List<List<string>> ConvertToLists(this string file)
        {
            string keyComma = "#comma>";
            string keyNewline = "#newline>";
            string keyDouble = "#double>";
            string text = File.ReadAllText(file).Replace("\"\"", keyDouble);
            string[] matches = Regex.Split(text, @"(\""[^""]+\"")");
            var sb = new StringBuilder();
            foreach (string match in matches)
            {
                if (match.StartsWith("\"") && match.EndsWith("\""))
                {
                    sb.Append(match.Trim('"').Replace(",", keyComma).Replace("\r\n", keyNewline).Replace("\n", keyNewline));
                }
                else
                {
                    sb.Append(match);
                }
            }

            List<string> lines = sb.ToString().ToLines();
            var lists = new List<List<string>>(lines.Count);
            foreach (string line in lines)
            {
                string[] cells = line.Split(',');
                var row = new List<string>(cells.Length);
                foreach (string cell in cells)
                {
                    row.Add(cell.Replace(keyComma, ",").Replace(keyNewline, "\n").Replace(keyDouble, "\""));
                }
                lists.Add(row);
            }

            return lists;
        }
    }
}
