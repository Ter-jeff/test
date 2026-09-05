using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CommonLib.Extension
{
    public static partial class EpplusExtensions
    {
        public static void ExportWorkBook2Txt(this ExcelWorksheet excelWorksheet, string outPath)
        {
            if (!Directory.Exists(outPath))
            {
                Directory.CreateDirectory(outPath);
            }

            int maxRow = excelWorksheet.Dimension.End.Row;
            int maxCol = excelWorksheet.Dimension.End.Column;
            if (maxRow > 99999)
            {
                maxRow = 30000;
            }

            if (maxCol > 500)
            {
                maxCol = 300;
            }

            string? contentCheck = "";
            string txtFileName = Path.Combine(outPath, excelWorksheet.Name + ".txt");
            StreamWriter sw = File.CreateText(txtFileName);
            const int tabStart = 1;
            for (int iRow = 1; iRow <= maxRow; iRow++)
            {
                for (int iCol = 1; iCol <= maxCol; iCol++)
                {
                    if (iCol > tabStart)
                    {
                        sw.Write("\t");
                    }

                    sw.Write(excelWorksheet.Cells[iRow, iCol].Value);
                    if (excelWorksheet.Cells[iRow, iCol].Value != null)
                    {
                        if (!string.IsNullOrEmpty(excelWorksheet.Cells[iRow, iCol].Value.ToString()) && string.IsNullOrEmpty(contentCheck))
                        {
                            contentCheck = excelWorksheet.Cells[iRow, iCol].Value.ToString();
                        }
                    }
                }
                sw.WriteLine();
            }
            sw.Close();
            if (string.IsNullOrEmpty(contentCheck))
            {
                File.Delete(txtFileName);
            }
        }

        public static string GetCellLine(this ExcelWorksheet excelWorksheet, int startRow = 1)
        {
            string line = "";
            if (excelWorksheet.Dimension == null)
            {
                return line;
            }

            int endCol = excelWorksheet.Dimension.End.Column;
            for (int i = 1; i <= endCol; i++)
            {
                object value = excelWorksheet.Cells[startRow, i].Value;
                value ??= "";
                string word = i == 1 ? "" : "\t";
                line = line + word + value;

            }
            return line;
        }

        public static string GetCellValue(this ExcelWorksheet excelWorksheet, int row, int col)
        {
            if (col <= 0 || row <= 0)
            {
                return "";
            }

            return excelWorksheet.GetMergeCellValue(row, col);
        }

        public static string GetMergeCellValue(this ExcelWorksheet excelWorksheet, int row, int col)
        {
            string mergedCell = excelWorksheet.MergedCells[row, col];
            if (string.IsNullOrEmpty(mergedCell))
            {
                return GetNonMergeCellValue(excelWorksheet, row, col);
            }

            ExcelRange cell = excelWorksheet.Cells[new ExcelAddress(mergedCell).Start.Row, new ExcelAddress(mergedCell).Start.Column];

            if (cell.Value is string)
            {
                return cell.Text.Trim();
            }

            if (cell.Value is double o)
            {
                return o.ToString("G15", CultureInfo.InvariantCulture).Trim();
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

        private static string GetNonMergeCellValue(ExcelWorksheet excelWorksheet, int row, int col)
        {
            if (excelWorksheet.Cells[row, col].Value is string)
            {
                return excelWorksheet.Cells[row, col].Value?.ToString()?.Trim() ?? "";
            }

            if (excelWorksheet.Cells[row, col].Value is double value)
            {
                return value.ToString("G15", CultureInfo.InvariantCulture);
            }

            if (excelWorksheet.Cells[row, col].Value == null)
            {
                return excelWorksheet.Cells[row, col].Text.Trim();
            }

            if (excelWorksheet.Cells[row, col].Value is bool)
            {
                return excelWorksheet.Cells[row, col].Value?.ToString()?.Trim() ?? "";
            }

            return excelWorksheet.Cells[row, col].Text != null ? excelWorksheet.Cells[row, col].Text.Trim() : string.Empty;
        }

        public static void FindCellByValue(this ExcelWorksheet excelWorksheet, ref int row, ref int col, string value)
        {
            if (excelWorksheet.Dimension == null)
            {
                return;
            }

            for (int i = 1; i <= excelWorksheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= excelWorksheet.Dimension.Columns; j++)
                {
                    string lCellContent = excelWorksheet.GetMergeCellValue(i, j);
                    if (lCellContent.EqualsIgnoreCase(value))
                    {
                        row = i;
                        col = j;
                        return;
                    }
                }
            }
        }

        public static string GetMergedCellValueAndAddress(this ExcelWorksheet excelWorksheet, int rowNumber, int columnNumber, out string address)
        {
            string range = excelWorksheet.MergedCells[rowNumber, columnNumber];
            return string.IsNullOrEmpty(range)
                ? GetCellValueAndAddress(excelWorksheet, rowNumber, columnNumber, out address)
                : GetCellValueAndAddress(excelWorksheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column, out address);
        }

        public static string GetCellValueAndAddress(this ExcelWorksheet excelWorksheet, int row, int column, out string address)
        {
            address = string.Empty;
            if (excelWorksheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            if (excelWorksheet.Cells[row, column] == null)
            {
                return "";
            }

            if (excelWorksheet.Cells[row, column].Value != null)
            {
                ExcelRange range = excelWorksheet.Cells[row, column];
                address = range.Address;
                return FormatCellValue(excelWorksheet.Cells[row, column].Value);
            }

            if (excelWorksheet.Cells[row, column].Text != null)
            {
                ExcelRange range = excelWorksheet.Cells[row, column];
                address = range.Address;
                return excelWorksheet.Cells[row, column].Text;
            }

            return "";
        }

        // Preserve legacy .NET Framework double/float formatting (~15/~7 digits) under
        // invariant culture. .NET Core 3.0+ switched double.ToString() to shortest-round-
        // trippable, which elongates values like 3.3 to 3.2999999999999998 and breaks
        // IGXL validation (MDLC0008 "Illegal Value In Cell") of the generated .igxl.
        private static string FormatCellValue(object value)
        {
            return value switch
            {
                double d => d.ToString("G15", CultureInfo.InvariantCulture),
                float f => f.ToString("G7", CultureInfo.InvariantCulture),
                decimal m => m.ToString(CultureInfo.InvariantCulture),
                DateTime dt => dt.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture),
                DateTimeOffset dto => dto.ToString("M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture),
                IFormattable fmt => fmt.ToString(null, CultureInfo.InvariantCulture),
                _ => value.ToString()
            } ?? "";
        }

        public static void ExportToTxt(this ExcelWorksheet excelWorksheet, string filePath, string delimiter = "\t", string[,]? pColumnNameFixedValue = null, List<int>? skipColumnList = null, Dictionary<int, List<string>>? extraRows = null)
        {
            if (excelWorksheet?.Dimension == null)
            {
                return;
            }

            int columnCount = excelWorksheet.Dimension.End.Column;
            int rowCount = excelWorksheet.Dimension.End.Row;
            string? dir = Path.GetDirectoryName(filePath);
            if (dir != null && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (rowCount > 0)
            {
                using var sw = new StreamWriter(filePath);
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

                        ExcelRange cellRange = excelWorksheet.Cells[i, j];
                        object cell = cellRange.Value;
                        string text;
                        if (cell == null)
                        {
                            text = !string.IsNullOrEmpty(cellRange.Formula) ? cellRange.Formula : "";
                        }
                        else if (cell is double value)
                        {
                            // Force G15 / InvariantCulture: .NET 5+ changed the
                            // default "G" specifier to G17 (round-trippable),
                            // producing strings like 0.10000000000000001 that
                            // IGXL rejects with MDLC0008 "Illegal Value In Cell".
                            text = value.ToString("G15", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            text = cellRange.Text ?? "";
                        }
                        arr.Add(text);
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

        public static List<int> SplitByEmptyRow(this ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null)
            {
                return [];
            }

            int startColNumber = excelWorksheet.Dimension.Start.Column;
            int startRowNumber = excelWorksheet.Dimension.Start.Row;
            int endColNumber = excelWorksheet.Dimension.End.Column;
            int endRowNumber = excelWorksheet.Dimension.End.Row + 1;
            var splitRows = new List<int>();
            for (int i = startRowNumber; i <= endRowNumber; i++)
            {
                IEnumerable<string> range = excelWorksheet.Cells[i, startColNumber, i, endColNumber].ToList().Select(x => x.Text);
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

        public static void InsertColumn(this ExcelWorksheet excelWorksheet, string title, int row, int column)
        {
            excelWorksheet.InsertColumn(column, 1);
            excelWorksheet.Cells[row, column].Value = title;
            excelWorksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.DeepPink);
        }

        public static DataTable? ReadSheetToDataSet(this ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null)
            {
                return null;
            }

            int columnCount = excelWorksheet.Dimension.End.Column;
            int rowCount = excelWorksheet.Dimension.End.Row;
            var dt = new DataTable(excelWorksheet.Name);
            if (rowCount > 0)
            {
                object? objCellValue;
                object? cellValue;
                for (int j = 0; j < columnCount; j++)
                {
                    objCellValue = excelWorksheet.Cells[1, j + 1].Value;
                    cellValue = objCellValue == null ? "" : objCellValue.ToString();
                    dt.Columns.Add(cellValue?.ToString(), typeof(object));
                }

                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        objCellValue = excelWorksheet.Cells[i, j].Value;
                        cellValue = objCellValue ?? "";
                        dr[j - 1] = cellValue;
                    }

                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        public static void OutLineColumnByCell(this ExcelWorksheet excelWorksheet, int colNum, int rowNum = 0)
        {
            const int count = 1;
            int startRow = rowNum != 0 ? rowNum : excelWorksheet.Dimension.Start.Row;
            var list = new Dictionary<int, int>();
            for (int i = startRow; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                object now = excelWorksheet.Cells[i, colNum].Value;
                object? next = i + 1 < excelWorksheet.Dimension.End.Row ? excelWorksheet.Cells[i + 1, colNum].Value : null;
                if (now != null && next != null)
                {
                    if (now.Equals(next))
                    {
                        int start = i;
                        do
                        {
                            i++;
                            next = i + 1 <= excelWorksheet.Dimension.End.Row ? excelWorksheet.Cells[i + 1, colNum].Value : null;
                        } while (now.Equals(next));

                        list.Add(start, i);
                    }
                }
            }

            foreach (KeyValuePair<int, int> item in list)
            {
                for (int i = item.Key; i < item.Value; i++)
                {
                    excelWorksheet.Row(i).OutlineLevel = count;
                    excelWorksheet.Row(i).Collapsed = true;
                }
            }

        }

        public static void MergeColumn(this ExcelWorksheet excelWorksheet, int colNum, bool isHeaders = true)
        {
            for (int i = excelWorksheet.Dimension.Start.Row; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                object now = excelWorksheet.Cells[i, colNum].Value;
                object? next = i + 1 <= excelWorksheet.Dimension.End.Row ? excelWorksheet.Cells[i + 1, colNum].Value : null;
                if (now != null && next != null && now.Equals(next))
                {
                    int start = i;
                    do
                    {
                        i++;
                        next = i + 1 <= excelWorksheet.Dimension.End.Row ? excelWorksheet.Cells[i + 1, colNum].Value : null;
                    } while (now.Equals(next));

                    excelWorksheet.Cells[start, colNum, i, colNum].Merge = true;
                    excelWorksheet.Cells[start, colNum, i, colNum].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
            }
        }

        public static void SetFormula(this ExcelWorksheet excelWorksheet, int colNum, bool isHeaders = true)
        {
            int start = isHeaders ? 2 : 1;
            for (int i = start; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                if (excelWorksheet.Cells[i, colNum].Value != null)
                {
                    SetHyperLinkFormat(excelWorksheet.Cells[i, colNum]);
                }
            }
        }

        public static void SetHeaderStyle(this ExcelWorksheet excelWorksheet)
        {
            excelWorksheet.Workbook.AddHeaderStyle();
            excelWorksheet.Cells[excelWorksheet.Dimension.Start.Row, excelWorksheet.Dimension.Start.Column,
                    excelWorksheet.Dimension.Start.Row, excelWorksheet.Dimension.End.Column].StyleName =
                "Header";
        }

        public static void ExportSheet(this ExcelWorksheet excelWorksheet, string txtFileName, string delimiter = "\t")
        {
            int maxRow = excelWorksheet.Dimension.Rows;
            int maxCol = excelWorksheet.Dimension.Columns;
            StreamWriter sw = File.CreateText(txtFileName);
            for (int iRow = 1; iRow <= maxRow; iRow++)
            {
                var row = new List<string>();
                for (int iCol = 1; iCol <= maxCol; iCol++)
                {
                    string value = excelWorksheet.Cells[iRow, iCol].Value?.ToString() ?? "";
                    row.Add(value);
                }

                sw.WriteLine(string.Join(delimiter, row));
            }

            sw.Close();
        }

        public static List<string> ConvertToLines(this ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null)
            {
                return [];
            }
            int maxRow = excelWorksheet.Dimension.Rows;
            int maxCol = excelWorksheet.Dimension.Columns;
            var lines = new List<string>();
            for (int iRow = 1; iRow <= maxRow; iRow++)
            {
                var row = new List<string>();
                for (int iCol = 1; iCol <= maxCol; iCol++)
                {
                    string value = excelWorksheet.Cells[iRow, iCol].Value?.ToString() ?? "";
                    row.Add(value);
                }

                lines.Add(string.Join("\t", row));
            }

            return lines;
        }

        private static void CopyCellsAcrossPackage(ExcelWorksheet source, ExcelWorksheet target)
        {
            if (source.Dimension != null)
            {
                int lastRow = source.Dimension.End.Row;
                int lastCol = source.Dimension.End.Column;

                // Bulk-copy all values in one operation (much faster than cell-by-cell)
                if (lastRow > 1 || lastCol > 1)
                {
                    if (source.Cells[1, 1, lastRow, lastCol].Value is object[,] allValues)
                    {
                        target.Cells[1, 1, lastRow, lastCol].Value = allValues;
                    }
                }
                else
                {
                    target.Cells[1, 1].Value = source.Cells[1, 1].Value;
                }

            }
        }

        public static Dictionary<string, int> GetHeaderOrder(ExcelWorksheet excelWorksheet, int startRow = 1)
        {
            var headerOrder = new Dictionary<string, int>(StringExtensions.IgnoreCase);
            if (excelWorksheet.Dimension == null)
            {
                return headerOrder;
            }

            int endCol = excelWorksheet.Dimension.End.Column;
            for (int i = 1; i <= endCol; i++)
            {
                if (excelWorksheet.Cells[startRow, i].Value != null)
                {
                    string header = excelWorksheet.Cells[startRow, i].Value?.ToString()?.Trim() ?? "";
                    headerOrder.TryAdd(header, i);
                }
            }
            return headerOrder;
        }

        public static string GetMergedCellValue(ExcelWorksheet excelWorksheet, int rowNumber, int columnNumber)
        {
            string range = excelWorksheet.MergedCells[rowNumber, columnNumber];
            return string.IsNullOrEmpty(range) ?
                GetCellValue(excelWorksheet, rowNumber, columnNumber) :
                GetCellValue(excelWorksheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column);
        }

        public static string GetCellFormula(ExcelWorksheet excelWorksheet, int row, int column)
        {
            if (excelWorksheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            string value = excelWorksheet.Cells[row, column].Formula;
            if (string.IsNullOrEmpty(value))
            {
                return GetCellValue(excelWorksheet, row, column);
            }
            else
            {
                return value;
            }
        }

        public static string GetCellText(ExcelWorksheet excelWorksheet, int row, int column)
        {
            if (excelWorksheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            string value = excelWorksheet.Cells[row, column].Text;
            if (value != null)
            {
                if (value.Contains('%'))
                {
                    return value.Trim();
                }

                return GetCellValue(excelWorksheet, row, column);
            }
            return "";
        }
    }
}
