using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Excel;

using OfficeOpenXml;

using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;
using Style = Microsoft.Office.Interop.Excel.Style;

namespace CommonLib.Utility
{
    public static class InteropExcel
    {
        public static void Txt2Bas(string filePath, string moduleName, List<string> commentList = null)
        {
            if (File.Exists(filePath))
            {
                string targetFileName = Path.GetDirectoryName(filePath) + @"\" + moduleName + ".bas";
                if (File.Exists(targetFileName))
                {
                    string targetText = File.ReadAllText(targetFileName);
                    string text = File.ReadAllText(filePath);
                    using (var sw = new StreamWriter(targetFileName, false))
                    {
                        sw.WriteLine(targetText);
                        sw.WriteLine(text);
                    }
                    File.Delete(filePath);
                }
                else
                {
                    string oldText = File.ReadAllText(filePath);
                    using (var sw = new StreamWriter(filePath, false))
                    {
                        sw.WriteLine("Attribute VB_Name = \"" + moduleName + "\"");
                        if (commentList != null)
                        {
                            foreach (string comment in commentList)
                            {
                                sw.WriteLine("'" + comment);
                            }
                        }

                        sw.WriteLine(oldText);
                    }
                    File.Move(filePath, targetFileName);
                }
            }
        }

        public static int PrintExcelRange<T>(Worksheet workSheet, Range cell, T[,] data)
        {
            int rowNum = cell.Row;
            int colNum = cell.Column;
            workSheet.Range[cell, workSheet.Cells[rowNum + data.GetLength(0) - 1, colNum + data.GetLength(1) - 1]]
                .Value = data;
            workSheet.Range[cell, workSheet.Cells[rowNum + data.GetLength(0) - 1, colNum + data.GetLength(1) - 1]].Value
                = workSheet.Range[cell, workSheet.Cells[rowNum + data.GetLength(0) - 1, colNum + data.GetLength(1) - 1]]
                    .Value;
            return rowNum + data.GetLength(0);
        }

        public static void PrintExcelColRange<T>(Worksheet workSheet, Range cell, List<List<T>> list)
        {
            if (list == null)
            {
                throw new Exception("It can not contain null array.");
            }

            foreach (List<T> a in list)
            {
                if (a == null)
                {
                    throw new Exception("It can not contain null array.");
                }

                if (a.Count != list[0].Count)
                {
                    throw new Exception("Input array should have the same length");
                }
            }

            int rowCnt = list[0].Count;
            int colCnt = list.Count;
            var result = new T[rowCnt, colCnt];

            for (int i = 0; i < rowCnt; i++)
            {
                for (int j = 0; j < colCnt; j++)
                {
                    result[i, j] = list[j][i];
                }
            }

            workSheet.Range[cell, workSheet.Cells[cell.Row + rowCnt - 1, cell.Column + colCnt - 1]].Value = result;
            workSheet.Range[cell, workSheet.Cells[cell.Row + rowCnt - 1, cell.Column + colCnt - 1]].Value = workSheet
                .Range[cell, workSheet.Cells[cell.Row + rowCnt - 1, cell.Column + colCnt - 1]].Value;
        }

        public static int PrintExcelRowRange<T>(Worksheet workSheet, Range cell, List<List<T>> list)
        {
            if (list == null)
            {
                throw new Exception("It can not contain null array.");
            }

            foreach (List<T> a in list)
            {
                if (a == null)
                {
                    throw new Exception("It can not contain null array.");
                }

                if (a.Count != list[0].Count)
                {
                    throw new Exception("Input array should have the same length");
                }
            }

            int rowCnt = list.Count;
            int colCnt = list[0].Count;
            var result = new T[rowCnt, colCnt];

            for (int i = 0; i < rowCnt; i++)
            {
                for (int j = 0; j < colCnt; j++)
                {
                    result[i, j] = list[i][j];
                }
            }

            workSheet.Range[cell, workSheet.Cells[cell.Row + rowCnt - 1, cell.Column + colCnt - 1]].Value = result;
            workSheet.Range[cell, workSheet.Cells[cell.Row + rowCnt - 1, cell.Column + colCnt - 1]].Value = workSheet
                .Range[cell, workSheet.Cells[cell.Row + rowCnt - 1, cell.Column + colCnt - 1]].Value;
            return cell.Row + rowCnt;
        }

        public static string GetHyperlink(string name, int row, int column, string friendlyName)
        {
            if (column == 0)
            {
                return "=HYPERLINK(\"#\'" + name + "\'!" + row + ":" + row + "\",\"" + friendlyName + "\")";
            }

            return "=HYPERLINK(\"#\'" + name + "\'!" + GetAddress(row, column) + "\",\"" + friendlyName + "\")";
        }

        public static string GetAddress(int row, int column, bool absolute = false)
        {
            if (row == 0 || column == 0)
            {
                return "#REF!";
            }

            if (absolute)
            {
                return "$" + GetColumnLetter(column) + "$" + row;
            }

            return GetColumnLetter(column) + row;
        }

        private static string GetColumnLetter(int colNum)
        {
            if (colNum < 1)
            {
                return "#REF!";
            }

            string sCol = "";
            do
            {
                sCol = (char)('A' + (colNum - 1) % 26) + sCol;
                colNum = (colNum - (colNum - 1) % 26) / 26;
            } while (colNum > 0);

            return sCol;
        }
    }

    public class Cell
    {
        public string SheetName;
        public int RowNum;
        public int ColNum;
        public string Formula;
    }

    public static class InteropExcelExtensions
    {
        private static readonly Regex _regex = new Regex("_xlfn.SINGLE", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex("_xlfn.CONCAT", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        #region workbook
        public static void AddMacro(this Workbook workbook, List<string> lines, string name)
        {
            int index = -1;
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                index++;
                if (worksheet.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                {
                    index++;
                    break;
                }
            }

            if (index != -1)
            {
                Microsoft.Vbe.Interop.VBComponent vBComponent = workbook.VBProject.VBComponents.Item(index + 1);
                Microsoft.Vbe.Interop.CodeModule codeModule = vBComponent.CodeModule;
                codeModule.AddFromString(string.Join(Environment.NewLine, lines));
                Marshal.ReleaseComObject(codeModule);
            }
        }

        public static Worksheet GetSheet(this Workbook workbook, string name)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        public static void AddTxt(this Workbook workbook, string fileName)
        {
            string[] lines = File.ReadAllLines(fileName);
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            Worksheet newSheet;
            if (workbook.IsSheetExist(sheetName))
            {
                newSheet = workbook.Worksheets[sheetName];
                workbook.Parent.DisplayAlerts = false;
            }
            else
            {
                newSheet = workbook.Worksheets.Add();
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
            newSheet.GetRange(1, 1, rowCount, columnCount).Value2 = values;

        }

        public static Worksheet AddSheet(this Workbook workbook, string name)
        {
            if (workbook.IsSheetExist(name))
            {
                workbook.Worksheets[name].Delete();
            }

            Worksheet newSheet = workbook.Worksheets.Add();
            newSheet.Name = name;
            return newSheet;
        }

        public static bool IsSheetExist(this Workbook workbook, string name)
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        public static void SaveAsXlsm(this Workbook workbook)
        {
            string oldFilePath = workbook.FullName;
            string extension = Path.GetExtension(oldFilePath);
            if (extension == null)
            {
                return;
            }

            if (extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string newFilePath = Path.ChangeExtension(oldFilePath, "xlsm");
            workbook.SaveAs(newFilePath, XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
                null, null, false, false, XlSaveAsAccessMode.xlNoChange,
                false, false);
            if (oldFilePath == null)
            {
                throw new ArgumentNullException();
            }

            File.Delete(oldFilePath);
        }

        public static void AddHeaderStyle(this Workbook workbook)
        {
            foreach (Style item in workbook.Styles)
            {
                if (item.Name == "Header")
                {
                    return;
                }
            }

            Style style = workbook.Styles.Add("Header");
            style.Interior.Color = Color.YellowGreen;
            style.IncludeNumber = true;
            style.IncludeFont = true;
            style.IncludeAlignment = true;
            style.IncludeBorder = true;
            style.IncludePatterns = true;
            style.IncludeProtection = true;
        }

        public static void AddErrorStyle(this Workbook workbook)
        {
            foreach (Style item in workbook.Styles)
            {
                if (item.Name == "Error")
                {
                    return;
                }
            }

            Style style = workbook.Styles.Add("Error");
            style.Interior.Color = Color.Red;
            style.Font.Color = Color.White;
            style.IncludeNumber = true;
            style.IncludeFont = true;
            style.IncludeAlignment = true;
            style.IncludeBorder = true;
            style.IncludePatterns = true;
            style.IncludeProtection = true;
        }

        #endregion

        #region worksheet
        public static List<Cell> FixFormula(this Worksheet worksheet)
        {
            var errorList = new List<Cell>();
            worksheet.AutoFilterMode = false;
            int startColNumber = worksheet.UsedRange.Column;
            int startRowNumber = worksheet.UsedRange.Row;
            int stopColNumber = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            int stopRowNumber = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            dynamic formulas = worksheet.Range[worksheet.Cells[startRowNumber, startColNumber], worksheet.Cells[stopRowNumber, stopColNumber]].Formula;
            for (int i = startRowNumber; i <= stopRowNumber; i++)
            {
                for (int j = startColNumber; j <= stopColNumber; j++)
                {
                    dynamic formula = formulas[i, j];
                    if (!string.IsNullOrEmpty(formula) && (formula.Contains("_xlfn.SINGLE")))
                    {
                        errorList.Add(new Cell { SheetName = worksheet.Name, RowNum = i, ColNum = j, Formula = formula });
                        formulas[i, j] = _regex.Replace(formula, "");
                    }
                    else if (!string.IsNullOrEmpty(formula) && (formula.Contains("_xlfn.CONCAT")))
                    {
                        errorList.Add(new Cell { SheetName = worksheet.Name, RowNum = i, ColNum = j, Formula = formula });
                        formulas[i, j] = _regex2.Replace(formula, "CONCATENATE");
                    }
                }
            }
            worksheet.Range[worksheet.Cells[startRowNumber, startColNumber], worksheet.Cells[stopRowNumber, stopColNumber]].Formula = formulas;
            worksheet.Calculate();
            return errorList;
        }

        public static int GetColumnIndexByHeader(this Worksheet worksheet, string firstHeader, string headerGroupName, out int headerRowNumber)
        {
            int startColNumber = worksheet.UsedRange.Column;
            int startRowNumber = worksheet.UsedRange.Row;
            int stopColNumber = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            int stopRowNumber = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            dynamic dataArray = worksheet.GetRange(startRowNumber, startColNumber, stopRowNumber, stopColNumber).Value2;
            headerRowNumber = -1;

            int rowNum = stopRowNumber > 10 ? 10 : stopRowNumber;
            int colNum = stopColNumber > 10 ? 10 : stopColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    dynamic header = dataArray[i, j] == null ? "" : dataArray[i, j].ToString().Trim();
                    if (header.Equals(firstHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        headerRowNumber = i;
                        break;
                    }
                }
            }

            if (headerRowNumber != -1)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    dynamic header = dataArray[headerRowNumber, j] == null ? "" : dataArray[headerRowNumber, j].ToString().Trim();
                    if (header.Equals(headerGroupName, StringComparison.OrdinalIgnoreCase))
                    {
                        return j;
                    }
                }
            }
            return -1;
        }

        public static int GetColumnIndexByHeader(this Worksheet worksheet, string headerGroupName)
        {
            int startColNumber = worksheet.UsedRange.Column;
            int startRowNumber = worksheet.UsedRange.Row;
            int stopColNumber = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            int stopRowNumber = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            dynamic dataArray = worksheet.GetRange(startRowNumber, startColNumber, stopRowNumber, stopColNumber).Value2;

            int rowNum = stopRowNumber > 10 ? 10 : stopRowNumber;
            int colNum = stopColNumber > 10 ? 10 : stopColNumber;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    dynamic header = dataArray[i, j] == null ? "" : dataArray[i, j].ToString().Trim();
                    if (header.Equals(headerGroupName, StringComparison.OrdinalIgnoreCase))
                    {
                        return j;
                    }
                }
            }

            return -1;
        }

        public static void SetHeaderStyle(this Worksheet worksheet)
        {
            var workbook = (Workbook)worksheet.Parent;
            workbook.AddHeaderStyle();
            worksheet.GetRange(worksheet.UsedRange.Row, worksheet.UsedRange.Column,
                    worksheet.UsedRange.Row, worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1).Style =
                "Header";
        }

        public static void SetOutline<T>(this Worksheet worksheet, List<T> list, Expression<Func<T, string>> lambdaExpression) where T : new()
        {
            const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance;
            int colNum = typeof(T).GetProperties(memberFlags).Count(p => p.DeclaringType == typeof(T));
            PropertyInfo first = typeof(T).GetProperties(memberFlags).First(p => p.DeclaringType == typeof(T));
            PropertyInfo last = typeof(T).GetProperties(memberFlags).Last(p => p.DeclaringType == typeof(T));
            Func<T, string> getter = lambdaExpression.Compile();

            worksheet.UsedRange.ClearOutline();
            var orderList = new List<T>();
            int rowNum = 1;
            foreach (IGrouping<string, T> group in list.GroupBy(getter).ToList())
            {
                orderList.AddRange(group);

                #region add summary row

                var summaryRow = new T();
                int startRow = rowNum + 1;
                int endRow = rowNum + group.Count();
                rowNum += group.Count() + 1;
                string formula = "=COUNTA(" + InteropExcel.GetAddress(startRow, colNum) + ":" +
                                 InteropExcel.GetAddress(endRow, colNum) + ")";
                first.SetValue(summaryRow, group.Select(getter).First(), null);
                last.SetValue(summaryRow, formula, null);
                orderList.Add(summaryRow);

                #endregion
            }

            var range = (Range)worksheet.Cells[1, 1];
            range.LoadFromCollection(orderList);
            worksheet.UsedRange.AutoOutline();
            worksheet.Outline.ShowLevels(2);
            worksheet.UsedRange.AutoFilter(1);
            worksheet.Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        public static Range GetRange(this Worksheet worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            return worksheet.Range[worksheet.Cells[fromRow, fromCol], worksheet.Cells[toRow, toCol]];
        }

        public static void MergeColumn(this Worksheet worksheet, int colNum, bool printHeaders = true)
        {
            dynamic array = worksheet.Range[worksheet.Cells[1, colNum],
                worksheet.Cells[worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1, colNum]].Value2;
            if (array.count == 1)
            {
                return;
            }

            int lastIndex = 0;
            int start = printHeaders ? 2 : 1;
            for (int i = start; i <= array.Length; i++)
            {
                for (int j = i + 1; j <= array.Length; j++)
                {
                    if (array[i, colNum] != array[j, colNum])
                    {
                        lastIndex = j - 1;
                        break;
                    }

                    if (j == array.Length)
                    {
                        lastIndex = j;
                        break;
                    }
                }

                if (lastIndex - i > 0)
                {
                    worksheet.Range[worksheet.Cells[i, colNum], worksheet.Cells[lastIndex, colNum]].Merge();
                }
            }
        }

        public static void ExportTxt(this Worksheet worksheet, string path)
        {
            string file = Path.Combine(path, worksheet.Name + ".txt");
            if (File.Exists(file))
            {
                File.Delete(file);
            }

            using (StreamWriter sw = File.CreateText(file))
            {
                if (worksheet.UsedRange.Count == 1)
                {
                    int rowCount = worksheet.UsedRange.Rows.Count;
                    int colCount = worksheet.UsedRange.Columns.Count;
                    object data = worksheet.UsedRange.Formula;
                    if (data != null)
                    {
                        for (int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {
                                string value = data == null ? "" : data.ToString();
                                sw.Write(value);
                                sw.Write("\t");
                            }
                            sw.Write(Environment.NewLine);
                        }
                    }
                }
                else
                {
                    int rowCount = worksheet.UsedRange.Rows.Count;
                    int colCount = worksheet.UsedRange.Columns.Count;
                    object[,] data = worksheet.UsedRange.Formula;
                    if (data != null)
                    {
                        for (int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {
                                string value = data[i, j] == null ? "" : data[i, j].ToString();
                                sw.Write(value);
                                sw.Write("\t");
                            }
                            sw.Write(Environment.NewLine);
                        }
                    }
                }
            }
        }

        public static ExcelWorksheet ConvertToExcelSheet(this Worksheet worksheet)
        {
            object[,] data = worksheet.UsedRange.Value2;
            int columnCount = data.GetLength(1);
            int rowCount = data.GetLength(0);
            var table = new DataTable();
            if (data != null)
            {
                for (int j = 1; j <= columnCount; j++)
                {
                    table.Columns.Add(data[1, j] == null ? "" : data[1, j].ToString());
                }
                for (int i = 1; i <= rowCount; i++)
                {
                    System.Data.DataRow row = table.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        row[j - 1] = data[i, j] == null ? "" : data[i, j].ToString();
                    }
                    table.Rows.Add(row);
                }
            }
            var excelPackage = new ExcelPackage();
            string sheetName = worksheet.Name;
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            sheet.Cells[1, 1].LoadFromDataTable(table, false);
            return sheet;
        }
        #endregion

        #region range

        public static Range LoadFromCollectionInverse<T>(this Range range, IEnumerable<T> collection,
            bool printHeaders = true)
        {
            return LoadFromCollection(range, collection, true, printHeaders);
        }

        public static void LoadFromCollection<T>(this Range range, IEnumerable<T> collection,
            bool printHeaders = true)
        {
            LoadFromCollection(range, collection, false, printHeaders);
        }

        private static Range LoadFromCollection<T>(Range range, IEnumerable<T> collection, bool isInverse,
            bool isPrintHeaders = true)
        {
            const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance;
            Type type = typeof(T);
            var members = new List<MemberInfo>();
            members.AddRange(type.GetProperties(memberFlags).Where(p => p.DeclaringType == typeof(T)).ToList());

            if (members.Count == 0)
            {
                throw new ArgumentException("Parameter Members must have at least one Property. Length is zero");
            }

            T[] enumerable = collection as T[] ?? collection.ToArray();
            object[,] values = new object[isPrintHeaders ? enumerable.Length + 1 : enumerable.Length, members.Count];
            int col = 0, row = 0;
            if (members.Count > 0 && isPrintHeaders)
            {
                foreach (MemberInfo t in members)
                {
                    string header;
                    if (t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() is
                        DescriptionAttribute)
                    {
                        var descriptionAttribute = (DescriptionAttribute)t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault();
                        header = descriptionAttribute.Description;
                    }
                    else
                    {
                        if (t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() is
                             DisplayNameAttribute)
                        {
                            var displayNameAttribute = (DisplayNameAttribute)t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault();
                            header = displayNameAttribute.DisplayName;
                        }
                        else
                        {
                            header = t.Name.Replace('_', ' ');
                        }
                    }

                    values[row, col++] = header;
                }

                row++;
            }

            if (!enumerable.Any() && (members.Count == 0 || !isPrintHeaders))
            {
                return null;
            }

            foreach (T item in enumerable)
            {
                col = 0;
                if (item is string || item is decimal || item is DateTime || item.GetType().IsPrimitive)
                {
                    values[row, col++] = item;
                }
                else
                {
                    int index = 0;
                    foreach (MemberInfo t in members)
                    {
                        if (t is PropertyInfo)
                        {
                            var propertyInfo = (PropertyInfo)t;
                            ParameterInfo[] parameters = propertyInfo.GetIndexParameters();
                            if (!parameters.Any())
                            {
                                values[row, col++] = propertyInfo.GetValue(item, null);
                            }
                            else
                            {
                                values[row, col++] = propertyInfo.GetValue(item, new object[] { index++ });
                            }
                        }
                        else if (t is FieldInfo)
                        {
                            var fieldInfo = (FieldInfo)t;
                            values[row, col++] = fieldInfo.GetValue(item);
                        }
                        else if (t is MethodInfo)
                        {
                            var methodInfo = (MethodInfo)t;
                            values[row, col++] = methodInfo.Invoke(item, null);
                        }
                    }
                }

                row++;
            }

            Worksheet workSheet = range.Worksheet;
            int fromRow = range.Row;
            int fromCol = range.Column;
            return Inverse(workSheet.GetRange(fromRow, fromCol, fromRow + row - 1, fromCol + col - 1)
                , values, isInverse);
        }

        public static Range LoadFromArrayInverse(this Range range, IEnumerable<object[]> data,
            List<string> headerList = null)
        {
            return LoadFromArray(range, data, true, headerList);
        }

        public static Range LoadFromArray(this Range range, IEnumerable<object[]> data,
            List<string> headerList = null)
        {
            return LoadFromArray(range, data, false, headerList);
        }

        private static Range LoadFromArray(Range range, IEnumerable<object[]> data, bool isInverse, List<string> headerList = null)
        {
            if (data == null)
            {
                throw new ArgumentNullException();
            }

            bool isPrintHeaders = !(headerList == null || headerList.Count == 0);

            var rowArray = new List<object[]>();
            int maxCol = 0;
            object[][] enumerable = data as object[][] ?? data.ToArray();
            foreach (object[] item in enumerable)
            {
                rowArray.Add(item);
                if (maxCol < item.Length)
                {
                    maxCol = item.Length;
                }
            }

            maxCol = isPrintHeaders ? maxCol + 1 : maxCol;
            int maxRow = headerList == null || headerList.Count == 0
                ? rowArray.Count
                : Math.Min(rowArray.Count, headerList.Count);

            if (rowArray.Count == 0)
            {
                return null;
            }

            object[,] values = new object[maxRow, maxCol];
            int col = 0, row = 0;
            if (maxCol > 0 && isPrintHeaders)
            {
                for (int i = 0; i < maxRow; i++)
                {
                    values[row, col++] = headerList[i];
                }
            }

            row = 0;
            foreach (object[] item in enumerable)
            {
                col = 0;
                foreach (object t in item)
                {
                    values[row, col++] = t;
                }
                row++;
            }

            Worksheet workSheet = range.Worksheet;
            int fromRow = range.Row;
            int fromCol = range.Column;

            return Inverse(workSheet.GetRange(fromRow, fromCol, fromRow + maxRow - 1, fromCol + maxCol - 1), values, isInverse);
        }

        private static Range Inverse(Range range, object[,] values, bool isInverse)
        {
            if (isInverse)
            {
                var app = new Application();
                Range inverseRange = range.Worksheet.GetRange(range.Row, range.Column,
                    range.Row + range.Columns.Count - 1, range.Column + range.Rows.Count - 1);
                inverseRange.Value2 = app.WorksheetFunction.Transpose(values);
                return inverseRange;
            }

            range.Value2 = values;
            return range;
        }

        public static void AddErrorStyle(this Range range)
        {
            range.Style = "Error";
        }

        public static void AddHeaderStyle(this Range range)
        {
            range.Style = "Header";
        }

        public static string Address(this Range range)
        {
            int startRow = range.Row;
            int stopRow = range.Row + range.Rows.Count - 1;
            int startCol = range.Column;
            int stopCol = range.Column + range.Columns.Count - 1;

            return "$" + GetColumnLetter(startCol) + "$" + startRow + ":" + "$" + GetColumnLetter(stopCol) + "$" + stopRow;

        }

        public static string GetColumnLetter(this Range range)
        {
            int startCol = range.Column;
            return GetColumnLetter(startCol);
        }

        private static string GetColumnLetter(int colNum)
        {
            if (colNum < 1)
            {
                return "#REF!";
            }

            string sCol = "";
            do
            {
                sCol = (char)('A' + (colNum - 1) % 26) + sCol;
                colNum = (colNum - (colNum - 1) % 26) / 26;
            } while (colNum > 0);

            return sCol;
        }

        public static void SetHyperLinkFormat(this Range range)
        {
            range.Formula = range.Value.ToString();
            range.Style.Font.UnderLine = true;
            range.Style.Font.Color.SetColor(Color.Blue);
        }
        #endregion
    }
}
