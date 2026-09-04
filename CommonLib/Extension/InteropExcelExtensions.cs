using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

namespace CommonLib.Extension
{
    public static class InteropExcelExtensions
    {
        #region workbook

        //public static Worksheet GetSheet(this Workbook workbook, string name)
        //{
        //    foreach (Worksheet worksheet in workbook.Worksheets)
        //    {
        //        if (worksheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
        //        {
        //            return worksheet;
        //        }
        //    }

        //    return null;
        //}

        public static Worksheet AddSheet(this Workbook workbook, string name)
        {
            if (workbook.IsSheetExist(name))
            {
                workbook.Worksheets[name].Delete();
            }

            Worksheet newSheet =
                workbook.Worksheets.Add(workbook.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
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

        #endregion

        #region worksheet

        public static string GetMergeCellValue(this Worksheet worksheet, int row, int col)
        {
            var range = (Range)worksheet.Cells[row, col];
            Range mergedCell = range.MergeArea;
            if (mergedCell == null)
            {
                if (worksheet.Cells[row, col] != null)
                {
                    if (worksheet.Cells[row, col].Formula != string.Empty)
                    {
                        return "=" + worksheet.Cells[row, col].Formula;
                    }

                    if (worksheet.Cells[row, col].Value != null)
                    {
                        double? d = worksheet.Cells[row, col].Value as double?;
                        if (d != null ||
                            worksheet.Cells[row, col].Value is bool)
                        {
                            return worksheet.Cells[row, col].Value.ToString();
                        }
                    }

                    return worksheet.Cells[row, col].Text;
                }

                return string.Empty;
            }

            if (range.Formula != string.Empty)
            {
                return "=" + range.Formula;
            }

            if (range.Value != null)
            {
                if (range.Value is double ||
                    range.Value is bool)
                {
                    return range.Value.ToString();
                }
            }

            return range.Text;
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
                                string value = data.ToString();
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

        public static int GetColumnIndexByHeader(this Worksheet worksheet, string firstHeader, string headerGroupName,
            out int headerRowNumber)
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
                    dynamic header = dataArray[headerRowNumber, j] == null
                        ? ""
                        : dataArray[headerRowNumber, j].ToString().Trim();
                    if (header.Equals(headerGroupName, StringComparison.OrdinalIgnoreCase))
                    {
                        return j;
                    }
                }
            }

            return -1;
        }

        public static Range GetRange(this Worksheet worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            return worksheet.Range[worksheet.Cells[fromRow, fromCol], worksheet.Cells[toRow, toCol]];
        }

        #endregion

        #region range

        public static Range LoadFromCollection<T>(this Range range, IEnumerable<T> collection,
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

            T[] enumerable = collection is T[]? (T[])collection : collection.ToArray();
            object[,] values = new object[isPrintHeaders ? enumerable.Length + 1 : enumerable.Length, members.Count];
            int col = 0, row = 0;
            if (members.Count > 0 && isPrintHeaders)
            {
                foreach (MemberInfo t in members)
                {
                    string header = "";
                    if (t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() is
                        DescriptionAttribute)
                    {
                        var descriptionAttribute =
                            (DescriptionAttribute)t.GetCustomAttributes(typeof(DescriptionAttribute), false)
                                .FirstOrDefault();
                        if (descriptionAttribute != null)
                        {
                            header = descriptionAttribute.Description;
                        }
                    }
                    else
                    {
                        if (t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() is
                            DisplayNameAttribute)
                        {
                            var displayNameAttribute =
                                (DisplayNameAttribute)t.GetCustomAttributes(typeof(DisplayNameAttribute), false)
                                    .FirstOrDefault();
                            if (displayNameAttribute != null)
                            {
                                header = displayNameAttribute.DisplayName;
                            }
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
                    values[row, col] = item;
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

            int fromRow = range.Row;
            int fromCol = range.Column;
            Range myRange = range.Worksheet.GetRange(fromRow, fromCol, fromRow + values.GetLength(0) - 1,
                fromCol + values.GetLength(1) - 1);
            myRange.Value = values;
            return myRange;
        }

        public static Range LoadFromArrays(this Range range, IEnumerable<object[]> data, List<string> headerList = null)
        {
            return LoadFromArrays(range, data, false, headerList);
        }

        private static Range LoadFromArrays(Range range, IEnumerable<object[]> data, bool isInverse,
            List<string> headerList = null)
        {
            if (data == null)
            {
                throw new ArgumentNullException();
            }

            bool isPrintHeaders = !(headerList == null || headerList.Count == 0);

            var rowArray = new List<object[]>();
            int maxRow = 0;
            object[][] enumerable = data is object[][]? (object[][])data : data.ToArray();
            foreach (object[] item in enumerable)
            {
                rowArray.Add(item);
                if (maxRow < item.Length)
                {
                    maxRow = item.Length;
                }
            }

            maxRow = isPrintHeaders ? maxRow + 1 : maxRow;
            int minCol = headerList == null || headerList.Count == 0
                ? rowArray.Count
                : Math.Min(rowArray.Count, headerList.Count);

            if (rowArray.Count == 0)
            {
                return null;
            }

            if (isInverse)
            {
                object[,] values = new object[maxRow, minCol];
                int col = 0;
                int row = 0;
                if (maxRow > 0 && isPrintHeaders)
                {
                    for (int i = 0; i < minCol; i++)
                    {
                        values[row, col++] = headerList[i];
                    }
                }

                col = 0;
                foreach (object[] item in enumerable)
                {
                    row = 0;
                    foreach (object t in item)
                    {
                        values[row++, col] = t;
                    }

                    col++;
                }

                int fromRow = range.Row;
                int fromCol = range.Column;
                Range myRange = range.Worksheet.GetRange(fromRow, fromCol, fromRow + values.GetLength(0) - 1,
                    fromCol + values.GetLength(1) - 1);
                myRange.Value = values;
                return myRange;
            }
            else
            {
                object[,] values = new object[minCol, maxRow];
                int col = 0;
                int row = 0;
                if (maxRow > 0 && isPrintHeaders)
                {
                    for (int i = 0; i < minCol; i++)
                    {
                        values[col++, row] = headerList[i];
                    }
                }

                col = 0;
                foreach (object[] item in enumerable)
                {
                    row = 0;
                    foreach (object t in item)
                    {
                        values[col, row++] = t;
                    }

                    col++;
                }

                int fromRow = range.Row;
                int fromCol = range.Column;
                Range myRange = range.Worksheet.GetRange(fromRow, fromCol, fromRow + values.GetLength(0) - 1,
                    fromCol + values.GetLength(1) - 1);
                myRange.Value = values;
                return myRange;
            }
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

        #endregion
    }
}
