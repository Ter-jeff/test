using System.Collections.Generic;
using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

namespace CommonLib.Extension
{
    public static partial class EpplusExtensions
    {
        public static void SetHyperLinkFormat(this ExcelRange excelRange)
        {
            if (excelRange.Value == null)
            {
                return;
            }
            excelRange.Formula = excelRange.Value.ToString();
            excelRange.Style.Font.UnderLine = true;
            excelRange.Style.Font.Color.SetColor(Color.Blue);
        }

        public static void FormatHeader(this ExcelRange excelRange, bool wrap = false)
        {
            int lastCol = excelRange.Worksheet.Dimension.End.Column;
            ExcelRange range = excelRange.Worksheet.Cells[excelRange.Start.Row, excelRange.Start.Column, excelRange.Start.Row,
                lastCol];
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);
            range.Style.Font.Color.SetColor(Color.White);
            range.Style.Font.Size = 12;
            range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            range.TryAutoFitColumns();
            if (wrap)
            {
                range.Style.WrapText = true;
            }
        }

        public static int PrintExcelRow<T>(this ExcelRange excelRange, T[] tArray)
        {
            ExcelWorksheet worksheet = excelRange.Worksheet;
            int startRow = excelRange.Start.Row;
            if (tArray.Length == 0)
            {
                return startRow;
            }
            int startCol = excelRange.Start.Column;
            var arrays = new List<object[]>();
            object[] array = new object[tArray.Length];
            for (int index = 0; index < tArray.Length; index++)
            {
                T item = tArray[index];
                array[index] = item == null ? "" : item;
            }

            arrays.Add(array);
            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
            return startRow + arrays.Count;
        }

        public static void PrintExcelCol<T>(this ExcelRange excelRange, T[] tArray)
        {
            ExcelWorksheet worksheet = excelRange.Worksheet;
            int startRow = excelRange.Start.Row;
            int startCol = excelRange.Start.Column;
            for (int i = 0; i < tArray.Length; i++)
            {
                worksheet.Cells[startRow + i, startCol].Value = tArray[i] == null ? "" : tArray[i]?.ToString();
            }
        }

        public static int PrintExcelRange<T>(this ExcelRange excelRange, T[,] tArray)
        {
            ExcelWorksheet worksheet = excelRange.Worksheet;
            int startRow = excelRange.Start.Row;
            int startCol = excelRange.Start.Column;
            int rows = tArray.GetLength(0);
            int cols = tArray.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    worksheet.Cells[startRow + i, startCol + j].Value = tArray[i, j] == null ? "" : tArray[i, j]?.ToString();
                }
            }
            return startRow + rows;
        }

        public static void PrintExcelColByList<T>(this ExcelRange excelRange, List<List<T>> list)
        {
            ExcelWorksheet worksheet = excelRange.Worksheet;
            int startRow = excelRange.Start.Row;
            int startCol = excelRange.Start.Column;
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

        public static int PrintExcelRowByList<T>(this ExcelRange excelRange, List<List<T>> list)
        {
            ExcelWorksheet worksheet = excelRange.Worksheet;
            int startRow = excelRange.Start.Row;
            int startCol = excelRange.Start.Column;
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

        public static void AddHyperLink(this ExcelRange excelRange, string sheetName2, int x2, int y2)
        {
            if (x2 > 0 && y2 > 0)
            {
                string cellPosBase = ExcelCellBase.GetAddress(x2, y2);
                excelRange.Hyperlink = new ExcelHyperLink(sheetName2 + "!" + cellPosBase, excelRange.Text);
                excelRange.StyleName = "HyperLink";
                excelRange.Style.Font.UnderLine = true;
            }
        }

        public static void SetPercentFormat(this ExcelRange excelRange, string value, bool fill = true)
        {
            excelRange.Formula = value;
            excelRange.Style.Numberformat.Format = "0.00%";

            if (!fill)
            {
                return;
            }

            IExcelConditionalFormattingDataBarGroup cf = excelRange.ConditionalFormatting.AddDatabar(Color.FromArgb(153, 0, 255));
            cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf.HighValue.Value = 1;
            cf.LowValue.Value = 0;

            foreach (System.Xml.XmlElement node in cf.Node.ChildNodes)
            {
                node.SetAttribute("minLength", "0");
                node.SetAttribute("maxLength", "100");
            }
        }
    }
}
