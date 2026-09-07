using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public abstract class IgxlSheetReader<TSheet> : IIgxlSheetReader<TSheet> where TSheet : IIgxlSheet
    {
        public virtual EnumSheetType SupportedType => EnumSheetType.DTUnknown;

        public abstract TSheet ReadSheet(ExcelWorksheet excelWorksheet);

        public abstract TSheet ReadSheet(Stream stream, string sheetName);

        public object Read(object source, string sheetName)
        {
            return source switch
            {
                ExcelWorksheet excelWorksheet => ReadSheet(excelWorksheet),
                Stream stream => ReadSheet(stream, sheetName),
                _ => throw new ArgumentException("Unsupported source type", nameof(source)),
            };
        }

        public TSheet GetSheet(string fileName)
        {
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            FileStream stream = File.OpenRead(fileName);
            return ReadSheet(stream, sheetName);
        }

        protected string GetCellValue(string[] line, int column)
        {
            if (column < line.Length)
            {
                return line.ElementAt(column);
            }

            return "";
        }

        protected string GetCellText(ExcelWorksheet excelWorksheet, int row, int column)
        {
            ExcelRange cell = excelWorksheet.Cells[row, column];
            if (cell == null)
            {
                return "";
            }

            if (!string.IsNullOrEmpty(cell.Formula))
            {
                return "=" + cell.Formula;
            }

            if (cell.Value is double doubleValue)
            {
                return doubleValue.ToString("G15", CultureInfo.InvariantCulture);
            }

            if (cell.Value is bool)
            {
                return cell.Value.ToString() ?? "";
            }

            return cell.Text ?? "";
        }

        protected string GetMergeCellValue(ExcelWorksheet excelWorksheet, int row, int column)
        {
            string range = excelWorksheet.MergedCells[row, column];
            if (string.IsNullOrEmpty(range))
            {
                return GetCellText(excelWorksheet, row, column);
            }

            var address = new ExcelAddress(range);
            return GetCellText(excelWorksheet, address.Start.Row, address.Start.Column);
        }

        protected T GetRowData<T>(ExcelWorksheet excelWorksheet, int row, int startColumnIndex, List<string> headList, Action<T, string, string> mapProperty) where T : new()
        {
            var rowObj = new T();
            (rowObj as dynamic).SheetName = excelWorksheet.Name;
            (rowObj as dynamic).RowNum = row;

            for (int i = startColumnIndex; i < startColumnIndex + headList.Count; i++)
            {
                string header = (i == 1) ? "ColumnA" : headList[i - startColumnIndex];
                string content = GetCellText(excelWorksheet, row, i);

                mapProperty(rowObj, CellDiff.FormatStringForCompare(header), content);
            }
            return rowObj;
        }

        protected T CreateRow<T>(string line, string sheetName, int row, Action<T, string[]> mapProperties) where T : IgxlRow, new()
        {
            string[] arr = line.Split('\t');
            T rowObj = new T
            {
                RowNum = row,
                SheetName = sheetName,
                ColumnA = GetCellValue(arr, 0)
            };

            mapProperties(rowObj, arr);
            return rowObj;
        }
    }
}
