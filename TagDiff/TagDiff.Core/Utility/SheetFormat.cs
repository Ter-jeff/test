using System;
using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

namespace TagDiff.Core.Utility
{
    public static class SheetFormat
    {
        public static void SetSheetFormat(ExcelWorksheet excelWorksheet)
        {
            ArgumentNullException.ThrowIfNull(excelWorksheet);

            if (excelWorksheet.Dimension == null)
            {
                return;
            }

            static void ApplyFill(IExcelConditionalFormattingRule rule, Color color)
            {
                rule.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rule.Style.Fill.BackgroundColor.Color = color;
            }

            void AddEqual(string range, string formulaText, Color color)
            {
                IExcelConditionalFormattingEqual rule = excelWorksheet.ConditionalFormatting.AddEqual(excelWorksheet.Cells[range]);
                rule.Formula = $"\"{formulaText}\"";
                ApplyFill(rule, color);
            }

            void AddContainsText(string range, string text, Color color)
            {
                IExcelConditionalFormattingContainsText rule = excelWorksheet.ConditionalFormatting.AddContainsText(excelWorksheet.Cells[range]);
                rule.Text = text;
                ApplyFill(rule, color);
            }

            AddEqual("A:A", "P", Color.LightGreen);
            AddEqual("A:A", "F", Color.Red);
            AddEqual("A:A", "M", Color.Yellow);

            AddContainsText(excelWorksheet.Dimension.Address, "=>", Color.Yellow);

            AddContainsText("A:A", "Missing", Color.Red);
            AddContainsText("1:1", "Missing", Color.Red);
        }

        public static string GetHyperlinkFormula(string sheetName, int row, int column, string friendlyName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (row <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(row), "Row must be >= 1.");
            }

            string cellRef = column <= 0 ? $"{row}:{row}" : $"R{row}C{column}";
            return $"=HYPERLINK(\"#'{sheetName}'!{cellRef}\",\"{friendlyName}\")";
        }
    }
}
