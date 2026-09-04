using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Font = System.Drawing.Font;

namespace CommonLib.Utility
{
    public static class EpplusOperation
    {
        private static readonly Regex _regex = new Regex(@"\=.*\*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool IsOpened(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            try
            {
                Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        public static Dictionary<string, int> GetHeaderOrder(ExcelWorksheet sheet, int startRow = 1)
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

        public static string FloorMinValue(string value, string pwrSupplyRes)
        {
            if (_regex.IsMatch(value))
            {
                string resultValue = value.Replace("=", "");
                double outD;
                double.TryParse(pwrSupplyRes, out outD);
                if (outD > 0.001)
                {
                    resultValue = "=FLOOR(" + resultValue + "," + outD + ")";
                }
                else
                {
                    resultValue = "=" + resultValue;
                }

                return resultValue;
            }
            return value;
        }

        public static string CeilingMaxValue(string value, string pwrSupplyRes)
        {
            if (_regex.IsMatch(value))
            {
                string resultValue = value.Replace("=", "");
                double outD;
                double.TryParse(pwrSupplyRes, out outD);
                if (outD > 0.001)
                {
                    resultValue = "=CEILING(" + resultValue + "," + outD + ")";
                }
                else
                {
                    resultValue = "=" + resultValue;
                }

                return resultValue;
            }
            return value;
        }

        public static string GetMergedCellValue(ExcelWorksheet sheet, int rowNumber, int columnNumber)
        {
            string range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null ?
                GetCellValue(sheet, rowNumber, columnNumber) :
                GetCellValue(sheet, (new ExcelAddress(range).Start.Row), (new ExcelAddress(range).Start.Column));
        }

        public static string GetMergedCellValueAutoRate(ExcelWorksheet sheet, int rowNumber, int columnNumber)
        {
            string range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null ?
                GetCellValueAutoRate(sheet, rowNumber, columnNumber) :
                GetCellValueAutoRate(sheet, (new ExcelAddress(range).Start.Row), (new ExcelAddress(range).Start.Column));
        }

        public static string GetCellFormula(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            string value = sheet.Cells[row, column].Formula;
            if (value.Equals(""))
            {
                return GetCellValue(sheet, row, column);
            }
            else
            {
                return value;
            }
        }

        public static string GetCellText(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            string value = sheet.Cells[row, column].Text;
            if (value != null)
            {
                if (value.Contains("%"))
                {
                    return value.Trim();
                }

                return GetCellValue(sheet, row, column);
            }
            return "";
        }

        public static string GetCellValue(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            //if (!string.IsNullOrEmpty(sheet.Cells[row, column].Formula))
            //{
            //    return sheet.Cells[row, column].Formula;
            //}
            if (sheet.Cells[row, column].Value != null)
            {
                return sheet.Cells[row, column].Value.ToString().Trim();
            }
            if (sheet.Cells[row, column].Text != null)
            {
                return sheet.Cells[row, column].Text.Trim();
            }
            return "";
        }

        public static string GetCellValueAutoRate(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null)
            {
                return "";
            }

            if (row <= 0 || column <= 0)
            {
                return "";
            }

            if (!string.IsNullOrEmpty(sheet.Cells[row, column].Formula))
            {
                return sheet.Cells[row, column].Formula;
            }
            if (sheet.Cells[row, column].Value != null)
            {
                return sheet.Cells[row, column].Value.ToString().Trim();
            }
            if (sheet.Cells[row, column].Text != null)
            {
                return sheet.Cells[row, column].Text.Trim();
            }
            return "";
        }

        public static string GetCellValueOld(ExcelWorksheet wSheet, int row, int column)
        {
            if (row <= 0 || column <= 0)
            {
                return "";
            }

            if (wSheet.Cells[row, column].Value != null)
            {
                return wSheet.Cells[row, column].Value.ToString().Trim();
            }

            return "";
        }

        private static string ReplaceDoubleBlank(string text)
        {
            string result = text;
            do
            {
                result = result.Replace("  ", " ");
            } while (result.IndexOf("  ", StringComparison.Ordinal) >= 0);
            return result;
        }

        private static string FormatCell(string text)
        {
            string result = text.Trim();

            result = ReplaceDoubleBlank(result);

            result = result.Replace(" ", "_");

            result = result.ToUpper();

            return result;
        }

        public static bool IsLiked(string pStrInput, string pStrPatten)
        {
            if (pStrPatten.IndexOf(@".*", StringComparison.Ordinal) >= 0 ||
                pStrPatten.IndexOf(@".+", StringComparison.Ordinal) >= 0 ||
                pStrPatten.IndexOf(@"|", StringComparison.Ordinal) >= 0)
            {
                bool value = Regex.IsMatch(FormatCell(pStrInput), FormatCell(pStrPatten));
                return value;
            }
            else
            {
                bool value = FormatCell(pStrInput) == FormatCell(pStrPatten);
                return value;
            }

        }

        public static void MergeCellsFillBackGroundColorList(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn, string value)
        {
            using (ExcelRange range = worksheet.Cells[startRow, startColumn, endRow, endColumn])
            {
                if (!range.Merge)
                {
                    range.Merge = true;
                }

                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(85, 107, 47));
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Font.Color.SetColor(Color.White);
                range.Style.Font.SetFromFont(new Font("Arial", 12));
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                range.AutoFitColumns();
                range.Value = value;
            }
        }

        public static void CreateDefaultNamedStyleInWorkBook(ref ExcelPackage ep, string epType) //預先對指定的Excel全Sheet建立Style
        {
            //Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); //這一行沒辦法加到NamedStyle裡面...?
            OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyleTitleRow = ep.Workbook.Styles.CreateNamedStyle("Title Row");
            namedStyleTitleRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            namedStyleTitleRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            namedStyleTitleRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyleTitleRow.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(107, 142, 35));
            namedStyleTitleRow.Style.Font.Color.SetColor(Color.White);
            //namedStyleTitleRow.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            if (epType == "TestFlowProfile")
            {
                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyleTitleRowTf = ep.Workbook.Styles.CreateNamedStyle("TF Title Row");
                namedStyleTitleRowTf.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleTitleRowTf.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleTitleRowTf.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleTitleRowTf.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 176, 240));
                namedStyleTitleRowTf.Style.Font.Color.SetColor(Color.White);
                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyleSubTitleRow = ep.Workbook.Styles.CreateNamedStyle("Sub Title Row");
                namedStyleSubTitleRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleSubTitleRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleSubTitleRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleSubTitleRow.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                namedStyleSubTitleRow.Style.Font.Color.SetColor(Color.White);
                //namedStyleTitleRow.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyleOddRow = ep.Workbook.Styles.CreateNamedStyle("Odd Row");
                namedStyleOddRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleOddRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleOddRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleOddRow.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyleEvenRow = ep.Workbook.Styles.CreateNamedStyle("Even Row");
                namedStyleEvenRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleEvenRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleEvenRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleEvenRow.Style.Fill.BackgroundColor.SetColor(Color.FloralWhite);

                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyleTestSettingHeader = ep.Workbook.Styles.CreateNamedStyle("Test Setting Header");
                namedStyleTestSettingHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleTestSettingHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleTestSettingHeader.Style.TextRotation = 180;

                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml namedStyleHighlightCell = ep.Workbook.Styles.CreateNamedStyle("Highlight");
                namedStyleHighlightCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleHighlightCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleHighlightCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleHighlightCell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            }
        }
    }
}
